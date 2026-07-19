/**
 * UOCRA — Actualizador de escalas salariales mano de obra
 * Scraper mensual: descarga el último PDF de UOCRA, extrae salarios con OCR
 * y actualiza la hoja "MANO DE OBRA" del Sheets de Materiales y Mano de Obra.
 *
 * SETUP (hacer una sola vez):
 *   1. Servicios → + → Drive API v2 → Agregar  (necesario para OCR de PDFs)
 *   2. Ejecutar crearTriggerMensualUOCRA()       (instala el trigger del día 1)
 *   3. Ejecutar actualizarSalariosUOCRA()        (prueba manual)
 *
 * ZONAS UOCRA (CCT 76/75 — construcción interior del país):
 *   Zona A: CABA, Bs.As., Santa Fe, Córdoba, Mendoza, Tucumán, Salta, Jujuy,
 *           Catamarca, Santiago del Estero, Entre Ríos, Chaco, Formosa,
 *           Corrientes, Misiones, La Rioja, San Luis, San Juan, La Pampa
 *   Zona B: Neuquén, Río Negro, Chubut
 *   Zona C: Santa Cruz
 *   Zona C Austral: Tierra del Fuego
 */

// ─── Configuración ────────────────────────────────────────────────────────────

// Script vinculado al Sheets — usa getActiveSpreadsheet(), no necesita ID
var MO_EMAIL  = 'estudiologarquitectura@gmail.com';
var UOCRA_URL = 'https://www.uocra.org/?s=nuevas-escalas-salariales&lang=1';

// ─── Función principal ────────────────────────────────────────────────────────

function actualizarSalariosUOCRA() {
  try {
    // 1. Scrapear página UOCRA
    // El HTML usa comillas simples: href='pdf/...'
    var html = UrlFetchApp.fetch(UOCRA_URL, {
      muteHttpExceptions: true,
      headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36' }
    }).getContentText();

    // Extraer pares título+ruta de cada fila de la tabla de escalas
    var entradas = [];
    var reEntrada = /class='tablas floatl'>([\s\S]*?)<\/p>[\s\S]*?href='(pdf\/[^']+\.pdf)'/g;
    var m;
    while ((m = reEntrada.exec(html)) !== null) {
      entradas.push({ titulo: m[1].trim(), path: m[2] });
    }

    if (!entradas.length) {
      Logger.log('UOCRA: No se encontraron PDFs en la página.');
      return;
    }

    // Prioridad 1: "Acuerdo 76/75" (tiene tablas con valores absolutos)
    // Prioridad 2: cualquier "76/75" (Paritaria, Homologación)
    // Fallback: primera entrada
    var entrada = null;
    for (var i = 0; i < entradas.length; i++) {
      var t = entradas[i].titulo;
      if (/Acuerdo/i.test(t) && /76[\/-]?75/i.test(t) && !/545|Paritaria|Homolog/i.test(t)) {
        entrada = entradas[i]; break;
      }
    }
    if (!entrada) {
      for (var i = 0; i < entradas.length; i++) {
        var t = entradas[i].titulo;
        if (/76[\/-]?75/i.test(t) && !/545/.test(t)) { entrada = entradas[i]; break; }
      }
    }
    if (!entrada) entrada = entradas[0];

    var pdfUrl = 'https://www.uocra.org/' + entrada.path;
    Logger.log('UOCRA: PDF seleccionado → ' + entrada.titulo + ' | ' + pdfUrl);

    // 2. Verificar si ya fue procesado este mes
    var props     = PropertiesService.getScriptProperties();
    var ultimoUrl = props.getProperty('UOCRA_LAST_PDF') || '';
    if (ultimoUrl === pdfUrl) {
      Logger.log('UOCRA: Sin novedades. Último PDF ya procesado: ' + pdfUrl);
      return;
    }

    // 3. Descargar PDF
    var pdfBlob = UrlFetchApp.fetch(pdfUrl, { muteHttpExceptions: true })
      .getBlob().setName('uocra_escalas.pdf');

    // 4. Subir a Drive con OCR (v3: el OCR es automático al convertir a Google Doc)
    var tempDoc = Drive.Files.create(
      { name: 'uocra_ocr_' + Date.now(), mimeType: 'application/vnd.google-apps.document' },
      pdfBlob,
      { fields: 'id' }
    );

    // 5. Extraer texto y eliminar doc temporal
    var texto = DocumentApp.openById(tempDoc.id).getBody().getText();
    DriveApp.getFileById(tempDoc.id).setTrashed(true);

    // 6. Parsear salarios
    props.setProperty('UOCRA_LAST_PDF', pdfUrl);
    var data  = parsearSalariosUOCRA(texto);
    var fecha = new Date().toLocaleDateString('es-AR', { timeZone: 'America/Argentina/Buenos_Aires' });

    if (data && data.basicos && Object.keys(data.basicos).length > 0) {
      // 7. Actualizar hoja
      actualizarHojaManoObra(data, fecha);

      // 8. Email de confirmación
      var lineas = ['Salarios básicos (Zona A — base todas las zonas):'];
      var bA = data.basicos['Zona A'] || {};
      Object.keys(bA).forEach(function(cat) {
        lineas.push('  • ' + cat + ': $' + Number(bA[cat]).toLocaleString('es-AR') + '/h');
      });
      if (data.adicionales && Object.keys(data.adicionales).length) {
        lineas.push('\nAdicionales por zona:');
        Object.keys(data.adicionales).forEach(function(zona) {
          lineas.push('  ' + zona + ':');
          var adics = data.adicionales[zona];
          Object.keys(adics).forEach(function(cat) {
            lineas.push('    + $' + Number(adics[cat]).toLocaleString('es-AR') + ' (' + cat + ')');
          });
        });
      }
      MailApp.sendEmail(MO_EMAIL,
        'UOCRA — Salarios actualizados ' + fecha,
        'Fuente: ' + pdfUrl + '\n\n' + lineas.join('\n')
      );
      Logger.log('UOCRA: Actualizado — ' + JSON.stringify(data));

    } else {
      // Parseo fallido: mandar texto para revisión manual
      MailApp.sendEmail(MO_EMAIL,
        'UOCRA — Nuevo PDF (revisión manual requerida)',
        'Se detectó un nuevo PDF pero no se pudo parsear automáticamente.\n' +
        'URL: ' + pdfUrl + '\n\n' +
        'Texto extraído (primeros 3000 caracteres):\n' + texto.substring(0, 3000)
      );
      Logger.log('UOCRA: Parseo fallido para ' + pdfUrl);
    }

  } catch (err) {
    Logger.log('UOCRA Error: ' + err.toString());
    MailApp.sendEmail(MO_EMAIL, 'UOCRA — Error al actualizar salarios', err.toString());
  }
}

// ─── Parser ───────────────────────────────────────────────────────────────────

function parsearSalariosUOCRA(texto) {
  /**
   * Las tablas del PDF "Acuerdo 76/75" tienen este formato OCR (una línea por mes):
   *   may-26 Oficial Especializado 6119 6119 673 6792 6119 3275 9394 6119 6119 12238
   *          Oficial 5235 5235 579 5813 5235 3573 8808 5235 5235 10469 ...
   *
   * Columnas (10 números por categoría):
   *   [A_basic, B_basic=A, B_addl, B_total, C_basic=A, C_addl, C_total, CA_basic=A, CA_addl, CA_total]
   *
   * Retorna: { basicos: { 'Zona A': {cat: val, ...}, ... }, adicionales: { 'Zona B': {...}, ... } }
   */

  // Descartar ANEXO II (CCT 577/10 — canalización/instalaciones, no aplica a construcción general)
  var idxII = texto.search(/ANEXO\s*II/i);
  var textoI = idxII > 0 ? texto.substring(0, idxII) : texto;

  // Encontrar el último bloque mensual (el más reciente en el documento)
  var reMes = /\b(?:ene|feb|mar|abr|may|jun|jul|ago|sep|oct|nov|dic)-\d{2}\b/gi;
  var lastMes = null, m;
  while ((m = reMes.exec(textoI)) !== null) lastMes = m;
  if (!lastMes) return { basicos: {}, adicionales: {} };

  // Extraer el bloque desde ese mes hasta pie de tabla (SNR, firmas)
  var bloque = textoI.substring(lastMes.index);
  var endIdx = bloque.search(/\bMas\b|\bMás\b|Suma no rem|Ayudante Zona|Sereno Zona|Ing\.\s*[A-Z]/i);
  if (endIdx > 0) bloque = bloque.substring(0, endIdx);

  // Definición de categorías — orden: procesar OE antes que O para no confundirlos
  var catDefs = [
    { name: 'Oficial especializado', re: /Oficial\s*Especializado/i             },
    { name: 'Oficial',               re: /\bOficial\b(?!\s*Esp)/i               },
    { name: 'Medio oficial',         re: /Medio\s*Oficial|1[\s\/\\]2\s*Oficial|½\s*Oficial/i },
    { name: 'Ayudante',              re: /\bAyudante\b/i                         },
    { name: 'Sereno (mensual)',       re: /\bSereno\b/i                           }
  ];

  // Encontrar posición de cada categoría en el bloque
  var found = [];
  catDefs.forEach(function(cat) {
    var m = cat.re.exec(bloque);
    if (m) found.push({ name: cat.name, start: m.index, end: m.index + m[0].length });
  });
  found.sort(function(a, b) { return a.start - b.start; });

  var basicos     = {};
  var adicionales = {};
  var zonas = ['Zona A', 'Zona B', 'Zona C', 'Zona C Austral'];

  for (var i = 0; i < found.length; i++) {
    var segmento = bloque.substring(found[i].end, i + 1 < found.length ? found[i + 1].start : bloque.length);
    var nums = (segmento.match(/\d+/g) || []).map(Number).filter(function(n) { return n > 0; });
    if (nums.length < 9) continue;

    // nums: [A_basic, B_basic, B_addl, B_total, C_basic, C_addl, C_total, CA_basic, CA_addl, CA_total]
    var basic  = nums[0];
    var addlB  = nums[2];
    var addlC  = nums[5];
    var addlCA = nums[8];

    var cat = found[i].name;
    zonas.forEach(function(z) { if (!basicos[z]) basicos[z] = {}; });
    ['Zona B', 'Zona C', 'Zona C Austral'].forEach(function(z) { if (!adicionales[z]) adicionales[z] = {}; });

    basicos['Zona A'][cat]             = basic;
    basicos['Zona B'][cat]             = basic;
    basicos['Zona C'][cat]             = basic;
    basicos['Zona C Austral'][cat]     = basic;
    adicionales['Zona B'][cat]         = addlB;
    adicionales['Zona C'][cat]         = addlC;
    adicionales['Zona C Austral'][cat] = addlCA;
  }

  return { basicos: basicos, adicionales: adicionales };
}

// ─── Actualizador de hoja ─────────────────────────────────────────────────────

function actualizarHojaManoObra(data, fecha) {
  /**
   * Detecta dinámicamente los encabezados de columna en la hoja "MANO DE OBRA"
   * y escribe el básico y el adicional por categoría y zona.
   *
   * Columnas que actualiza (las demás son fórmulas en el Sheets):
   *   Salario Básico Zona "A"
   *   Salario Básico Zona "B"  + Adicional zona desfavorable Zona "B"
   *   Salario Básico Zona "C"  + Adicional zona desfavorable Zona "C"
   *   Salario Básico Zona "C" Austral + Adicional zona desfavorable Zona "C" Austral
   */
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName('MANO DE OBRA')
          || ss.getSheetByName('mano de obra')
          || ss.getSheetByName('Mano de Obra')
          || ss.getSheetByName('MO');
  if (!hoja) throw new Error('Hoja "MANO DE OBRA" no encontrada en Sheets.');

  var datos = hoja.getDataRange().getValues();

  // Detectar fila de encabezados y mapear columnas por nombre
  var colCat = 0;
  var colMap = {};
  var filaHeader = -1;

  for (var r = 0; r < datos.length; r++) {
    var encontro = false;
    for (var c = 0; c < datos[r].length; c++) {
      // Normalizar: minúsculas, sin comillas tipográficas
      var t = (datos[r][c] || '').toString().toLowerCase().replace(/["""'']/g, '').trim();
      if (!t) continue;

      if (/^categor/.test(t))                                                  { colCat = c; encontro = true; }
      if (/salario.*b[aá]s.*zona.*a\b/.test(t))                               { colMap.basico_A  = c; encontro = true; }
      if (/salario.*b[aá]s.*zona.*b\b/.test(t))                               { colMap.basico_B  = c; encontro = true; }
      if (/salario.*b[aá]s.*zona.*c.*austral/.test(t))                        { colMap.basico_CA = c; encontro = true; }
      else if (/salario.*b[aá]s.*zona.*c\b/.test(t))                          { colMap.basico_C  = c; encontro = true; }
      if (/adicional.*zona.*b\b|zona.*b.*adicional/.test(t))                  { colMap.adicional_B  = c; encontro = true; }
      if (/adicional.*zona.*c.*austral|zona.*c.*austral.*adicional/.test(t))  { colMap.adicional_CA = c; encontro = true; }
      else if (/adicional.*zona.*c\b|zona.*c.*adicional/.test(t))             { colMap.adicional_C  = c; encontro = true; }
    }
    if (encontro) { filaHeader = r; break; }
  }

  var zonaColDef = [
    { zona: 'Zona A',         colBasico: 'basico_A',  colAdic: null           },
    { zona: 'Zona B',         colBasico: 'basico_B',  colAdic: 'adicional_B'  },
    { zona: 'Zona C',         colBasico: 'basico_C',  colAdic: 'adicional_C'  },
    { zona: 'Zona C Austral', colBasico: 'basico_CA', colAdic: 'adicional_CA' }
  ];

  var filaInicio = filaHeader >= 0 ? filaHeader + 1 : 1;
  var actualizados = [];

  for (var r = filaInicio; r < datos.length; r++) {
    var catCelda = (datos[r][colCat] || '').toString().trim();
    if (!catCelda) continue;

    zonaColDef.forEach(function(def) {
      var bZona = (data.basicos     || {})[def.zona] || {};
      var aZona = (data.adicionales || {})[def.zona] || {};

      // Scoring para evitar que "Oficial" matchee "Oficial especializado":
      //   3 = exacto, 2 = celda contiene clave, 1 = clave contiene celda (fallback)
      var bestCat = null, bestScore = -1;
      for (var cat in bZona) {
        var a = catCelda.toLowerCase(), b = cat.toLowerCase();
        var score = -1;
        if (a === b)               score = 3;
        else if (a.indexOf(b) !== -1) score = 2;
        else if (b.indexOf(a) !== -1) score = 1;
        if (score > bestScore) { bestScore = score; bestCat = cat; }
      }
      if (bestScore < 0) return;

      var cat  = bestCat;
      var colB = colMap[def.colBasico];
      if (colB !== undefined) {
        hoja.getRange(r + 1, colB + 1).setValue(bZona[cat]);
        actualizados.push(def.zona + ' ' + cat + ' $' + bZona[cat]);
      }
      if (def.colAdic && aZona[cat] !== undefined) {
        var colA = colMap[def.colAdic];
        if (colA !== undefined) {
          hoja.getRange(r + 1, colA + 1).setValue(aZona[cat]);
          actualizados.push(def.zona + ' ' + cat + ' +$' + aZona[cat]);
        }
      }
    });
  }

  Logger.log('UOCRA actualizado: ' + actualizados.join(' | '));
  return actualizados;
}

// ─── Trigger mensual (ejecutar UNA VEZ) ──────────────────────────────────────

function crearTriggerMensualUOCRA() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'actualizarSalariosUOCRA') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('actualizarSalariosUOCRA')
    .timeBased()
    .onMonthDay(1)
    .atHour(9)
    .inTimezone('America/Argentina/Buenos_Aires')
    .create();
  Logger.log('Trigger mensual instalado: corre el día 1 de cada mes a las 9 AM.');
}

// ─── Forzar re-proceso del último PDF (para pruebas) ─────────────────────────

function resetearUltimoPDF() {
  PropertiesService.getScriptProperties().deleteProperty('UOCRA_LAST_PDF');
  Logger.log('UOCRA_LAST_PDF borrado. La próxima ejecución procesará el PDF actual.');
}

// ─── Debug: mostrar texto OCR del PDF actual ─────────────────────────────────

function debugOCR() {
  // Probar distintos PDFs para ver cuál tiene tablas salariales OCR-legibles
  // Cambiar esta URL para probar otros PDFs de la lista
  var pdfUrl = 'https://www.uocra.org/pdf/24d49b_acuerdoabril2026.pdf'; // Acuerdo 76/75 - abril 2026
  // var pdfUrl = 'https://www.uocra.org/pdf/bba99d_homologacionabril2026.pdf'; // Homologación 76/75 - abril 2026
  var pdfBlob = UrlFetchApp.fetch(pdfUrl, { muteHttpExceptions: true })
    .getBlob().setName('uocra_debug.pdf');

  var tempDoc = Drive.Files.create(
    { name: 'uocra_ocr_debug_' + Date.now(), mimeType: 'application/vnd.google-apps.document' },
    pdfBlob,
    { fields: 'id' }
  );

  var texto = DocumentApp.openById(tempDoc.id).getBody().getText();
  DriveApp.getFileById(tempDoc.id).setTrashed(true);

  Logger.log('Texto OCR (' + texto.length + ' chars):');
  // Mostrar en bloques de 4000
  for (var i = 0; i < texto.length; i += 4000) {
    Logger.log('[' + i + '-' + Math.min(i+4000, texto.length) + ']:\n' + texto.substring(i, i + 4000));
  }
}

// ─── Debug: listar los PDFs disponibles ──────────────────────────────────────

function debugUOCRAPage() {
  var html = UrlFetchApp.fetch(UOCRA_URL, {
    muteHttpExceptions: true,
    headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36' }
  }).getContentText();

  var reEntrada = /class='tablas floatl'>([\s\S]*?)<\/p>[\s\S]*?href='(pdf\/[^']+\.pdf)'/g;
  var m, i = 0;
  while ((m = reEntrada.exec(html)) !== null) {
    Logger.log((++i) + '. ' + m[1].trim() + '\n   → ' + m[2]);
  }
  Logger.log('Total: ' + i + ' entradas');
}
