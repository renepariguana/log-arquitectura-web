// ============================================================
// PRESUPUESTO — Standalone Apps Script (pagina-web)
// Maneja todas las acciones de presupuesto.html
// ============================================================

var AP_ID = '1mzWe4dXvMYvRDJA9MIqjJDy3U2kZY133qKyICT6W43s';
var MM_ID = '1z8t4fvp0urZpCm2EoW8wkNOcufnlCWRXnp65tYDZ4TE';

// ─────────────────────────────────────────────
// ROUTER
// ─────────────────────────────────────────────

function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) || '';
  if (action === 'items')        return getItems();
  if (action === 'catalogo')     return getCatalogo();
  if (action === 'manoobra')     return getManoObra(e.parameter.zona || 'A');
  if (action === 'nuevoitem')    return crearNuevoItem(e.parameter.data || '{}');
  if (action === 'analisisitem') return getAnalisisItem(e.parameter.codigo || '', e.parameter.nombre || '');
  if (action === 'cascada')      return getCascada();
  return ContentService.createTextOutput('OK').setMimeType(ContentService.MimeType.TEXT);
}

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents || '{}');
    if (data.action === 'registrarusuario') return registrarUsuario(data);
  } catch(err) {}
  return json({ error: 'Acción no reconocida' });
}

// ─────────────────────────────────────────────
// USUARIOS — Registra o actualiza usuario en la planilla
// ─────────────────────────────────────────────

function registrarUsuario(data) {
  try {
    var ss    = SpreadsheetApp.openById(AP_ID);
    var hoja  = ss.getSheetByName('Usuarios');

    // Crear la hoja si no existe
    if (!hoja) {
      hoja = ss.insertSheet('Usuarios');
      hoja.appendRow(['Email', 'Nombre', 'Foto URL', 'Primer acceso', 'Último acceso']);
      hoja.getRange(1, 1, 1, 5).setFontWeight('bold');
    }

    var email   = data.email   || '';
    var nombre  = data.nombre  || '';
    var foto    = data.foto    || '';
    var ahora   = new Date().toLocaleString('es-AR');

    // Buscar si el usuario ya existe (columna A = email)
    var valores = hoja.getDataRange().getValues();
    for (var i = 1; i < valores.length; i++) {
      if (valores[i][0] === email) {
        // Actualizar último acceso y nombre (por si cambió)
        hoja.getRange(i + 1, 2).setValue(nombre);
        hoja.getRange(i + 1, 3).setValue(foto);
        hoja.getRange(i + 1, 5).setValue(ahora);
        return json({ ok: true, nuevo: false });
      }
    }

    // Usuario nuevo — agregar fila
    hoja.appendRow([email, nombre, foto, ahora, ahora]);
    return json({ ok: true, nuevo: true });

  } catch(err) {
    return json({ error: err.message });
  }
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─────────────────────────────────────────────
// GET ITEMS — Rubros e ítems de ANALISIS DE PRECIOS
// ─────────────────────────────────────────────

function getItems() {
  try {
    var apSS = SpreadsheetApp.openById(AP_ID);
    var hoja = apSS.getSheetByName('Inicio') || apSS.getSheetByName('BASE DE DATOS');
    if (!hoja) throw new Error('Hoja Inicio no encontrada en ANALISIS DE PRECIOS');

    var datos  = hoja.getDataRange().getValues();
    var rubros = [];
    var items  = [];

    // Saltar encabezados buscando la primera fila con código numérico
    var inicio = 0;
    for (var ri = 0; ri < datos.length; ri++) {
      if (/^\d+\.?\d*$/.test((datos[ri][0] || '').toString().trim())) { inicio = ri; break; }
    }

    for (var ri = inicio; ri < datos.length; ri++) {
      var codigo = (datos[ri][0] || '').toString().trim();
      var nombre = (datos[ri][1] || '').toString().trim();
      var unidad = (datos[ri][2] || '').toString().trim();
      var precio = parseFloat(datos[ri][3]) || 0;

      if (!codigo || !nombre || !/^\d+\.?\d*$/.test(codigo)) continue;

      if (/^\d+$/.test(codigo)) {
        rubros.push({ codigo: codigo, nombre: nombre });
      } else {
        items.push({ codigo: codigo, nombre: nombre, unidad: unidad, precio: precio, rubroId: codigo.split('.')[0] });
      }
    }

    return ContentService
      .createTextOutput(JSON.stringify({ rubros: rubros, items: items }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ─────────────────────────────────────────────
// GET CATALOGO — Materiales de MATERIALES Y MANO DE OBRA
// ─────────────────────────────────────────────

function getCatalogo() {
  try {
    var mmSS = SpreadsheetApp.openById(MM_ID);
    var hoja = mmSS.getSheetByName('Materiales');
    if (!hoja) throw new Error('Hoja Materiales no encontrada');

    var datos = hoja.getDataRange().getValues();

    // Buscar fila de encabezado dinámicamente (la tabla puede no empezar en fila 1)
    var headerRow = -1, colMat = -1, colUn = -1, colPre = -1;
    for (var r = 0; r < datos.length; r++) {
      for (var c = 0; c < datos[r].length; c++) {
        if ((datos[r][c] || '').toString().toUpperCase().trim() === 'MATERIALES') {
          headerRow = r; colMat = c; break;
        }
      }
      if (headerRow >= 0) break;
    }
    if (headerRow < 0) throw new Error('Columna MATERIALES no encontrada en hoja Materiales');

    var headers = datos[headerRow];
    for (var c = 0; c < headers.length; c++) {
      var h = (headers[c] || '').toString().toUpperCase().trim();
      if (h === 'UNIDAD') colUn = c;
      if (h === 'PRECIO UNITARIO ($)') colPre = c;
    }

    var materiales = [];
    for (var r = headerRow + 1; r < datos.length; r++) {
      var nombre = (datos[r][colMat] || '').toString().trim();
      if (!nombre) continue;
      var unidad = colUn >= 0 ? (datos[r][colUn] || '').toString().trim() : '';
      var precioRaw = colPre >= 0 ? datos[r][colPre] : 0;
      var precio = typeof precioRaw === 'number' ? precioRaw
        : parseFloat((precioRaw || '').toString().replace(/\$/g, '').replace(/\./g, '').replace(',', '.')) || 0;
      materiales.push({ nombre: nombre, unidad: unidad, precio: precio });
    }

    return ContentService.createTextOutput(JSON.stringify({ materiales: materiales }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ─────────────────────────────────────────────
// GET MANO DE OBRA — Lee categorías y precios por zona desde AP_ID
// zona: 'A' | 'B' | 'C' | 'C_Austral'
// Columnas en MANO DE OBRA (desde col A=1): A=3, B=6, C=9, C_Austral=12
// ─────────────────────────────────────────────

function getManoObra(zona) {
  try {
    var ZONA_COL = { 'A': 2, 'B': 5, 'C': 8, 'C_Austral': 11 }; // 0-based
    var colPrecioOffset = ZONA_COL[zona] !== undefined ? ZONA_COL[zona] : ZONA_COL['A'];

    var apSS = SpreadsheetApp.openById(AP_ID);
    var hoja = apSS.getSheetByName('MANO DE OBRA');
    if (!hoja) throw new Error('Hoja MANO DE OBRA no encontrada en AP_ID');

    var datos = hoja.getDataRange().getValues();

    // Encontrar fila encabezado (tiene CATEGORIA en col A)
    var headerRow = -1, colCat = -1, colUn = -1;
    for (var r = 0; r < datos.length; r++) {
      for (var c = 0; c < datos[r].length; c++) {
        var v = (datos[r][c] || '').toString().toUpperCase().trim();
        if (v === 'CATEGORIA' || v === 'CATEGORÍA') {
          headerRow = r; colCat = c; break;
        }
      }
      if (headerRow >= 0) break;
    }
    if (headerRow < 0) throw new Error('Columna CATEGORIA no encontrada en MANO DE OBRA');

    // Columna UN: buscar en la fila encabezado
    for (var c = 0; c < datos[headerRow].length; c++) {
      var h = (datos[headerRow][c] || '').toString().toUpperCase().trim();
      if (h === 'UN' || h === 'UNIDAD') { colUn = c; break; }
    }

    var categorias = [];
    for (var r = headerRow + 1; r < datos.length; r++) {
      var nombre = (datos[r][colCat] || '').toString().trim();
      if (!nombre) continue;
      var unidad = colUn >= 0 ? (datos[r][colUn] || '').toString().trim() : '$/Hs';
      var precioRaw = datos[r][colPrecioOffset];
      var precio = typeof precioRaw === 'number' ? precioRaw
        : parseFloat((precioRaw || '').toString().replace(/\$/g, '').replace(/\./g, '').replace(',', '.')) || 0;
      categorias.push({ nombre: nombre, unidad: unidad, precio: precio });
    }

    return ContentService.createTextOutput(JSON.stringify({ categorias: categorias, zona: zona }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ─────────────────────────────────────────────
// CREAR NUEVO ITEM
// ─────────────────────────────────────────────

function crearNuevoItem(dataStr) {
  try {
    var p       = JSON.parse(dataStr);
    var rubroId = (p.rubroId || '').toString().trim();
    var nombre  = (p.nombre  || '').toString().trim();
    var unidad  = (p.unidad  || '').toString().trim();

    if (!rubroId || !nombre || !unidad) throw new Error('Faltan datos obligatorios');

    var ap   = SpreadsheetApp.openById(AP_ID);
    var hoja = ap.getSheetByName('Inicio');
    if (!hoja) throw new Error('Hoja Inicio no encontrada');

    var datos = hoja.getDataRange().getValues();

    // Calcular costo total
    var costo = 0;
    (p.materiales || []).forEach(function(m) { costo += (parseFloat(m.cantidad) || 0) * (parseFloat(m.precio) || 0); });
    (p.manoObra  || []).forEach(function(m) { costo += (parseFloat(m.cantidad) || 0) * (parseFloat(m.precio) || 0); });
    (p.equipos   || []).forEach(function(m) { costo += (parseFloat(m.cantidad) || 0) * (parseFloat(m.precio) || 0); });

    // Obtener nombre del rubro para la pestaña
    var rubroNombre = '';
    for (var i = 0; i < datos.length; i++) {
      var cod = (datos[i][0] || '').toString().trim();
      if (/^\d+$/.test(cod) && cod === rubroId) {
        rubroNombre = (datos[i][1] || '').toString();
        break;
      }
    }

    // Encontrar fila de inserción (última fila del rubro) y siguiente número
    var insertRow   = -1;
    var lastDecimal = 0;

    for (var i = 0; i < datos.length; i++) {
      var codigo = (datos[i][0] || '').toString().trim();
      if (codigo.indexOf(rubroId + '.') === 0) {
        insertRow = i + 1; // 1-indexed
        var decimal = parseFloat((codigo.split('.')[1]) || '0') || 0;
        if (decimal > lastDecimal) lastDecimal = decimal;
      }
    }

    // Si el rubro no tiene ítems aún, insertar justo después de la fila del rubro
    if (insertRow < 0) {
      for (var i = 0; i < datos.length; i++) {
        var cod = (datos[i][0] || '').toString().trim();
        if (/^\d+$/.test(cod) && cod === rubroId) {
          insertRow = i + 1; // 1-indexed
          break;
        }
      }
    }

    if (insertRow < 0) throw new Error('Rubro ' + rubroId + ' no encontrado en Inicio');

    var nuevoCodigo = rubroId + '.' + (lastDecimal + 1);

    // Insertar fila en Inicio con fondo amarillo
    hoja.insertRowBefore(insertRow + 1);
    hoja.getRange(insertRow + 1, 1, 1, 4).setValues([[nuevoCodigo, nombre, unidad, costo]]);
    hoja.getRange(insertRow + 1, 1, 1, 4).setBackground('#FFF9C4');

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok', codigo: nuevoCodigo, costo: costo }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ─────────────────────────────────────────────
// GET ANALISIS ITEM — Lee la pestaña del ítem en AP
// ─────────────────────────────────────────────

function getAnalisisItem(codigo, nombre) {
  try {
    var apSS = SpreadsheetApp.openById(AP_ID);

    // Buscar pestaña: primero por código exacto, luego por nombre normalizado
    var hoja = apSS.getSheetByName(codigo);
    if (!hoja && nombre) {
      var normNombre = norm_(nombre);
      apSS.getSheets().forEach(function(h) {
        if (!hoja && norm_(h.getName()).indexOf(normNombre) !== -1) hoja = h;
      });
    }
    if (!hoja) {
      return ContentService
        .createTextOutput(JSON.stringify({ error: 'Sin pestaña para ' + codigo }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    var datos = hoja.getDataRange().getValues();

    var TITULOS = ['MATERIALES', 'MANO DE OBRA', 'EQUIPOS'];
    var resultKeys = ['materiales', 'manoobra', 'equipos'];
    var resultado = { materiales: [], manoobra: [], equipos: [] };

    TITULOS.forEach(function(titulo, s) {
      // Localizar fila de título que sea cabecera de tabla real (con CANTIDAD o PRECIO cerca)
      var titleRowIdx = -1;
      for (var r = 0; r < datos.length; r++) {
        if (!datos[r].some(function(c) { return (c||'').toString().toUpperCase().trim() === titulo; })) continue;
        var esTabla = false;
        for (var rr = r; rr < Math.min(r + 3, datos.length); rr++) {
          if (datos[rr].some(function(h) {
            var hUp = (h||'').toString().toUpperCase().trim();
            return hUp === 'CANTIDAD' || hUp === 'HS' || hUp === 'HORAS' || hUp.indexOf('PRECIO') !== -1;
          })) { esTabla = true; break; }
        }
        if (esTabla) { titleRowIdx = r; break; }
      }
      if (titleRowIdx < 0) return;

      var colHeaderRowIdx = titleRowIdx;
      var colNombre = -1, colUnidad = -1, colCantidad = -1, colPrecio = -1;

      for (var r = titleRowIdx; r < Math.min(titleRowIdx + 4, datos.length); r++) {
        var tieneEnc = datos[r].some(function(h) {
          var hUp = (h||'').toString().toUpperCase().trim();
          return hUp === 'CANTIDAD' || hUp === 'HS' || hUp === 'HORAS' || hUp.indexOf('PRECIO') !== -1;
        });
        if (tieneEnc) { colHeaderRowIdx = r; break; }
      }

      // Mapear columnas
      datos[colHeaderRowIdx].forEach(function(h, c) {
        var hUp = (h||'').toString().toUpperCase().trim();
        if (colNombre < 0 && (hUp === titulo || hUp === 'DESCRIPCION' || hUp === 'DESCRIPCIÓN' ||
            hUp === 'INSUMO' || hUp === 'ITEM' || hUp === 'CATEGORÍA' || hUp === 'CATEGORIA' ||
            hUp === 'MATERIALES' || hUp === 'MANO DE OBRA' || hUp === 'EQUIPOS')) colNombre = c;
        if (colUnidad < 0   && (hUp === 'UNIDAD' || hUp === 'UN' || hUp === 'UND')) colUnidad = c;
        if (colCantidad < 0 && (hUp === 'CANTIDAD' || hUp === 'HS' || hUp === 'HORAS'))       colCantidad = c;
        if (colPrecio < 0   && hUp.indexOf('PRECIO') !== -1) colPrecio = c;
      });
      // Fallback nombre: primera columna no vacía del encabezado
      if (colNombre < 0) {
        datos[colHeaderRowIdx].forEach(function(h, c) { if (colNombre < 0 && (h||'') !== '') colNombre = c; });
      }

      // Límites de las otras secciones para no pasarlas
      var otros = TITULOS.filter(function(t) { return t !== titulo; });
      otros.push('SUBTOTAL', 'COSTO DEL ITEM', 'TOTAL');

      var filas = [];
      for (var r = colHeaderRowIdx + 1; r < datos.length; r++) {
        var esFin = datos[r].some(function(c) {
          var v = (c||'').toString().toUpperCase().trim();
          return otros.some(function(t) { return v.indexOf(t) !== -1; });
        });
        if (esFin) break;
        var nom = colNombre >= 0 ? (datos[r][colNombre]||'').toString().trim() : '';
        if (!nom) continue;
        filas.push({
          nombre  : nom,
          unidad  : colUnidad   >= 0 ? (datos[r][colUnidad]||'').toString().trim() : '',
          cantidad: colCantidad >= 0 ? (parseFloat(datos[r][colCantidad]) || 0) : 0,
          precio  : colPrecio   >= 0 ? (parseFloat(datos[r][colPrecio])   || 0) : 0
        });
      }
      resultado[resultKeys[s]] = filas;
    });

    return ContentService
      .createTextOutput(JSON.stringify(resultado))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function norm_(str) {
  return str.toString().toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').trim();
}

// ─────────────────────────────────────────────
// CREAR PESTAÑA A PARTIR DE LA PLANTILLA "TABLA"
// ─────────────────────────────────────────────

function crearPestanaItem_(ap, codigo, nombre, unidad, rubroId, rubroNombre, materiales, manoObra, equipos) {
  try {
    // Buscar la hoja plantilla
    var template = ap.getSheetByName('TABLA')
                || ap.getSheetByName('Tabla')
                || ap.getSheetByName('tabla');
    if (!template) {
      Logger.log('Plantilla TABLA no encontrada — se omite la creación de pestaña');
      return;
    }

    // Copiar la plantilla al final del archivo
    var nueva = template.copyTo(ap);
    nueva.setName(codigo);
    ap.moveActiveSheet(ap.getNumSheets());

    // Eliminar rangos con nombre duplicados que crea Sheets al copiar
    ap.getNamedRanges().forEach(function(rango) {
      if (rango.getName().slice(-2) === '_2' &&
          rango.getRange().getSheet().getName() === nueva.getName()) {
        rango.remove();
      }
    });

    // Completar celdas de encabezado del ítem
    try { nueva.getRange('A3').setValue(rubroId || ''); } catch(e) {}
    try { nueva.getRange('B3').setValue(rubroNombre || ''); } catch(e) {}
    try { nueva.getRange('A4').setValue(codigo); } catch(e) {}
    try { nueva.getRange('B4').setValue(nombre); } catch(e) {}
    try { nueva.getRange('C4').setValue(unidad); } catch(e) {}

    // Leer toda la hoja de una vez para analizar la estructura
    var datos   = nueva.getDataRange().getValues();
    var numCols = nueva.getLastColumn();

    var secciones = [
      { titulo: 'MATERIALES',   filas: materiales || [] },
      { titulo: 'MANO DE OBRA', filas: manoObra   || [] },
      { titulo: 'EQUIPOS',      filas: equipos     || [] }
    ];

    // Mapear estructura de cada sección (índices 0-based en el array datos)
    secciones.forEach(function(sec) {
      sec.titleRowIdx     = -1;
      sec.colHeaderRowIdx = -1;
      sec.subtotalRowIdx  = -1;
      sec.colNombre   = 0;
      sec.colUnidad   = 0;
      sec.colCantidad = 0;
      sec.colPrecio   = 0;

      // Encontrar fila con el título de la sección
      for (var r = 0; r < datos.length; r++) {
        var found = datos[r].some(function(c) {
          return (c || '').toString().toUpperCase().trim() === sec.titulo;
        });
        if (found) { sec.titleRowIdx = r; break; }
      }
      if (sec.titleRowIdx < 0) return;

      // La siguiente fila puede ser la de encabezados de columna
      // Si esa fila tiene CANTIDAD o PRECIO → es col header
      // Si no → el título mismo actúa como col header
      var candidatoColHeader = sec.titleRowIdx + 1;
      var fila = datos[candidatoColHeader] || [];
      var tieneEncabezados = fila.some(function(h) {
        var hUp = (h || '').toString().toUpperCase().trim();
        return hUp === 'CANTIDAD' || hUp.indexOf('PRECIO') !== -1;
      });
      sec.colHeaderRowIdx = tieneEncabezados ? candidatoColHeader : sec.titleRowIdx;

      // Extraer índices de columnas desde la fila de encabezados (0-based)
      var hdrs = datos[sec.colHeaderRowIdx] || [];
      hdrs.forEach(function(h, c) {
        var hUp = (h || '').toString().toUpperCase().trim();
        if (!sec.colNombre   && (hUp === sec.titulo || hUp === 'DESCRIPCION' || hUp === 'DESCRIPCIÓN' || hUp === 'INSUMO' || hUp === 'ITEM')) sec.colNombre = c + 1;
        if (!sec.colUnidad   && (hUp === 'UNIDAD'   || hUp === 'UN'  || hUp === 'UND')) sec.colUnidad   = c + 1;
        if (!sec.colCantidad && hUp === 'CANTIDAD')  sec.colCantidad = c + 1;
        if (!sec.colPrecio   && hUp.indexOf('PRECIO') !== -1 && hUp.indexOf('UNITARIO') !== -1) sec.colPrecio = c + 1;
      });
      // Fallback nombre: primera columna no vacía del encabezado
      if (!sec.colNombre) {
        hdrs.forEach(function(h, c) { if (!sec.colNombre && h !== '') sec.colNombre = c + 1; });
      }

      // Buscar la primera fila SUBTOTAL después de los encabezados de columna
      for (var r = sec.colHeaderRowIdx + 1; r < datos.length; r++) {
        var foundSub = datos[r].some(function(c) {
          return (c || '').toString().toUpperCase().trim().indexOf('SUBTOTAL') !== -1;
        });
        if (foundSub) { sec.subtotalRowIdx = r; break; }
      }
    });

    // Rellenar tablas, procesando en orden y ajustando offset por filas insertadas
    var rowOffset = 0;

    secciones.forEach(function(sec) {
      if (sec.titleRowIdx < 0 || sec.subtotalRowIdx < 0) return;

      // Convertir a filas 1-indexed de la hoja real (aplicando offset)
      var dataStart1   = sec.colHeaderRowIdx + 2 + rowOffset; // primera fila de datos
      var subtotalRow1 = sec.subtotalRowIdx  + 1 + rowOffset; // fila SUBTOTAL
      var dataCount    = subtotalRow1 - dataStart1;           // filas plantilla disponibles

      // Borrar contenido de las filas plantilla
      if (dataCount > 0) {
        nueva.getRange(dataStart1, 1, dataCount, numCols).clearContent();
      }

      // Insertar filas extra si tenemos más datos que filas plantilla
      var extraRows = sec.filas.length - dataCount;
      if (extraRows > 0) {
        nueva.insertRowsBefore(subtotalRow1, extraRows);
        rowOffset += extraRows;
      }

      // Escribir los datos fila a fila
      sec.filas.forEach(function(fila, idx) {
        var rowNum = dataStart1 + idx;
        if (sec.colNombre)   nueva.getRange(rowNum, sec.colNombre).setValue(fila.nombre   || '');
        if (sec.colUnidad)   nueva.getRange(rowNum, sec.colUnidad).setValue(fila.unidad   || '');
        if (sec.colCantidad) nueva.getRange(rowNum, sec.colCantidad).setValue(parseFloat(fila.cantidad) || 0);
        if (sec.colPrecio)   nueva.getRange(rowNum, sec.colPrecio).setValue(parseFloat(fila.precio)    || 0);
      });
    });

    Logger.log('Pestaña creada: ' + codigo);

  } catch (err) {
    Logger.log('Error en crearPestanaItem_: ' + err.toString());
    // No relanzar el error — si la pestaña falla, igual devolvemos OK desde crearNuevoItem
  }
}

// ─────────────────────────────────────────────
// GET CASCADA — Porcentajes desde pestaña CF
// ─────────────────────────────────────────────

function getCascada() {
  try {
    var apSS = SpreadsheetApp.openById(AP_ID);
    // Buscar la pestaña CF por nombre; si no, por gid 999397180
    var hoja = apSS.getSheetByName('CF');
    if (!hoja) {
      var sheets = apSS.getSheets();
      for (var i = 0; i < sheets.length; i++) {
        if (sheets[i].getSheetId() === 999397180) { hoja = sheets[i]; break; }
      }
    }
    if (!hoja) throw new Error('Pestaña CF no encontrada');

    var datos = hoja.getDataRange().getValues();

    // Mapeo de keywords → clave resultado
    var KEYWORDS = [
      { key: 'GASTOS GENERALES',            campo: 'gastosGenerales' },
      { key: 'GASTOS GEN',                  campo: 'gastosGenerales' },
      { key: 'BENEFICIOS',                  campo: 'beneficios' },
      { key: 'COSTO FINANCIERO',            campo: 'costoFinanciero' },
      { key: 'IVA',                         campo: 'iva' },
      { key: 'INGRESOS BRUTOS',             campo: 'ingresosBrutos' },
      { key: 'IMPUESTOS MUNICIPALES',       campo: 'impMunicipales' },
      { key: 'DEBITO Y CREDITO',            campo: 'impDebitoCredito' },
      { key: 'DEBITO',                      campo: 'impDebitoCredito' }
    ];

    var result = {
      gastosGenerales: 0,
      beneficios: 0,
      costoFinanciero: 0,
      iva: 0,
      ingresosBrutos: 0,
      impMunicipales: 0,
      impDebitoCredito: 0
    };

    for (var r = 0; r < datos.length; r++) {
      var fila = datos[r];
      for (var c = 0; c < fila.length; c++) {
        var celda = (fila[c] || '').toString().toUpperCase().trim();
        if (!celda) continue;
        for (var k = 0; k < KEYWORDS.length; k++) {
          if (celda.indexOf(KEYWORDS[k].key) === -1) continue;
          // Buscar el porcentaje en las columnas siguientes de la misma fila
          for (var cc = c + 1; cc < fila.length; cc++) {
            var v = fila[cc];
            if (v === '' || v === null || v === undefined) continue;
            var n = parseFloat(v);
            if (isNaN(n)) continue;
            // Google Sheets almacena % como decimal (0.15) o como número (15)
            if (n >= 0 && n <= 1) {
              result[KEYWORDS[k].campo] = n;
            } else if (n > 1 && n <= 100) {
              result[KEYWORDS[k].campo] = n / 100;
            }
            break;
          }
          break; // no seguir buscando keywords en esta celda
        }
      }
    }

    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
