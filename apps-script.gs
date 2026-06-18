/**
 * LOG arquitectura — Google Apps Script
 * 1. POST  → recibe formularios de contacto → hoja "Solicitudes"
 * 2. GET ?action=proyectos → devuelve lista de imágenes de la carpeta Drive
 * 3. Actualiza el costo m² en B2 con trigger diario
 *
 * INSTRUCCIONES:
 * 1. Pegá este código en Extensiones > Apps Script
 * 2. Guardá (Ctrl+S)
 * 3. Ejecutá "crearTriggerDiario" UNA vez (▶ Run) para el costo diario
 * 4. Implementá como Aplicación web (NUEVA implementación cada vez que editás):
 *    - Ejecutar como: Yo
 *    - Quién puede acceder: Cualquier persona
 * 5. Copiá la URL nueva y pegála en index.html donde dice GSHEET_URL
 */

var SHEET_NAME          = 'Solicitudes';
var PROYECTOS_FOLDER_ID = '1ZRTrI0wXUfZgctGbdPeVHWhi2pcQ4ck1';
var VIDEOS_FOLDER_ID    = '13lIAjEn1qt_DQ0-ABJavZkRC0gJQekQL';
var COSTO_URL           = 'https://arquitecturayconstrucciondigital.com/';

// ─── GET: proyectos desde Drive ──────────────────────────────────────────────

function doGet(e) {
  if (e && e.parameter && e.parameter.action === 'proyectos') {
    try {
      var folder  = DriveApp.getFolderById(PROYECTOS_FOLDER_ID);
      var files   = folder.getFiles();
      var projects = [];

      while (files.hasNext()) {
        var file = files.next();
        if (!file.getMimeType().match(/^image\//)) continue;

        // Nombre: "2026-Casa Pima (Córdoba)" o "2026-Casa Pima (Córdoba).jpg"
        var title = file.getName().replace(/\.(jpe?g|png|gif|webp)$/i, '');
        var m = title.match(/^(\d{4})[- ](.+?)\s*\(([^)]+)\)\s*$/);
        if (!m) continue;

        var id = file.getId();
        projects.push({
          name     : m[2].trim(),
          location : m[3].replace(/\./g, ',').trim(),
          year     : m[1],
          cover    : 'https://lh3.googleusercontent.com/d/' + id,
          images   : ['https://lh3.googleusercontent.com/d/' + id]
        });
      }

      // Más recientes primero, luego por nombre
      projects.sort(function(a, b) {
        if (b.year !== a.year) return b.year > a.year ? 1 : -1;
        return a.name.localeCompare(b.name, 'es');
      });

      return ContentService
        .createTextOutput(JSON.stringify(projects))
        .setMimeType(ContentService.MimeType.JSON);

    } catch (err) {
      return ContentService
        .createTextOutput(JSON.stringify({ error: err.toString() }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  if (e && e.parameter && e.parameter.action === 'videos') {
    try {
      var folder = DriveApp.getFolderById(VIDEOS_FOLDER_ID);
      var files  = folder.getFiles();
      var videos = [];

      while (files.hasNext()) {
        var file = files.next();
        if (!file.getMimeType().match(/^video\//)) continue;
        videos.push({
          id  : file.getId(),
          name: file.getName().replace(/\.[^.]+$/, ''),
          url : 'https://drive.google.com/uc?id=' + file.getId()
        });
      }

      return ContentService
        .createTextOutput(JSON.stringify(videos))
        .setMimeType(ContentService.MimeType.JSON);

    } catch (err) {
      return ContentService
        .createTextOutput(JSON.stringify({ error: err.toString() }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  return ContentService
    .createTextOutput('LOG arquitectura — Apps Script activo.')
    .setMimeType(ContentService.MimeType.TEXT);
}

// ─── POST: formulario → hoja Solicitudes o Mensajes ─────────────────────────

function doPost(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var p  = e.parameter;

    if (p.hoja === 'mensajes') {
      // Sección Contacto → envío por email al estudio
      var asunto = 'Nuevo mensaje desde logarquitectura.com.ar — ' + (p.nombre || 'Sin nombre');
      var cuerpo =
        'Nombre:    ' + (p.nombre    || '-') + '\n' +
        'Email:     ' + (p.email     || '-') + '\n' +
        'Teléfono:  ' + (p.telefono  || '-') + '\n' +
        'Fecha:     ' + (p.fecha     || '-') + '\n\n' +
        'Mensaje:\n' + (p.mensaje || '-');
      MailApp.sendEmail('estudiologarquitectura@gmail.com', asunto, cuerpo);

      // Correo de confirmación al cliente
      if (p.email) {
        var asuntoCliente = 'Recibimos tu mensaje — LOG arquitectura';
        var htmlCliente =
          '<div style="font-family:Arial,sans-serif; max-width:560px; margin:0 auto; color:#222;">' +
          '<p>Hola ' + (p.nombre || 'allí') + ',</p>' +
          '<p>Gracias por escribirnos. Recibimos tu mensaje y nos pondremos en contacto con vos en las próximas 24 horas.</p>' +
          '<p>Mientras tanto, podés conocer más de nuestro trabajo en: <a href="https://logarquitectura.com.ar" style="color:#0D0D0D;">logarquitectura.com.ar</a></p>' +
          '<hr style="border:none; border-top:1px solid #e0e0e0; margin:24px 0;">' +
          '<img src="https://lh3.googleusercontent.com/d/1Yn29ePkBM8POROzWUPBljjrRJIEWQlSZ" alt="LOG arquitectura" style="height:40px; width:auto; margin-bottom:6px;"><br>' +
          '<strong style="font-size:13px; color:#0D0D0D; letter-spacing:0.08em;">LOG arquitectura</strong><br>' +
          '<span style="font-size:11px; color:#888;">Estudio de Diseño Arquitectónico</span><br>' +
          '<span style="font-size:12px; color:#555;">Tucumán, Argentina &nbsp;·&nbsp; estudiologarquitectura@gmail.com &nbsp;·&nbsp; +54 381 517 1638</span>' +
          '</div>';
        MailApp.sendEmail({ to: p.email, subject: asuntoCliente, htmlBody: htmlCliente });
      }
    } else {
      // Modal servicios → hoja Solicitudes
      var sheet = ss.getSheetByName(SHEET_NAME);
      if (!sheet) {
        sheet = ss.insertSheet(SHEET_NAME);
        sheet.appendRow(['Fecha', 'Servicio', 'Nombre', 'Email', 'Teléfono', 'Tipo de Proyecto', 'M²', 'Mensaje']);
        sheet.getRange(1, 1, 1, 8).setFontWeight('bold');
      }
      sheet.appendRow([
        p.fecha     || new Date().toLocaleString('es-AR'),
        p.servicio  || '',
        p.nombre    || '',
        p.email     || '',
        p.telefono  || '',
        p.tipo      || '',
        p.m2        || '',
        p.mensaje   || ''
      ]);

      // Correo de confirmación al cliente
      if (p.email) {
        var asuntoCliente = 'Recibimos tu solicitud — LOG arquitectura';
        var htmlCliente =
          '<div style="font-family:Arial,sans-serif; max-width:560px; margin:0 auto; color:#222;">' +
          '<p>Hola ' + (p.nombre || 'allí') + ',</p>' +
          '<p>Gracias por contactarte con LOG arquitectura. Recibimos tu solicitud de <strong>' + (p.servicio || 'servicio') + '</strong> y ya estamos revisando los detalles de tu proyecto.</p>' +
          '<p>Nos pondremos en contacto con vos en las próximas 24 horas para darte una respuesta personalizada.</p>' +
          '<p>Mientras tanto, podés conocer más de nuestro trabajo en: <a href="https://logarquitectura.com.ar" style="color:#0D0D0D;">logarquitectura.com.ar</a></p>' +
          '<hr style="border:none; border-top:1px solid #e0e0e0; margin:24px 0;">' +
          '<img src="https://lh3.googleusercontent.com/d/1Yn29ePkBM8POROzWUPBljjrRJIEWQlSZ" alt="LOG arquitectura" style="height:40px; width:auto; margin-bottom:6px;"><br>' +
          '<strong style="font-size:13px; color:#0D0D0D; letter-spacing:0.08em;">LOG arquitectura</strong><br>' +
          '<span style="font-size:11px; color:#888;">Estudio de Diseño Arquitectónico</span><br>' +
          '<span style="font-size:12px; color:#555;">Tucumán, Argentina &nbsp;·&nbsp; estudiologarquitectura@gmail.com &nbsp;·&nbsp; +54 381 517 1638</span>' +
          '</div>';
        MailApp.sendEmail({
          to: p.email,
          subject: asuntoCliente,
          htmlBody: htmlCliente
        });
      }
    }

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ─── Costo m² diario ─────────────────────────────────────────────────────────

function actualizarCostoM2() {
  try {
    var html  = UrlFetchApp.fetch(COSTO_URL, { muteHttpExceptions: true }).getContentText();
    var match = html.match(/Costo\s+m[²2]\s*[\|l]\s*\$([\d\.,]+)/i);
    if (!match) { Logger.log('No se encontró el costo.'); return; }

    var costo = '$' + match[1];
    var fecha = new Date().toLocaleDateString('es-AR', { timeZone: 'America/Argentina/Buenos_Aires' });
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];

    sheet.getRange('A2').setValue('Costo m² construcción');
    sheet.getRange('B2').setValue(costo);
    sheet.getRange('C2').setValue('Actualizado: ' + fecha);

    Logger.log('Costo actualizado: ' + costo + ' (' + fecha + ')');
  } catch (err) {
    Logger.log('Error: ' + err.toString());
  }
}

// ─── Trigger diario (ejecutá una vez) ────────────────────────────────────────

function crearTriggerDiario() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'actualizarCostoM2') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('actualizarCostoM2')
    .timeBased().atHour(8).everyDays(1)
    .inTimezone('America/Argentina/Buenos_Aires')
    .create();
  Logger.log('Trigger diario creado — actualizarCostoM2 corre cada día a las 8 AM.');
}
