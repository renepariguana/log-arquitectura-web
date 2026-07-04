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
var CLIENTES_FOLDER_ID  = '1SqXMjyeBs4uHa7E0dpwe408oQKL45545';
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

  if (e && e.parameter && e.parameter.action === 'items') {
    try {
      var AP_ID = '1mzWe4dXvMYvRDJA9MIqjJDy3U2kZY133qKyICT6W43s';
      var apSS  = SpreadsheetApp.openById(AP_ID);
      var hoja  = apSS.getSheetByName('BASE DE DATOS')
               || apSS.getSheetByName('Inicio')
               || apSS.getSheetByName('base de datos');
      if (!hoja) throw new Error('Hoja BASE DE DATOS no encontrada en ANALISIS DE PRECIOS');

      var datos  = hoja.getDataRange().getValues();
      var rubros = [];
      var items  = [];

      // Encontrar primer fila con código numérico (saltar encabezados)
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

  if (e && e.parameter && e.parameter.action === 'login') {
    try {
      var ss      = SpreadsheetApp.getActiveSpreadsheet();
      var hoja    = ss.getSheetByName('Clientes');
      if (!hoja) throw new Error('Hoja Clientes no encontrada');

      var datos   = hoja.getDataRange().getValues();
      // Fila 0 = encabezados, buscar desde fila 1
      for (var i = 1; i < datos.length; i++) {
        var nombre   = (datos[i][0] || '').toString().trim();
        var email    = (datos[i][1] || '').toString().trim().toLowerCase();
        var dni      = (datos[i][2] || '').toString().trim();
        var estado    = (datos[i][4] || '').toString().trim().toLowerCase();
        var linkDrive = (datos[i][5] || '').toString().trim();

        if (email !== (e.parameter.email || '').toLowerCase().trim()) continue;
        if (estado === 'inactivo') {
          return ContentService
            .createTextOutput(JSON.stringify({ ok: false, error: 'Cuenta inactiva' }))
            .setMimeType(ContentService.MimeType.JSON);
        }
        if (dni !== (e.parameter.dni || '').trim()) {
          return ContentService
            .createTextOutput(JSON.stringify({ ok: false, error: 'DNI incorrecto' }))
            .setMimeType(ContentService.MimeType.JSON);
        }

        // Extraer ID de la carpeta del link de Drive
        var folderMatch = linkDrive.match(/[-\w]{25,}/);
        var folderId = folderMatch ? folderMatch[0] : '';

        return ContentService
          .createTextOutput(JSON.stringify({ ok: true, nombre: nombre, folderId: folderId }))
          .setMimeType(ContentService.MimeType.JSON);
      }

      return ContentService
        .createTextOutput(JSON.stringify({ ok: false, error: 'Email no encontrado' }))
        .setMimeType(ContentService.MimeType.JSON);

    } catch (err) {
      return ContentService
        .createTextOutput(JSON.stringify({ ok: false, error: err.toString() }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  if (e && e.parameter && e.parameter.action === 'loginGoogle') {
    try {
      var email  = (e.parameter.email  || '').toLowerCase().trim();
      var nombre = (e.parameter.nombre || '').toString().trim();
      if (!email) throw new Error('Email requerido');
      var ss = SpreadsheetApp.getActiveSpreadsheet();

      // Buscar en Clientes
      var hojaClientes = ss.getSheetByName('Clientes');
      if (!hojaClientes) {
        hojaClientes = ss.insertSheet('Clientes');
        hojaClientes.appendRow(['Nombre', 'Email', 'DNI', '', 'Estado', 'LinkDrive']);
        hojaClientes.getRange(1, 1, 1, 6).setFontWeight('bold');
      }
      var folderId = '', esCliente = false;
      var datosC = hojaClientes.getDataRange().getValues();
      for (var i = 1; i < datosC.length; i++) {
        var emailC  = (datosC[i][1] || '').toString().trim().toLowerCase();
        var estadoC = (datosC[i][4] || '').toString().trim().toLowerCase();
        if (emailC === email) {
          if (estadoC === 'inactivo') break;
          nombre    = (datosC[i][0] || nombre).toString().trim();
          var link  = (datosC[i][5] || '').toString().trim();
          var fm    = link.match(/[-\w]{25,}/);
          folderId  = fm ? fm[0] : '';
          esCliente = true;
          break;
        }
      }

      // Si no existe → crear carpeta en Drive y registrar en Clientes
      if (!esCliente) {
        var CARPETA_CLIENTES_ID = '1SqXMjyeBs4uHa7E0dpwe408oQKL45545';
        var carpetaPadre = DriveApp.getFolderById(CARPETA_CLIENTES_ID);
        var nombreCarpeta = nombre || email;
        var nuevaCarpeta  = carpetaPadre.createFolder(nombreCarpeta);
        folderId  = nuevaCarpeta.getId();
        var linkDrive = 'https://drive.google.com/drive/folders/' + folderId;
        hojaClientes.appendRow([nombre, email, '', '', 'activo', linkDrive]);
        esCliente = true;
      }

      // Buscar en Suscriptores
      // Estructura: A=CARPETA, B=Nombre, C=Email, D=SUSCRIPCION, E=INICIO, F=FIN, G=TIEMPO RESTANTE, H=NRO SUSCRIPTOR
      var hojaSub = ss.getSheetByName('Suscriptores');
      var suscripcion = false;
      var nroSuscriptor = '';
      if (hojaSub) {
        var datosS = hojaSub.getDataRange().getValues();
        for (var j = 1; j < datosS.length; j++) {
          var emailS  = (datosS[j][2] || '').toString().trim().toLowerCase(); // Col C
          var estadoS = (datosS[j][3] || '').toString().trim().toLowerCase(); // Col D
          if (emailS === email && (estadoS === 'activa' || estadoS === 'activo' || estadoS === 'approved' || estadoS === 'active')) {
            suscripcion = true;
            nroSuscriptor = (datosS[j][7] || '').toString().trim(); // Col H
            if (!nroSuscriptor) {
              // Contar cuántos suscriptores ya tienen número asignado
              var count = 0;
              for (var k = 1; k < datosS.length; k++) {
                if ((datosS[k][7] || '').toString().trim()) count++;
              }
              nroSuscriptor = String(count + 1).padStart(3, '0');
              hojaSub.getRange(j + 1, 8).setValue(nroSuscriptor);
            }
            break;
          }
        }
      }

      return ContentService
        .createTextOutput(JSON.stringify({ ok: true, nombre: nombre, folderId: folderId, esCliente: esCliente, suscripcion: suscripcion, nroSuscriptor: nroSuscriptor }))
        .setMimeType(ContentService.MimeType.JSON);

    } catch (err) {
      return ContentService
        .createTextOutput(JSON.stringify({ ok: false, error: err.toString() }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  if (e && e.parameter && e.parameter.action === 'archivos') {
    try {
      var folderId = e.parameter.folderId || '';
      if (!folderId) throw new Error('folderId requerido');

      var clienteFolder = DriveApp.getFolderById(folderId);
      var subFolders    = clienteFolder.getFolders();
      var proyectos     = [];

      while (subFolders.hasNext()) {
        var sub     = subFolders.next();
        var fils    = sub.getFiles();
        var pdfs    = [];
        var coverId = '';
        while (fils.hasNext()) {
          var file = fils.next();
          var mime = file.getMimeType();
          if (mime === 'application/pdf') {
            var nombreCompleto = file.getName().replace(/\.pdf$/i, '');
            var prefijo = sub.getName() + '_';
            var nombreMostrar = nombreCompleto.indexOf(prefijo) === 0
              ? nombreCompleto.slice(prefijo.length)
              : nombreCompleto;
            pdfs.push({ id: file.getId(), nombre: nombreMostrar });
          } else if (!coverId && mime.match(/^image\//)) {
            coverId = file.getId();
          }
        }
        pdfs.sort(function(a, b) { return a.nombre.localeCompare(b.nombre, 'es'); });
        proyectos.push({ proyecto: sub.getName(), coverId: coverId, pdfs: pdfs });
      }

      proyectos.sort(function(a, b) { return a.proyecto.localeCompare(b.proyecto, 'es'); });

      return ContentService
        .createTextOutput(JSON.stringify(proyectos))
        .setMimeType(ContentService.MimeType.JSON);

    } catch (err) {
      return ContentService
        .createTextOutput(JSON.stringify({ error: err.toString() }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  if (e && e.parameter && e.parameter.action === 'precios') {
    try {
      var sheet  = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Servicios');
      var rango  = sheet.getRange('E10:F13').getValues();
      var result = {};
      for (var i = 0; i < rango.length; i++) {
        var nombre = rango[i][0].toString().trim().toLowerCase();
        var precio = Number(rango[i][1]);
        if (nombre && precio) result[nombre] = precio;
      }
      return ContentService
        .createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    } catch(err) {
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

    if (p.action === 'crear-preferencia') {
      var ACCESS_TOKEN = PropertiesService.getScriptProperties().getProperty('MP_ACCESS_TOKEN');
      var plan  = p.plan  || '';
      var email = p.email || '';

      // Leer precio desde hoja Servicios E10:F13
      var sheetServ = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Servicios');
      var rango = sheetServ.getRange('E10:F13').getValues();
      var monto = 0;
      for (var ri = 0; ri < rango.length; ri++) {
        if (rango[ri][0].toString().trim().toLowerCase() === plan.toLowerCase()) {
          monto = Number(rango[ri][1]);
          break;
        }
      }

      var preference = {
        items: [{ title: 'Suscripción LOG arquitectura — ' + plan, quantity: 1, unit_price: monto, currency_id: 'ARS' }],
        payer: { email: email },
        back_urls: { success: 'https://logarquitectura.com.ar', failure: 'https://logarquitectura.com.ar', pending: 'https://logarquitectura.com.ar' },
        auto_return: 'approved'
      };

      var mpRes = UrlFetchApp.fetch('https://api.mercadopago.com/checkout/preferences', {
        method: 'post',
        contentType: 'application/json',
        headers: { 'Authorization': 'Bearer ' + ACCESS_TOKEN },
        payload: JSON.stringify(preference),
        muteHttpExceptions: true
      });

      var mpData = JSON.parse(mpRes.getContentText());
      var url = mpData.sandbox_init_point || mpData.init_point || '';

      return ContentService
        .createTextOutput(JSON.stringify({ url: url, monto: monto }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (p.action === 'procesar-pago') {
      var ACCESS_TOKEN = PropertiesService.getScriptProperties().getProperty('MP_ACCESS_TOKEN');
      var formData     = JSON.parse(p.formData || '{}');
      var response     = UrlFetchApp.fetch('https://api.mercadopago.com/v1/payments', {
        method: 'post',
        contentType: 'application/json',
        headers: {
          'Authorization': 'Bearer ' + ACCESS_TOKEN,
          'X-Idempotency-Key': Utilities.getUuid()
        },
        payload: JSON.stringify(formData),
        muteHttpExceptions: true
      });
      var result = JSON.parse(response.getContentText());
      // Actualizar estado en hoja Suscriptores si pago aprobado
      if (result.status === 'approved' || result.status === 'in_process') {
        var sheetSub = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Suscriptores');
        if (sheetSub) {
          var email = (result.payer && result.payer.email) ? result.payer.email.toLowerCase() : '';
          var datos = sheetSub.getDataRange().getValues();
          for (var r = 1; r < datos.length; r++) {
            if (datos[r][5].toString().toLowerCase() === email && datos[r][9] === 'pendiente') {
              sheetSub.getRange(r + 1, 10).setValue(result.status);
              break;
            }
          }
        }
      }
      return ContentService
        .createTextOutput(JSON.stringify({ status: result.status, id: result.id }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (p.action === 'suscripcion') {
      var sheetSub = ss.getSheetByName('Suscriptores');
      if (!sheetSub) {
        sheetSub = ss.insertSheet('Suscriptores');
        sheetSub.appendRow(['Fecha', 'Nombre', 'Apellidos', 'País', 'Provincia', 'Email', 'Usuario', 'Contraseña', 'Plan', 'Estado']);
        sheetSub.getRange(1, 1, 1, 10).setFontWeight('bold');
      }
      sheetSub.appendRow([
        new Date().toLocaleString('es-AR'),
        p.nombre    || '',
        p.apellidos || '',
        p.pais      || '',
        p.provincia || '',
        p.email     || '',
        p.usuario   || '',
        p.pass      || '',
        p.plan      || '',
        'pendiente'
      ]);
      return ContentService
        .createTextOutput(JSON.stringify({ status: 'ok' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

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

// ─── Auto-crear carpeta Drive al agregar cliente ──────────────────────────────

function autoCrearCarpeta(e) {
  var sheet = e.range.getSheet();
  if (sheet.getName() !== 'Clientes') return;

  var row = e.range.getRow();
  if (row < 2) return;

  var nombre     = sheet.getRange(row, 1).getValue().toString().trim(); // col A: Nombre
  var linkActual = sheet.getRange(row, 6).getValue().toString().trim(); // col F: LINK

  if (!nombre || linkActual) return;

  var carpetaPadre = DriveApp.getFolderById(CLIENTES_FOLDER_ID);
  var nuevaCarpeta = carpetaPadre.createFolder(nombre);
  sheet.getRange(row, 6).setValue(nuevaCarpeta.getUrl());
}

// ─── Configurar Access Token de MercadoPago (ejecutá una vez) ────────────────

function configurarMPToken() {
  PropertiesService.getScriptProperties().setProperty(
    'MP_ACCESS_TOKEN',
    'TEST-4519992620088034-070312-dcb30786a1159aa0285ab249a61f81fd-151587833'
  );
  Logger.log('MP Access Token guardado.');
}

function crearTriggerClientes() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'autoCrearCarpeta') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('autoCrearCarpeta')
    .forSpreadsheet(ss)
    .onEdit()
    .create();
  Logger.log('Trigger autoCrearCarpeta instalado.');
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
