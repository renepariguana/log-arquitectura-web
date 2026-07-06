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

  if (e && e.parameter && e.parameter.action === 'costo') {
    try {
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Servicios');
      var costo = sheet.getRange('C3').getValue().toString().trim();
      return ContentService
        .createTextOutput(JSON.stringify({ costo: costo }))
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

      // Estructura Clientes: A=Nº, B=Nombre, C=Email, D=CARPETA, E=SUSCRIPCION, F=Inicio, G=Fin suscripcion, H=Tiempo restante, I=LINK DRIVE
      // Estructura actual del Sheet: A=Nº, B=Nombre, C=Email, D=Suscripcion, E=LINK DRIVE
      var hojaClientes = ss.getSheetByName('Clientes');
      if (!hojaClientes) {
        hojaClientes = ss.insertSheet('Clientes');
        hojaClientes.appendRow(['Nº', 'Nombre', 'Email', 'Suscripcion', 'LINK DRIVE']);
        hojaClientes.getRange(1, 1, 1, 5).setFontWeight('bold');
      }

      var folderId = '', esCliente = false, suscripcion = false, nroSuscriptor = '';
      var filaEncontrada = -1;
      var datosC = hojaClientes.getDataRange().getValues();

      for (var i = 1; i < datosC.length; i++) {
        var emailC = (datosC[i][2] || '').toString().trim().toLowerCase(); // Col C: Email
        if (emailC !== email) continue;

        nroSuscriptor  = (datosC[i][0] || '').toString().trim();           // Col A: Nº
        nombre         = (datosC[i][1] || nombre).toString().trim();        // Col B: Nombre
        var textoSusc  = (datosC[i][3] || '').toString().trim().toLowerCase(); // Col D: Suscripcion
        suscripcion    = (textoSusc === 'activa' || textoSusc === 'activo' || textoSusc === 'active' || textoSusc === 'approved');

        var linkDriveC = (datosC[i][4] || '').toString().trim();           // Col E: LINK DRIVE
        var fm = linkDriveC.match(/[-\w]{25,}/);
        folderId       = fm ? fm[0] : '';
        esCliente      = true;
        filaEncontrada = i + 1;
        break;
      }

      // Validar que el folderId almacenado realmente exista en Drive
      if (esCliente && folderId) {
        try { DriveApp.getFolderById(folderId); } catch(e) { folderId = ''; }
      }

      // Formato de nombre de carpeta: "001 — Nombre"
      function nombreCarpeta(nro, nom) {
        return nro ? nro + ' — ' + nom : nom;
      }

      // Encontrado pero sin carpeta Drive → buscar existente o crear
      if (esCliente && !folderId) {
        var carpetaPadre = DriveApp.getFolderById(CLIENTES_FOLDER_ID);
        var buscaNro = carpetaPadre.getFoldersByName(nombreCarpeta(nroSuscriptor, nombre));
        var buscaNom = carpetaPadre.getFoldersByName(nombre || email);
        var carpeta;
        if (buscaNro.hasNext()) {
          carpeta = buscaNro.next();
        } else if (buscaNom.hasNext()) {
          carpeta = buscaNom.next();
          if (nroSuscriptor) carpeta.setName(nombreCarpeta(nroSuscriptor, nombre));
        } else {
          carpeta = carpetaPadre.createFolder(nombreCarpeta(nroSuscriptor, nombre));
        }
        folderId = carpeta.getId();
        hojaClientes.getRange(filaEncontrada, 5).setValue('https://drive.google.com/drive/folders/' + folderId);
      }

      // No existe en Clientes → auto-numerar, crear carpeta, agregar fila
      if (!esCliente) {
        var siguiente = 1;
        for (var k = 1; k < datosC.length; k++) {
          var nroK = parseInt(datosC[k][0]);
          if (!isNaN(nroK) && nroK >= siguiente) siguiente = nroK + 1;
        }
        nroSuscriptor = String(siguiente).padStart(3, '0');

        var carpetaPadre2 = DriveApp.getFolderById(CLIENTES_FOLDER_ID);
        var nomCarpeta2   = nombreCarpeta(nroSuscriptor, nombre || email);
        var buscaNro2     = carpetaPadre2.getFoldersByName(nomCarpeta2);
        var buscaNom2     = carpetaPadre2.getFoldersByName(nombre || email);
        var carpeta2;
        if (buscaNro2.hasNext()) {
          carpeta2 = buscaNro2.next();
        } else if (buscaNom2.hasNext()) {
          carpeta2 = buscaNom2.next();
          carpeta2.setName(nomCarpeta2);
        } else {
          carpeta2 = carpetaPadre2.createFolder(nomCarpeta2);
        }
        folderId = carpeta2.getId();
        var linkNuevo = 'https://drive.google.com/drive/folders/' + folderId;
        hojaClientes.appendRow([nroSuscriptor, nombre, email, 'pendiente', linkNuevo]);
        esCliente = true;
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

  if (e && e.parameter && e.parameter.action === 'manoobra') {
    try {
      var AP_SS  = SpreadsheetApp.openById('1mzWe4dXvMYvRDJA9MIqjJDy3U2kZY133qKyICT6W43s');
      var hojaM  = AP_SS.getSheetByName('mano de obra')
                || AP_SS.getSheetByName('Mano de Obra')
                || AP_SS.getSheetByName('MO');
      if (!hojaM) throw new Error('Hoja "mano de obra" no encontrada');

      var datosM = hojaM.getDataRange().getValues();
      var items  = [];
      var header = datosM[0] || [];

      for (var ri = 1; ri < datosM.length; ri++) {
        var fila = datosM[ri];
        if (!fila[0] && !fila[1]) continue;
        var obj = {};
        for (var ci = 0; ci < header.length; ci++) {
          var key = (header[ci] || 'col' + ci).toString().trim().toLowerCase()
            .replace(/\s+/g, '_').replace(/[^a-z0-9_]/g, '');
          obj[key] = fila[ci];
        }
        items.push(obj);
      }

      return ContentService
        .createTextOutput(JSON.stringify(items))
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
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Servicios');

    sheet.getRange('C3').setValue(costo);

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

  var nro        = sheet.getRange(row, 1).getValue().toString().trim(); // col A: Nº
  var nombre     = sheet.getRange(row, 2).getValue().toString().trim(); // col B: Nombre
  var linkActual = sheet.getRange(row, 5).getValue().toString().trim(); // col E: LINK DRIVE

  if (!nombre || linkActual) return;

  var nomCarpeta   = nro ? nro + ' — ' + nombre : nombre;
  var carpetaPadre = DriveApp.getFolderById(CLIENTES_FOLDER_ID);
  var existentes   = carpetaPadre.getFoldersByName(nomCarpeta);
  var carpeta      = existentes.hasNext() ? existentes.next() : carpetaPadre.createFolder(nomCarpeta);
  sheet.getRange(row, 5).setValue(carpeta.getUrl()); // col E: LINK DRIVE
}

// ─── Configurar Access Token de MercadoPago (ejecutá una vez) ────────────────

function configurarMPToken() {
  PropertiesService.getScriptProperties().setProperty(
    'MP_ACCESS_TOKEN',
    'TEST-4519992620088034-070312-dcb30786a1159aa0285ab249a61f81fd-151587833'
  );
  Logger.log('MP Access Token guardado.');
}

// ─── UOCRA — Actualizar escalas salariales mano de obra ──────────────────────
// REQUISITO PREVIO: en el proyecto Apps Script ir a
//   Servicios → Agregar servicio → Drive API v2 → Agregar
// Luego ejecutar crearTriggerMensualUOCRA() una sola vez.

var UOCRA_SHEET_ID  = '1mzWe4dXvMYvRDJA9MIqjJDy3U2kZY133qKyICT6W43s';
var UOCRA_EMAIL     = 'estudiologarquitectura@gmail.com';
var UOCRA_PAGE_URL  = 'https://www.uocra.org/?s=nuevas-escalas-salariales&lang=1';
// Tucumán pertenece a Zona A (CCT 76/75 — construcción interior del país)
// Zona A: CABA, Bs.As., Santa Fe, Córdoba, Mendoza, Tucumán y otras provincias del NOA/NEA/centro
// Zona B: Neuquén, Río Negro, Chubut | Zona C: Santa Cruz | Zona C Austral: Tierra del Fuego
var UOCRA_ZONA      = 'A';
var UOCRA_CCT       = '76'; // buscar PDF del CCT 76/75 (construcción interior)

function actualizarSalariosUOCRA() {
  try {
    // 1. Scrapear página UOCRA — extraer todos los links a PDF
    var html = UrlFetchApp.fetch(UOCRA_PAGE_URL, { muteHttpExceptions: true }).getContentText();
    var pdfLinks = [];
    var reLink = /href="([^"]*\.pdf[^"]*)"/gi;
    var m;
    while ((m = reLink.exec(html)) !== null) pdfLinks.push(m[1]);

    if (!pdfLinks.length) {
      Logger.log('UOCRA: No se encontraron PDFs en la página.');
      return;
    }

    // Buscar primero un PDF del CCT 76/75 (aplica a Tucumán, interior del país)
    // Si no se encuentra, usar el primero disponible como fallback
    var pdfPath = '';
    for (var i = 0; i < pdfLinks.length; i++) {
      if (/76[-_]?75|76[-_]75|cct[-_]?76/i.test(pdfLinks[i])) {
        pdfPath = pdfLinks[i];
        break;
      }
    }
    if (!pdfPath) pdfPath = pdfLinks[0]; // fallback al primero

    var pdfUrl  = pdfPath.startsWith('http')
      ? pdfPath
      : 'https://www.uocra.org/' + pdfPath.replace(/^\//, '');

    // 2. Verificar si ya fue procesado
    var props    = PropertiesService.getScriptProperties();
    var ultimoUrl = props.getProperty('UOCRA_LAST_PDF') || '';
    if (ultimoUrl === pdfUrl) {
      Logger.log('UOCRA: Sin novedades. Último PDF procesado: ' + pdfUrl);
      return;
    }

    // 3. Descargar PDF
    var pdfBlob = UrlFetchApp.fetch(pdfUrl, { muteHttpExceptions: true })
      .getBlob().setName('uocra_escalas.pdf');

    // 4. Subir a Drive con OCR (convierte imagen escaneada a texto)
    var tempDoc = Drive.Files.insert(
      { title: 'uocra_ocr_' + Date.now(), mimeType: 'application/vnd.google-apps.document' },
      pdfBlob,
      { ocr: true, ocrLanguage: 'es' }
    );

    // 5. Extraer texto del documento
    var texto = DocumentApp.openById(tempDoc.id).getBody().getText();
    DriveApp.getFileById(tempDoc.id).setTrashed(true);

    // 6. Parsear salarios y actualizar hoja
    props.setProperty('UOCRA_LAST_PDF', pdfUrl);
    var salarios = parsearSalariosUOCRA(texto);
    var fecha    = new Date().toLocaleDateString('es-AR', { timeZone: 'America/Argentina/Buenos_Aires' });

    var hayDatos = salarios && salarios.basicos && Object.keys(salarios.basicos).length > 0;
    if (hayDatos) {
      actualizarHojaManoObra(salarios, fecha);
      // Armar resumen del email
      var lineasEmail = ['Salarios básicos (iguales para todas las zonas):'];
      var basicosA = salarios.basicos['Zona A'] || {};
      Object.keys(basicosA).forEach(function(cat) {
        lineasEmail.push('  • ' + cat + ': $' + Number(basicosA[cat]).toLocaleString('es-AR') + '/h');
      });
      if (salarios.adicionales && Object.keys(salarios.adicionales).length) {
        lineasEmail.push('\nAdicionales por zona:');
        Object.keys(salarios.adicionales).forEach(function(zona) {
          lineasEmail.push('  ' + zona + ':');
          var adics = salarios.adicionales[zona];
          Object.keys(adics).forEach(function(cat) {
            lineasEmail.push('    + $' + Number(adics[cat]).toLocaleString('es-AR') + ' (' + cat + ')');
          });
        });
      }
      MailApp.sendEmail(UOCRA_EMAIL,
        'UOCRA — Salarios actualizados ' + fecha,
        'Fuente: ' + pdfUrl + '\n\n' + lineasEmail.join('\n')
      );
      Logger.log('UOCRA: Actualizado — ' + JSON.stringify(salarios));
    } else {
      MailApp.sendEmail(UOCRA_EMAIL,
        'UOCRA — Nuevo PDF detectado (revisión manual)',
        'Se detectó un nuevo PDF pero no se pudo parsear automáticamente.\n' +
        'URL: ' + pdfUrl + '\n\n' +
        'Texto extraído:\n' + texto.substring(0, 3000)
      );
      Logger.log('UOCRA: Parseo fallido para ' + pdfUrl);
    }

  } catch (err) {
    Logger.log('UOCRA Error: ' + err.toString());
    MailApp.sendEmail(UOCRA_EMAIL, 'UOCRA — Error al actualizar salarios', err.toString());
  }
}

function parsearSalariosUOCRA(texto) {
  // Retorna: { basicos: { 'Zona A': {cat: val}, ... }, adicionales: { 'Zona B': {cat: val}, ... } }
  // CCT 76/75: Zona A/B/C/C Austral  |  CCT 545/08 fallback: Zona I/II/III/IV
  var NCAT = 5;
  var categorias = ['Ayudante', 'Medio oficial', 'Oficial', 'Oficial especializado', 'Sereno (mensual)'];

  var zonaDefs = [
    { key: 'Zona A',         re: /zona\s*[Aa]\b/,                hayAdic: false },
    { key: 'Zona B',         re: /zona\s*[Bb]\b/,                hayAdic: true  },
    { key: 'Zona C Austral', re: /zona\s*[Cc]\s*austral/i,       hayAdic: true  },
    { key: 'Zona C',         re: /zona\s*[Cc]\b(?!\s*austral)/i, hayAdic: true  },
    { key: 'Zona A',         re: /zona\s*[I1l]\b(?!\s*[I1Vv])/,  hayAdic: false },
    { key: 'Zona B',         re: /zona\s*II\b/i,                 hayAdic: true  },
    { key: 'Zona C',         re: /zona\s*III\b/i,                hayAdic: true  },
    { key: 'Zona C Austral', re: /zona\s*IV\b/i,                 hayAdic: true  }
  ];

  var lineas     = texto.split('\n');
  var basicos    = {};
  var adicionales = {};

  zonaDefs.forEach(function(z) {
    var k = z.key;
    if (basicos[k] && (!z.hayAdic || adicionales[k])) return;

    var ultimaFila = -1;
    for (var i = 0; i < lineas.length; i++) {
      if (z.re.test(lineas[i]) && /[\d.,]{3,}/.test(lineas[i])) ultimaFila = i;
    }
    if (ultimaFila === -1) return;

    var filaTexto = lineas.slice(ultimaFila, ultimaFila + 3).join(' ');
    var reNum = /\b\d{1,3}(?:[.,]\d{3})+(?:[.,]\d+)?\b|\b\d{3,}\b/g;
    var nums = [], m;
    while ((m = reNum.exec(filaTexto)) !== null) {
      var val = parseFloat(m[0].replace(/\./g, '').replace(',', '.'));
      if (val >= 50 && val < 100000000) nums.push(val);
    }
    if (!nums.length) return;

    // Si hay > NCAT+4 números, la fila tiene intercalados básico+adicional por categoría
    var intercalado = z.hayAdic && nums.length > NCAT + 4;
    var bObj = {}, aObj = {};

    for (var i = 0; i < NCAT; i++) {
      var bi = intercalado ? i * 2     : i;
      var ai = intercalado ? i * 2 + 1 : -1;
      if (bi < nums.length) bObj[categorias[i]] = nums[bi];
      if (ai !== -1 && ai < nums.length) aObj[categorias[i]] = nums[ai];
    }

    if (!basicos[k])      basicos[k] = bObj;
    if (z.hayAdic && Object.keys(aObj).length && !adicionales[k]) adicionales[k] = aObj;
  });

  return { basicos: basicos, adicionales: adicionales };
}

function actualizarHojaManoObra(data, fecha) {
  // data = { basicos: { 'Zona A': {cat:val}, ... }, adicionales: { 'Zona B': {cat:val}, ... } }
  var ss   = SpreadsheetApp.openById(UOCRA_SHEET_ID);
  var hoja = ss.getSheetByName('MANO DE OBRA')
          || ss.getSheetByName('mano de obra')
          || ss.getSheetByName('Mano de Obra')
          || ss.getSheetByName('MO');
  if (!hoja) throw new Error('Hoja "MANO DE OBRA" no encontrada.');

  var datos = hoja.getDataRange().getValues();

  // Detectar la fila de encabezados sub-columna (tiene "CATEGORIA" y "Salario Básico")
  // y mapear cada tipo de columna a su índice.
  // Estructura esperada:
  //   "Salario Básico Zona A"         → basico_A
  //   "Salario Básico Zona B"         → basico_B
  //   "Adicional zona desfavorable... B" → adicional_B
  //   "Salario Básico Zona C Austral" → basico_CA
  //   "Adicional ... C Austral"       → adicional_CA
  //   "Salario Básico Zona C"         → basico_C  (detectar DESPUÉS de CA para no confundir)
  //   "Adicional ... C"               → adicional_C

  var colCat  = 0;
  var colMap  = {}; // { 'basico_A': col, 'basico_B': col, 'adicional_B': col, ... }
  var filaHeader = -1;

  for (var r = 0; r < datos.length; r++) {
    var fila = datos[r];
    var encontro = false;
    for (var c = 0; c < fila.length; c++) {
      var t = (fila[c] || '').toString().toLowerCase().replace(/["""]/g, '').trim();
      if (!t) continue;

      if (/^categor/.test(t))                                         { colCat = c; encontro = true; }

      // Básicos — C Austral ANTES que C para evitar match parcial
      if (/salario.*b[aá]s.*zona.*a\b/.test(t))                      { colMap.basico_A  = c; encontro = true; }
      if (/salario.*b[aá]s.*zona.*b\b/.test(t))                      { colMap.basico_B  = c; encontro = true; }
      if (/salario.*b[aá]s.*zona.*c.*austral/.test(t))               { colMap.basico_CA = c; encontro = true; }
      else if (/salario.*b[aá]s.*zona.*c\b/.test(t))                 { colMap.basico_C  = c; encontro = true; }

      // Adicionales
      if (/adicional.*zona.*b\b|zona.*b.*adicional/.test(t))         { colMap.adicional_B  = c; encontro = true; }
      if (/adicional.*zona.*c.*austral|zona.*c.*austral.*adicional/.test(t)) { colMap.adicional_CA = c; encontro = true; }
      else if (/adicional.*zona.*c\b|zona.*c.*adicional/.test(t))    { colMap.adicional_C  = c; encontro = true; }
    }
    if (encontro) { filaHeader = r; break; }
  }

  var zonaColDef = [
    { zona: 'Zona A',         colBasico: 'basico_A',  colAdic: null         },
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
      var basicosZona = (data.basicos || {})[def.zona] || {};
      var adicsZona   = (data.adicionales || {})[def.zona] || {};

      for (var cat in basicosZona) {
        var a = catCelda.toLowerCase(), b = cat.toLowerCase();
        if (a.indexOf(b) === -1 && b.indexOf(a) === -1) continue;

        // Escribir básico
        var colB = colMap[def.colBasico];
        if (colB !== undefined) {
          hoja.getRange(r + 1, colB + 1).setValue(basicosZona[cat]);
          actualizados.push(def.zona + ' ' + cat + ' básico $' + basicosZona[cat]);
        }
        // Escribir adicional (Zona B/C/C Austral)
        if (def.colAdic && adicsZona[cat] !== undefined) {
          var colA = colMap[def.colAdic];
          if (colA !== undefined) {
            hoja.getRange(r + 1, colA + 1).setValue(adicsZona[cat]);
            actualizados.push(def.zona + ' ' + cat + ' adicional $' + adicsZona[cat]);
          }
        }
        break;
      }
    });
  }

  Logger.log('UOCRA: ' + actualizados.join(' | '));
  return actualizados;
}

// Ejecutar UNA VEZ para instalar el trigger mensual
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
  Logger.log('Trigger mensual creado: actualizarSalariosUOCRA corre el día 1 de cada mes a las 9 AM.');
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
