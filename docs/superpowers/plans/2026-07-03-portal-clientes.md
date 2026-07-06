# Portal de Clientes — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Agregar una sección "Clientes" al sitio donde cada cliente se loguea con email + DNI y ve/descarga sus PDFs desde Google Drive.

**Architecture:** Login vía Apps Script que verifica email+DNI contra la hoja "Clientes" del Google Sheet activo. Si coincide, devuelve el folderId de Drive del cliente. El portal lista los PDFs de esa carpeta y los muestra con un visor embebido de Google Drive.

**Tech Stack:** HTML + Tailwind CSS (CDN) + Google Apps Script + Google Drive

## Global Constraints

- Mismo stack que el resto del sitio: HTML puro, Tailwind via CDN, DM Sans, colores cream `#F7F5F2` y off-black `#0D0D0D`
- No agregar dependencias externas nuevas
- El Apps Script usa `SpreadsheetApp.getActiveSpreadsheet()` — la hoja "Clientes" debe existir en ese mismo spreadsheet
- Columnas hoja "Clientes": Nombre | Proyecto | Email | Contraseña (DNI) | Link Drive | Estado
- La carpeta Drive de cada cliente se identifica por su ID (extraído del Link Drive)
- Cada vez que se modifica log-arquitectura.gs hay que publicar una **nueva versión** del Web App y actualizar `GSHEET_URL` en index.html

---

### Task 1: Apps Script — endpoint login + listar archivos

**Files:**
- Modify: `log-arquitectura.gs`

**Interfaces:**
- Produce: `doPost({ action:"login", email, dni })` → `{ ok:true, nombre, folderId }` | `{ ok:false, error }`
- Produce: `doGet?action=archivos&folderId=XXX` → `[{ id, nombre }]`

- [ ] **Step 1: Agregar acción `login` al `doPost`**

En `log-arquitectura.gs`, dentro de la función `doPost`, antes del bloque `if (p.hoja === 'mensajes')`, agregar:

```javascript
if (p.action === 'login') {
  try {
    var ss      = SpreadsheetApp.getActiveSpreadsheet();
    var hoja    = ss.getSheetByName('Clientes');
    if (!hoja) throw new Error('Hoja Clientes no encontrada');

    var datos   = hoja.getDataRange().getValues();
    // Fila 0 = encabezados, buscar desde fila 1
    for (var i = 1; i < datos.length; i++) {
      var nombre   = (datos[i][0] || '').toString().trim();
      var email    = (datos[i][2] || '').toString().trim().toLowerCase();
      var dni      = (datos[i][3] || '').toString().trim();
      var linkDrive = (datos[i][4] || '').toString().trim();
      var estado   = (datos[i][5] || '').toString().trim().toLowerCase();

      if (email !== (p.email || '').toLowerCase().trim()) continue;
      if (estado === 'inactivo') {
        return ContentService
          .createTextOutput(JSON.stringify({ ok: false, error: 'Cuenta inactiva' }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      if (dni !== (p.dni || '').trim()) {
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
```

- [ ] **Step 2: Agregar acción `archivos` al `doGet`**

En `log-arquitectura.gs`, dentro de `doGet`, agregar un nuevo bloque `if` antes del `return` final:

```javascript
if (e && e.parameter && e.parameter.action === 'archivos') {
  try {
    var folderId = e.parameter.folderId || '';
    if (!folderId) throw new Error('folderId requerido');

    var folder = DriveApp.getFolderById(folderId);
    var files  = folder.getFiles();
    var pdfs   = [];

    while (files.hasNext()) {
      var file = files.next();
      if (file.getMimeType() !== 'application/pdf') continue;
      pdfs.push({
        id    : file.getId(),
        nombre: file.getName().replace(/\.pdf$/i, '')
      });
    }

    pdfs.sort(function(a, b) { return a.nombre.localeCompare(b.nombre, 'es'); });

    return ContentService
      .createTextOutput(JSON.stringify(pdfs))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
```

- [ ] **Step 3: Publicar nueva versión del Apps Script**

En Google Apps Script:
1. Implementar → Administrar implementaciones → Nueva versión
2. Copiar la nueva URL
3. Actualizar `GSHEET_URL` en index.html (línea ~1239)

- [ ] **Step 4: Commit**

```bash
git add log-arquitectura.gs
git commit -m "feat: agregar login y listar archivos al Apps Script"
```

---

### Task 2: Nav — botón Clientes (desktop + mobile)

**Files:**
- Modify: `index.html` (líneas ~290-321)

- [ ] **Step 1: Agregar botón en nav desktop**

Después del botón "Cotizar" (línea ~300), antes del `</div>` del `nav-links`, agregar:

```html
      <button onclick="openClientesLogin()"
         class="text-[11px] tracking-[0.15em] uppercase text-gray-500 hover:text-[#0D0D0D] transition-colors leading-none border-none bg-transparent cursor-pointer">
        Clientes
      </button>
```

- [ ] **Step 2: Agregar botón en menú mobile**

Después del botón "Cotizar" mobile (línea ~320), antes del `</div>` del mobile-menu, agregar:

```html
    <button onclick="toggleMobileMenu(); openClientesLogin();"
       class="text-[13px] tracking-[0.2em] uppercase text-gray-600 border-none bg-transparent cursor-pointer">
      Clientes
    </button>
```

- [ ] **Step 3: Verificar manualmente**

Abrir index.html en el browser y confirmar que el botón "Clientes" aparece en la nav desktop y en el menú mobile.

---

### Task 3: Modal de login

**Files:**
- Modify: `index.html` (agregar antes de `</body>`)

- [ ] **Step 1: Agregar HTML del modal de login**

Antes del cierre `</body>`, agregar:

```html
<!-- ─── MODAL LOGIN CLIENTES ─────────────────────────────────── -->
<div id="modal-clientes-login" style="display:none; position:fixed; inset:0; z-index:80; background:rgba(0,0,0,0.7); align-items:center; justify-content:center;">
  <div style="background:#fff; width:100%; max-width:420px; margin:16px; padding:40px 32px 36px; font-family:'DM Sans',sans-serif;">
    <div style="display:flex; justify-content:space-between; align-items:flex-start; margin-bottom:28px;">
      <div>
        <p style="font-size:10px; letter-spacing:0.22em; text-transform:uppercase; color:#888; margin:0 0 6px;">LOG arquitectura</p>
        <h2 style="font-size:22px; font-weight:300; color:#0D0D0D; margin:0;">Área de Clientes</h2>
      </div>
      <button onclick="closeClientesLogin()" style="background:none; border:none; cursor:pointer; font-size:22px; color:#888; line-height:1; padding:0;">&times;</button>
    </div>

    <div style="margin-bottom:18px;">
      <label style="display:block; font-size:10px; letter-spacing:0.18em; text-transform:uppercase; color:#888; margin-bottom:6px;">Email</label>
      <input id="cl-email" type="email" placeholder="tu@email.com"
        style="width:100%; border:none; border-bottom:1px solid #ddd; padding:8px 0; font-size:14px; font-family:'DM Sans',sans-serif; outline:none; box-sizing:border-box;">
    </div>
    <div style="margin-bottom:28px;">
      <label style="display:block; font-size:10px; letter-spacing:0.18em; text-transform:uppercase; color:#888; margin-bottom:6px;">DNI</label>
      <input id="cl-dni" type="text" placeholder="Sin puntos ni espacios"
        style="width:100%; border:none; border-bottom:1px solid #ddd; padding:8px 0; font-size:14px; font-family:'DM Sans',sans-serif; outline:none; box-sizing:border-box;">
    </div>

    <p id="cl-error" style="display:none; color:#c0392b; font-size:12px; margin-bottom:16px;"></p>

    <button id="cl-btn" onclick="clientesLogin()"
      style="width:100%; background:#0D0D0D; color:#fff; border:none; padding:14px; font-size:11px; letter-spacing:0.18em; text-transform:uppercase; cursor:pointer; font-family:'DM Sans',sans-serif;">
      Ingresar
    </button>
  </div>
</div>
```

- [ ] **Step 2: Verificar manualmente**

Abrir index.html → clic en "Clientes" → confirmar que aparece el modal con los campos Email y DNI.

---

### Task 4: Portal del cliente (vista post-login)

**Files:**
- Modify: `index.html` (agregar después del modal login)

- [ ] **Step 1: Agregar HTML del portal**

Después del modal de login, agregar:

```html
<!-- ─── PORTAL CLIENTES ──────────────────────────────────────── -->
<div id="portal-clientes" style="display:none; position:fixed; inset:0; z-index:80; background:#F7F5F2; overflow-y:auto; font-family:'DM Sans',sans-serif;">
  <!-- Header -->
  <div style="background:#0D0D0D; padding:20px 32px; display:flex; justify-content:space-between; align-items:center;">
    <div>
      <p style="font-size:10px; letter-spacing:0.22em; text-transform:uppercase; color:#888; margin:0 0 4px;">LOG arquitectura</p>
      <h2 id="portal-saludo" style="font-size:18px; font-weight:300; color:#fff; margin:0;">Hola, Cliente</h2>
    </div>
    <button onclick="clientesCerrarSesion()"
      style="background:none; border:1px solid #555; color:#aaa; padding:8px 18px; font-size:10px; letter-spacing:0.15em; text-transform:uppercase; cursor:pointer; font-family:'DM Sans',sans-serif;">
      Cerrar sesión
    </button>
  </div>

  <!-- Contenido -->
  <div style="max-width:900px; margin:0 auto; padding:40px 24px;">
    <p style="font-size:10px; letter-spacing:0.2em; text-transform:uppercase; color:#888; margin-bottom:24px;">Tus proyectos</p>
    <div id="portal-archivos" style="display:grid; grid-template-columns: repeat(auto-fill, minmax(240px, 1fr)); gap:16px;">
      <!-- Se llena dinámicamente -->
    </div>
    <p id="portal-vacio" style="display:none; color:#888; font-size:14px;">No hay archivos disponibles por el momento.</p>
    <p id="portal-cargando" style="color:#888; font-size:13px;">Cargando archivos...</p>
  </div>
</div>

<!-- ─── VISOR PDF ─────────────────────────────────────────────── -->
<div id="modal-pdf" style="display:none; position:fixed; inset:0; z-index:90; background:rgba(0,0,0,0.9); flex-direction:column;">
  <div style="display:flex; justify-content:space-between; align-items:center; padding:14px 24px; background:#0D0D0D;">
    <p id="pdf-titulo" style="color:#fff; font-size:13px; margin:0; font-family:'DM Sans',sans-serif;"></p>
    <button onclick="cerrarPdf()" style="background:none; border:none; color:#fff; font-size:24px; cursor:pointer; font-family:'DM Sans',sans-serif;">&times;</button>
  </div>
  <iframe id="pdf-frame" src="" style="flex:1; border:none; width:100%;"></iframe>
</div>
```

- [ ] **Step 2: Verificar HTML**

Abrir index.html → no debe haber errores de consola relacionados con los nuevos elementos.

---

### Task 5: JavaScript — login, portal y visor

**Files:**
- Modify: `index.html` (bloque `<script>` al final, antes de `</body>`)

- [ ] **Step 1: Agregar funciones JavaScript**

Dentro del bloque `<script>` existente (antes del cierre `</script>` final), agregar:

```javascript
// ── PORTAL CLIENTES ──────────────────────────────────────────────

function openClientesLogin() {
  // Si ya hay sesión activa, ir directo al portal
  var sesion = sessionStorage.getItem('cl_sesion');
  if (sesion) {
    var datos = JSON.parse(sesion);
    mostrarPortal(datos.nombre, datos.folderId);
    return;
  }
  var m = document.getElementById('modal-clientes-login');
  m.style.display = 'flex';
  document.getElementById('cl-email').value = '';
  document.getElementById('cl-dni').value   = '';
  document.getElementById('cl-error').style.display = 'none';
  setTimeout(function(){ document.getElementById('cl-email').focus(); }, 100);
}

function closeClientesLogin() {
  document.getElementById('modal-clientes-login').style.display = 'none';
}

function clientesLogin() {
  var email = document.getElementById('cl-email').value.trim();
  var dni   = document.getElementById('cl-dni').value.trim();
  var err   = document.getElementById('cl-error');
  var btn   = document.getElementById('cl-btn');

  if (!email || !dni) {
    err.textContent = 'Completá todos los campos.';
    err.style.display = 'block';
    return;
  }

  btn.textContent = 'Verificando...';
  btn.disabled = true;
  err.style.display = 'none';

  var params = new URLSearchParams({ action: 'login', email: email, dni: dni });
  fetch(GSHEET_URL, { method: 'POST', mode: 'no-cors', body: params });

  // Con no-cors no podemos leer la respuesta. Usamos GET con parámetros.
  fetch(GSHEET_URL + '?action=login&email=' + encodeURIComponent(email) + '&dni=' + encodeURIComponent(dni))
    .then(function(r){ return r.json(); })
    .then(function(data) {
      btn.textContent = 'Ingresar';
      btn.disabled = false;
      if (data.ok) {
        sessionStorage.setItem('cl_sesion', JSON.stringify({ nombre: data.nombre, folderId: data.folderId }));
        closeClientesLogin();
        mostrarPortal(data.nombre, data.folderId);
      } else {
        err.textContent = data.error || 'Email o DNI incorrecto.';
        err.style.display = 'block';
      }
    })
    .catch(function() {
      btn.textContent = 'Ingresar';
      btn.disabled = false;
      err.textContent = 'Error de conexión. Intentá de nuevo.';
      err.style.display = 'block';
    });
}

function mostrarPortal(nombre, folderId) {
  document.getElementById('portal-saludo').textContent = 'Hola, ' + nombre;
  document.getElementById('portal-clientes').style.display = 'block';
  document.getElementById('portal-archivos').innerHTML = '';
  document.getElementById('portal-vacio').style.display = 'none';
  document.getElementById('portal-cargando').style.display = 'block';

  fetch(GSHEET_URL + '?action=archivos&folderId=' + encodeURIComponent(folderId))
    .then(function(r){ return r.json(); })
    .then(function(pdfs) {
      document.getElementById('portal-cargando').style.display = 'none';
      if (!Array.isArray(pdfs) || pdfs.length === 0) {
        document.getElementById('portal-vacio').style.display = 'block';
        return;
      }
      var grid = document.getElementById('portal-archivos');
      pdfs.forEach(function(pdf) {
        var card = document.createElement('div');
        card.style.cssText = 'background:#fff; padding:20px; display:flex; flex-direction:column; gap:14px;';
        card.innerHTML =
          '<div style="display:flex; align-items:center; gap:10px;">' +
            '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#0D0D0D" stroke-width="1.5"><path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z"/><polyline points="14 2 14 8 20 8"/></svg>' +
            '<span style="font-size:13px; color:#0D0D0D; font-weight:400; flex:1; line-height:1.3;">' + pdf.nombre + '</span>' +
          '</div>' +
          '<div style="display:flex; gap:8px;">' +
            '<button onclick="verPdf(\'' + pdf.id + '\',\'' + pdf.nombre.replace(/'/g,"\'") + '\')" ' +
              'style="flex:1; background:#0D0D0D; color:#fff; border:none; padding:10px; font-size:10px; letter-spacing:0.15em; text-transform:uppercase; cursor:pointer; font-family:\'DM Sans\',sans-serif;">Ver</button>' +
            '<a href="https://drive.google.com/uc?export=download&id=' + pdf.id + '" target="_blank" ' +
              'style="flex:1; background:#fff; color:#0D0D0D; border:1px solid #0D0D0D; padding:10px; font-size:10px; letter-spacing:0.15em; text-transform:uppercase; cursor:pointer; font-family:\'DM Sans\',sans-serif; text-decoration:none; text-align:center; display:flex; align-items:center; justify-content:center;">Descargar</a>' +
          '</div>';
        grid.appendChild(card);
      });
    })
    .catch(function() {
      document.getElementById('portal-cargando').style.display = 'none';
      document.getElementById('portal-vacio').style.display = 'block';
      document.getElementById('portal-vacio').textContent = 'Error al cargar archivos.';
    });
}

function clientesCerrarSesion() {
  sessionStorage.removeItem('cl_sesion');
  document.getElementById('portal-clientes').style.display = 'none';
}

function verPdf(id, nombre) {
  document.getElementById('pdf-titulo').textContent = nombre;
  document.getElementById('pdf-frame').src = 'https://drive.google.com/file/d/' + id + '/preview';
  document.getElementById('modal-pdf').style.display = 'flex';
}

function cerrarPdf() {
  document.getElementById('pdf-frame').src = '';
  document.getElementById('modal-pdf').style.display = 'none';
}
```

**Nota importante:** El login usa `doGet` con parámetros (en lugar de `doPost`) porque el Apps Script necesitar una nueva acción GET para login. Agregar en `doGet` de log-arquitectura.gs:

```javascript
if (e && e.parameter && e.parameter.action === 'login') {
  // (mismo código que el doPost login de Task 1, pero usando e.parameter en lugar de p)
  try {
    var emailParam = (e.parameter.email || '').toLowerCase().trim();
    var dniParam   = (e.parameter.dni   || '').trim();
    var ss         = SpreadsheetApp.getActiveSpreadsheet();
    var hoja       = ss.getSheetByName('Clientes');
    if (!hoja) throw new Error('Hoja Clientes no encontrada');
    var datos = hoja.getDataRange().getValues();
    for (var i = 1; i < datos.length; i++) {
      var nombre    = (datos[i][0] || '').toString().trim();
      var email     = (datos[i][2] || '').toString().trim().toLowerCase();
      var dni       = (datos[i][3] || '').toString().trim();
      var linkDrive = (datos[i][4] || '').toString().trim();
      var estado    = (datos[i][5] || '').toString().trim().toLowerCase();
      if (email !== emailParam) continue;
      if (estado === 'inactivo') return ContentService.createTextOutput(JSON.stringify({ ok: false, error: 'Cuenta inactiva' })).setMimeType(ContentService.MimeType.JSON);
      if (dni !== dniParam)     return ContentService.createTextOutput(JSON.stringify({ ok: false, error: 'Email o DNI incorrecto.' })).setMimeType(ContentService.MimeType.JSON);
      var folderMatch = linkDrive.match(/[-\w]{25,}/);
      var folderId    = folderMatch ? folderMatch[0] : '';
      return ContentService.createTextOutput(JSON.stringify({ ok: true, nombre: nombre, folderId: folderId })).setMimeType(ContentService.MimeType.JSON);
    }
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: 'Email o DNI incorrecto.' })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: err.toString() })).setMimeType(ContentService.MimeType.JSON);
  }
}
```

- [ ] **Step 2: Verificar flujo completo manualmente**

1. Clic en "Clientes" → aparece modal de login
2. Ingresar email + DNI de un cliente de prueba en el Sheet → clic "Ingresar"
3. Verificar que aparece el portal con "Hola, [Nombre]"
4. Verificar que aparecen las tarjetas de PDFs
5. Clic "Ver" → visor PDF se abre con el PDF embebido
6. Clic "Descargar" → descarga el PDF
7. Clic "Cerrar sesión" → vuelve al sitio normal
8. Clic "Clientes" de nuevo → debe volver al portal directamente (sesión activa)

- [ ] **Step 3: Commit**

```bash
git add index.html
git commit -m "feat: agregar portal de clientes con login y visor PDF"
```

---

### Task 6: Deploy

- [ ] **Step 1: Publicar nueva versión Apps Script**

En Google Apps Script: Implementar → Administrar implementaciones → editar → Nueva versión → Implementar. Copiar URL.

- [ ] **Step 2: Actualizar GSHEET_URL en index.html**

Reemplazar la URL en la línea `var GSHEET_URL = '...'` con la nueva URL del Apps Script.

- [ ] **Step 3: Push a GitHub**

```bash
git add index.html log-arquitectura.gs
git commit -m "feat: actualizar GSHEET_URL para portal de clientes"
git push https://TOKEN@github.com/renepariguana/log-arquitectura-web.git main
```

GitHub Actions despliega automáticamente a Hostinger.

- [ ] **Step 4: Verificar en logarquitectura.com.ar**

Probar el flujo completo en el sitio en vivo.
