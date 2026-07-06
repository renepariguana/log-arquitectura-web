# Identificación de participantes por rol — Plan de Implementación

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Que en el modal de login el usuario elija su rol (Cliente / Proyectista / Mano de obra / Proveedor) antes de entrar con Google, y que según el rol quede registrado en su pestaña del Google Sheets.

**Architecture:** El portal (`index.html`) ya se loguea vía Google GSI → `action=loginGoogle` (Apps Script) → `mostrarPortal(...)`. Se agrega un selector de rol al modal, se pasa el rol al `loginGoogle`, y el backend rutea: clientes al flujo actual; los otros 3 roles a `registrarParticipante()` que hace upsert por email en su pestaña.

**Tech Stack:** HTML + JS vanilla (sin build), Google Apps Script, Google Sheets. Sin test runner → verificación manual en navegador.

## Global Constraints

- Reutilizar `esc()` para cualquier dato de usuario/servidor inyectado en `innerHTML`.
- No tocar el modo embebido (`?embedded=1`) ni `presupuesto.html`.
- El flujo de **clientes** (carpeta Drive, suscripción, proyectos) debe quedar **intacto**.
- Editar `apps-script.gs` exige **re-publicar una NUEVA implementación** del Web App (Implementar → Administrar implementaciones → Editar → Nueva versión). La `GSHEET_URL` no cambia si se actualiza la implementación existente.
- Roles válidos (valor interno): `cliente`, `proyectista`, `manoDeObra`, `proveedor`.
- Mapa rol → pestaña: `proyectista → Proyectistas`, `manoDeObra → Mano de obra`, `proveedor → Proveedores`. Cliente → pestaña `Clientes` (flujo actual).

---

### Task 1: Restaurar el nombre de la pestaña `Clientes` (fix del bug del rename)

**Files:**
- Ninguno (acción manual en Google Sheets). Sin commit.

**Contexto:** El Apps Script hace `getSheetByName('Clientes')`. La pestaña fue renombrada a "Usuarios", por lo que el próximo login crearía una `Clientes` vacía e ignoraría los datos reales.

- [ ] **Step 1: Renombrar la pestaña**

En el Google Sheets del portal (LOG arquitectura), doble clic en la pestaña **"Usuarios"** → renombrar a **`Clientes`** (respetando mayúscula inicial, sin espacios).

- [ ] **Step 2: Verificar el login de cliente**

Abrir el portal → "Ingresa a tu cuenta" → login con Google del cliente existente (`renepariguana@gmail.com`).
Esperado: entra y muestra sus proyectos; en la pestaña `Clientes` NO se duplica su fila.

---

### Task 2: Backend — routing por rol + `registrarParticipante()`

**Files:**
- Modify: `apps-script.gs` (bloque `action === 'loginGoogle'` dentro de `doGet`, y nueva función a nivel de archivo)

**Interfaces:**
- Produce: `action=loginGoogle&email=&nombre=&rol=&foto=` → si `rol ∈ {proyectista,manoDeObra,proveedor}` responde `{ ok:true, nombre, esCliente:false, suscripcion:false, folderId:'', rol:<pestaña> }`; si `rol` es `cliente`/vacío, responde igual que hoy.

- [ ] **Step 1: Agregar el ruteo por rol al inicio de `loginGoogle`**

En `apps-script.gs`, dentro de `if (... action === 'loginGoogle')`, justo después de `var ss = SpreadsheetApp.getActiveSpreadsheet();`, insertar:

```javascript
      // ── Ruteo por rol: los 3 sectores de servicio van a su propia pestaña ──
      var rol  = (e.parameter.rol  || '').toString().trim();
      var foto = (e.parameter.foto || '').toString().trim();
      var PESTANA_POR_ROL = { proyectista: 'Proyectistas', manoDeObra: 'Mano de obra', proveedor: 'Proveedores' };
      if (PESTANA_POR_ROL[rol]) {
        return registrarParticipante(ss, PESTANA_POR_ROL[rol], email, nombre, foto);
      }
      // rol === 'cliente' o vacío → continúa el flujo de Clientes de siempre
```

- [ ] **Step 2: Agregar la función `registrarParticipante`**

En `apps-script.gs`, a nivel de archivo (por ejemplo, arriba de `function autoCrearCarpeta(e)`), agregar:

```javascript
// ── Registro de participantes no-cliente (proyectistas, mano de obra, proveedores) ──
function registrarParticipante(ss, nombrePestana, email, nombre, foto) {
  try {
    var hoja = ss.getSheetByName(nombrePestana);
    if (!hoja) {
      hoja = ss.insertSheet(nombrePestana);
      hoja.appendRow(['Nombre', 'Email', 'Foto', 'Creado', 'Actualizado']);
      hoja.getRange(1, 1, 1, 5).setFontWeight('bold');
    }
    var ahora = new Date().toLocaleString('es-AR');
    var datos = hoja.getDataRange().getValues();
    var fila  = -1;
    for (var i = 1; i < datos.length; i++) {
      if ((datos[i][1] || '').toString().trim().toLowerCase() === email) { fila = i + 1; break; }
    }
    if (fila > 0) {
      if (nombre) hoja.getRange(fila, 1).setValue(nombre);
      if (foto)   hoja.getRange(fila, 3).setValue(foto);
      hoja.getRange(fila, 5).setValue(ahora);           // Actualizado
    } else {
      hoja.appendRow([nombre, email, foto, ahora, ahora]);
    }
    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, nombre: nombre, esCliente: false, suscripcion: false, folderId: '', rol: nombrePestana }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
```

- [ ] **Step 3: Agregar la función de setup que crea las 4 pestañas**

En `apps-script.gs`, a nivel de archivo (junto a `registrarParticipante`), agregar:

```javascript
// ── Setup: crea las pestañas de participantes con encabezados (ejecutar UNA vez) ──
function crearPestanasParticipantes() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ['Proyectistas', 'Mano de obra', 'Proveedores'].forEach(function(nombre) {
    if (!ss.getSheetByName(nombre)) {
      var hoja = ss.insertSheet(nombre);
      hoja.appendRow(['Nombre', 'Email', 'Foto', 'Creado', 'Actualizado']);
      hoja.getRange(1, 1, 1, 5).setFontWeight('bold');
    }
  });
  Logger.log('Pestañas de participantes verificadas/creadas.');
}
```

> `Clientes` ya existe (tras el rename de la Task 1); esta función crea las otras 3.

- [ ] **Step 4: Ejecutar `crearPestanasParticipantes` una vez**

En el editor de Apps Script: seleccionar la función `crearPestanasParticipantes` en el desplegable → ▶ Ejecutar (autorizar permisos si lo pide).
Esperado: en el Sheets aparecen las pestañas **`Proyectistas`**, **`Mano de obra`** y **`Proveedores`**, cada una con encabezados `Nombre · Email · Foto · Creado · Actualizado`. Junto con `Clientes`, quedan las **4 categorías**.

- [ ] **Step 5: Publicar nueva versión del Web App**

Apps Script → Implementar → Administrar implementaciones → Editar (lápiz) → Versión: Nueva → Implementar. (Mantiene la misma `GSHEET_URL`.)

- [ ] **Step 6: Verificar el endpoint con una URL directa**

En el navegador, reemplazando `<GSHEET_URL>`:
`<GSHEET_URL>?action=loginGoogle&email=test.mano@ejemplo.com&nombre=Test%20Mano&rol=manoDeObra&foto=`
Esperado: JSON `{ "ok": true, ... "esCliente": false, "rol": "Mano de obra" }`, y en la pestaña **"Mano de obra"** aparece una fila `Test Mano | test.mano@ejemplo.com | | <fecha> | <fecha>`.

- [ ] **Step 7: Verificar que cliente sigue igual**

`<GSHEET_URL>?action=loginGoogle&email=renepariguana@gmail.com&nombre=Rene&rol=cliente`
Esperado: responde con `esCliente:true` y su `folderId` (flujo actual intacto).

- [ ] **Step 8: Commit**

```bash
git add apps-script.gs
git commit -m "feat(portal): registrar participantes por rol en loginGoogle"
```

---

### Task 3: Frontend — selector de rol en el modal

**Files:**
- Modify: `index.html` (`#cl-vista-login`, líneas ~1326-1330, y bloque JS cerca de `renderGoogleBtn()` ~1579)

**Interfaces:**
- Produce: variable global `participanteRolElegido` (string) y función `clElegirRol(rol, el)`; el botón de Google (`#cl-google-wrap`) permanece oculto hasta elegir un rol.

- [ ] **Step 1: Reemplazar el contenido de `#cl-vista-login`**

Reemplazar el bloque actual:

```html
      <div id="cl-vista-login">
        <p style="font-size:12px; color:#888; text-align:center; margin:0 0 24px;">Iniciá sesión para ver tus proyectos o acceder al Presupuestador.</p>
        <div id="cl-google-btn" style="display:flex; justify-content:center; margin-bottom:16px;"></div>
        <p id="cl-google-error" style="display:none; color:#c0392b; font-size:12px; text-align:center; margin-top:8px;"></p>
      </div>
```

por:

```html
      <div id="cl-vista-login">
        <p style="font-size:12px; color:#888; text-align:center; margin:0 0 18px;">Ingresá a la plataforma según lo que hacés:</p>

        <div id="cl-roles" style="display:grid; grid-template-columns:1fr 1fr; gap:10px; margin-bottom:16px;">
          <button type="button" class="cl-rol-card" data-rol="cliente"     onclick="clElegirRol('cliente', this)"     style="border:1px solid #e0e0e0; background:#fff; color:#0D0D0D; font-family:'DM Sans',sans-serif; font-size:13px; padding:14px 12px; cursor:pointer;">🧑 Cliente</button>
          <button type="button" class="cl-rol-card" data-rol="proyectista" onclick="clElegirRol('proyectista', this)" style="border:1px solid #e0e0e0; background:#fff; color:#0D0D0D; font-family:'DM Sans',sans-serif; font-size:13px; padding:14px 12px; cursor:pointer;">🏗️ Proyectista</button>
          <button type="button" class="cl-rol-card" data-rol="manoDeObra"  onclick="clElegirRol('manoDeObra', this)"  style="border:1px solid #e0e0e0; background:#fff; color:#0D0D0D; font-family:'DM Sans',sans-serif; font-size:13px; padding:14px 12px; cursor:pointer;">👷 Mano de obra</button>
          <button type="button" class="cl-rol-card" data-rol="proveedor"   onclick="clElegirRol('proveedor', this)"   style="border:1px solid #e0e0e0; background:#fff; color:#0D0D0D; font-family:'DM Sans',sans-serif; font-size:13px; padding:14px 12px; cursor:pointer;">🧱 Proveedor</button>
        </div>

        <p id="cl-rol-ayuda" style="font-size:11px; color:#888; text-align:center; margin:0 0 16px;">Elegí una opción para continuar</p>

        <div id="cl-google-wrap" style="display:none;">
          <div id="cl-google-btn" style="display:flex; justify-content:center; margin-bottom:16px;"></div>
        </div>
        <p id="cl-google-error" style="display:none; color:#c0392b; font-size:12px; text-align:center; margin-top:8px;"></p>
      </div>
```

- [ ] **Step 2: Agregar la lógica `clElegirRol` + variable global**

Justo antes de `function renderGoogleBtn() {` (línea ~1579), insertar:

```javascript
    var participanteRolElegido = '';

    function clElegirRol(rol, el) {
      participanteRolElegido = rol;
      var cards = document.querySelectorAll('#cl-roles .cl-rol-card');
      for (var i = 0; i < cards.length; i++) {
        var activa = cards[i] === el;
        cards[i].style.borderColor = activa ? '#0D0D0D' : '#e0e0e0';
        cards[i].style.background  = activa ? '#0D0D0D' : '#fff';
        cards[i].style.color       = activa ? '#fff'    : '#0D0D0D';
      }
      var ayuda = document.getElementById('cl-rol-ayuda');
      if (ayuda) ayuda.style.display = 'none';
      var wrap = document.getElementById('cl-google-wrap');
      if (wrap) wrap.style.display = 'block';
      renderGoogleBtn();
    }

    function clResetRol() {
      participanteRolElegido = '';
      var cards = document.querySelectorAll('#cl-roles .cl-rol-card');
      for (var i = 0; i < cards.length; i++) {
        cards[i].style.borderColor = '#e0e0e0';
        cards[i].style.background  = '#fff';
        cards[i].style.color       = '#0D0D0D';
      }
      var ayuda = document.getElementById('cl-rol-ayuda');
      if (ayuda) ayuda.style.display = 'block';
      var wrap = document.getElementById('cl-google-wrap');
      if (wrap) wrap.style.display = 'none';
    }
```

- [ ] **Step 3: Resetear el selector al mostrar la vista de login**

En `function clMostrarLogin()`, agregar `clResetRol();` como primera línea del cuerpo (para que al reabrir el modal el rol arranque sin elegir y el botón de Google oculto).

- [ ] **Step 4: Verificar el modal**

Abrir el portal → "Ingresa a tu cuenta".
Esperado: se ven 4 tarjetas; el botón de Google NO aparece; el texto dice "Elegí una opción para continuar". Al tocar una tarjeta: queda resaltada en negro, desaparece la ayuda y aparece el botón de Google.

- [ ] **Step 5: Commit**

```bash
git add index.html
git commit -m "feat(portal): selector de rol en el modal de login"
```

---

### Task 4: Frontend — pasar rol + foto a `loginGoogle` y bloquear sin rol

**Files:**
- Modify: `index.html` (`function onGoogleLoginPortal(response)`, ~1629-1664)

**Interfaces:**
- Consume: `participanteRolElegido` (Task 3), endpoint con `&rol=&foto=` (Task 2).

- [ ] **Step 1: Bloquear el login si no hay rol elegido**

Al inicio de `onGoogleLoginPortal(response)`, después de decodificar `payload/emailGoogle/nombreGoogle` (tras el `try/catch` que lee las credenciales), insertar:

```javascript
      if (!participanteRolElegido) {
        var errRol = document.getElementById('cl-google-error');
        if (errRol) { errRol.textContent = 'Elegí una opción (Cliente, Proyectista, Mano de obra o Proveedor) para continuar.'; errRol.style.display = 'block'; }
        return;
      }
```

- [ ] **Step 2: Agregar `rol` y `foto` a la URL del `fetch`**

Reemplazar la línea del fetch:

```javascript
      fetch(GSHEET_URL + '?action=loginGoogle&email=' + encodeURIComponent(emailGoogle) + '&nombre=' + encodeURIComponent(nombreGoogle))
```

por:

```javascript
      fetch(GSHEET_URL + '?action=loginGoogle&email=' + encodeURIComponent(emailGoogle) + '&nombre=' + encodeURIComponent(nombreGoogle) + '&rol=' + encodeURIComponent(participanteRolElegido) + '&foto=' + encodeURIComponent(payload.picture || ''))
```

- [ ] **Step 3: Verificar end-to-end (mano de obra)**

Portal → elegir **Mano de obra** → login con una cuenta Google de prueba (que no sea cliente).
Esperado: entra al portal (vista genérica / "portal-vacio"); en el Sheets, pestaña **Mano de obra**, aparece su fila con nombre, email y foto.

- [ ] **Step 4: Verificar end-to-end (cliente, regresión)**

Portal → elegir **Cliente** → login con `renepariguana@gmail.com`.
Esperado: entra y ve sus proyectos como siempre; no se duplica su fila en `Clientes`.

- [ ] **Step 5: Commit**

```bash
git add index.html
git commit -m "feat(portal): enviar rol elegido a loginGoogle y exigir selección"
```

---

## Notas de verificación final (regresión)

- Login de cliente existente → ve proyectos (sin cambios).
- Suscripción / presupuestador embebido → sin cambios.
- Reabrir el modal tras cerrarlo → el rol vuelve a estar sin elegir y el botón de Google oculto.
- Pestañas `Proyectistas` / `Mano de obra` / `Proveedores` se crean solas al primer registro de cada rol; el trigger `autoCrearCarpeta` (solo actúa sobre `Clientes`) no interfiere.
