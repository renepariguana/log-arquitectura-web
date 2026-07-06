# Identificación de participantes por rol — Diseño Técnico

*Fecha: 2026-07-06*
*Proyecto: LOG arquitectura — portal / `pagina-web`*

## Resumen

En el modal de "Ingresa a tu cuenta", **antes** de loguearse con Google, la persona
elige **qué es**: Cliente, Proyectista, Mano de obra o Proveedor. Luego se loguea con
Google y entra al portal. Según el rol elegido, su registro se guarda en **una de 4
pestañas** del Google Sheets del portal (`Clientes`, `Proyectistas`, `Mano de obra`,
`Proveedores`), vía Apps Script. La identidad es el **email**. Esta base se va poblando
sola con cada login y es el cimiento de los pasos siguientes (portal adaptado por rol,
subida de archivos, interacción, cotización).

Este spec cubre **solo el Paso 1** (versión mínima): selector de rol en el modal +
registro del participante en la pestaña de su rol + entrada al portal.

### Visión (norte)

LOG arquitectura evoluciona de "el portal de un estudio" a una **plataforma abierta**
donde participan múltiples estudios (proyectistas), mano de obra, proveedores y clientes.
Todos entran por el mismo formulario, sin cuentas privilegiadas.

### Tipos de participante (→ pestaña destino)

| Rol | Quién es | Pestaña |
|---|---|---|
| 🧑 **Cliente** | Encarga la obra; entra a ver lo que se construye en su proyecto. | `Clientes` |
| 🏗️ **Proyectista** | Hace los planos: arquitecto, ingeniero, MMO, estudio. | `Proyectistas` |
| 👷 **Mano de obra** | Construye: albañil, plomero, electricista, etc. | `Mano de obra` |
| 🧱 **Proveedor** | Provee los materiales. | `Proveedores` |

## Alcance

### Dentro de alcance (Paso 1)
- Rediseño del modal de login: 4 tarjetas de rol antes del botón de Google; el botón se
  habilita al elegir uno.
- El rol elegido se manda al backend sumándolo al `action=loginGoogle` existente.
- **Fix del rename:** la pestaña de clientes debe llamarse `Clientes` (lo que el código
  espera). Hoy está renombrada a "Usuarios" — se vuelve a llamar `Clientes`.
- `loginGoogle` enruta por rol: clientes → flujo actual intacto; los otros 3 roles →
  registro (upsert por email) en su pestaña, creándola si no existe.
- Entrar al portal (comportamiento actual; el portal aún no cambia por rol).

### Fuera de alcance (pasos futuros)
- **Portal adaptado por rol** (qué ve cada rol) — paso inmediato siguiente.
- **Subida de archivos desde el portal** (el usuario sube los suyos).
- Sub-rubro, ubicación, WhatsApp; carpetas de Drive para roles no-cliente.
- Sidebar de interacción, asignación a proyectos, cotización, Manos a la Obra.

---

## Flujo

```
Clic en "Ingresa a tu cuenta"
        │
        ▼
  Modal: ¿Qué sos?  [Cliente] [Proyectista] [Mano de obra] [Proveedor]
        │  (elegir rol habilita el botón de Google)
        ▼
  Login con Google (GSI — onGoogleLoginPortal, flujo actual)
        │
        ▼
  action=loginGoogle&...&rol=<rol>   (se agrega el rol al fetch actual)
        │
        ├── rol = cliente  → flujo actual: pestaña Clientes, carpeta, suscripción
        └── rol ≠ cliente  → upsert por email en Proyectistas / Mano de obra / Proveedores
        │
        ▼
  mostrarPortal(...)   → entra al portal
```

- El rol se guarda en una variable JS (`participanteRolElegido`) al tocar la tarjeta y
  se agrega a la URL de `loginGoogle`.
- **Upsert por email** en la pestaña destino: si el email ya está, actualiza
  `Actualizado`; si no, agrega la fila con `Creado`.

---

## El modal (rediseño de `#modal-clientes-login`, vista `#cl-vista-login`)

Mockup (mobile-first, estilo actual: DM Sans, blanco/negro):

```
┌───────────────────────────────────────┐
│  LOG ARQUITECTURA                   ✕  │
│                                        │
│  Ingresá a la plataforma según lo      │
│  que hacés:                            │
│                                        │
│   ┌──────────┐  ┌──────────────┐       │
│   │ 🧑 Cliente│  │ 🏗️ Proyectista│      │
│   └──────────┘  └──────────────┘       │
│   ┌──────────────┐  ┌───────────┐      │
│   │ 👷 Mano de obra│  │ 🧱 Proveedor│    │
│   └──────────────┘  └───────────┘      │
│                                        │
│   (elegí una opción para continuar)    │
│                                        │
│      [ Inicia sesión con Google  G ]   │
└───────────────────────────────────────┘
```

- Las 4 tarjetas son seleccionables (una a la vez); la elegida queda resaltada.
- El botón de Google arranca **deshabilitado/oculto**; se habilita al elegir un rol.
- Se reemplaza el copy actual centrado en clientes por uno que incluye a todos los roles.
- El resto de las vistas del modal (suscripción, datos, pago) no se tocan.

**Validación:** no se puede iniciar el login de Google sin un rol elegido.

---

## Datos que se guardan — 4 pestañas por rol

- **`Clientes`** = la pestaña que ya existe (hoy renombrada a "Usuarios"; **se vuelve a
  llamar `Clientes`**). Mantiene su estructura y lógica actuales (columnas
  `Nº · Nombre · Email · Suscripcion · LINK DRIVE`, carpeta de Drive, suscripción,
  proyectos). **Sin cambios de comportamiento para clientes.**
- **`Proyectistas` / `Mano de obra` / `Proveedores`** = pestañas nuevas, creadas con
  encabezados la primera vez que se registra alguien de ese rol. Estructura mínima:

| Columna | Campo | Ejemplo |
|---|---|---|
| A | `Nombre` (de Google) | `Juan Pérez` |
| B | `Email` (clave única) | `juan@gmail.com` |
| C | `Foto` (URL de Google) | `https://lh3...` |
| D | `Creado` | fecha/hora (primera vez) |
| E | `Actualizado` | fecha/hora (último login) |

La clave de identidad es el **email**. (Sub-rubro, ubicación, WhatsApp y carpeta de
Drive de estos roles se agregan en pasos futuros.)

---

## Cambios en Apps Script (`apps-script.gs`)

Requiere **re-publicar una nueva versión** del Web App tras editar.

1. **Fix del rename:** renombrar la pestaña `Usuarios` → `Clientes` en el Sheet (lo que
   `getSheetByName('Clientes')` espera). No requiere cambiar código para esta parte.
   > Sin este fix, el próximo login crea una pestaña `Clientes` vacía e ignora los datos
   > reales — hay que hacerlo sí o sí.
2. **`action=loginGoogle` acepta `&rol=`** (`e.parameter.rol`), lo valida contra la lista
   `['cliente','proyectista','manoDeObra','proveedor']` y enruta:
   - `rol === 'cliente'` o vacío (retrocompat) → **flujo actual intacto** (pestaña
     `Clientes`, carpeta, suscripción). La respuesta no cambia.
   - `rol` de los otros 3 → función nueva `registrarParticipante(rol, email, nombre, foto)`:
     `getSheetByName(<pestaña>)` (o `insertSheet` con encabezados), busca la fila por
     email; si existe actualiza `Actualizado`, si no hace `appendRow([nombre, email,
     foto, ahora, ahora])`. Devuelve `{ ok:true, esCliente:false, suscripcion:false,
     nombre:nombre, rol:rol, folderId:'' }`.
3. Mapa rol → pestaña: `{ proyectista:'Proyectistas', manoDeObra:'Mano de obra',
   proveedor:'Proveedores' }`.

---

## Integración con `index.html` (verificado en el código)

- **Modal:** la vista de login vive en `#modal-clientes-login` / `#cl-vista-login`
  (funciones `clMostrarLogin()`, `renderGoogleBtn()` línea ~1578, botón en
  `#cl-google-btn`). Ahí se agregan las 4 tarjetas de rol y la lógica de habilitación,
  guardando la selección en `participanteRolElegido`.
- **Login:** `onGoogleLoginPortal(response)` (línea ~1629) decodifica el JWT →
  `fetch(GSHEET_URL + '?action=loginGoogle...')` → `sessionStorage.cl_sesion` →
  `mostrarPortal(...)`. `GSHEET_URL` está en la línea ~1208.
- **Cambio:** agregar `+ '&rol=' + encodeURIComponent(participanteRolElegido)` a la URL
  del `fetch` de `loginGoogle`. **Sin endpoint ni request aparte.**
- `mostrarPortal(nombre, folderId, suscripcion, esCliente, nroSuscriptor, picture)` ya
  maneja no-clientes: si `!esCliente || !folderId` muestra `#portal-vacio`. Los roles
  nuevos caen ahí por ahora (portal genérico). El portal por rol es el paso siguiente.
- Reutilizar `esc()` para datos inyectados en el DOM.
- No tocar el modo embebido (`?embedded=1`) ni `presupuesto.html`.

---

## Manejo de errores

- **Sin rol elegido:** el botón de Google no dispara el login; ayuda "elegí una opción
  para continuar".
- **`rol` inválido en el backend:** `{ ok:false, error }`; el front muestra el error
  existente y no entra.
- **`action=loginGoogle` falla:** comportamiento actual sin cambios.

---

## Testing

`index.html` + `apps-script.gs` sin build ni test runner → validación **manual**:

1. Modal → 4 tarjetas visibles; botón de Google deshabilitado hasta elegir un rol.
2. Elegir **Mano de obra** → login → se crea la pestaña `Mano de obra` (si no existía) y
   una fila con nombre, email, foto, `Creado`, `Actualizado`.
3. Elegir **Proveedor** con otro email → fila en `Proveedores`.
4. Elegir **Proyectista** → fila en `Proyectistas`.
5. Elegir **Cliente** → va a `Clientes`, con carpeta/suscripción/proyectos como hoy.
6. Re-login del mismo email+rol → la fila se **actualiza** (no se duplica); `Actualizado`
   cambia, `Creado` se mantiene.
7. Un cliente existente sigue viendo sus proyectos como antes (regresión del flujo actual).

---

## Criterios de éxito

- El modal ofrece los 4 roles y exige elegir uno antes del login.
- Cada login registra al participante en la pestaña correcta según su rol, sin duplicar
  por email.
- El flujo de clientes (carpeta, suscripción, proyectos) sigue intacto.
- Queda arreglado el bug del rename `Usuarios`/`Clientes`.
- La base (4 pestañas) queda lista para adaptar el portal por rol en el próximo paso.
