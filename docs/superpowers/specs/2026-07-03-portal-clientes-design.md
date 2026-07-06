# Portal de Clientes — LOG arquitectura

## Resumen

Sección privada del sitio donde cada cliente accede con su email y DNI para ver y descargar los PDFs de sus proyectos, almacenados en Google Drive.

## Autenticación

- Botón "Clientes" en la barra de navegación
- Modal con campos: Email + DNI
- Apps Script verifica contra la hoja "Clientes" del Google Sheet
- Si coincide: devuelve nombre del cliente + ID de carpeta Drive
- Si no coincide: mensaje "Email o DNI incorrecto"
- Sesión guardada en `sessionStorage` (persiste mientras la pestaña esté abierta)

## Portal

- Encabezado: "Hola, [Nombre]"
- Listado de PDFs de la carpeta Drive del cliente (dinámico, se actualiza al subir archivos)
- Cada PDF muestra: nombre del archivo + botón Ver + botón Descargar
- Visor PDF embebido con iframe de Google Drive
- Botón "Cerrar sesión"

## Apps Script — nuevas acciones

- `doPost { action: "login", email, dni }` → busca en hoja "Clientes", devuelve `{ nombre, folderId }` o error
- `doGet?action=archivos&folderId=XXX` → lista PDFs de la carpeta Drive del cliente

## Google Sheets — hoja "Clientes"

Columnas: Nombre | Proyecto | Email | Contraseña (DNI) | Link Drive | Estado

## Stack

HTML + Tailwind CSS + Google Apps Script + Google Drive (igual que el resto del sitio)
