# Base de datos — LOG arquitectura

## Google Sheets

La base de datos de solicitudes se almacena en Google Sheets.

### Pasos para conectar:

1. Abrí tu Google Sheet en: https://sheets.google.com
2. Andá a **Extensiones > Apps Script**
3. Pegá el contenido del archivo `log-arquitectura.gs` (está en la carpeta raíz del proyecto)
4. Guardá (Ctrl+S)
5. Hacé clic en **Implementar > Nueva implementación**
   - Tipo: Aplicación web
   - Ejecutar como: Yo (tu cuenta)
   - Quién puede acceder: Cualquier persona
6. Autorizá los permisos
7. Copiá la URL generada
8. Pegala en `index.html` donde dice:
   ```
   var GSHEET_URL = 'PEGAR_AQUI_URL_DEL_APPS_SCRIPT';
   ```

### Columnas de la hoja "Solicitudes":

| Fecha | Servicio | Nombre | Email | Tipo de Proyecto | M² | Mensaje |
|-------|----------|--------|-------|------------------|----|---------|
