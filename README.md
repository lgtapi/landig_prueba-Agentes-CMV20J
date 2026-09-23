# Reporte de fallos – Agente CMV20J (Ambar 360 grados)

- `index.html` + `logo.png`: landing estática que se publica en Vercel.
- `apps-script/Code.gs`: backend en Google Apps Script (guarda en Google Sheets y las fotos en Drive).

## 1. Backend (Google Apps Script)
1. Crea una Google Sheet → Extensiones → Apps Script.
2. Pega `apps-script/Code.gs` (no se necesita el archivo Index en Apps Script).
3. Ejecuta `setup` una vez y acepta permisos.
4. Implementar → Nueva implementación → Aplicación web:
   - Ejecutar como: **Yo**
   - Quién tiene acceso: **Cualquier usuario** (sin cuenta de Google; obligatorio para que Vercel pueda enviar).
5. Copia la URL que termina en `/exec`.

## 2. Frontend (Vercel)
1. En `index.html`, reemplaza `PEGA_AQUI_LA_URL_DE_TU_APPS_SCRIPT` por la URL `/exec`.
2. Publica la carpeta (sin la subcarpeta `apps-script` si prefieres):
   - Con CLI: `npm i -g vercel`, luego `vercel` y `vercel --prod` dentro de la carpeta.
   - O súbela a un repositorio de GitHub e impórtala en vercel.com → Add New → Project (Framework: Other, sin build).
