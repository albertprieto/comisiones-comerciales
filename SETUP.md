# Setup automatización semanal

## 1. Secrets a crear en GitHub

Ve a: **Repo → Settings → Secrets and variables → Actions → New repository secret**

Crea estos 6 secrets (los nombres deben coincidir EXACTAMENTE):

| Nombre              | Valor |
|---------------------|-------|
| `ODOO_URL`          | URL base de Odoo, ej: `https://industrialshields-prod.odoo.com` |
| `ODOO_DB`           | Database name, ej: `industrialshields-prod` |
| `ODOO_USER`         | `apm@industrialshields.com` |
| `ODOO_PASSWORD`     | Password o API key de tu usuario Odoo |
| `PAGOS_CSV_URL`     | URL pública CSV de la hoja "Pagos Comisiones" (ver paso 2) |
| `GMAIL_USER`        | `apm@industrialshields.com` |
| `GMAIL_APP_PASS`    | App Password de 16 chars (ver paso 3) |

## 2. Publicar la hoja Pagos como CSV

1. Abre https://docs.google.com/spreadsheets/d/1urDlTjSZaxcapOiXWz4GUHz9Y4MJiBFT4yP6JRJy3sI/edit
2. **Archivo → Compartir → Publicar en la web**
3. Selecciona la hoja correcta + formato **CSV**
4. Click **Publicar** → Acepta el aviso
5. Copia la URL generada (acaba en `pub?output=csv`) y pégala como secret `PAGOS_CSV_URL`

## 3. Generar Gmail App Password

1. Ve a https://myaccount.google.com/apppasswords
2. App: "Mail"  ·  Device: "GitHub Actions Comisiones"
3. Generar
4. Copia el password de 16 caracteres (sin espacios) y pégalo como secret `GMAIL_APP_PASS`

> El App Password es un password único para esta app (no es tu password normal).
> Solo se ve una vez; si lo pierdes, generas uno nuevo.

## 4. Probar el workflow manualmente

1. **Repo → Actions → Weekly commercial update**
2. Click **"Run workflow"** → branch main → **Run**
3. Espera a que termine (~3-5 min)
4. Verás en el log si todo OK; los HTMLs se actualizarán y los emails se enviarán

## 5. Schedule

Una vez configurado, corre automáticamente:
- **Cada viernes a las 14:30 hora Madrid** (con doble cron por verano/invierno)
- Si quieres cambiar la hora, edita `.github/workflows/weekly-update.yml`

## Lista de destinatarios del email

Editable en `scripts/send_emails.py` → diccionario `RECIPIENTS`. Por defecto:
- Eloi Davila Lopez (edl@)
- Gerard Montero Martínez (gmm@)
- Jordi Hernandez (jhs@)
- Josep Massó (jmp@)
