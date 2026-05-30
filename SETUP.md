# Setup — Automatización con GitHub Actions

## Qué hace esto

Cada domingo a las 3:00 AM (hora UTC), GitHub Actions ejecuta el scraper
automáticamente, borra la hoja "Contratos" del Google Sheet y la reescribe
con todos los contratos del año en curso. Sin duplicados, sin intervención manual.

---

## Paso 1 — Crear Service Account en Google Cloud

1. Ve a https://console.cloud.google.com
2. Crea un proyecto nuevo (o usa uno existente)
3. Activa las APIs necesarias:
   - **Google Sheets API**
   - **Google Drive API**
4. Ve a **IAM y administración → Cuentas de servicio**
5. Clic en **Crear cuenta de servicio**
   - Nombre: `contratos-bot` (o el que quieras)
   - Clic en **Crear y continuar** → **Listo**
6. Entra en la cuenta de servicio creada → pestaña **Claves**
7. Clic en **Agregar clave → Crear clave nueva → JSON**
8. Se descarga un archivo `.json` — **guárdalo, lo necesitas ahora**

---

## Paso 2 — Compartir el Google Sheet con el bot

1. Abre el archivo `.json` descargado y copia el campo `client_email`
   (algo como `contratos-bot@tu-proyecto.iam.gserviceaccount.com`)
2. Abre tu Google Sheet
3. Clic en **Compartir** → pega el email del bot → rol **Editor** → Aceptar

---

## Paso 3 — Subir el repo a GitHub

```bash
cd /home/jesus/Documentos/Github/Selenium/contratos_menores_velez_malaga
git init
git add .
git commit -m "Añadir automation con GitHub Actions"
# Crea el repo en github.com y luego:
git remote add origin https://github.com/TU_USUARIO/contratos-menores-velez.git
git push -u origin main
```

> El repo puede ser **privado** (tienes 2000 min/mes gratis) o **público** (ilimitado).

---

## Paso 4 — Añadir secretos en GitHub

1. En tu repo de GitHub → **Settings → Secrets and variables → Actions**
2. Clic en **New repository secret** y añade estos dos:

### `GOOGLE_CREDENTIALS`
Contenido: el JSON completo del archivo descargado en el Paso 1.
Ábrelo con un editor de texto, selecciona todo y pégalo.

### `SPREADSHEET_ID`
Valor: el ID de tu Google Sheet (los caracteres entre `/d/` y `/edit` en la URL)

### `PERFIL_URL` (opcional)
Solo si quieres usarlo con otro ayuntamiento distinto de Vélez-Málaga.

#### Cómo obtener la URL de tu ayuntamiento

1. Ve a [contrataciondelestado.es](https://contrataciondelestado.es)
2. **Perfil Contratante → Buscar perfiles**
3. En "Organización contratante" → Seleccionar → Sector Público → Entidades Locales → tu provincia → Ayuntamientos → selecciona el tuyo → Añadir
4. Pulsa **Buscar** y entra en el órgano principal (normalmente "Junta de Gobierno del Ayuntamiento de...")
5. Pestaña **Perfil del Contratante** → sección **Datos Generales**
6. Localiza el campo **"Enlace directo vía hiperenlace"**
7. Clic derecho sobre el enlace → **Copiar dirección del enlace**

> ⚠️ Usa siempre el campo "Enlace directo vía hiperenlace", no la URL del navegador.

**Con GitHub Actions:** Settings → Secrets → Actions → añade `PERFIL_URL` con esa URL

**En local (Linux/Mac):**
```bash
export PERFIL_URL="URL_COPIADA"
python scraper.py
```

**En local (fish shell):**
```bash
set -x PERFIL_URL "URL_COPIADA"
python scraper.py
```

---

## Paso 5 — Verificar que funciona

1. Ve a tu repo en GitHub → pestaña **Actions**
2. Clic en **Actualizar Contratos Menores** → **Run workflow**
3. Espera ~5 minutos y comprueba que el Google Sheet se ha actualizado

---

## Frecuencia

Por defecto: **cada domingo a las 3:00 UTC**.

Para cambiarlo, edita la línea `cron:` en `.github/workflows/update_sheet.yml`:

```yaml
# Cada día a las 3am
- cron: '0 3 * * *'

# Cada lunes
- cron: '0 3 * * 1'

# El 1 de cada mes
- cron: '0 3 1 * *'
```

---

## Garantía anti-duplicados

El scraper siempre:
1. Consulta desde `01-01-AÑO_ACTUAL` hasta hoy
2. **Borra completamente** la hoja antes de escribir
3. Reescribe todo desde cero

No hay ningún mecanismo de append, por lo que es imposible duplicar datos.
