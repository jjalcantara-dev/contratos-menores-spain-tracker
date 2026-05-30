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

## Paso 3 — Obtener el código en GitHub

No necesitas descargar nada ni tocar código. Haz un **fork** de este repositorio:

1. Si no tienes cuenta en GitHub, créate una en [github.com](https://github.com) (es gratis)
2. Entra en la página del repositorio en GitHub
3. Pulsa el botón **Fork** que aparece arriba a la derecha, junto al nombre del repositorio
4. En la pantalla que aparece, deja el nombre como está y pulsa **Create fork**

GitHub copia todo el repositorio en tu cuenta. A partir de ahí trabajas sobre tu copia, no sobre la original. El código y el workflow de GitHub Actions ya están incluidos — no hay que configurar nada más de lo que se describe en los pasos siguientes.

> El repo puede ser **privado** (tienes 2000 min/mes gratis de Actions) o **público** (ilimitado).

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

#### Cómo obtener el enlace directo del Perfil del Contratante

El programa necesita la URL del campo **"Enlace directo vía hiperenlace"** dentro de la Plataforma de Contratación del Sector Público.

**Paso 1 — Acceder al buscador de perfiles**

Entrar en [contrataciondelestado.es](https://contrataciondelestado.es) e ir a:

**Perfil Contratante → Buscar perfiles**

**Paso 2 — Seleccionar el ayuntamiento**

En el selector de **Organización contratante**:

1. Pulsar **Seleccionar**
2. Desplegar: Sector Público → Entidades Locales → tu comunidad → tu provincia → Ayuntamientos
3. Seleccionar tu ayuntamiento y pulsar **Añadir**

**Paso 3 — Realizar la búsqueda**

Pulsar **Buscar**. Aparecerá una tabla con los órganos de contratación. Seleccionar el órgano principal, normalmente:

**Junta de Gobierno del Ayuntamiento de [Municipio]**

**Paso 4 — Abrir el Perfil del Contratante**

Dentro del órgano de contratación, pulsar la pestaña **Perfil del Contratante** → sección **Datos Generales**.

**Paso 5 — Copiar el enlace directo**

Localizar el campo **"Enlace directo vía hiperenlace"**, hacer clic derecho sobre el enlace y seleccionar **Copiar dirección del enlace**.

> ⚠️ No usar la URL del navegador. Usar siempre el campo "Enlace directo vía hiperenlace", que es la referencia permanente del perfil.

La URL obtenida tendrá este aspecto:

```
https://contrataciondelestado.es/wps/poc?uri=deeplink:perfilContratante&idBp=XXXXXX%3D%3D
```

Pégala tal cual, sin modificar ningún carácter.

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

## Cargar años anteriores

Por defecto el scraper solo recoge el año en curso. Para rellenar años históricos hay que ejecutarlo manualmente con las variables `FECHA_DESDE`, `FECHA_HASTA` y `SHEET_NAME`.

**Linux/Mac:**
```bash
export SPREADSHEET_ID="tu-id"
export GOOGLE_CREDENTIALS="$(cat credenciales.json)"
export FECHA_DESDE="01-01-2023"
export FECHA_HASTA="31-12-2023"
export SHEET_NAME="Contratos 2023"
python scraper.py
```

**fish shell:**
```bash
set -x SPREADSHEET_ID "tu-id"
set -x GOOGLE_CREDENTIALS (cat credenciales.json)
set -x FECHA_DESDE "01-01-2023"
set -x FECHA_HASTA "31-12-2023"
set -x SHEET_NAME "Contratos 2023"
python scraper.py
```

Repite el proceso cambiando el año para cada pestaña que quieras generar (2018, 2019, 2020...). Cada ejecución crea la pestaña si no existe, o la sobreescribe si ya está.

---

## Garantía anti-duplicados

El scraper siempre:
1. Consulta desde `01-01-AÑO_ACTUAL` hasta hoy
2. **Borra completamente** la hoja antes de escribir
3. Reescribe todo desde cero

No hay ningún mecanismo de append, por lo que es imposible duplicar datos.
