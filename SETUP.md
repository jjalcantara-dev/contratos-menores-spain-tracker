# Setup — Automatización con GitHub Actions

## ¿Solo quieres probarlo rápido?

Si no quieres configurar Google Sheets ni GitHub, usa la versión local. Solo necesitas Python y Firefox instalados:

```bash
pip install selenium webdriver-manager openpyxl
python scraper_xlsx.py
```

Genera `output/contratos_AÑO.xlsx` y `output/contratos_AÑO.log`. Para otro municipio añade `PERFIL_URL` (ver [Cómo obtener la URL](#cómo-obtener-el-enlace-directo-del-perfil-del-contratante)). Para otro año, consulta [Cargar años anteriores](#cargar-años-anteriores).

---

## Qué hace la automatización con GitHub Actions

Cada domingo a las 3:00 AM (hora UTC), GitHub Actions ejecuta el scraper automáticamente, borra la hoja del Google Sheet y la reescribe con todos los contratos del año en curso. Sin duplicados, sin intervención manual.

---

## Paso 1 — Fork del repositorio

No necesitas descargar nada ni tocar código. Haz un **fork** de este repositorio:

1. Si no tienes cuenta en GitHub, créate una en [github.com](https://github.com) (es gratis)
2. Entra en la página del repositorio en GitHub
3. Pulsa el botón **Fork** que aparece arriba a la derecha, junto al nombre del repositorio
4. En la pantalla que aparece, deja el nombre como está y pulsa **Create fork**

GitHub copia todo el repositorio en tu cuenta. El código y el workflow de GitHub Actions ya están incluidos — no hay que configurar nada más de lo que se describe en los pasos siguientes.

> El repo puede ser **privado** (tienes 2000 min/mes gratis de Actions) o **público** (ilimitado).

---

## Paso 2 — Crear un Google Sheet

Crea una hoja de cálculo vacía en [Google Sheets](https://sheets.google.com). El scraper creará y formateará las pestañas automáticamente.

Apunta el **ID** de la hoja: son los caracteres entre `/d/` y `/edit` en la URL.

```
https://docs.google.com/spreadsheets/d/ESTE_ES_EL_ID/edit
```

---

## Paso 3 — Crear Service Account en Google Cloud

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

## Paso 4 — Compartir el Google Sheet con el bot

1. Abre el archivo `.json` descargado y copia el campo `client_email`
   (algo como `contratos-bot@tu-proyecto.iam.gserviceaccount.com`)
2. Abre tu Google Sheet
3. Clic en **Compartir** → pega el email del bot → rol **Editor** → Aceptar

---

## Paso 5 — Añadir secretos en GitHub

1. En tu repo de GitHub → **Settings → Secrets and variables → Actions**
2. Clic en **New repository secret** y añade estos:

### `GOOGLE_CREDENTIALS`
Contenido: el JSON completo del archivo descargado en el Paso 3.
Ábrelo con un editor de texto, selecciona todo y pégalo.

### `SPREADSHEET_ID`
Valor: el ID de tu Google Sheet anotado en el Paso 2.

### `PERFIL_URL` (opcional)
Solo si quieres usarlo con otro ayuntamiento distinto de Vélez-Málaga.

#### Cómo obtener el enlace directo del Perfil del Contratante

El programa necesita la URL del campo **"Enlace directo vía hiperenlace"** dentro de la Plataforma de Contratación del Sector Público.

**1 — Acceder al buscador de perfiles**

Entrar en [contrataciondelestado.es](https://contrataciondelestado.es) e ir a:

**Perfil Contratante → Buscar perfiles**

**2 — Seleccionar el ayuntamiento**

En el selector de **Organización contratante**:

1. Pulsar **Seleccionar**
2. Desplegar: Sector Público → Entidades Locales → tu comunidad → tu provincia → Ayuntamientos
3. Seleccionar tu ayuntamiento y pulsar **Añadir**

**3 — Realizar la búsqueda**

Pulsar **Buscar**. Aparecerá una tabla con los órganos de contratación. Seleccionar el órgano principal, normalmente:

**Junta de Gobierno del Ayuntamiento de [Municipio]**

**4 — Abrir el Perfil del Contratante**

Dentro del órgano de contratación, pulsar la pestaña **Perfil del Contratante** → sección **Datos Generales**.

**5 — Copiar el enlace directo**

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
python scraper_xlsx.py
```

**En local (fish shell):**
```bash
set -x PERFIL_URL "URL_COPIADA"
python scraper_xlsx.py
```

---

## Paso 6 — Verificar que funciona

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

Por defecto el scraper solo recoge el año en curso. Para rellenar años históricos ejecuta manualmente con las variables `FECHA_DESDE`, `FECHA_HASTA` y `SHEET_NAME`.

> `SHEET_NAME` es el año en ambos casos (`"2023"`, no `"Contratos 2023"`). El prefijo se añade automáticamente.

**Opción A — Local con xlsx (sin credenciales de Google):**

```bash
# Linux/Mac
export FECHA_DESDE="01-01-2023"
export FECHA_HASTA="31-12-2023"
export SHEET_NAME="2023"
python scraper_xlsx.py
```

```bash
# fish shell
set -x FECHA_DESDE "01-01-2023"
set -x FECHA_HASTA "31-12-2023"
set -x SHEET_NAME "2023"
python scraper_xlsx.py
```

Genera `output/contratos_2023.xlsx`. Repite cambiando el año.

**Opción B — Directo a Google Sheets:**

```bash
# Linux/Mac
export SPREADSHEET_ID="tu-id"
export GOOGLE_CREDENTIALS="$(cat credenciales.json)"
export FECHA_DESDE="01-01-2023"
export FECHA_HASTA="31-12-2023"
export SHEET_NAME="2023"
python scraper.py
```

```bash
# fish shell
set -x SPREADSHEET_ID "tu-id"
set -x GOOGLE_CREDENTIALS (cat credenciales.json)
set -x FECHA_DESDE "01-01-2023"
set -x FECHA_HASTA "31-12-2023"
set -x SHEET_NAME "2023"
python scraper.py
```

Cada ejecución crea la pestaña si no existe, o la sobreescribe si ya está.

---

## Garantía anti-duplicados

El scraper siempre:
1. Consulta desde `01-01-AÑO` hasta `31-12-AÑO` (o hoy si es el año en curso)
2. **Borra completamente** la hoja antes de escribir
3. Reescribe todo desde cero

No hay ningún mecanismo de append, por lo que es imposible duplicar datos.
