"""
Scraper automático de contratos menores - Vélez-Málaga
Diseñado para ejecutarse sin intervención manual (headless, sin input()).
Escribe directamente en Google Sheets, borrando antes los datos anteriores
para garantizar que nunca haya duplicados.
"""

import os
import re
import sys
import json
import time
import logging
from datetime import date
from collections import defaultdict

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.firefox import GeckoDriverManager

import gspread
from gspread.exceptions import WorksheetNotFound
from google.oauth2.service_account import Credentials

# ---------------------------------------------------------------------------
# Configuración
# ---------------------------------------------------------------------------

URL_PERFIL = (
    "https://contrataciondelestado.es/wps/poc"
    "?uri=deeplink:perfilContratante&idBp=VW8fwBSzF%2FEQK2TEfXGy%2BA%3D%3D"
)

FECHA_DESDE   = f"01-01-{date.today().year}"
SPREADSHEET_ID = os.environ.get("SPREADSHEET_ID")
if not SPREADSHEET_ID:
    raise EnvironmentError("Variable SPREADSHEET_ID no definida.")
SHEET_NAME    = f"Contratos {date.today().year}"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)
log = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Google Sheets
# ---------------------------------------------------------------------------

def get_worksheet():
    creds_raw = os.environ.get("GOOGLE_CREDENTIALS")
    if not creds_raw:
        raise EnvironmentError("Variable GOOGLE_CREDENTIALS no definida.")

    creds_dict = json.loads(creds_raw)
    creds = Credentials.from_service_account_info(
        creds_dict,
        scopes=[
            "https://spreadsheets.google.com/feeds",
            "https://www.googleapis.com/auth/drive",
        ],
    )
    gc = gspread.Client(auth=creds)
    sh = gc.open_by_key(SPREADSHEET_ID)

    try:
        ws = sh.worksheet(SHEET_NAME)
    except WorksheetNotFound:
        ws = sh.add_worksheet(title=SHEET_NAME, rows=1000, cols=10)

    return ws


def escribir_en_sheets(ws, ranking, total_global, paginas_con_error):
    log.info("Borrando hoja y reescribiendo datos...")

    hoy = date.today().strftime("%d/%m/%Y")
    aviso = f"⚠ Páginas con error: {paginas_con_error}" if paginas_con_error else ""

    # Fila de metadatos
    meta = [["Última actualización:", hoy, f"Total adjudicatarios: {len(ranking)}", aviso]]

    cabecera = [["#", "Adjudicatario", "Nº Contratos", "Proyectos", "Total (€)"]]

    filas_datos = []
    for i, (empresa, datos) in enumerate(ranking, 1):
        proyectos_str = " | ".join(
            f"{desc} ({valor:.2f} €)" for desc, valor in datos["proyectos"]
        )
        filas_datos.append([
            i,
            empresa,
            len(datos["proyectos"]),
            proyectos_str,
            round(datos["total"], 2),
        ])

    fila_total = [["", "TOTAL GENERAL", "", "", round(total_global, 2)]]

    todas = meta + [[]] + cabecera + filas_datos + [[]] + fila_total

    ws.clear()
    ws.update(values=todas, range_name="A1", value_input_option="USER_ENTERED")
    log.info(f"Hoja actualizada: {len(ranking)} adjudicatarios — {total_global:,.2f} €")

# ---------------------------------------------------------------------------
# Scraper
# ---------------------------------------------------------------------------

def init_driver():
    options = Options()
    options.add_argument("--headless")
    options.add_argument("--width=1920")
    options.add_argument("--height=1080")
    service = Service(GeckoDriverManager().install())
    return webdriver.Firefox(service=service, options=options)


def scrape():
    log.info(f"Iniciando scrape desde {FECHA_DESDE}")

    # Inicializar fuera del try para que siempre estén definidas
    total_global = 0.0
    stats = defaultdict(lambda: {"total": 0.0, "proyectos": []})
    paginas_con_error = []

    driver = init_driver()
    wait = WebDriverWait(driver, 20)

    try:
        driver.get(URL_PERFIL)

        tab = wait.until(EC.element_to_be_clickable((
            By.XPATH,
            "//*[@id='viewns_Z7_AVEQAI930GRPE02BR764FO30G0_:perfilComp:linkPrepContratosMenores']"
        )))
        tab.click()
        log.info("Tab 'Contratos Menores' pulsado.")

        campo_fecha = wait.until(EC.presence_of_element_located((
            By.XPATH,
            "//*[@id='viewns_Z7_AVEQAI930GRPE02BR764FO30G0_:form1:textMinFecAnuncioMAQ2']"
        )))
        campo_fecha.clear()
        campo_fecha.send_keys(FECHA_DESDE)

        boton = wait.until(EC.element_to_be_clickable((
            By.XPATH,
            "//*[@id='viewns_Z7_AVEQAI930GRPE02BR764FO30G0_:form1:busReasProc18']"
        )))
        boton.click()
        log.info("Búsqueda iniciada...")

        wait.until(EC.presence_of_element_located((
            By.XPATH, "//td[contains(@class,'tdImporte')]"
        )))

        pagina = 1
        while True:
            log.info(f"Procesando página {pagina}...")
            time.sleep(2)

            try:
                filas = driver.find_elements(
                    By.XPATH,
                    "//table//tr[.//td[contains(@class,'tdImporte')]]"
                )
                for fila in filas:
                    celda_importe = fila.find_elements(By.XPATH, ".//td[contains(@class,'tdImporte')]")
                    celda_desc    = fila.find_elements(By.XPATH, ".//td[contains(@class,'tdTipoContratoLicOC')]")
                    celda_adj     = fila.find_elements(By.XPATH, ".//td[contains(@class,'tdFecha') and contains(@class,'textAlignLeft')]")

                    if not celda_importe:
                        continue

                    m = re.search(r'[\d,]+(?:\.\d+)?', celda_importe[0].text.strip())
                    if not m:
                        continue
                    try:
                        valor = float(m.group(0).replace(",", ""))
                    except ValueError:
                        continue

                    descripcion = celda_desc[0].text.strip() if celda_desc else "Sin descripción"
                    empresa     = celda_adj[-1].text.strip() if celda_adj else "DESCONOCIDO"

                    total_global           += valor
                    stats[empresa]["total"] += valor
                    stats[empresa]["proyectos"].append((descripcion, valor))

                log.info(f"  Página {pagina} OK — acumulado: {total_global:,.2f} €")

            except Exception as e:
                log.error(f"  ERROR en página {pagina}: {e}")
                paginas_con_error.append(pagina)

            # Siguiente página
            try:
                siguiente = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((
                        By.ID,
                        "viewns_Z7_AVEQAI930GRPE02BR764FO30G0_:form1:siguienteLink"
                    ))
                )
                driver.execute_script("arguments[0].scrollIntoView();", siguiente)
                time.sleep(0.5)
                siguiente.click()
                pagina += 1
            except Exception:
                log.info("No hay más páginas.")
                break

    finally:
        driver.quit()

    if paginas_con_error:
        log.warning(f"Páginas con error: {paginas_con_error} — los datos pueden estar incompletos.")

    log.info(f"TOTAL FINAL: {total_global:,.2f} € — {len(stats)} adjudicatarios")
    ranking = sorted(stats.items(), key=lambda x: x[1]["total"], reverse=True)
    return ranking, total_global, paginas_con_error


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    ranking, total_global, paginas_con_error = scrape()

    if not ranking:
        log.error("No se obtuvo ningún dato. Abortando escritura en Sheets.")
        sys.exit(1)

    ws = get_worksheet()
    escribir_en_sheets(ws, ranking, total_global, paginas_con_error)
    log.info("Proceso completado.")
