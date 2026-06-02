import os
import re
import sys
import time
import logging
from datetime import date
from collections import defaultdict

URL_PERFIL  = os.environ.get(
    "PERFIL_URL",
    "https://contrataciondelestado.es/wps/poc?uri=deeplink:perfilContratante&idBp=VW8fwBSzF%2FEQK2TEfXGy%2BA%3D%3D"
)
FECHA_DESDE = os.environ.get("FECHA_DESDE", f"01-01-{date.today().year}")
FECHA_HASTA = os.environ.get("FECHA_HASTA", "")
YEAR        = os.environ.get("SHEET_NAME", str(date.today().year))


def agregar_por_año(todas_filas):
    """
    Recibe una lista de registros {año, empresa, num_contratos, proyectos, total}
    y devuelve dos dicts de agregación reutilizables por xlsx y Sheets.

    Returns:
        año_resumen:    {year: {total, contratos, adjudicatarios}}
        empresa_totals: {empresa: {total, contratos, años}}
    """
    año_resumen    = {}
    empresa_totals = {}

    for f in todas_filas:
        year = f["año"]
        if year not in año_resumen:
            año_resumen[year] = {"total": 0.0, "contratos": 0, "adjudicatarios": 0}
        año_resumen[year]["total"]          += f["total"]
        año_resumen[year]["contratos"]      += f["num_contratos"]
        año_resumen[year]["adjudicatarios"] += 1

        e = f["empresa"]
        if e not in empresa_totals:
            empresa_totals[e] = {"total": 0.0, "contratos": 0, "años": set()}
        empresa_totals[e]["total"]     += f["total"]
        empresa_totals[e]["contratos"] += f["num_contratos"]
        empresa_totals[e]["años"].add(year)

    return año_resumen, empresa_totals


def parse_años(args):
    """
    Parsea argumentos CLI para determinar qué años scrapear.

    Sin args       → [año actual]
    "2024"         → [2024]
    "2012-2025"    → [2012, 2013, ..., 2025]
    "2012" "2025"  → [2012, 2013, ..., 2025]
    """
    if not args:
        return [int(YEAR)]
    if len(args) == 1:
        m = re.match(r"^(\d{4})-(\d{4})$", args[0])
        if m:
            start, end = int(m.group(1)), int(m.group(2))
            if start > end:
                sys.exit(f"Error: año inicio ({start}) > año fin ({end})")
            return list(range(start, end + 1))
        try:
            return [int(args[0])]
        except ValueError:
            sys.exit(f"Error: argumento no válido '{args[0]}'. Usa: año, año-año, o año año")
    if len(args) == 2:
        try:
            start, end = int(args[0]), int(args[1])
        except ValueError:
            sys.exit("Error: los argumentos deben ser años (ej. 2012 2025)")
        if start > end:
            sys.exit(f"Error: año inicio ({start}) > año fin ({end})")
        return list(range(start, end + 1))
    sys.exit("Uso: python scraper_xlsx.py [año_inicio] [año_fin]")

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.firefox import GeckoDriverManager

log = logging.getLogger(__name__)


def init_driver():
    options = Options()
    options.add_argument("--headless")
    options.add_argument("--width=1920")
    options.add_argument("--height=1080")
    service = Service(GeckoDriverManager().install())
    return webdriver.Firefox(service=service, options=options)


def scrape(url_perfil, fecha_desde, fecha_hasta=""):
    log.info(f"Iniciando scrape desde {fecha_desde}")

    total_global      = 0.0
    stats             = defaultdict(lambda: {"total": 0.0, "proyectos": []})
    paginas_con_error = []

    driver = init_driver()
    wait   = WebDriverWait(driver, 20)

    try:
        driver.get(url_perfil)

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
        campo_fecha.send_keys(fecha_desde)

        if fecha_hasta:
            campo_hasta = driver.find_element(
                By.XPATH,
                "//*[@id='viewns_Z7_AVEQAI930GRPE02BR764FO30G0_:form1:textMaxFecAnuncioMAQ']"
            )
            campo_hasta.clear()
            campo_hasta.send_keys(fecha_hasta)
            log.info(f"Fecha hasta: {fecha_hasta}")

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

                    # La plataforma muestra importes sin separador de miles ("88501,34 €" o "88501.34 €").
                    # El patrón captura dígitos y comas; replace(",","") convierte la coma decimal.
                    m = re.search(r'[\d,]+(?:\.\d+)?', celda_importe[0].text.strip())
                    if not m:
                        continue
                    try:
                        valor = float(m.group(0).replace(",", ""))
                    except ValueError:
                        continue

                    descripcion = celda_desc[0].text.strip() if celda_desc else "Sin descripción"
                    empresa     = celda_adj[-1].text.strip()  if celda_adj  else "DESCONOCIDO"

                    total_global           += valor
                    stats[empresa]["total"] += valor
                    stats[empresa]["proyectos"].append((descripcion, valor))

                log.info(f"  Página {pagina} OK — acumulado: {total_global:,.2f} €")

            except Exception as e:
                log.error(f"  ERROR en página {pagina}: {e}")
                paginas_con_error.append(pagina)

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
