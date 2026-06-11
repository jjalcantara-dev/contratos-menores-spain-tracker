import os
import re
import sys
import time
import logging
import unicodedata
from dataclasses import dataclass
from datetime import date
from collections import defaultdict
from collections.abc import Sequence
from typing import NamedTuple, TypedDict

URL_PERFIL  = os.environ.get(
    "PERFIL_URL",
    "https://contrataciondelestado.es/wps/poc?uri=deeplink:perfilContratante&idBp=VW8fwBSzF%2FEQK2TEfXGy%2BA%3D%3D"
)
FECHA_DESDE = os.environ.get("FECHA_DESDE", f"01-01-{date.today().year}")
FECHA_HASTA = os.environ.get("FECHA_HASTA", "")
YEAR        = os.environ.get("SHEET_NAME", str(date.today().year))

# Nombres de pestañas — fuente única de verdad compartida por xlsx y Sheets
TAB_ESTADISTICAS   = "Estadísticas"
TAB_REGISTRO_TOTAL = "Registro Total"
PREFIJO_CONTRATO   = "Contratos"
PATRON_TAB_AÑO     = re.compile(rf"^{PREFIJO_CONTRATO} (\d{{4}})$")

# ---------------------------------------------------------------------------
# Tipos compartidos
# ---------------------------------------------------------------------------

class FilaContrato(TypedDict):
    año:           int
    empresa:       str
    num_contratos: int
    proyectos:     str
    total:         float


class ResumenAño(TypedDict):
    total:          float
    contratos:      int
    adjudicatarios: int


class ResumenEmpresa(TypedDict):
    total:     float
    contratos: int
    años:      set[int]


class ResultadoAgregado(NamedTuple):
    año_resumen:    dict[int, ResumenAño]
    empresa_totals: dict[str, ResumenEmpresa]


class ResultadoScrape(NamedTuple):
    ranking:           list
    total_global:      float
    paginas_con_error: list[int]


@dataclass
class DatosEstadisticas:
    años_ordenados: list
    año_resumen:    dict
    top_empresas:   list   # [(empresa, ResumenEmpresa), ...]
    n_años:         int
    n_top:          int

# ---------------------------------------------------------------------------
# Funciones de agregación
# ---------------------------------------------------------------------------

def _normalizar_empresa(nombre: str) -> str:
    """Clave de agrupación normalizada para una empresa.
    'AXAPLAY, S.L.' y 'AXAPLAY S.L.' → 'AXAPLAY SL'
    'HERMANOS GUIRADO, S.C.' y 'HERMANOS GUIRADO SC' → 'HERMANOS GUIRADO SC'
    """
    # Eliminar acentos (NFD descompone, Mn es la categoría de diacríticos)
    texto = unicodedata.normalize("NFD", nombre.upper())
    texto = "".join(c for c in texto if unicodedata.category(c) != "Mn")
    # Quitar puntos y comas que varían entre registros del mismo adjudicatario
    texto = texto.replace(".", "").replace(",", "")
    return " ".join(texto.split())


def _parse_importe(text: str) -> float | None:
    """Extrae el importe numérico del texto de una celda de la plataforma.
    Maneja el formato sin separador de miles: '88501,34 €' o '88501.34 €'.
    """
    m = re.search(r'[\d,]+(?:\.\d+)?', text.strip())
    if not m:
        return None
    try:
        return float(m.group(0).replace(",", ""))
    except ValueError:
        return None


def normalizar_fila(row: Sequence, year: int) -> FilaContrato | None:
    """Convierte una fila en bruto en FilaContrato.
    Devuelve None si la fila es de cabecera, total o tiene datos inválidos.
    Compatible con openpyxl (iter_rows values_only) y gspread (get_all_values).
    """
    empresa = row[1] if len(row) > 1 else None
    if not empresa or str(empresa) in ("Adjudicatario", "TOTAL GENERAL"):
        return None

    total_raw = row[4] if len(row) > 4 else None
    if total_raw is None or total_raw == "":
        return None
    try:
        total = float(total_raw)
    except (ValueError, TypeError):
        return None

    try:
        num_contratos = int(row[2]) if len(row) > 2 and row[2] not in (None, "") else 0
    except (ValueError, TypeError):
        num_contratos = 0

    return {
        "año":           year,
        "empresa":       str(empresa),
        "num_contratos": num_contratos,
        "proyectos":     str(row[3]) if len(row) > 3 and row[3] else "",
        "total":         total,
    }


def agregar_por_año(todas_filas: list[FilaContrato]) -> ResultadoAgregado:
    """
    Recibe una lista de FilaContrato y devuelve resúmenes por año y por empresa.
    Reutilizable por scraper_xlsx y scraper.
    """
    año_resumen:    dict = {}
    empresa_totals: dict = {}

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

    return ResultadoAgregado(año_resumen, empresa_totals)


def preparar_estadisticas(todas_filas: list[FilaContrato], top_n: int = 10) -> DatosEstadisticas:
    """Agrega filas y construye el DTO listo para renderizar en Estadísticas."""
    resultado      = agregar_por_año(todas_filas)
    años_ordenados = sorted(resultado.año_resumen.keys())
    top_empresas   = sorted(
        resultado.empresa_totals.items(),
        key=lambda x: x[1]["total"],
        reverse=True,
    )[:top_n]
    return DatosEstadisticas(
        años_ordenados = años_ordenados,
        año_resumen    = resultado.año_resumen,
        top_empresas   = top_empresas,
        n_años         = len(años_ordenados),
        n_top          = len(top_empresas),
    )

# ---------------------------------------------------------------------------
# Parseo de argumentos CLI
# ---------------------------------------------------------------------------

def parse_años(args: list[str]) -> list[int]:
    """
    Sin args       → [año actual]
    "2024"         → [2024]
    "2012-2025"    → [2012, ..., 2025]
    "2012" "2025"  → [2012, ..., 2025]
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

# ---------------------------------------------------------------------------
# Scraper Selenium
# ---------------------------------------------------------------------------

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.firefox import GeckoDriverManager

log = logging.getLogger(__name__)


def init_driver() -> webdriver.Firefox:
    options = Options()
    options.add_argument("--headless")
    options.add_argument("--width=1920")
    options.add_argument("--height=1080")
    service = Service(GeckoDriverManager().install())
    driver = webdriver.Firefox(service=service, options=options)
    driver.set_page_load_timeout(120)
    return driver


def _get_with_retry(driver, url: str, max_intentos: int = 3, espera: int = 15) -> None:
    for intento in range(1, max_intentos + 1):
        try:
            driver.get(url)
            return
        except Exception as exc:
            log.warning(f"Intento {intento}/{max_intentos} fallido al cargar {url}: {exc}")
            if intento == max_intentos:
                raise
            time.sleep(espera)


def scrape(url_perfil: str, fecha_desde: str, fecha_hasta: str = "") -> tuple:
    log.info(f"Iniciando scrape desde {fecha_desde}")

    total_global                 = 0.0
    stats                        = defaultdict(lambda: {"total": 0.0, "proyectos": []})
    canon_nombre: dict[str, str] = {}   # clave normalizada → nombre original (primera aparición)
    paginas_con_error: list[int] = []

    driver = init_driver()
    wait   = WebDriverWait(driver, 20)

    try:
        _get_with_retry(driver, url_perfil)

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

                    valor = _parse_importe(celda_importe[0].text)
                    if valor is None:
                        continue

                    descripcion = celda_desc[0].text.strip() if celda_desc else "Sin descripción"
                    empresa = celda_adj[-1].text.strip() if celda_adj else "DESCONOCIDO"
                    clave   = _normalizar_empresa(empresa)
                    if clave not in canon_nombre:
                        canon_nombre[clave] = empresa   # conserva el primer nombre visto

                    total_global              += valor
                    stats[clave]["total"]     += valor
                    stats[clave]["proyectos"].append((descripcion, valor))

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
    ranking = sorted(
        [(canon_nombre[clave], datos) for clave, datos in stats.items()],
        key=lambda x: x[1]["total"],
        reverse=True,
    )
    return ResultadoScrape(ranking, total_global, paginas_con_error)
