"""
Versión local del scraper — genera un archivo .xlsx sin configurar Google Sheets.
Uso:
    python scraper_xlsx.py
"""

import os
import sys
import logging
from datetime import date
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter

from scraper_core import scrape, URL_PERFIL, FECHA_DESDE, FECHA_HASTA, YEAR

OUTPUT_DIR = Path("output")

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler(OUTPUT_DIR / f"contratos_{YEAR}.log", encoding="utf-8"),
    ],
)
log = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Exportar a xlsx
# ---------------------------------------------------------------------------

VERDE_OSCURO = "1F5C2E"
VERDE_CLARO  = "E8F5E9"


def escribir_xlsx(ranking, total_global, paginas_con_error):
    OUTPUT_DIR.mkdir(exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = f"Contratos {YEAR}"

    fill_verde  = PatternFill("solid", fgColor=VERDE_OSCURO)
    fill_claro  = PatternFill("solid", fgColor=VERDE_CLARO)
    font_blanco = Font(bold=True, color="FFFFFF", size=11)
    center      = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left        = Alignment(horizontal="left",   vertical="center", wrap_text=True)

    hoy   = date.today().strftime("%d/%m/%Y")
    aviso = f"Páginas con error: {paginas_con_error}" if paginas_con_error else ""
    ws.append([f"Última actualización: {hoy}", f"Total adjudicatarios: {len(ranking)}", aviso])
    ws.append([])

    cabecera = ["#", "Adjudicatario", "Nº Contratos", "Proyectos", "Total (€)"]
    ws.append(cabecera)
    fila_cabecera = ws.max_row
    for col in range(1, 6):
        cell = ws.cell(fila_cabecera, col)
        cell.fill = fill_verde
        cell.font = font_blanco
        cell.alignment = center

    for i, (empresa, datos) in enumerate(ranking, 1):
        proyectos_str = " | ".join(
            f"{desc} ({valor:.2f} €)" for desc, valor in datos["proyectos"]
        )
        ws.append([i, empresa, len(datos["proyectos"]), proyectos_str, round(datos["total"], 2)])
        fila_actual = ws.max_row
        for col in range(1, 6):
            cell = ws.cell(fila_actual, col)
            if i % 2 == 0:
                cell.fill = fill_claro
            cell.alignment = left

    ws.append([])
    ws.append(["", "TOTAL GENERAL", "", "", round(total_global, 2)])
    fila_total = ws.max_row
    for col in range(1, 6):
        cell = ws.cell(fila_total, col)
        cell.fill = fill_verde
        cell.font = font_blanco

    for col, ancho in [(1, 5), (2, 40), (3, 14), (4, 80), (5, 16)]:
        ws.column_dimensions[get_column_letter(col)].width = ancho

    ws.freeze_panes = f"A{fila_cabecera + 1}"

    nombre = OUTPUT_DIR / f"contratos_{YEAR}.xlsx"
    wb.save(nombre)
    log.info(f"Archivo generado: {nombre}")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    ranking, total_global, paginas_con_error = scrape(URL_PERFIL, FECHA_DESDE, FECHA_HASTA)

    if not ranking:
        log.error("No se obtuvo ningún dato. Abortando.")
        sys.exit(1)

    escribir_xlsx(ranking, total_global, paginas_con_error)
    log.info("Proceso completado.")
