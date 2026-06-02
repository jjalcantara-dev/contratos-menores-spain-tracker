"""
Versión local del scraper — genera un archivo .xlsx sin configurar Google Sheets.
Uso:
    python scraper_xlsx.py                # año actual
    python scraper_xlsx.py 2023           # año concreto
    python scraper_xlsx.py 2012 2025      # rango completo
    python scraper_xlsx.py 2012-2025      # misma sintaxis con guion
"""

import re
import sys
import logging
from datetime import date
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter

from scraper_core import (
    scrape, URL_PERFIL, FECHA_DESDE, FECHA_HASTA, YEAR,
    parse_años, agregar_por_año,
)

OUTPUT_DIR  = Path("output")
NOMBRE_XLSX = OUTPUT_DIR / "contratos.xlsx"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler(OUTPUT_DIR / f"contratos_{YEAR}.log", encoding="utf-8"),
    ],
)
log = logging.getLogger(__name__)

VERDE_OSCURO = "1F5C2E"
VERDE_CLARO  = "E8F5E9"


# ---------------------------------------------------------------------------
# Helpers internos
# ---------------------------------------------------------------------------

def _estilos():
    return (
        PatternFill("solid", fgColor=VERDE_OSCURO),
        PatternFill("solid", fgColor=VERDE_CLARO),
        Font(bold=True, color="FFFFFF", size=11),
        Alignment(horizontal="center", vertical="center", wrap_text=True),
        Alignment(horizontal="left",   vertical="center", wrap_text=True),
    )


def _tabs_año(wb):
    """Devuelve [(year_int, worksheet), ...] ordenados cronológicamente."""
    resultado = []
    for name in wb.sheetnames:
        m = re.match(r"^Contratos (\d{4})$", name)
        if m:
            resultado.append((int(m.group(1)), wb[name]))
    resultado.sort()
    return resultado


def _leer_datos_año(ws_year, year):
    """Lee filas de datos (excluyendo meta/cabecera/total) de un tab de año xlsx."""
    filas = []
    for row in ws_year.iter_rows(min_row=4, values_only=True):
        empresa = row[1] if len(row) > 1 else None
        if not empresa or empresa == "TOTAL GENERAL":
            continue
        total = row[4] if len(row) > 4 else None
        if total is None:
            continue
        filas.append({
            "año":           year,
            "empresa":       empresa,
            "num_contratos": row[2] or 0,
            "proyectos":     row[3] or "",
            "total":         float(total),
        })
    return filas


def _open_or_create_wb():
    if NOMBRE_XLSX.exists():
        return load_workbook(NOMBRE_XLSX)
    wb = Workbook()
    wb.remove(wb.active)
    return wb


def _reordenar_tabs(wb):
    """Ordena: año_actual → Estadísticas → años anteriores (desc) → Registro Total"""
    names     = wb.sheetnames
    hoy_año   = str(date.today().year)
    tab_actual = f"Contratos {hoy_año}"

    year_names = sorted(
        [n for n in names if re.match(r"^Contratos \d{4}$", n)],
        reverse=True,
    )

    desired = []
    if tab_actual in year_names:
        desired.append(tab_actual)
        year_names = [n for n in year_names if n != tab_actual]
    elif year_names:
        desired.append(year_names[0])  # más reciente disponible
        year_names = year_names[1:]

    if "Estadísticas" in names:
        desired.append("Estadísticas")
    desired.extend(year_names)
    if "Registro Total" in names:
        desired.append("Registro Total")

    sheet_map  = {ws.title: ws for ws in wb.worksheets}
    wb._sheets = [sheet_map[n] for n in desired if n in sheet_map]


# ---------------------------------------------------------------------------
# Tab de año
# ---------------------------------------------------------------------------

def _escribir_tab_año(wb, year, ranking, total_global, paginas_con_error):
    """Escribe/sobreescribe el tab de un año en el workbook (sin guardar)."""
    tab_nombre = f"Contratos {year}"
    if tab_nombre in wb.sheetnames:
        wb.remove(wb[tab_nombre])
    ws = wb.create_sheet(title=tab_nombre)

    fill_verde, fill_claro, font_blanco, center, left = _estilos()

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


def escribir_xlsx(ranking, total_global, paginas_con_error):
    """Punto de entrada para año único (usa YEAR del entorno). Guarda el archivo."""
    OUTPUT_DIR.mkdir(exist_ok=True)
    wb = _open_or_create_wb()
    _escribir_tab_año(wb, YEAR, ranking, total_global, paginas_con_error)
    actualizar_registro_total(wb)
    actualizar_estadisticas(wb)
    _reordenar_tabs(wb)
    wb.save(NOMBRE_XLSX)
    log.info(f"Archivo actualizado: {NOMBRE_XLSX}")


# ---------------------------------------------------------------------------
# Tab Registro Total
# ---------------------------------------------------------------------------

def actualizar_registro_total(wb):
    """Reconstruye 'Registro Total' leyendo todos los tabs de año."""
    year_tabs = _tabs_año(wb)
    if not year_tabs:
        return

    if "Registro Total" in wb.sheetnames:
        wb.remove(wb["Registro Total"])
    ws = wb.create_sheet("Registro Total")

    fill_verde, fill_claro, font_blanco, center, left = _estilos()

    todas_filas = []
    for year, ws_year in year_tabs:
        todas_filas.extend(_leer_datos_año(ws_year, year))
    todas_filas.sort(key=lambda r: r["total"], reverse=True)

    hoy  = date.today().strftime("%d/%m/%Y")
    años = ", ".join(str(y) for y, _ in year_tabs)
    ws.append([f"Última actualización: {hoy}", f"Total registros: {len(todas_filas)}", f"Años: {años}"])
    ws.append([])

    cabecera = ["Año", "Adjudicatario", "Nº Contratos", "Proyectos", "Total (€)"]
    ws.append(cabecera)
    fila_cab = ws.max_row
    for col in range(1, 6):
        cell = ws.cell(fila_cab, col)
        cell.fill = fill_verde
        cell.font = font_blanco
        cell.alignment = center

    for i, f in enumerate(todas_filas, 1):
        ws.append([f["año"], f["empresa"], f["num_contratos"], f["proyectos"], round(f["total"], 2)])
        fila_actual = ws.max_row
        for col in range(1, 6):
            cell = ws.cell(fila_actual, col)
            if i % 2 == 0:
                cell.fill = fill_claro
            cell.alignment = left

    ws.append([])
    total_global = sum(f["total"] for f in todas_filas)
    ws.append(["", "TOTAL GENERAL", "", "", round(total_global, 2)])
    fila_total = ws.max_row
    for col in range(1, 6):
        cell = ws.cell(fila_total, col)
        cell.fill = fill_verde
        cell.font = font_blanco

    for col, ancho in [(1, 8), (2, 40), (3, 14), (4, 80), (5, 16)]:
        ws.column_dimensions[get_column_letter(col)].width = ancho
    ws.freeze_panes = f"A{fila_cab + 1}"

    log.info(f"Tab 'Registro Total' actualizado: {len(todas_filas)} registros de {len(year_tabs)} año(s)")


# ---------------------------------------------------------------------------
# Tab Estadísticas (resúmenes + gráficas)
# ---------------------------------------------------------------------------

def actualizar_estadisticas(wb):
    """Reconstruye 'Estadísticas' con tablas resumen y gráficas por año/adjudicatario."""
    year_tabs = _tabs_año(wb)
    if not year_tabs:
        return

    todas_filas = []
    for year, ws_year in year_tabs:
        todas_filas.extend(_leer_datos_año(ws_year, year))

    año_resumen, empresa_totals = agregar_por_año(todas_filas)

    if "Estadísticas" in wb.sheetnames:
        wb.remove(wb["Estadísticas"])
    ws = wb.create_sheet("Estadísticas")

    fill_verde, fill_claro, font_blanco, center, left = _estilos()
    font_titulo = Font(bold=True, size=12)

    # --- Tabla 1: Resumen por año ---
    ws.append(["Resumen por año"])
    ws.cell(ws.max_row, 1).font = font_titulo

    ws.append(["Año", "Total (€)", "Nº Contratos", "Nº Adjudicatarios"])
    fila_cab_años = ws.max_row
    for col in range(1, 5):
        c = ws.cell(fila_cab_años, col)
        c.fill = fill_verde
        c.font = font_blanco
        c.alignment = center

    años_ordenados = sorted(año_resumen.keys())
    for i, year in enumerate(años_ordenados):
        d = año_resumen[year]
        ws.append([year, round(d["total"], 2), d["contratos"], d["adjudicatarios"]])
        if i % 2 == 0:
            for col in range(1, 5):
                ws.cell(ws.max_row, col).fill = fill_claro
    fila_fin_años = ws.max_row

    ws.append([])

    # --- Tabla 2: Top 10 adjudicatarios histórico ---
    ws.append(["Top 10 adjudicatarios (histórico)"])
    ws.cell(ws.max_row, 1).font = font_titulo

    ws.append(["Empresa", "Total (€)", "Nº Contratos", "Años activos"])
    fila_cab_top = ws.max_row
    for col in range(1, 5):
        c = ws.cell(fila_cab_top, col)
        c.fill = fill_verde
        c.font = font_blanco
        c.alignment = center

    top_empresas = sorted(empresa_totals.items(), key=lambda x: x[1]["total"], reverse=True)[:10]
    for i, (empresa, d) in enumerate(top_empresas):
        años_str = ", ".join(str(y) for y in sorted(d["años"]))
        ws.append([empresa, round(d["total"], 2), d["contratos"], años_str])
        if i % 2 == 0:
            for col in range(1, 5):
                ws.cell(ws.max_row, col).fill = fill_claro
    fila_fin_top = ws.max_row

    for col, ancho in [(1, 40), (2, 16), (3, 16), (4, 20)]:
        ws.column_dimensions[get_column_letter(col)].width = ancho

    cats_años = Reference(ws, min_col=1, min_row=fila_cab_años + 1, max_row=fila_fin_años)

    # --- Gráfica 1: Gasto total por año ---
    c1 = BarChart()
    c1.type, c1.title, c1.style = "col", "Gasto total por año (€)", 10
    c1.width, c1.height = 15, 10
    c1.y_axis.title, c1.x_axis.title = "€", "Año"
    c1.add_data(Reference(ws, min_col=2, min_row=fila_cab_años, max_row=fila_fin_años), titles_from_data=True)
    c1.set_categories(cats_años)
    ws.add_chart(c1, "F2")

    # --- Gráfica 2: Nº contratos por año ---
    c2 = BarChart()
    c2.type, c2.title, c2.style = "col", "Nº de contratos adjudicados por año", 10
    c2.width, c2.height = 15, 10
    c2.y_axis.title, c2.x_axis.title = "Contratos", "Año"
    c2.add_data(Reference(ws, min_col=3, min_row=fila_cab_años, max_row=fila_fin_años), titles_from_data=True)
    c2.set_categories(cats_años)
    ws.add_chart(c2, "F22")

    # --- Gráfica 3: Nº adjudicatarios distintos por año ---
    c3 = BarChart()
    c3.type, c3.title, c3.style = "col", "Nº de adjudicatarios distintos por año", 10
    c3.width, c3.height = 15, 10
    c3.y_axis.title, c3.x_axis.title = "Adjudicatarios", "Año"
    c3.add_data(Reference(ws, min_col=4, min_row=fila_cab_años, max_row=fila_fin_años), titles_from_data=True)
    c3.set_categories(cats_años)
    ws.add_chart(c3, "V2")

    # --- Gráfica 4: Top 10 adjudicatarios (barras horizontales) ---
    c4 = BarChart()
    c4.type, c4.title, c4.style = "bar", "Top 10 adjudicatarios por importe total (€)", 10
    c4.width, c4.height = 20, 14
    c4.x_axis.title = "€"
    c4.add_data(Reference(ws, min_col=2, min_row=fila_cab_top, max_row=fila_fin_top), titles_from_data=True)
    c4.set_categories(Reference(ws, min_col=1, min_row=fila_cab_top + 1, max_row=fila_fin_top))
    ws.add_chart(c4, f"A{fila_fin_top + 3}")

    log.info(f"Tab 'Estadísticas' actualizado: {len(años_ordenados)} año(s), top {len(top_empresas)} adjudicatarios")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    OUTPUT_DIR.mkdir(exist_ok=True)
    wb = _open_or_create_wb()

    if sys.argv[1:]:
        # Modo rango: python scraper_xlsx.py 2012 2025  (o 2012-2025)
        años = parse_años(sys.argv[1:])
        hoy_año = date.today().year

        scrapeado = False
        for year in años:
            log.info(f"{'─' * 50}")
            log.info(f"Procesando año {year}...")
            f_hasta = "" if year >= hoy_año else f"31-12-{year}"
            ranking, total_global, paginas_con_error = scrape(
                URL_PERFIL, f"01-01-{year}", f_hasta
            )
            if not ranking:
                log.warning(f"Año {year}: sin datos publicados, se omite.")
                continue
            _escribir_tab_año(wb, str(year), ranking, total_global, paginas_con_error)
            log.info(f"Año {year}: {len(ranking)} adjudicatarios — {total_global:,.2f} €")
            scrapeado = True

        if not scrapeado:
            log.error("No se obtuvieron datos para ningún año solicitado.")
            sys.exit(1)
    else:
        # Modo año único (respeta FECHA_DESDE / FECHA_HASTA / SHEET_NAME del entorno)
        ranking, total_global, paginas_con_error = scrape(URL_PERFIL, FECHA_DESDE, FECHA_HASTA)
        if not ranking:
            log.error("No se obtuvo ningún dato. Abortando.")
            sys.exit(1)
        _escribir_tab_año(wb, YEAR, ranking, total_global, paginas_con_error)

    actualizar_registro_total(wb)
    actualizar_estadisticas(wb)
    _reordenar_tabs(wb)
    wb.save(NOMBRE_XLSX)
    log.info("Proceso completado.")
