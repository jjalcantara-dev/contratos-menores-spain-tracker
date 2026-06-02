"""
Scraper automático de contratos menores - Vélez-Málaga
Diseñado para ejecutarse sin intervención manual (headless, sin input()).
Escribe directamente en Google Sheets, borrando antes los datos anteriores
para garantizar que nunca haya duplicados.

Uso:
    python scraper.py                 # año actual
    python scraper.py 2023            # año concreto
    python scraper.py 2012 2025       # rango completo
    python scraper.py 2012-2025       # misma sintaxis con guion
"""

import os
import re
import sys
import json
import logging
from datetime import date

import gspread
from gspread.exceptions import WorksheetNotFound
from google.oauth2.service_account import Credentials

from scraper_core import scrape, URL_PERFIL, FECHA_DESDE, FECHA_HASTA, YEAR, parse_años, agregar_por_año

# ---------------------------------------------------------------------------
# Configuración
# ---------------------------------------------------------------------------

SPREADSHEET_ID = os.environ.get("SPREADSHEET_ID")
if not SPREADSHEET_ID:
    raise EnvironmentError("Variable SPREADSHEET_ID no definida.")
SHEET_NAME = f"Contratos {YEAR}"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)
log = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Google Sheets — conexión
# ---------------------------------------------------------------------------

def get_spreadsheet():
    creds_raw = os.environ.get("GOOGLE_CREDENTIALS")
    if not creds_raw:
        raise EnvironmentError("Variable GOOGLE_CREDENTIALS no definida.")
    creds = Credentials.from_service_account_info(
        json.loads(creds_raw),
        scopes=[
            "https://spreadsheets.google.com/feeds",
            "https://www.googleapis.com/auth/drive",
        ],
    )
    gc = gspread.Client(auth=creds)
    return gc.open_by_key(SPREADSHEET_ID)


BLANCO = {"red": 1, "green": 1, "blue": 1}


def get_worksheet():
    sh = get_spreadsheet()
    try:
        return sh.worksheet(SHEET_NAME)
    except WorksheetNotFound:
        return sh.add_worksheet(title=SHEET_NAME, rows=1000, cols=10)


def _get_ws_año(sh, year):
    """Obtiene o crea el worksheet para un año."""
    name = f"Contratos {year}"
    try:
        return sh.worksheet(name)
    except WorksheetNotFound:
        return sh.add_worksheet(title=name, rows=1000, cols=10)

# ---------------------------------------------------------------------------
# Utilidades de color y formato
# ---------------------------------------------------------------------------

def _hex_to_rgb(hex_color):
    h = hex_color.lstrip("#")
    return {
        "red":   int(h[0:2], 16) / 255,
        "green": int(h[2:4], 16) / 255,
        "blue":  int(h[4:6], 16) / 255,
    }


def _cell_fmt(sheet_id, start_row, end_row, bg, bold=False, font_color=None,
              font_size=None, n_cols=5):
    fmt = {"backgroundColor": bg, "textFormat": {"bold": bold}}
    if font_color:
        fmt["textFormat"]["foregroundColor"] = font_color
    if font_size:
        fmt["textFormat"]["fontSize"] = font_size
    return {
        "repeatCell": {
            "range": {
                "sheetId":          sheet_id,
                "startRowIndex":    start_row - 1,
                "endRowIndex":      end_row,
                "startColumnIndex": 0,
                "endColumnIndex":   n_cols,
            },
            "cell":   {"userEnteredFormat": fmt},
            "fields": "userEnteredFormat(backgroundColor,textFormat)",
        }
    }


def _col_width(sheet_id, col_idx, px):
    return {
        "updateDimensionProperties": {
            "range": {
                "sheetId":    sheet_id,
                "dimension":  "COLUMNS",
                "startIndex": col_idx,
                "endIndex":   col_idx + 1,
            },
            "properties": {"pixelSize": px},
            "fields":     "pixelSize",
        }
    }


def _source_range(sheet_id, start_row, end_row, start_col, end_col):
    return {
        "sheetId":          sheet_id,
        "startRowIndex":    start_row,
        "endRowIndex":      end_row,
        "startColumnIndex": start_col,
        "endColumnIndex":   end_col,
    }


def _chart_request(title, chart_type, sheet_id, cat_range, series_range,
                   anchor_row, anchor_col, width=500, height=320):
    # BAR (horizontal) usa BOTTOM_AXIS; COLUMN/LINE usan LEFT_AXIS
    target_axis = "BOTTOM_AXIS" if chart_type == "BAR" else "LEFT_AXIS"
    return {
        "addChart": {
            "chart": {
                "spec": {
                    "title": title,
                    "basicChart": {
                        "chartType":      chart_type,
                        "legendPosition": "NO_LEGEND",
                        "domains": [{"domain": {"sourceRange": {"sources": [cat_range]}}}],
                        "series": [{
                            "series": {"sourceRange": {"sources": [series_range]}},
                            "targetAxis": target_axis,
                        }],
                        "headerCount": 1,
                    },
                },
                "position": {
                    "overlayPosition": {
                        "anchorCell": {
                            "sheetId":     sheet_id,
                            "rowIndex":    anchor_row,
                            "columnIndex": anchor_col,
                        },
                        "widthPixels":  width,
                        "heightPixels": height,
                    }
                },
            }
        }
    }

# ---------------------------------------------------------------------------
# Tab del año
# ---------------------------------------------------------------------------

def escribir_en_sheets(ws, ranking, total_global, paginas_con_error):
    log.info("Borrando hoja y reescribiendo datos...")

    hoy   = date.today().strftime("%d/%m/%Y")
    aviso = f"⚠ Páginas con error: {paginas_con_error}" if paginas_con_error else ""

    meta      = [["Última actualización:", hoy, f"Total adjudicatarios: {len(ranking)}", aviso]]
    cabecera  = [["#", "Adjudicatario", "Nº Contratos", "Proyectos", "Total (€)"]]

    filas_datos = []
    for i, (empresa, datos) in enumerate(ranking, 1):
        proyectos_str = " | ".join(
            f"{desc} ({valor:.2f} €)" for desc, valor in datos["proyectos"]
        )
        filas_datos.append([i, empresa, len(datos["proyectos"]), proyectos_str, round(datos["total"], 2)])

    fila_total = [["", "TOTAL GENERAL", "", "", round(total_global, 2)]]
    todas = meta + [[]] + cabecera + filas_datos + [[]] + fila_total

    ws.clear()
    ws.update(values=todas, range_name="A1", value_input_option="USER_ENTERED")

    fila_cabecera  = 3
    fila_total_idx = 3 + len(filas_datos) + 2

    verde_oscuro = _hex_to_rgb("1F5C2E")
    verde_claro  = _hex_to_rgb("E8F5E9")
    

    sh       = ws.spreadsheet
    sheet_id = ws.id

    requests = []

    requests.append(_cell_fmt(sheet_id, fila_cabecera, fila_cabecera,
                               verde_oscuro, bold=True, font_color=BLANCO, font_size=11))
    for i, fila in enumerate(range(fila_cabecera + 1, fila_cabecera + 1 + len(filas_datos))):
        color = verde_claro if i % 2 == 0 else BLANCO
        requests.append(_cell_fmt(sheet_id, fila, fila, color))
    requests.append(_cell_fmt(sheet_id, fila_total_idx, fila_total_idx,
                               verde_oscuro, bold=True, font_color=BLANCO))

    for idx, px in [(0, 40), (1, 300), (2, 100), (3, 550), (4, 130)]:
        requests.append(_col_width(sheet_id, idx, px))

    requests.append({"updateSheetProperties": {
        "properties": {"sheetId": sheet_id, "gridProperties": {"frozenRowCount": fila_cabecera}},
        "fields": "gridProperties.frozenRowCount",
    }})
    requests.append({"repeatCell": {
        "range": {
            "sheetId": sheet_id, "startRowIndex": fila_cabecera, "endRowIndex": fila_total_idx,
            "startColumnIndex": 4, "endColumnIndex": 5,
        },
        "cell": {"userEnteredFormat": {"numberFormat": {"type": "NUMBER", "pattern": '#,##0.00 "€"'}}},
        "fields": "userEnteredFormat.numberFormat",
    }})
    requests.append({"setBasicFilter": {
        "filter": {
            "range": {
                "sheetId": sheet_id,
                "startRowIndex": fila_cabecera - 1,
                "startColumnIndex": 0,
                "endColumnIndex": 5,
            }
        }
    }})

    sh.batch_update({"requests": requests})
    log.info(f"Hoja formateada y actualizada: {len(ranking)} adjudicatarios — {total_global:,.2f} €")

# ---------------------------------------------------------------------------
# Tab Registro Total
# ---------------------------------------------------------------------------

def _leer_datos_año_sheets(ws_year, year):
    """Lee filas de datos de un tab de año en Google Sheets."""
    # UNFORMATTED_VALUE evita recibir los totales con formato de celda ("88.501,34 €")
    # lo que haría fallar float() silenciosamente
    values = ws_year.get_all_values(value_render_option="UNFORMATTED_VALUE")
    filas = []
    for row in values[3:]:   # fila 1=meta, 2=vacía, 3=cabecera → datos desde índice 3
        empresa = row[1] if len(row) > 1 else ""
        if not empresa or empresa in ("Adjudicatario", "TOTAL GENERAL"):
            continue
        total_raw = row[4] if len(row) > 4 else ""
        try:
            total = float(total_raw) if total_raw != "" else 0.0
        except (ValueError, TypeError):
            continue
        try:
            num_contratos = int(row[2]) if len(row) > 2 and row[2] != "" else 0
        except (ValueError, TypeError):
            num_contratos = 0
        filas.append({
            "año":           year,
            "empresa":       empresa,
            "num_contratos": num_contratos,
            "proyectos":     row[3] if len(row) > 3 else "",
            "total":         total,
        })
    return filas


def _tabs_año_sheets(sh):
    """Devuelve [(year_int, worksheet), ...] para todos los tabs de año."""
    resultado = []
    for ws in sh.worksheets():
        m = re.match(r"^Contratos (\d{4})$", ws.title)
        if m:
            resultado.append((int(m.group(1)), ws))
    resultado.sort()
    return resultado


def escribir_registro_total(sh):
    """Crea/actualiza 'Registro Total' leyendo todos los tabs de año."""
    year_sheets = _tabs_año_sheets(sh)
    if not year_sheets:
        return

    try:
        ws_total = sh.worksheet("Registro Total")
        sh.del_worksheet(ws_total)
    except WorksheetNotFound:
        pass
    ws_total = sh.add_worksheet(title="Registro Total", rows=5000, cols=6)

    todas_filas = []
    for year, ws_year in year_sheets:
        todas_filas.extend(_leer_datos_año_sheets(ws_year, year))
    todas_filas.sort(key=lambda r: r["total"], reverse=True)

    hoy  = date.today().strftime("%d/%m/%Y")
    años = ", ".join(str(y) for y, _ in year_sheets)

    meta     = [[f"Última actualización: {hoy}", f"Total registros: {len(todas_filas)}", f"Años: {años}"]]
    cabecera = [["Año", "Adjudicatario", "Nº Contratos", "Proyectos", "Total (€)"]]
    filas_out = [
        [str(f["año"]), f["empresa"], f["num_contratos"], f["proyectos"], f["total"]]
        for f in todas_filas
    ]
    total_global = sum(f["total"] for f in todas_filas)
    fila_total   = [["", "TOTAL GENERAL", "", "", round(total_global, 2)]]

    todas = meta + [[]] + cabecera + filas_out + [[]] + fila_total
    ws_total.clear()
    ws_total.update(range_name="A1", values=todas, value_input_option="USER_ENTERED")

    sheet_id     = ws_total.id
    verde_oscuro = _hex_to_rgb("1F5C2E")
    verde_claro  = _hex_to_rgb("E8F5E9")
    
    fila_cab     = 3
    fila_tot_idx = 3 + len(todas_filas) + 2

    requests = []
    requests.append(_cell_fmt(sheet_id, fila_cab, fila_cab,
                               verde_oscuro, bold=True, font_color=BLANCO, font_size=11, n_cols=6))
    for i, fila in enumerate(range(fila_cab + 1, fila_cab + 1 + len(todas_filas))):
        color = verde_claro if i % 2 == 0 else BLANCO
        requests.append(_cell_fmt(sheet_id, fila, fila, color, n_cols=6))
    requests.append(_cell_fmt(sheet_id, fila_tot_idx, fila_tot_idx,
                               verde_oscuro, bold=True, font_color=BLANCO, n_cols=6))
    for idx, px in [(0, 60), (1, 300), (2, 100), (3, 550), (4, 130)]:
        requests.append(_col_width(sheet_id, idx, px))
    requests.append({"updateSheetProperties": {
        "properties": {"sheetId": sheet_id, "gridProperties": {"frozenRowCount": fila_cab}},
        "fields": "gridProperties.frozenRowCount",
    }})
    requests.append({"repeatCell": {
        "range": {
            "sheetId": sheet_id, "startRowIndex": fila_cab, "endRowIndex": fila_tot_idx,
            "startColumnIndex": 4, "endColumnIndex": 5,
        },
        "cell": {"userEnteredFormat": {"numberFormat": {"type": "NUMBER", "pattern": '#,##0.00 "€"'}}},
        "fields": "userEnteredFormat.numberFormat",
    }})
    sh.batch_update({"requests": requests})
    log.info(f"Tab 'Registro Total' actualizado: {len(todas_filas)} registros de {len(year_sheets)} año(s)")

# ---------------------------------------------------------------------------
# Tab Estadísticas (tablas resumen + gráficas)
# ---------------------------------------------------------------------------

def _reordenar_tabs_sheets(sh):
    """Ordena: año_actual → Estadísticas → años anteriores (desc) → Registro Total"""
    all_ws   = sh.worksheets()
    ws_map   = {ws.title: ws for ws in all_ws}
    hoy_año  = str(date.today().year)
    tab_actual = f"Contratos {hoy_año}"

    year_titles = sorted(
        [t for t in ws_map if re.match(r"^Contratos \d{4}$", t)],
        reverse=True,
    )

    desired = []
    if tab_actual in year_titles:
        desired.append(tab_actual)
        year_titles = [t for t in year_titles if t != tab_actual]
    elif year_titles:
        desired.append(year_titles[0])
        year_titles = year_titles[1:]

    if "Estadísticas" in ws_map:
        desired.append("Estadísticas")
    desired.extend(year_titles)
    if "Registro Total" in ws_map:
        desired.append("Registro Total")

    ordered = [ws_map[t] for t in desired if t in ws_map]
    ordered += [ws for ws in all_ws if ws.title not in set(desired)]
    sh.reorder_worksheets(ordered)


def escribir_estadisticas(sh):
    """Crea/actualiza 'Estadísticas' con resúmenes y gráficas."""
    year_sheets = _tabs_año_sheets(sh)
    if not year_sheets:
        return

    todas_filas = []
    for year, ws_year in year_sheets:
        todas_filas.extend(_leer_datos_año_sheets(ws_year, year))

    año_resumen, empresa_totals = agregar_por_año(todas_filas)

    try:
        ws_est = sh.worksheet("Estadísticas")
        sh.del_worksheet(ws_est)
    except WorksheetNotFound:
        pass
    ws_est = sh.add_worksheet(title="Estadísticas", rows=200, cols=20)

    años_ordenados = sorted(año_resumen.keys())
    n_años         = len(años_ordenados)
    top_empresas   = sorted(empresa_totals.items(), key=lambda x: x[1]["total"], reverse=True)[:10]
    n_top          = len(top_empresas)

    fila_t1_cab        = 2
    fila_t1_data_end   = 2 + n_años
    fila_t2_cab        = fila_t1_data_end + 3
    fila_t2_data_end   = fila_t1_data_end + 3 + n_top

    tabla1_header = [["Resumen por año"]]
    tabla1_cab    = [["Año", "Total (€)", "Nº Contratos", "Nº Adjudicatarios"]]
    tabla1_datos  = [
        [year, round(d["total"], 2), d["contratos"], d["adjudicatarios"]]
        for year, d in [(y, año_resumen[y]) for y in años_ordenados]
    ]
    tabla2_header = [["Top 10 adjudicatarios (histórico)"]]
    tabla2_cab    = [["Empresa", "Total (€)", "Nº Contratos", "Años activos"]]
    tabla2_datos  = [
        [empresa, round(d["total"], 2), d["contratos"],
         ", ".join(str(y) for y in sorted(d["años"]))]
        for empresa, d in top_empresas
    ]

    todas = (
        tabla1_header + tabla1_cab + tabla1_datos +
        [[]] +
        tabla2_header + tabla2_cab + tabla2_datos
    )
    ws_est.clear()
    ws_est.update(range_name="A1", values=todas, value_input_option="USER_ENTERED")

    sheet_id     = ws_est.id
    verde_oscuro = _hex_to_rgb("1F5C2E")
    verde_claro  = _hex_to_rgb("E8F5E9")
    

    requests = []
    requests.append(_cell_fmt(sheet_id, fila_t1_cab, fila_t1_cab,
                               verde_oscuro, bold=True, font_color=BLANCO, n_cols=4))
    requests.append(_cell_fmt(sheet_id, fila_t2_cab, fila_t2_cab,
                               verde_oscuro, bold=True, font_color=BLANCO, n_cols=4))
    for i, fila in enumerate(range(fila_t1_cab + 1, fila_t1_data_end + 1)):
        requests.append(_cell_fmt(sheet_id, fila, fila,
                                   verde_claro if i % 2 == 0 else BLANCO, n_cols=4))
    for i, fila in enumerate(range(fila_t2_cab + 1, fila_t2_data_end + 1)):
        requests.append(_cell_fmt(sheet_id, fila, fila,
                                   verde_claro if i % 2 == 0 else BLANCO, n_cols=4))
    for idx, px in [(0, 250), (1, 130), (2, 100), (3, 150)]:
        requests.append(_col_width(sheet_id, idx, px))

    t1_cab_0      = fila_t1_cab - 1
    t1_data_end_0 = fila_t1_data_end
    t2_cab_0      = fila_t2_cab - 1
    t2_data_end_0 = fila_t2_data_end

    cat1   = _source_range(sheet_id, t1_cab_0, t1_data_end_0, 0, 1)
    ser_b  = _source_range(sheet_id, t1_cab_0, t1_data_end_0, 1, 2)
    ser_c  = _source_range(sheet_id, t1_cab_0, t1_data_end_0, 2, 3)
    ser_d  = _source_range(sheet_id, t1_cab_0, t1_data_end_0, 3, 4)
    cat2   = _source_range(sheet_id, t2_cab_0, t2_data_end_0, 0, 1)
    ser_t2 = _source_range(sheet_id, t2_cab_0, t2_data_end_0, 1, 2)

    requests.append(_chart_request("Gasto total por año (€)", "COLUMN",
                                   sheet_id, cat1, ser_b, 0, 5, 500, 320))
    requests.append(_chart_request("Nº contratos adjudicados por año", "COLUMN",
                                   sheet_id, cat1, ser_c, 22, 5, 500, 320))
    requests.append(_chart_request("Nº adjudicatarios distintos por año", "COLUMN",
                                   sheet_id, cat1, ser_d, 44, 5, 500, 320))
    requests.append(_chart_request("Top 10 adjudicatarios por importe total (€)", "BAR",
                                   sheet_id, cat2, ser_t2, 66, 5, 600, 420))

    sh.batch_update({"requests": requests})
    log.info(f"Tab 'Estadísticas' actualizado: {n_años} año(s), top {n_top} adjudicatarios, 4 gráficas")

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    sh = get_spreadsheet()

    if sys.argv[1:]:
        # Modo rango: python scraper.py 2012 2025  (o 2012-2025)
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
            ws = _get_ws_año(sh, year)
            escribir_en_sheets(ws, ranking, total_global, paginas_con_error)
            scrapeado = True

        if not scrapeado:
            log.error("No se obtuvieron datos para ningún año solicitado.")
            sys.exit(1)
    else:
        # Modo año actual (comportamiento original, respeta vars de entorno)
        try:
            ws = sh.worksheet(SHEET_NAME)
        except WorksheetNotFound:
            ws = sh.add_worksheet(title=SHEET_NAME, rows=1000, cols=10)

        ranking, total_global, paginas_con_error = scrape(URL_PERFIL, FECHA_DESDE, FECHA_HASTA)
        if not ranking:
            log.error("No se obtuvo ningún dato. Abortando escritura en Sheets.")
            sys.exit(1)
        escribir_en_sheets(ws, ranking, total_global, paginas_con_error)

    escribir_registro_total(sh)
    escribir_estadisticas(sh)
    _reordenar_tabs_sheets(sh)
    log.info("Proceso completado.")
