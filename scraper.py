"""
Scraper automático de contratos menores - Vélez-Málaga
Escribe directamente en Google Sheets (sin intervención manual).

Uso:
    python scraper.py                 # año actual
    python scraper.py 2023            # año concreto
    python scraper.py 2012 2025       # rango completo
    python scraper.py 2012-2025       # misma sintaxis con guion
"""

from __future__ import annotations

import os
import sys
import json
import logging
from datetime import date

import gspread
from gspread.exceptions import WorksheetNotFound
from google.oauth2.service_account import Credentials

from scraper_core import (
    scrape, URL_PERFIL, FECHA_DESDE, FECHA_HASTA, YEAR,
    parse_años, preparar_estadisticas, normalizar_fila,
    TAB_ESTADISTICAS, TAB_REGISTRO_TOTAL, PREFIJO_CONTRATO, PATRON_TAB_AÑO,
    FilaContrato, DatosEstadisticas,
)

# ---------------------------------------------------------------------------
# Configuración
# ---------------------------------------------------------------------------

SPREADSHEET_ID = os.environ.get("SPREADSHEET_ID")
if not SPREADSHEET_ID:
    raise EnvironmentError("Variable SPREADSHEET_ID no definida.")
SHEET_NAME = f"{PREFIJO_CONTRATO} {YEAR}"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)
log = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Constantes de presentación (calculadas una sola vez)
# ---------------------------------------------------------------------------

def _hex_to_rgb(hex_color: str) -> dict:
    h = hex_color.lstrip("#")
    return {
        "red":   int(h[0:2], 16) / 255,
        "green": int(h[2:4], 16) / 255,
        "blue":  int(h[4:6], 16) / 255,
    }


BLANCO           = {"red": 1.0, "green": 1.0, "blue": 1.0}
VERDE_OSCURO_RGB = _hex_to_rgb("1F5C2E")
VERDE_CLARO_RGB  = _hex_to_rgb("E8F5E9")

# Anchos de columna (píxeles) para cada tipo de pestaña
COL_PX_AÑO   = [(0, 40),  (1, 300), (2, 100), (3, 550), (4, 130)]
COL_PX_TOTAL = [(0, 60),  (1, 300), (2, 100), (3, 550), (4, 130)]
COL_PX_EST   = [(0, 250), (1, 130), (2, 100), (3, 150)]

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


def get_worksheet():
    sh = get_spreadsheet()
    try:
        return sh.worksheet(SHEET_NAME)
    except WorksheetNotFound:
        return sh.add_worksheet(title=SHEET_NAME, rows=1000, cols=10)


def _get_ws_año(sh, year: int):
    name = f"{PREFIJO_CONTRATO} {year}"
    try:
        return sh.worksheet(name)
    except WorksheetNotFound:
        return sh.add_worksheet(title=name, rows=1000, cols=10)


def _borrar_y_crear_hoja(sh, titulo: str, rows: int = 1000, cols: int = 10):
    """Elimina la hoja si existe y la recrea vacía."""
    try:
        sh.del_worksheet(sh.worksheet(titulo))
    except WorksheetNotFound:
        pass
    return sh.add_worksheet(title=titulo, rows=rows, cols=cols)

# ---------------------------------------------------------------------------
# Helpers de la Sheets API
# ---------------------------------------------------------------------------

def _cell_fmt(
    sheet_id: int,
    start_row: int,
    end_row: int,
    bg: dict,
    bold: bool = False,
    font_color: dict | None = None,
    font_size: int | None = None,
    n_cols: int = 5,
) -> dict:
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


def _col_width(sheet_id: int, col_idx: int, px: int) -> dict:
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


def _request_banding(sheet_id: int, fila_cab: int, fila_fin_datos: int, n_cols: int) -> dict:
    """Colores alternos mediante una sola request (O(1)) en lugar de O(n) repeatCell.
    Usar cuando la tabla tiene cientos/miles de filas.
    """
    return {
        "addBanding": {
            "bandedRange": {
                "range": {
                    "sheetId":          sheet_id,
                    "startRowIndex":    fila_cab - 1,    # 0-based, incluye cabecera
                    "endRowIndex":      fila_fin_datos,  # exclusivo, hasta última fila de datos
                    "startColumnIndex": 0,
                    "endColumnIndex":   n_cols,
                },
                "rowProperties": {
                    "headerColor":     VERDE_OSCURO_RGB,
                    "firstBandColor":  BLANCO,
                    "secondBandColor": VERDE_CLARO_RGB,
                },
            }
        }
    }


def _source_range(sheet_id: int, r0: int, r1: int, c0: int, c1: int) -> dict:
    return {
        "sheetId":          sheet_id,
        "startRowIndex":    r0,
        "endRowIndex":      r1,
        "startColumnIndex": c0,
        "endColumnIndex":   c1,
    }


def _chart_request(
    title: str,
    chart_type: str,
    sheet_id: int,
    cat_range: dict,
    series_range: dict,
    anchor_row: int,
    anchor_col: int,
    width: int = 500,
    height: int = 320,
) -> dict:
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
                            "series":     {"sourceRange": {"sources": [series_range]}},
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


def _requests_tabla_coloreada(
    sheet_id: int,
    fila_cab: int,
    n_filas: int,
    fila_total: int,
    n_cols: int = 5,
) -> list[dict]:
    """
    Genera requests de color para: cabecera verde + filas alternas + total verde.
    Reutilizable en todas las pestañas con ese patrón.
    """
    return [
        _cell_fmt(sheet_id, fila_cab, fila_cab,
                  VERDE_OSCURO_RGB, bold=True, font_color=BLANCO, font_size=11, n_cols=n_cols),
        *[
            _cell_fmt(sheet_id, fila, fila,
                      VERDE_CLARO_RGB if i % 2 == 0 else BLANCO, n_cols=n_cols)
            for i, fila in enumerate(range(fila_cab + 1, fila_cab + 1 + n_filas))
        ],
        _cell_fmt(sheet_id, fila_total, fila_total,
                  VERDE_OSCURO_RGB, bold=True, font_color=BLANCO, n_cols=n_cols),
    ]

# ---------------------------------------------------------------------------
# Tab del año
# ---------------------------------------------------------------------------

def escribir_en_sheets(ws, ranking: list, total_global: float, paginas_con_error: list) -> None:
    log.info("Borrando hoja y reescribiendo datos...")

    hoy   = date.today().strftime("%d/%m/%Y")
    aviso = f"⚠ Páginas con error: {paginas_con_error}" if paginas_con_error else ""

    filas_datos = [
        [i, empresa, len(datos["proyectos"]),
         " | ".join(f"{d} ({v:.2f} €)" for d, v in datos["proyectos"]),
         round(datos["total"], 2)]
        for i, (empresa, datos) in enumerate(ranking, 1)
    ]
    fila_cabecera  = 3
    fila_total_idx = 3 + len(filas_datos) + 2

    todas = (
        [["Última actualización:", hoy, f"Total adjudicatarios: {len(ranking)}", aviso]] +
        [[]] +
        [["#", "Adjudicatario", "Nº Contratos", "Proyectos", "Total (€)"]] +
        filas_datos +
        [[]] +
        [["", "TOTAL GENERAL", "", "", round(total_global, 2)]]
    )
    ws.clear()
    ws.update(values=todas, range_name="A1", value_input_option="USER_ENTERED")

    sh       = ws.spreadsheet
    sheet_id = ws.id

    requests = [
        *_requests_tabla_coloreada(sheet_id, fila_cabecera, len(filas_datos), fila_total_idx),
        *[_col_width(sheet_id, idx, px) for idx, px in COL_PX_AÑO],
        {"updateSheetProperties": {
            "properties": {"sheetId": sheet_id, "gridProperties": {"frozenRowCount": fila_cabecera}},
            "fields": "gridProperties.frozenRowCount",
        }},
        {"repeatCell": {
            "range": {
                "sheetId": sheet_id, "startRowIndex": fila_cabecera, "endRowIndex": fila_total_idx,
                "startColumnIndex": 4, "endColumnIndex": 5,
            },
            "cell": {"userEnteredFormat": {"numberFormat": {"type": "NUMBER", "pattern": '#,##0.00 "€"'}}},
            "fields": "userEnteredFormat.numberFormat",
        }},
        {"setBasicFilter": {
            "filter": {
                "range": {
                    "sheetId":          sheet_id,
                    "startRowIndex":    fila_cabecera - 1,
                    "startColumnIndex": 0,
                    "endColumnIndex":   5,
                }
            }
        }},
    ]
    sh.batch_update({"requests": requests})
    log.info(f"Hoja formateada y actualizada: {len(ranking)} adjudicatarios — {total_global:,.2f} €")

# ---------------------------------------------------------------------------
# Tab Registro Total
# ---------------------------------------------------------------------------

def _leer_datos_año_sheets(ws_year, year: int) -> list[FilaContrato]:
    """Lee filas de datos de un tab de año en Google Sheets."""
    # UNFORMATTED_VALUE evita recibir totales con formato de celda ("88.501,34 €")
    values = ws_year.get_all_values(value_render_option="UNFORMATTED_VALUE")
    return [
        fila for row in values[3:]  # fila 1=meta, 2=vacía, 3=cabecera → datos desde índice 3
        if (fila := normalizar_fila(row, year)) is not None
    ]


def _tabs_año_sheets(sh) -> list[tuple[int, object]]:
    resultado = []
    for ws in sh.worksheets():
        m = PATRON_TAB_AÑO.match(ws.title)
        if m:
            resultado.append((int(m.group(1)), ws))
    resultado.sort()
    return resultado


def escribir_registro_total(sh) -> None:
    """Crea/actualiza 'Registro Total' leyendo todos los tabs de año."""
    year_sheets = _tabs_año_sheets(sh)
    if not year_sheets:
        return

    ws_total = _borrar_y_crear_hoja(sh, TAB_REGISTRO_TOTAL, rows=5000, cols=6)

    todas_filas: list[FilaContrato] = []
    for year, ws_year in year_sheets:
        todas_filas.extend(_leer_datos_año_sheets(ws_year, year))
    todas_filas.sort(key=lambda r: r["total"], reverse=True)

    hoy  = date.today().strftime("%d/%m/%Y")
    años = ", ".join(str(y) for y, _ in year_sheets)
    fila_cab     = 3
    fila_tot_idx = 3 + len(todas_filas) + 2

    todas = (
        [[f"Última actualización: {hoy}", f"Total registros: {len(todas_filas)}", f"Años: {años}"]] +
        [[]] +
        [["Año", "Adjudicatario", "Nº Contratos", "Proyectos", "Total (€)"]] +
        [[str(f["año"]), f["empresa"], f["num_contratos"], f["proyectos"], f["total"]]
         for f in todas_filas] +
        [[]] +
        [["", "TOTAL GENERAL", "", "", round(sum(f["total"] for f in todas_filas), 2)]]
    )
    ws_total.clear()
    ws_total.update(values=todas, range_name="A1", value_input_option="USER_ENTERED")

    sheet_id = ws_total.id
    # Banding (1 request) en lugar de O(n) repeatCell — crítico con miles de filas
    requests = [
        _request_banding(sheet_id, fila_cab, fila_cab + len(todas_filas), n_cols=6),
        _cell_fmt(sheet_id, fila_cab,     fila_cab,     VERDE_OSCURO_RGB, bold=True, font_color=BLANCO, font_size=11, n_cols=6),
        _cell_fmt(sheet_id, fila_tot_idx, fila_tot_idx, VERDE_OSCURO_RGB, bold=True, font_color=BLANCO, n_cols=6),
        *[_col_width(sheet_id, idx, px) for idx, px in COL_PX_TOTAL],
        {"updateSheetProperties": {
            "properties": {"sheetId": sheet_id, "gridProperties": {"frozenRowCount": fila_cab}},
            "fields": "gridProperties.frozenRowCount",
        }},
        {"repeatCell": {
            "range": {
                "sheetId": sheet_id, "startRowIndex": fila_cab, "endRowIndex": fila_tot_idx,
                "startColumnIndex": 4, "endColumnIndex": 5,
            },
            "cell": {"userEnteredFormat": {"numberFormat": {"type": "NUMBER", "pattern": '#,##0.00 "€"'}}},
            "fields": "userEnteredFormat.numberFormat",
        }},
        {"setBasicFilter": {
            "filter": {
                "range": {
                    "sheetId":          sheet_id,
                    "startRowIndex":    fila_cab - 1,
                    "startColumnIndex": 0,
                    "endColumnIndex":   5,
                }
            }
        }},
    ]
    sh.batch_update({"requests": requests})
    log.info(f"Tab '{TAB_REGISTRO_TOTAL}' actualizado: {len(todas_filas)} registros de {len(year_sheets)} año(s)")

# ---------------------------------------------------------------------------
# Tab Estadísticas — construcción, formato y gráficas separados
# ---------------------------------------------------------------------------

def _requests_formato_estadisticas(
    sheet_id: int,
    fila_t1_cab: int,
    fila_t1_data_end: int,
    fila_t2_cab: int,
    fila_t2_data_end: int,
) -> list[dict]:
    """Requests de color, anchos y formato euro para las dos tablas de Estadísticas."""
    reqs: list[dict] = [
        _cell_fmt(sheet_id, fila_t1_cab, fila_t1_cab, VERDE_OSCURO_RGB, bold=True, font_color=BLANCO, n_cols=4),
        _cell_fmt(sheet_id, fila_t2_cab, fila_t2_cab, VERDE_OSCURO_RGB, bold=True, font_color=BLANCO, n_cols=4),
        *[
            _cell_fmt(sheet_id, fila, fila, VERDE_CLARO_RGB if i % 2 == 0 else BLANCO, n_cols=4)
            for i, fila in enumerate(range(fila_t1_cab + 1, fila_t1_data_end + 1))
        ],
        *[
            _cell_fmt(sheet_id, fila, fila, VERDE_CLARO_RGB if i % 2 == 0 else BLANCO, n_cols=4)
            for i, fila in enumerate(range(fila_t2_cab + 1, fila_t2_data_end + 1))
        ],
        *[_col_width(sheet_id, idx, px) for idx, px in COL_PX_EST],
    ]
    for start, end in [(fila_t1_cab, fila_t1_data_end), (fila_t2_cab, fila_t2_data_end)]:
        reqs.append({"repeatCell": {
            "range": {
                "sheetId": sheet_id, "startRowIndex": start, "endRowIndex": end,
                "startColumnIndex": 1, "endColumnIndex": 2,
            },
            "cell": {"userEnteredFormat": {"numberFormat": {"type": "NUMBER", "pattern": '#,##0.00 "€"'}}},
            "fields": "userEnteredFormat.numberFormat",
        }})
    return reqs


def _añadir_graficas_sheets(
    requests: list[dict],
    sheet_id: int,
    t1_cab_0: int,
    t1_end_0: int,
    t2_cab_0: int,
    t2_end_0: int,
) -> None:
    """Añade los 4 requests addChart a la lista existente."""
    cat1  = _source_range(sheet_id, t1_cab_0, t1_end_0, 0, 1)
    ser_b = _source_range(sheet_id, t1_cab_0, t1_end_0, 1, 2)
    ser_c = _source_range(sheet_id, t1_cab_0, t1_end_0, 2, 3)
    ser_d = _source_range(sheet_id, t1_cab_0, t1_end_0, 3, 4)
    cat2  = _source_range(sheet_id, t2_cab_0, t2_end_0, 0, 1)
    ser_t = _source_range(sheet_id, t2_cab_0, t2_end_0, 1, 2)

    for title, chart_type, cat, ser, anchor_row, w, h in [
        ("Gasto total por año (€)",              "COLUMN", cat1, ser_b, 0,  500, 320),
        ("Nº contratos adjudicados por año",      "COLUMN", cat1, ser_c, 22, 500, 320),
        ("Nº adjudicatarios distintos por año",   "COLUMN", cat1, ser_d, 44, 500, 320),
        ("Top 10 adjudicatarios por importe (€)", "BAR",    cat2, ser_t, 66, 600, 420),
    ]:
        requests.append(_chart_request(title, chart_type, sheet_id, cat, ser, anchor_row, 5, w, h))


def escribir_estadisticas(sh) -> None:
    """Crea/actualiza 'Estadísticas' con resúmenes y gráficas."""
    year_sheets = _tabs_año_sheets(sh)
    if not year_sheets:
        return

    todas_filas: list[FilaContrato] = []
    for year, ws_year in year_sheets:
        todas_filas.extend(_leer_datos_año_sheets(ws_year, year))
    datos = preparar_estadisticas(todas_filas)

    ws_est = _borrar_y_crear_hoja(sh, TAB_ESTADISTICAS, rows=200, cols=20)

    n_años, n_top    = datos.n_años, datos.n_top
    fila_t1_cab      = 2
    fila_t1_data_end = 2 + n_años
    fila_t2_cab      = fila_t1_data_end + 3
    fila_t2_data_end = fila_t1_data_end + 3 + n_top

    todas = (
        [["Resumen por año"]] +
        [["Año", "Total (€)", "Nº Contratos", "Nº Adjudicatarios"]] +
        [
            [y, round(datos.año_resumen[y]["total"], 2),
             datos.año_resumen[y]["contratos"],
             datos.año_resumen[y]["adjudicatarios"]]
            for y in datos.años_ordenados
        ] +
        [[]] +
        [["Top 10 adjudicatarios (histórico)"]] +
        [["Empresa", "Total (€)", "Nº Contratos", "Años activos"]] +
        [
            [empresa, round(d["total"], 2), d["contratos"],
             ", ".join(str(y) for y in sorted(d["años"]))]
            for empresa, d in datos.top_empresas
        ]
    )
    ws_est.clear()
    ws_est.update(values=todas, range_name="A1", value_input_option="USER_ENTERED")

    sheet_id = ws_est.id
    requests = _requests_formato_estadisticas(
        sheet_id, fila_t1_cab, fila_t1_data_end, fila_t2_cab, fila_t2_data_end
    )
    _añadir_graficas_sheets(
        requests, sheet_id,
        fila_t1_cab - 1, fila_t1_data_end,
        fila_t2_cab - 1, fila_t2_data_end,
    )
    sh.batch_update({"requests": requests})
    log.info(f"Tab '{TAB_ESTADISTICAS}' actualizado: {n_años} año(s), top {n_top} adjudicatarios, 4 gráficas")

# ---------------------------------------------------------------------------
# Ordenación de pestañas
# ---------------------------------------------------------------------------

def _reordenar_tabs_sheets(sh) -> None:
    """Ordena: Estadísticas → Registro Total → años desc (2026, 2025, ...)"""
    all_ws = sh.worksheets()
    ws_map = {ws.title: ws for ws in all_ws}

    year_titles = sorted(
        [t for t in ws_map if PATRON_TAB_AÑO.match(t)],
        reverse=True,
    )
    desired = [t for t in [TAB_ESTADISTICAS, TAB_REGISTRO_TOTAL] if t in ws_map]
    desired.extend(year_titles)

    ordered  = [ws_map[t] for t in desired if t in ws_map]
    ordered += [ws for ws in all_ws if ws.title not in set(desired)]
    sh.reorder_worksheets(ordered)

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    sh = get_spreadsheet()

    if sys.argv[1:]:
        años    = parse_años(sys.argv[1:])
        hoy_año = date.today().year

        scrapeado = False
        for year in años:
            log.info(f"{'─' * 50}")
            log.info(f"Procesando año {year}...")
            f_hasta = "" if year >= hoy_año else f"31-12-{year}"
            resultado = scrape(URL_PERFIL, f"01-01-{year}", f_hasta)
            if not resultado.ranking:
                log.warning(f"Año {year}: sin datos publicados, se omite.")
                continue
            ws = _get_ws_año(sh, year)
            escribir_en_sheets(ws, resultado.ranking, resultado.total_global, resultado.paginas_con_error)
            scrapeado = True

        if not scrapeado:
            log.error("No se obtuvieron datos para ningún año solicitado.")
            sys.exit(1)
    else:
        try:
            ws = sh.worksheet(SHEET_NAME)
        except WorksheetNotFound:
            ws = sh.add_worksheet(title=SHEET_NAME, rows=1000, cols=10)

        resultado = scrape(URL_PERFIL, FECHA_DESDE, FECHA_HASTA)
        if not resultado.ranking:
            log.error("No se obtuvo ningún dato. Abortando escritura en Sheets.")
            sys.exit(1)
        escribir_en_sheets(ws, resultado.ranking, resultado.total_global, resultado.paginas_con_error)

    escribir_registro_total(sh)
    escribir_estadisticas(sh)
    _reordenar_tabs_sheets(sh)
    log.info("Proceso completado.")
