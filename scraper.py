"""
Scraper automático de contratos menores - Vélez-Málaga
Diseñado para ejecutarse sin intervención manual (headless, sin input()).
Escribe directamente en Google Sheets, borrando antes los datos anteriores
para garantizar que nunca haya duplicados.
"""

import os
import sys
import json
import logging
from datetime import date

import gspread
from gspread.exceptions import WorksheetNotFound
from google.oauth2.service_account import Credentials

from scraper_core import scrape, URL_PERFIL, FECHA_DESDE, FECHA_HASTA, YEAR

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


def _hex_to_rgb(hex_color):
    h = hex_color.lstrip("#")
    return {
        "red":   int(h[0:2], 16) / 255,
        "green": int(h[2:4], 16) / 255,
        "blue":  int(h[4:6], 16) / 255,
    }


def escribir_en_sheets(ws, ranking, total_global, paginas_con_error):
    log.info("Borrando hoja y reescribiendo datos...")

    hoy = date.today().strftime("%d/%m/%Y")
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
    ws.update(range_name="A1", values=todas, value_input_option="USER_ENTERED")

    # --- Formato ---
    fila_cabecera  = 3   # fila 1=meta, 2=vacía, 3=cabecera
    fila_total_idx = 3 + len(filas_datos) + 2  # cabecera + datos + fila vacía + total

    verde_oscuro = _hex_to_rgb("1F5C2E")
    verde_claro  = _hex_to_rgb("E8F5E9")
    blanco       = {"red": 1, "green": 1, "blue": 1}

    fmt_cabecera = {
        "backgroundColor": verde_oscuro,
        "textFormat": {"bold": True, "foregroundColor": {"red": 1, "green": 1, "blue": 1}, "fontSize": 11},
        "horizontalAlignment": "CENTER",
    }
    fmt_total = {
        "backgroundColor": verde_oscuro,
        "textFormat": {"bold": True, "foregroundColor": {"red": 1, "green": 1, "blue": 1}},
    }

    sh       = ws.spreadsheet
    sheet_id = ws.id

    def cell_fmt(start_row, end_row, bg, bold=False, font_color=None, font_size=None):
        fmt = {"backgroundColor": bg, "textFormat": {"bold": bold}}
        if font_color:
            fmt["textFormat"]["foregroundColor"] = font_color
        if font_size:
            fmt["textFormat"]["fontSize"] = font_size
        return {
            "repeatCell": {
                "range": {"sheetId": sheet_id, "startRowIndex": start_row - 1, "endRowIndex": end_row, "startColumnIndex": 0, "endColumnIndex": 5},
                "cell": {"userEnteredFormat": fmt},
                "fields": "userEnteredFormat(backgroundColor,textFormat)",
            }
        }

    blanco_rgb     = {"red": 1, "green": 1, "blue": 1}
    requests = []

    # Cabecera
    requests.append(cell_fmt(fila_cabecera, fila_cabecera, verde_oscuro, bold=True, font_color=blanco_rgb, font_size=11))

    # Filas alternas
    for i, fila in enumerate(range(fila_cabecera + 1, fila_cabecera + 1 + len(filas_datos))):
        color = verde_claro if i % 2 == 0 else blanco
        requests.append(cell_fmt(fila, fila, color))

    # Fila total
    requests.append(cell_fmt(fila_total_idx, fila_total_idx, verde_oscuro, bold=True, font_color=blanco_rgb))

    # Anchos de columna
    for idx, px in [(0, 40), (1, 300), (2, 100), (3, 550), (4, 130)]:
        requests.append({"updateDimensionProperties": {
            "range": {"sheetId": sheet_id, "dimension": "COLUMNS", "startIndex": idx, "endIndex": idx + 1},
            "properties": {"pixelSize": px}, "fields": "pixelSize"
        }})

    # Congelar cabecera
    requests.append({"updateSheetProperties": {
        "properties": {"sheetId": sheet_id, "gridProperties": {"frozenRowCount": fila_cabecera}},
        "fields": "gridProperties.frozenRowCount"
    }})

    # Formato numérico columna E: 1.234,56 €
    requests.append({"repeatCell": {
        "range": {"sheetId": sheet_id, "startRowIndex": fila_cabecera, "endRowIndex": fila_total_idx, "startColumnIndex": 4, "endColumnIndex": 5},
        "cell": {"userEnteredFormat": {"numberFormat": {"type": "NUMBER", "pattern": '#,##0.00 "€"'}}},
        "fields": "userEnteredFormat.numberFormat"
    }})

    # Filtros en cabecera
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
# Main
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    ranking, total_global, paginas_con_error = scrape(URL_PERFIL, FECHA_DESDE, FECHA_HASTA)

    if not ranking:
        log.error("No se obtuvo ningún dato. Abortando escritura en Sheets.")
        sys.exit(1)

    ws = get_worksheet()
    escribir_en_sheets(ws, ranking, total_global, paginas_con_error)
    log.info("Proceso completado.")
