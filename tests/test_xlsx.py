import os
import sys
import pytest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

import scraper_xlsx
from openpyxl import load_workbook

RANKING = [
    ("Empresa A", {"total": 5000.0,  "proyectos": [("Obra 1", 3000.0), ("Obra 2", 2000.0)]}),
    ("Empresa B", {"total": 3200.50, "proyectos": [("Servicio X", 3200.50)]}),
    ("Empresa C", {"total": 800.0,   "proyectos": [("Suministro Z", 800.0)]}),
]
TOTAL_GLOBAL = 9000.50


@pytest.fixture(autouse=True)
def limpiar_archivo(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    monkeypatch.setattr(scraper_xlsx, "YEAR", "2025")
    yield
    # tmp_path se limpia automáticamente


def test_archivo_creado():
    scraper_xlsx.escribir_xlsx(RANKING, TOTAL_GLOBAL, [])
    assert Path("output/contratos_2025.xlsx").exists()


def test_titulo_hoja():
    scraper_xlsx.escribir_xlsx(RANKING, TOTAL_GLOBAL, [])
    wb = load_workbook("output/contratos_2025.xlsx")
    assert "Contratos 2025" in wb.sheetnames


def test_cabecera_correcta():
    scraper_xlsx.escribir_xlsx(RANKING, TOTAL_GLOBAL, [])
    wb = load_workbook("output/contratos_2025.xlsx")
    ws = wb.active
    # Fila 3 es la cabecera (1=meta, 2=vacía, 3=cabecera)
    cabecera = [ws.cell(3, col).value for col in range(1, 6)]
    assert cabecera == ["#", "Adjudicatario", "Nº Contratos", "Proyectos", "Total (€)"]


def test_numero_filas_datos():
    scraper_xlsx.escribir_xlsx(RANKING, TOTAL_GLOBAL, [])
    wb = load_workbook("output/contratos_2025.xlsx")
    ws = wb.active
    # Filas 4,5,6 son los datos
    empresas = [ws.cell(3 + i, 2).value for i in range(1, 4)]
    assert empresas == ["Empresa A", "Empresa B", "Empresa C"]


def test_orden_ranking():
    scraper_xlsx.escribir_xlsx(RANKING, TOTAL_GLOBAL, [])
    wb = load_workbook("output/contratos_2025.xlsx")
    ws = wb.active
    totales = [ws.cell(3 + i, 5).value for i in range(1, 4)]
    assert totales == sorted(totales, reverse=True)


def test_fila_total():
    scraper_xlsx.escribir_xlsx(RANKING, TOTAL_GLOBAL, [])
    wb = load_workbook("output/contratos_2025.xlsx")
    ws = wb.active
    # Última fila con valor
    ultima_fila = ws.max_row
    assert ws.cell(ultima_fila, 2).value == "TOTAL GENERAL"
    assert ws.cell(ultima_fila, 5).value == round(TOTAL_GLOBAL, 2)


def test_numero_contratos_por_empresa():
    scraper_xlsx.escribir_xlsx(RANKING, TOTAL_GLOBAL, [])
    wb = load_workbook("output/contratos_2025.xlsx")
    ws = wb.active
    # Empresa A tiene 2 contratos
    assert ws.cell(4, 3).value == 2
    # Empresa B tiene 1 contrato
    assert ws.cell(5, 3).value == 1


def test_sin_paginas_error_no_aviso():
    scraper_xlsx.escribir_xlsx(RANKING, TOTAL_GLOBAL, [])
    wb = load_workbook("output/contratos_2025.xlsx")
    ws = wb.active
    assert not ws.cell(1, 3).value


def test_con_paginas_error_muestra_aviso():
    scraper_xlsx.escribir_xlsx(RANKING, TOTAL_GLOBAL, [3, 5])
    wb = load_workbook("output/contratos_2025.xlsx")
    ws = wb.active
    assert "3" in str(ws.cell(1, 3).value)
    assert "5" in str(ws.cell(1, 3).value)
