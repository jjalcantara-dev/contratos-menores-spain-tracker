import os
import sys
import pytest
from unittest.mock import MagicMock, patch, call
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

# scraper.py requiere SPREADSHEET_ID al importar
os.environ.setdefault("SPREADSHEET_ID", "fake-id-for-tests")

import scraper

RANKING = [
    ("Empresa A", {"total": 5000.0,  "proyectos": [("Obra 1", 3000.0), ("Obra 2", 2000.0)]}),
    ("Empresa B", {"total": 3200.50, "proyectos": [("Servicio X", 3200.50)]}),
]
TOTAL_GLOBAL = 8200.50


class TestHexToRgb:
    def test_blanco(self):
        assert scraper._hex_to_rgb("FFFFFF") == {"red": 1.0, "green": 1.0, "blue": 1.0}

    def test_negro(self):
        assert scraper._hex_to_rgb("000000") == {"red": 0.0, "green": 0.0, "blue": 0.0}

    def test_verde_oscuro(self):
        resultado = scraper._hex_to_rgb("1F5C2E")
        assert round(resultado["red"],   3) == round(0x1F / 255, 3)
        assert round(resultado["green"], 3) == round(0x5C / 255, 3)
        assert round(resultado["blue"],  3) == round(0x2E / 255, 3)

    def test_con_hash(self):
        assert scraper._hex_to_rgb("#FFFFFF") == scraper._hex_to_rgb("FFFFFF")


class TestEscribirEnSheets:
    def _make_ws(self):
        ws = MagicMock()
        ws.id = 42
        ws.spreadsheet = MagicMock()
        return ws

    def test_limpia_hoja_antes_de_escribir(self):
        ws = self._make_ws()
        scraper.escribir_en_sheets(ws, RANKING, TOTAL_GLOBAL, [])
        ws.clear.assert_called_once()

    def test_llama_update(self):
        ws = self._make_ws()
        scraper.escribir_en_sheets(ws, RANKING, TOTAL_GLOBAL, [])
        ws.update.assert_called_once()

    def test_datos_contienen_empresas(self):
        ws = self._make_ws()
        scraper.escribir_en_sheets(ws, RANKING, TOTAL_GLOBAL, [])
        args = ws.update.call_args
        todas = args.kwargs.get("values") or args[1].get("values") or args[0][1]
        contenido = str(todas)
        assert "Empresa A" in contenido
        assert "Empresa B" in contenido

    def test_total_general_en_datos(self):
        ws = self._make_ws()
        scraper.escribir_en_sheets(ws, RANKING, TOTAL_GLOBAL, [])
        args = ws.update.call_args
        todas = args.kwargs.get("values") or args[1].get("values") or args[0][1]
        contenido = str(todas)
        assert "TOTAL GENERAL" in contenido
        assert str(round(TOTAL_GLOBAL, 2)) in contenido

    def test_aviso_paginas_error(self):
        ws = self._make_ws()
        scraper.escribir_en_sheets(ws, RANKING, TOTAL_GLOBAL, [2, 4])
        args = ws.update.call_args
        todas = args.kwargs.get("values") or args[1].get("values") or args[0][1]
        assert any("2" in str(cell) and "4" in str(cell) for fila in todas for cell in fila if cell)

    def test_batch_update_llamado(self):
        ws = self._make_ws()
        scraper.escribir_en_sheets(ws, RANKING, TOTAL_GLOBAL, [])
        ws.spreadsheet.batch_update.assert_called_once()
