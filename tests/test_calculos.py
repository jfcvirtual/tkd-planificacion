"""Pruebas: los resultados deben coincidir con las fórmulas del libro de Excel."""
import datetime as dt
import io
import sys
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import openpyxl
import pytest

from motor.calculos import (Deportista, Parametros, calcular_plan, canilleras, fc_maxima,
                            indice_sobreentrenamiento, mapa_semanas, vpe)

CIEN = dict(cross=90, escalera=70, patadas=200, saco=150, canillera=120,
            combate_condicionado=16, combate_libre=12)
DEP = Deportista("Prueba", 27, "M", 60, 64, 60, CIEN)
INICIO = dt.date(2026, 11, 23)


def plan(ciclo):
    return calcular_plan(Parametros(ciclo, INICIO, DEP))


def test_formulas_basicas():
    assert fc_maxima(27, "M") == 193
    assert fc_maxima(27, "F") == 199
    assert vpe(70, 193, 64) == pytest.approx(154.3)
    assert canilleras(60)[2] == pytest.approx(150)
    assert indice_sobreentrenamiento(60, 76) == (16, True)
    assert indice_sobreentrenamiento(60, 75) == (15, False)


def test_anual_semana_1_igual_a_la_hoja_EFG():
    s = plan("ANUAL")["semanas"][0]
    assert s["EFG"]["pct_volumen"] == 38
    assert s["EFG"]["volumenes"]["cross"] == pytest.approx(34.2)
    assert s["EFG"]["vpe"] == pytest.approx(154.3)
    assert s["desde"] == INICIO and s["etapa"] == "Preparatoria General"


def test_semestral_comprime_por_etapa():
    mapa = mapa_semanas("SEMESTRAL")
    assert len(mapa) == 24
    assert mapa[0] == ("Preparatoria General", 1, 2)
    assert mapa[12] == ("Precompetitiva", 25, 25)
    assert mapa[20] == ("Transición", 37, 39)
    s = plan("SEMESTRAL")["semanas"][1]
    assert s["EFG"]["pct_volumen"] == pytest.approx(47.5)
    assert s["EFG"]["vpe"] == pytest.approx(158.815)


def test_semanas_sin_intensidad_no_tienen_vpe():
    s = plan("SEMESTRAL")["semanas"][23]
    assert s["EFE"]["pct_intensidad"] == 0 and s["EFE"]["vpe"] is None


def test_fechas_de_etapas_semestral():
    et = {e["etapa"]: e for e in plan("SEMESTRAL")["etapas"]}
    assert et["Precompetitiva"]["desde"] == dt.date(2027, 2, 15)
    assert et["Precompetitiva"]["hasta"] == dt.date(2027, 3, 14)
    assert sum(e["semanas"] for e in et.values()) == 24


def test_pdf_y_excel_se_generan():
    from motor.excel import libro_xlsm, libro_xlsx
    from motor.pdf import generar_pdf
    p = Parametros("SEMESTRAL", INICIO, DEP)
    assert generar_pdf(calcular_plan(p)).startswith(b"%PDF")
    wb = openpyxl.load_workbook(io.BytesIO(libro_xlsx(p)))
    assert wb["DATOS"]["B4"].value == "SEMESTRAL" and "MENÚ" not in wb.sheetnames
    assert wb["PLAN_EFG"].row_dimensions[30].hidden
    z = zipfile.ZipFile(io.BytesIO(libro_xlsm(p)))
    assert "xl/vbaProject.bin" in z.namelist()
    assert b"macroEnabled" in z.read("[Content_Types].xml")
