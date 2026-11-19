"""Cálculos del plan de entrenamiento (las mismas fórmulas del libro de Excel).

    FC máxima        = 220 - edad (hombres)  ·  226 - edad (mujeres)
    100 % esfuerzo   = FC máxima - VPR
    VPE              = % intensidad × (FC máxima - VPR) / 100 + VPR
    Volumen          = cantidad al 100 % × % volumen / 100

Ciclo SEMESTRAL: cada semana de una etapa se corresponde, en proporción, con un
grupo de semanas de la MISMA etapa del ciclo anual; sus porcentajes son el
promedio (sin contar ceros) de ese grupo.
"""
from __future__ import annotations

import datetime as dt
import json
from dataclasses import dataclass, field
from pathlib import Path

DATOS = Path(__file__).resolve().parent.parent / "datos"


def _cargar(nombre: str) -> dict:
    with open(DATOS / nombre, encoding="utf-8") as f:
        return json.load(f)


CURVAS = _cargar("curvas.json")
PERIODIZACION = _cargar("periodizacion.json")
MICROCICLOS = _cargar("microciclos.json")
ETAPAS: list[str] = PERIODIZACION["etapas"]

# Formación -> ejercicios (clave interna, nombre visible)
EJERCICIOS = {
    "EFG": [("cross", "Cross"), ("escalera", "Escalera")],
    "EFE": [("patadas", "Patadas"), ("saco", "Saco"), ("canillera", "Canillera")],
    "EFC": [("combate_condicionado", "Combate condicionado"), ("combate_libre", "Combate libre")],
}
NOMBRE_FORMACION = {
    "EFG": "EJERCICIOS DE FORMACIÓN GENERAL (EFG)",
    "EFE": "EJERCICIOS DE FORMACIÓN ESPECIAL (EFE)",
    "EFC": "EJERCICIOS DE FORMACIÓN COMPETITIVA (EFC)",
}
SEMANAS = {"ANUAL": 48, "SEMESTRAL": 24}


@dataclass
class Deportista:
    nombre: str
    edad: float
    sexo: str            # "M" o "F"
    peso: float          # kg
    vpr: float           # pulso en reposo
    vpb: float           # pulso basal
    cien: dict[str, float] = field(default_factory=dict)   # cantidades al 100 %


@dataclass
class Parametros:
    ciclo: str                  # "ANUAL" o "SEMESTRAL"
    inicio: dt.date
    deportista: Deportista


# ----------------------------------------------------------------------------
def fc_maxima(edad: float, sexo: str) -> float:
    return (226 if str(sexo).upper() == "F" else 220) - edad


def vpe(pct_intensidad: float, fcmax: float, vpr: float) -> float:
    return (fcmax - vpr) * pct_intensidad / 100 + vpr


def canilleras(peso: float) -> dict[int, float]:
    """Peso de cada canillera (gramos por pierna) = (peso × 0,25 / 2) × % × 1000."""
    return {p: (peso * 0.25 / 2) * (p / 100) * 1000 for p in (2, 3, 4, 5)}


def _promedio_positivos(valores: list[float]) -> float:
    pos = [v for v in valores if v and v > 0]
    return sum(pos) / len(pos) if pos else 0.0


def mapa_semanas(ciclo: str, anual: list[str] | None = None,
                 semestral: list[str] | None = None) -> list[tuple[str, int, int]]:
    """Para cada semana del plan: (etapa, primera, última semana anual que la alimentan)."""
    anual = anual or PERIODIZACION["ANUAL"]
    semestral = semestral or PERIODIZACION["SEMESTRAL"]
    if ciclo == "ANUAL":
        return [(anual[k - 1], k, k) for k in range(1, 49)]
    salida = []
    for k in range(1, len(semestral) + 1):
        etapa = semestral[k - 1]
        j = k - semestral.index(etapa)              # posición dentro de la etapa (1..n)
        n = semestral.count(etapa)
        if etapa not in anual:
            salida.append((etapa, 0, 0))
            continue
        a0 = anual.index(etapa) + 1
        m = anual.count(etapa)
        primera = a0 + int((j - 1) * m / n)
        ultima = max(primera, a0 + int(j * m / n) - 1)
        salida.append((etapa, primera, ultima))
    return salida


def calcular_plan(p: Parametros) -> dict:
    """Devuelve todo lo necesario para mostrar, imprimir y exportar el plan."""
    d = p.deportista
    fcmax = fc_maxima(d.edad, d.sexo)
    mapa = mapa_semanas(p.ciclo)
    semanas = []
    for k, (etapa, a, b) in enumerate(mapa, start=1):
        desde = p.inicio + dt.timedelta(days=7 * (k - 1))
        fila = {"semana": k, "mes": (k - 1) // 4 + 1, "desde": desde,
                "hasta": desde + dt.timedelta(days=6), "etapa": etapa}
        for form, ejer in EJERCICIOS.items():
            if a == 0:
                pv = pi = 0.0
            else:
                pv = _promedio_positivos(CURVAS[form]["volumen"][a - 1:b])
                pi = _promedio_positivos(CURVAS[form]["intensidad"][a - 1:b])
            fila[form] = {
                "pct_volumen": pv,
                "pct_intensidad": pi,
                "volumenes": {clave: d.cien.get(clave, 0) * pv / 100 for clave, _ in ejer},
                "vpe": vpe(pi, fcmax, d.vpr) if pi > 0 else None,
            }
        semanas.append(fila)

    etapas = []
    for nombre in ETAPAS:
        nums = [s["semana"] for s in semanas if s["etapa"] == nombre]
        micro = next((c for c, m in MICROCICLOS.items() if isinstance(m, dict) and m.get("etapa") == nombre), None)
        if nums:
            etapas.append({"etapa": nombre, "semanas": len(nums), "primera": nums[0], "ultima": nums[-1],
                           "desde": semanas[nums[0] - 1]["desde"], "hasta": semanas[nums[-1] - 1]["hasta"],
                           "microciclo": micro})
        else:
            etapas.append({"etapa": nombre, "semanas": 0, "primera": None, "ultima": None,
                           "desde": None, "hasta": None, "microciclo": micro})

    return {
        "parametros": p,
        "fc_maxima": fcmax,
        "esfuerzo_100": fcmax - d.vpr,
        "fin": semanas[-1]["hasta"],
        "semanas": semanas,
        "etapas": etapas,
        "canilleras": canilleras(d.peso),
    }


def indice_sobreentrenamiento(acostado: float, de_pie: float) -> tuple[float, bool]:
    """Diferencia (de pie - acostado) y si es señal de alerta (> 15 pulsaciones)."""
    dif = de_pie - acostado
    return dif, dif > 15


def lunes_de(fecha: dt.date) -> dt.date:
    return fecha - dt.timedelta(days=fecha.weekday())
