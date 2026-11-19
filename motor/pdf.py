"""Genera el PDF del plan de entrenamiento con reportlab (no necesita Excel).

Páginas: resumen · EFG · EFE · EFC · microciclos (horizontal) · registro del índice
de sobreentrenamiento.
"""
from __future__ import annotations

import io

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import letter, landscape
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import (BaseDocTemplate, Frame, NextPageTemplate, PageBreak, PageTemplate,
                                Paragraph, Spacer, Table, TableStyle)

from .calculos import EJERCICIOS, MICROCICLOS, NOMBRE_FORMACION, SEMANAS

AZUL = colors.HexColor("#1F3864")
CELESTE = colors.HexColor("#DDEBF7")
GRIS = colors.HexColor("#F2F2F2")
BORDE = colors.HexColor("#BFBFBF")
ROJO = colors.HexColor("#C00000")
COLOR_ETAPA = {
    "Preparatoria General": colors.HexColor("#E2EFDA"),
    "Preparatoria Especial": colors.HexColor("#FFF2CC"),
    "Precompetitiva": colors.HexColor("#FCE4D6"),
    "Competitiva": colors.HexColor("#F8CBAD"),
    "Transición": colors.HexColor("#EDEDED"),
}

E_TIT = ParagraphStyle("tit", fontName="Helvetica-Bold", fontSize=15, textColor=colors.white,
                       alignment=TA_CENTER, leading=18)
E_SUB = ParagraphStyle("sub", fontName="Helvetica-Oblique", fontSize=9.5, textColor=AZUL,
                       alignment=TA_CENTER)
E_SEC = ParagraphStyle("sec", fontName="Helvetica-Bold", fontSize=10, textColor=colors.white,
                       alignment=TA_CENTER)
E_TXT = ParagraphStyle("txt", fontName="Helvetica", fontSize=9, leading=11.5)
E_CEL = ParagraphStyle("cel", fontName="Helvetica", fontSize=9, leading=11)
E_TIT13 = ParagraphStyle("tit13", parent=E_TIT, fontSize=12.5, leading=15)
E_HDR = ParagraphStyle("hdr", fontName="Helvetica-Bold", fontSize=7.4, leading=8.6, textColor=AZUL, alignment=TA_CENTER)
E_CEN = ParagraphStyle("cen", parent=E_TXT, alignment=TA_CENTER)
E_NOTA = ParagraphStyle("nota", fontName="Helvetica-Oblique", fontSize=8, leading=10, textColor=colors.grey)


def _f(fecha) -> str:
    return fecha.strftime("%d/%m/%Y") if fecha else ""


def _n(valor, dec=1) -> str:
    if valor is None:
        return "-"
    texto = f"{valor:,.{dec}f}"
    return texto.replace(",", "X").replace(".", ",").replace("X", ".")


def _banda(texto: str, ancho: float, estilo=E_TIT, alto=None) -> Table:
    t = Table([[Paragraph(texto, estilo)]], colWidths=[ancho], rowHeights=alto)
    t.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), AZUL),
                           ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                           ("TOPPADDING", (0, 0), (-1, -1), 4), ("BOTTOMPADDING", (0, 0), (-1, -1), 5)]))
    return t


def _tabla(datos, anchos, encabezado=True, alto_filas=None, extra=None) -> Table:
    t = Table(datos, colWidths=anchos, rowHeights=alto_filas, repeatRows=1 if encabezado else 0)
    estilo = [("GRID", (0, 0), (-1, -1), 0.4, BORDE),
              ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
              ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
              ("FONTSIZE", (0, 0), (-1, -1), 9),
              ("TOPPADDING", (0, 0), (-1, -1), 2), ("BOTTOMPADDING", (0, 0), (-1, -1), 2)]
    if encabezado:
        estilo += [("BACKGROUND", (0, 0), (-1, 0), CELESTE),
                   ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                   ("TEXTCOLOR", (0, 0), (-1, 0), AZUL)]
    t.setStyle(TableStyle(estilo + (extra or [])))
    return t


def _datos_deportista(plan) -> str:
    p = plan["parametros"]
    d = p.deportista
    return (f"<b>Deportista:</b> {d.nombre} &nbsp;&nbsp; <b>Ciclo:</b> {p.ciclo} ({SEMANAS[p.ciclo]} semanas) "
            f"&nbsp;&nbsp; <b>Inicio:</b> {_f(p.inicio)}<br/>"
            f"<b>Edad:</b> {_n(d.edad, 0)} años &nbsp;&nbsp; <b>Sexo:</b> "
            f"{'Femenino' if d.sexo == 'F' else 'Masculino'} &nbsp;&nbsp; <b>VPR:</b> {_n(d.vpr, 0)} ppm "
            f"&nbsp;&nbsp; <b>FC máx:</b> {_n(plan['fc_maxima'], 0)} ppm")


# ----------------------------------------------------------------------------
def _resumen(plan, ancho):
    p = plan["parametros"]
    d = p.deportista
    h = [_banda("PLAN DE ENTRENAMIENTO PERSONALIZADO - TAEKWONDO", ancho, alto=26),
         Spacer(1, 3), Paragraph("Volúmenes e intensidades semanales según las curvas de Matveyev", E_SUB),
         Spacer(1, 10), _banda("DATOS DEL DEPORTISTA", ancho, E_SEC)]
    filas = [
        ["Nombre:", d.nombre, "Tipo de ciclo:", f"{p.ciclo} ({SEMANAS[p.ciclo]} semanas)"],
        ["Edad:", f"{_n(d.edad, 0)} años", "Fecha de inicio:", _f(p.inicio)],
        ["Sexo:", "Femenino" if d.sexo == "F" else "Masculino", "Fecha final:", _f(plan["fin"])],
        ["Peso corporal:", f"{_n(d.peso)} kg", "FC máxima estimada:", f"{_n(plan['fc_maxima'], 0)} ppm"],
        ["VPR (pulso en reposo):", f"{_n(d.vpr, 0)} ppm", "100 % de esfuerzo:", f"{_n(plan['esfuerzo_100'], 0)} ppm"],
        ["VPB (pulso basal):", f"{_n(d.vpb, 0)} ppm", "", ""],
    ]
    a = ancho / 4
    h.append(_tabla(filas, [a * 0.9, a * 1.1, a * 0.9, a * 1.1], encabezado=False, extra=[
        ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"), ("FONTNAME", (2, 0), (2, -1), "Helvetica-Bold"),
        ("ALIGN", (0, 0), (0, -1), "RIGHT"), ("ALIGN", (2, 0), (2, -1), "RIGHT"),
        ("BACKGROUND", (1, 0), (1, -1), GRIS), ("BACKGROUND", (3, 0), (3, -2), GRIS),
        ("FONTNAME", (1, 0), (1, 0), "Helvetica-Bold")]))
    h += [Spacer(1, 10), _banda("ETAPAS DEL CICLO", ancho, E_SEC)]
    filas = [["Etapa", "Semanas", "Semana\ninicial", "Semana\nfinal", "Desde", "Hasta", "Microciclo"]]
    extra = [("ALIGN", (1, 0), (-1, -1), "CENTER")]
    for i, e in enumerate(plan["etapas"], start=1):
        micro = MICROCICLOS[e["microciclo"]]["hoja"] if e["microciclo"] else "(según calendario)"
        filas.append([e["etapa"], e["semanas"], e["primera"] or "", e["ultima"] or "",
                      _f(e["desde"]), _f(e["hasta"]), micro])
        extra.append(("BACKGROUND", (0, i), (-1, i), COLOR_ETAPA[e["etapa"]]))
    h.append(_tabla(filas, [ancho * x for x in (.22, .09, .1, .1, .13, .13, .23)], extra=extra))
    h += [Spacer(1, 10), _banda("CANTIDADES AL 100 % DE ESFUERZO", ancho, E_SEC)]
    filas, fila = [], []
    for form, ejer in EJERCICIOS.items():
        for clave, nombre in ejer:
            fila += [f"{nombre} ({form}):", _n(d.cien.get(clave, 0))]
            if len(fila) == 4:
                filas.append(fila); fila = []
    if fila:
        filas.append(fila + ["", ""])
    h.append(_tabla(filas, [a * 1.2, a * 0.8, a * 1.2, a * 0.8], encabezado=False, extra=[
        ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"), ("FONTNAME", (2, 0), (2, -1), "Helvetica-Bold"),
        ("ALIGN", (0, 0), (0, -1), "RIGHT"), ("ALIGN", (2, 0), (2, -1), "RIGHT"),
        ("ALIGN", (1, 0), (1, -1), "CENTER"), ("ALIGN", (3, 0), (3, -1), "CENTER")]))
    h += [Spacer(1, 10), _banda("PESO DE CADA CANILLERA (gramos por pierna)", ancho, E_SEC)]
    can = plan["canilleras"]
    h.append(_tabla([[f"Peso: {_n(d.peso)} kg"] + [f"{k} %" for k in can],
                     [""] + [f"{_n(v, 0)} g" for v in can.values()]],
                    [ancho * .2] + [ancho * .2] * 4, extra=[
                        ("SPAN", (0, 0), (0, 1)), ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                        ("FONTNAME", (1, 1), (-1, 1), "Helvetica-Bold"), ("FONTSIZE", (1, 1), (-1, 1), 11)]))
    h += [Spacer(1, 10), _banda("FÓRMULAS UTILIZADAS", ancho, E_SEC), Spacer(1, 3)]
    for t in ["FC máxima = 220 - edad (hombres) · 226 - edad (mujeres)",
              "100 % de esfuerzo = FC máxima - VPR",
              "VPE (pulso de entrenamiento) = % Intensidad × (FC máxima - VPR) / 100 + VPR",
              "Volumen de cada ejercicio = cantidad al 100 % × % Volumen / 100",
              "Ciclo semestral: cada etapa comprime proporcionalmente la misma etapa del ciclo anual."]:
        h.append(Paragraph("• " + t, E_TXT))
    h += [Spacer(1, 10), _banda("ÍNDICE DE SOBREENTRENAMIENTO (control diario)", ancho, E_SEC), Spacer(1, 3),
          Paragraph("Cada mañana, tomar la frecuencia cardiaca ACOSTADO y luego inmediatamente al LEVANTARSE "
                    "(de pie). Si la diferencia supera las 15 pulsaciones es señal de alerta: revisar descanso, "
                    "trasnochos, consumo de alcohol y cargas de entrenamiento; si se repite, consultar al "
                    "médico. Registrar los valores en la planilla de la última página.", E_TXT)]
    return h


def _tabla_plan(plan, form, ancho):
    ejer = EJERCICIOS[form]
    h = [_banda("PLAN DE ENTRENAMIENTO - " + NOMBRE_FORMACION[form], ancho, E_TIT13), Spacer(1, 4),
         Paragraph(_datos_deportista(plan), E_TXT),
         Paragraph("<b>Al 100 %:</b> " + " &nbsp;&nbsp; ".join(
             f"{n}: {_n(plan['parametros'].deportista.cien.get(c, 0))}" for c, n in ejer), E_TXT),
         Spacer(1, 5)]
    filas = [[Paragraph(t, E_HDR) for t in
              ["Sem.", "Mes", "Desde", "Hasta", "Etapa", "% Vol."] + [f"{n} (volumen)" for _, n in ejer]
              + ["% Int.", "VPE (ppm)"]]]
    extra = [("ALIGN", (0, 0), (-1, -1), "CENTER"), ("FONTSIZE", (0, 1), (-1, -1), 7.6),
             ("FONTSIZE", (0, 0), (-1, 0), 7.6), ("TOPPADDING", (0, 1), (-1, -1), 1),
             ("BOTTOMPADDING", (0, 1), (-1, -1), 1.5), ("FONTNAME", (0, 1), (0, -1), "Helvetica-Bold"),
             ("FONTNAME", (-1, 1), (-1, -1), "Helvetica-Bold"), ("TEXTCOLOR", (-1, 1), (-1, -1), ROJO)]
    for i, s in enumerate(plan["semanas"], start=1):
        f = s[form]
        filas.append([s["semana"], s["mes"], _f(s["desde"]), _f(s["hasta"]), s["etapa"], _n(f["pct_volumen"], 0)]
                     + [_n(v) for v in f["volumenes"].values()]
                     + [_n(f["pct_intensidad"], 0), _n(f["vpe"], 0)])
        extra.append(("BACKGROUND", (0, i), (-1, i), COLOR_ETAPA[s["etapa"]]))
    n = len(ejer)
    fijos = [.06, .05, .11, .11, .19, .08]
    resto = (1 - sum(fijos) - .1 - .07) / n
    anchos = [ancho * x for x in fijos + [resto] * n + [.1, .07]]
    alto = 12.4 if len(plan["semanas"]) > 24 else 17
    h.append(_tabla(filas, anchos, alto_filas=[24] + [alto] * len(plan["semanas"]), extra=extra))
    return h


def _microciclo(plan, codigo, ancho):
    m = MICROCICLOS[codigo]
    d = plan["parametros"].deportista
    et = next(e for e in plan["etapas"] if e["etapa"] == m["etapa"])
    h = [_banda("MICROCICLO SEMANAL - " + m["titulo"], ancho, alto=26), Spacer(1, 5),
         Paragraph(_datos_deportista(plan), E_TXT)]
    if et["semanas"]:
        dur = (f"<b>Duración:</b> {et['semanas']} semanas (de la semana {et['primera']} a la {et['ultima']} "
               f"del ciclo) &nbsp;&nbsp; <b>Desde:</b> {_f(et['desde'])} &nbsp;&nbsp; <b>Hasta:</b> {_f(et['hasta'])}")
    else:
        dur = "<b>Esta etapa no forma parte del ciclo elegido.</b>"
    h += [Paragraph(dur, E_TXT), Spacer(1, 5)]
    t = Table([[Paragraph(f"<b><i>{m['correlacion']}</i></b>", E_TXT)]], colWidths=[ancho])
    t.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), CELESTE)]))
    h += [t, Spacer(1, 6)]
    filas = [["DÍA", "MAÑANA", "NOCHE"]]
    for dia in m["dias"]:
        filas.append([dia["dia"], Paragraph(dia["manana"], E_CEL), Paragraph(dia["noche"], E_CEL)])
    h.append(_tabla(filas, [ancho * .12, ancho * .44, ancho * .44], extra=[
        ("BACKGROUND", (0, 0), (-1, 0), AZUL), ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("ALIGN", (0, 0), (-1, 0), "CENTER"), ("FONTNAME", (0, 1), (0, -1), "Helvetica-Bold"),
        ("BACKGROUND", (0, 1), (0, -1), GRIS), ("TOPPADDING", (0, 1), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 1), (-1, -1), 4)]))
    pesas = m.get("pesas")
    if pesas:
        h += [Spacer(1, 8), _banda(pesas["titulo"], ancho, E_SEC)]
        if "pasos" in pesas:
            filas = [[f"Paso {i}", f"Un peso tal que permita {r} repetici{'ón' if r == 1 else 'ones'}"]
                     for i, r in enumerate(pesas["pasos"], start=1)]
            h.append(_tabla(filas, [ancho * .12, ancho * .88], encabezado=False, extra=[
                ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"), ("BACKGROUND", (0, 0), (0, -1), GRIS),
                ("FONTSIZE", (0, 0), (-1, -1), 8.5), ("TOPPADDING", (0, 0), (-1, -1), 1),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 1)]))
            h.append(Paragraph(f"<font color='#C00000'><b>{pesas['nota']}</b></font>", E_TXT))
        else:
            h.append(Paragraph(pesas["texto"], E_CEL))
    if m.get("canilleras"):
        can = plan["canilleras"]
        h += [Spacer(1, 8), _banda("PESO DE CADA CANILLERA (gramos por pierna, según el peso corporal)", ancho, E_SEC),
              _tabla([[f"{_n(d.peso)} kg"] + [f"{k} %:  {_n(v, 0)} g" for k, v in can.items()]],
                     [ancho * .2] * 5, encabezado=False, extra=[
                         ("ALIGN", (0, 0), (-1, -1), "CENTER"), ("FONTNAME", (0, 0), (-1, -1), "Helvetica-Bold"),
                         ("FONTSIZE", (0, 0), (-1, -1), 10.5), ("BACKGROUND", (0, 0), (0, 0), GRIS)])]
    h += [Spacer(1, 8), Paragraph("Pulso de entrenamiento (VPE) de cada semana: ver las tablas EFG, EFE y EFC. "
                                  "Índice de sobreentrenamiento: tomar el pulso acostado y luego de pie; si la "
                                  "diferencia supera 15 pulsaciones es señal de alerta.", E_NOTA)]
    return h


def _planilla_indice(plan, ancho, registros=None):
    d = plan["parametros"].deportista
    h = [_banda("ÍNDICE DE SOBREENTRENAMIENTO", ancho, alto=24), Spacer(1, 5),
         Paragraph(f"<b>Deportista:</b> {d.nombre}", E_TXT), Spacer(1, 3),
         Paragraph("Tomar el pulso ACOSTADO al despertar y luego inmediatamente DE PIE. Si la diferencia es "
                   "mayor de 15 pulsaciones es señal de alerta: revisar descanso, trasnochos, alcohol y cargas; "
                   "si se repite, consultar al médico.", E_NOTA), Spacer(1, 6)]
    filas = [["Fecha", "Pulso acostado\n(ppm)", "Pulso de pie\n(ppm)", "Diferencia", "Resultado"]]
    extra = []
    for i, r in enumerate(registros or [], start=1):
        dif = r["de_pie"] - r["acostado"]
        filas.append([_f(r["fecha"]), _n(r["acostado"], 0), _n(r["de_pie"], 0), _n(dif, 0),
                      "ALERTA: posible sobreentrenamiento" if dif > 15 else "Normal"])
        if dif > 15:
            extra.append(("TEXTCOLOR", (0, i), (-1, i), ROJO))
    while len(filas) < 33:
        filas.append(["", "", "", "", ""])
    h.append(_tabla(filas, [ancho * x for x in (.16, .17, .17, .14, .36)],
                    alto_filas=[26] + [18.5] * (len(filas) - 1),
                    extra=[("ALIGN", (0, 0), (-1, -1), "CENTER")] + extra))
    return h


# ----------------------------------------------------------------------------
def generar_pdf(plan, registros_indice=None) -> bytes:
    """Devuelve el PDF del plan completo como bytes."""
    buf = io.BytesIO()
    margen = 1.3 * cm
    nombre = plan["parametros"].deportista.nombre

    def pie(canvas, doc):
        canvas.saveState()
        canvas.setFont("Helvetica", 7.5)
        canvas.setFillColor(colors.grey)
        w, _ = canvas._pagesize
        canvas.drawString(margen, 0.7 * cm, f"Plan de entrenamiento - {nombre}")
        canvas.drawRightString(w - margen, 0.7 * cm, f"Página {doc.page}")
        canvas.restoreState()

    doc = BaseDocTemplate(buf, pagesize=letter, leftMargin=margen, rightMargin=margen,
                          topMargin=margen, bottomMargin=margen,
                          title=f"Plan de entrenamiento - {nombre}", author="Plantilla de Justo Fuentes")
    vw, vh = letter
    hw, hh = landscape(letter)
    doc.addPageTemplates([
        PageTemplate("vertical", [Frame(margen, margen, vw - 2 * margen, vh - 2 * margen, id="v")],
                     onPage=pie, pagesize=letter),
        PageTemplate("horizontal", [Frame(margen, margen, hw - 2 * margen, hh - 2 * margen, id="h")],
                     onPage=pie, pagesize=landscape(letter)),
    ])
    av, ah = vw - 2 * margen, hw - 2 * margen
    h = _resumen(plan, av)
    for form in EJERCICIOS:
        h += [PageBreak()] + _tabla_plan(plan, form, av)
    h += [NextPageTemplate("horizontal")]
    for codigo in ("EPG", "EPE", "EPC"):
        h += [PageBreak()] + _microciclo(plan, codigo, ah)
    h += [NextPageTemplate("vertical"), PageBreak()] + _planilla_indice(plan, av, registros_indice)
    doc.build(h)
    return buf.getvalue()
