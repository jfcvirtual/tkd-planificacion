"""Genera el libro de Excel del plan (.xlsx sin macros o .xlsm con el menú/macro).

Parte de la plantilla original de Justo Fuentes (datos/plantilla_base.xlsx), que conserva
las hojas FORMULAS, EFG, EFE, EFC, los gráficos de Matveyev y la imagen de la correlación.
Los cálculos dentro del libro son FÓRMULAS de Excel (iguales a motor/calculos.py).
"""
import datetime
import io

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, Protection
from openpyxl.formatting.rule import FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.properties import PageSetupProperties

from .calculos import DATOS, MICROCICLOS, PERIODIZACION, SEMANAS, Parametros
from .vba import inyectar_vba

PLANTILLA = DATOS / "plantilla_base.xlsx"



# ------------------------------------------------------------------ estilos
AZUL = "1F3864"; CELESTE = "DDEBF7"; GRIS = "F2F2F2"; BORDE = "BFBFBF"
thin = Side(style="thin", color=BORDE)
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
F_TIT = Font(name="Calibri", size=16, bold=True, color="FFFFFF")
F_SUB = Font(name="Calibri", size=11, italic=True, color="1F3864")
F_HDR = Font(name="Calibri", size=10, bold=True, color="1F3864")
F_LBL = Font(name="Calibri", size=11, bold=True)
F_TXT = Font(name="Calibri", size=11)
FILL_T = PatternFill("solid", fgColor=AZUL)
FILL_H = PatternFill("solid", fgColor=CELESTE)
FILL_G = PatternFill("solid", fgColor=GRIS)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
RIGHT = Alignment(horizontal="right", vertical="center")
UNLOCK = Protection(locked=False)
FMT_FECHA = "dd/mm/yyyy"

ETAPAS = ["Preparatoria General", "Preparatoria Especial", "Precompetitiva", "Competitiva", "Transición"]
COLOR_ETAPA = ["E2EFDA", "FFF2CC", "FCE4D6", "F8CBAD", "EDEDED"]
MICRO = {"Preparatoria General": "MICROCICLO_EPG", "Preparatoria Especial": "MICROCICLO_EPE",
         "Precompetitiva": "MICROCICLO_EPC"}


def titulo(ws, rng, texto, alto=30, font=F_TIT):
    ws.merge_cells(rng)
    c = ws[rng.split(":")[0]]
    c.value, c.font, c.fill, c.alignment = texto, font, FILL_T, CENTER
    ws.row_dimensions[c.row].height = alto


def encab(ws, rng, texto):
    if ":" in rng:
        ws.merge_cells(rng)
    c = ws[rng.split(":")[0]]
    c.value, c.font, c.fill, c.alignment = texto, Font(bold=True, color="FFFFFF"), FILL_T, CENTER
    ws.row_dimensions[c.row].height = 20


def celda(ws, ref, valor, font=F_TXT, fill=None, align=CENTER, fmt=None, borde=True):
    c = ws[ref]
    c.value = valor
    c.font = font
    if fill: c.fill = fill
    c.alignment = align
    if fmt: c.number_format = fmt
    if borde: c.border = BOX
    return c


def pagina(ws, area, horizontal=False, alto_una=True):
    ws.print_area = area
    ws.page_setup.orientation = "landscape" if horizontal else "portrait"
    ws.page_setup.paperSize = ws.PAPERSIZE_LETTER
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1 if alto_una else 0
    ws.print_options.horizontalCentered = True
    ws.page_margins.left = ws.page_margins.right = 0.4
    ws.page_margins.top = ws.page_margins.bottom = 0.5
    ws.oddFooter.center.text = "Página &P de &N"
    ws.oddFooter.center.size = 8



def construir_workbook(p: Parametros) -> openpyxl.Workbook:
    """Arma el libro completo con los datos del deportista ya escritos en la hoja DATOS."""
    dep = p.deportista
    c = dep.cien
    wb = openpyxl.load_workbook(PLANTILLA)

    # ================================================================== DATOS
    D = wb.create_sheet("DATOS")
    D.sheet_properties.tabColor = "70AD47"
    D.column_dimensions["A"].width = 38; D.column_dimensions["B"].width = 26; D.column_dimensions["C"].width = 60
    titulo(D, "A1:C1", "DATOS DEL DEPORTISTA (los llena el MENÚ - opción 1)", 26, Font(size=13, bold=True, color="FFFFFF"))
    lunes = p.inicio
    entradas = [  # fila, etiqueta, valor, formato, nota
        (3, "Nombre del deportista", dep.nombre, "@", ""),
        (4, "Tipo de ciclo", p.ciclo, "@", "ANUAL (48 semanas) o SEMESTRAL (24 semanas)"),
        (5, "Fecha de inicio del ciclo", lunes, FMT_FECHA, "Se recomienda que sea un lunes"),
        (6, "Edad (años)", dep.edad, "0", ""),
        (7, "Sexo (M / F)", dep.sexo, "@", ""),
        (8, "Peso corporal (kg)", dep.peso, "0.0", "Para calcular el peso de las canilleras"),
        (9, "VPR - pulso en reposo (ppm)", dep.vpr, "0", ""),
        (10, "VPB - pulso basal (ppm)", dep.vpb, "0", "Al despertar, antes de levantarse"),
        (12, "Cross - cantidad al 100 %", c["cross"], "0.0", "Formación General (EFG)"),
        (13, "Escalera - cantidad al 100 %", c["escalera"], "0.0", "Formación General (EFG)"),
        (14, "Patadas - cantidad al 100 %", c["patadas"], "0.0", "Formación Especial (EFE)"),
        (15, "Saco - cantidad al 100 %", c["saco"], "0.0", "Formación Especial (EFE)"),
        (16, "Canillera - cantidad al 100 %", c["canillera"], "0.0", "Formación Especial (EFE)"),
        (17, "Combate condicionado - cantidad al 100 %", c["combate_condicionado"], "0.0", "Formación Competitiva (EFC)"),
        (18, "Combate libre - cantidad al 100 %", c["combate_libre"], "0.0", "Formación Competitiva (EFC)"),
    ]
    for f, lbl, val, fmt, nota in entradas:
        celda(D, f"A{f}", lbl, F_LBL, None, RIGHT)
        c = celda(D, f"B{f}", val, F_TXT, PatternFill("solid", fgColor="FFF2CC"), CENTER, fmt)
        c.protection = UNLOCK
        celda(D, f"C{f}", nota, Font(italic=True, color="7F7F7F"), None, LEFT, borde=False)
    encab(D, "A11:B11", "CANTIDADES AL 100 % DE ESFUERZO")
    encab(D, "A19:B19", "VALORES CALCULADOS (no modificar)")
    calc = [
        (20, "Número de semanas del ciclo", '=IF(B4="SEMESTRAL",24,48)', "0"),
        (21, "FC máxima estimada (ppm)", '=IF(B7="F",226,220)-B6', "0"),
        (22, "100 % de esfuerzo = FC máx - VPR (ppm)", "=B21-B9", "0"),
        (23, "Fecha final del ciclo", "=B5+7*B20-1", FMT_FECHA),
        (24, "Canillera 2 % (gramos por pierna)", "=((B8*0.25)/2)*0.02*1000", "0"),
        (25, "Canillera 3 % (gramos por pierna)", "=((B8*0.25)/2)*0.03*1000", "0"),
        (26, "Canillera 4 % (gramos por pierna)", "=((B8*0.25)/2)*0.04*1000", "0"),
        (27, "Canillera 5 % (gramos por pierna)", "=((B8*0.25)/2)*0.05*1000", "0"),
    ]
    for f, lbl, fx, fmt in calc:
        celda(D, f"A{f}", lbl, F_LBL, None, RIGHT)
        celda(D, f"B{f}", fx, F_TXT, FILL_G, CENTER, fmt)
    dv = DataValidation(type="list", formula1='"ANUAL,SEMESTRAL"', allow_blank=False); D.add_data_validation(dv); dv.add("B4")
    dv2 = DataValidation(type="list", formula1='"M,F"', allow_blank=False); D.add_data_validation(dv2); dv2.add("B7")
    pagina(D, "A1:C27")

    # ================================================================== PERIODIZACION
    P = wb.create_sheet("PERIODIZACION")
    P.sheet_properties.tabColor = "70AD47"
    for col, w in zip("ABCDEFGHIJKLMNO", [9, 24, 3, 9, 24, 3, 3, 8, 22, 6, 6, 6, 6, 8, 8]):
        P.column_dimensions[col].width = w
    titulo(P, "A1:E1", "PERIODIZACIÓN: ETAPA DE CADA SEMANA", 26, Font(size=13, bold=True, color="FFFFFF"))
    P.merge_cells("A2:E3")
    celda(P, "A2", "Puede cambiar la etapa de cada semana (celdas amarillas) escribiendo uno de estos nombres: "
          + ", ".join(ETAPAS) + ". Las semanas de una misma etapa deben ir seguidas.",
          Font(italic=True, size=9), None, LEFT, borde=False)
    P.row_dimensions[2].height = 22; P.row_dimensions[3].height = 22
    for ref, t in [("A5", "Semana"), ("B5", "Ciclo ANUAL"), ("D5", "Semana"), ("E5", "Ciclo SEMESTRAL")]:
        celda(P, ref, t, F_HDR, FILL_H)
    anual = PERIODIZACION["ANUAL"]
    semes = PERIODIZACION["SEMESTRAL"]
    amarillo = PatternFill("solid", fgColor="FFF2CC")
    dv3 = DataValidation(type="list", formula1='"' + ",".join(ETAPAS) + '"', allow_blank=False)
    P.add_data_validation(dv3)
    for k in range(1, 49):
        r = k + 5
        celda(P, f"A{r}", k, F_TXT)
        c = celda(P, f"B{r}", anual[k - 1], F_TXT, amarillo); c.protection = UNLOCK; dv3.add(c.coordinate)
        if k <= 24:
            celda(P, f"D{r}", k, F_TXT)
            c = celda(P, f"E{r}", semes[k - 1], F_TXT, amarillo); c.protection = UNLOCK; dv3.add(c.coordinate)
    # Ayudas de cálculo (H:O)
    P.merge_cells("H4:O4")
    celda(P, "H4", "Cálculo interno (no modificar): semanas del ciclo anual que corresponden a cada semana del plan",
          Font(italic=True, size=8, color="7F7F7F"), None, LEFT, borde=False)
    for ref, t in zip(["H5", "I5", "J5", "K5", "L5", "M5", "N5", "O5"],
                      ["Sem.", "Etapa del plan", "Pos.", "n", "a0", "m", "Desde", "Hasta"]):
        celda(P, ref, t, Font(size=8, bold=True, color="7F7F7F"), FILL_G)
    S = 'DATOS!$B$4="SEMESTRAL"'
    for k in range(1, 49):
        r = k + 5
        fx = {
            "H": k,
            "I": f'=IF(H{r}>DATOS!$B$20,"",IF({S},INDEX($E$6:$E$29,MIN(H{r},24))&"",INDEX($B$6:$B$53,H{r})&""))',
            "J": f'=IF(I{r}="","",IF({S},H{r}-MATCH(I{r},$E$6:$E$29,0)+1,1))',
            "K": f'=IF(I{r}="","",IF({S},COUNTIF($E$6:$E$29,I{r}),1))',
            "L": f'=IF(I{r}="","",IFERROR(MATCH(I{r},$B$6:$B$53,0),0))',
            "M": f'=IF(I{r}="","",COUNTIF($B$6:$B$53,I{r}))',
            "N": f'=IF(I{r}="","",IF({S},IF(M{r}=0,0,L{r}+INT((J{r}-1)*M{r}/K{r})),H{r}))',
            "O": f'=IF(I{r}="","",IF({S},IF(M{r}=0,0,MAX(N{r},L{r}+INT(J{r}*M{r}/K{r})-1)),H{r}))',
        }
        for col, v in fx.items():
            celda(P, f"{col}{r}", v, Font(size=8, color="7F7F7F"))
    pagina(P, "A1:E53")

    ACT = "PERIODIZACION!$I$6:$I$53"   # etapa de cada semana del plan activo


    def f_desde_hasta(etapa_ref):
        """Fórmulas: semanas, primera semana, última semana, fecha inicio, fecha fin de una etapa."""
        n = f'COUNTIF({ACT},{etapa_ref})'
        p = f'MATCH({etapa_ref},{ACT},0)'
        return (f'={n}',
                f'=IF({n}=0,"",{p})',
                f'=IF({n}=0,"",{p}+{n}-1)',
                f'=IF({n}=0,"",DATOS!$B$5+7*({p}-1))',
                f'=IF({n}=0,"",DATOS!$B$5+7*({p}+{n}-1)-1)')


    # ================================================================== PLAN_EFG / EFE / EFC
    PLANES = [
        ("PLAN_EFG", "EFG", "EJERCICIOS DE FORMACIÓN GENERAL (EFG)",
         [("Cross", 12), ("Escalera", 13)]),
        ("PLAN_EFE", "EFE", "EJERCICIOS DE FORMACIÓN ESPECIAL (EFE)",
         [("Patadas", 14), ("Saco", 15), ("Canillera", 16)]),
        ("PLAN_EFC", "EFC", "EJERCICIOS DE FORMACIÓN COMPETITIVA (EFC)",
         [("Combate condicionado", 17), ("Combate libre", 18)]),
    ]
    LET = "ABCDEFGHIJKLMNOP"
    for nombre, base, tit, ejer in PLANES:
        ws = wb.create_sheet(nombre)
        ws.sheet_properties.tabColor = "2F5597"
        ncol = 5 + 1 + len(ejer) + 2        # A..E, %vol, ejercicios, %int, VPE
        last = LET[ncol - 1]
        widths = [8, 6, 11, 11, 20, 10] + [13] * len(ejer) + [11, 10]
        for i, w in enumerate(widths):
            ws.column_dimensions[LET[i]].width = w
        titulo(ws, f"A1:{last}1", "PLAN DE ENTRENAMIENTO - " + tit, 28, Font(size=14, bold=True, color="FFFFFF"))
        # cabecera de datos
        celda(ws, "A2", "Deportista:", F_LBL, None, RIGHT, borde=False); ws.merge_cells("A2:B2")
        ws.merge_cells(f"C2:{LET[ncol - 4]}2"); celda(ws, "C2", '=DATOS!B3&""', Font(size=12, bold=True), None, LEFT, borde=False)
        celda(ws, f"{LET[ncol - 3]}2", "Ciclo:", F_LBL, None, RIGHT, borde=False)
        ws.merge_cells(f"{LET[ncol - 2]}2:{last}2")
        celda(ws, f"{LET[ncol - 2]}2", '=DATOS!B4&" ("&DATOS!B20&" sem.)"', F_TXT, None, LEFT, borde=False)
        celda(ws, "A3", "Edad:", F_LBL, None, RIGHT, borde=False); ws.merge_cells("A3:B3")
        celda(ws, "C3", '=DATOS!B6&" años - "&IF(DATOS!B7="F","Femenino","Masculino")', F_TXT, None, LEFT, borde=False)
        ws.merge_cells("C3:D3")
        celda(ws, "E3", '="VPR: "&DATOS!B9&"   FC máx: "&DATOS!B21&" ppm"', F_TXT, None, LEFT, borde=False)
        ws.merge_cells(f"E3:{LET[ncol - 4]}3")
        celda(ws, f"{LET[ncol - 3]}3", "Inicio:", F_LBL, None, RIGHT, borde=False)
        ws.merge_cells(f"{LET[ncol - 2]}3:{last}3")
        celda(ws, f"{LET[ncol - 2]}3", "=DATOS!B5", F_TXT, None, LEFT, FMT_FECHA, borde=False)
        celda(ws, "A4", "Al 100 %:", F_LBL, None, RIGHT, borde=False); ws.merge_cells("A4:B4")
        txt = "&\"     \"&".join(f'"{e}: "&DATOS!B{row}' for e, row in ejer)
        ws.merge_cells(f"C4:{last}4"); celda(ws, "C4", "=" + txt, F_TXT, None, LEFT, borde=False)
        # encabezados de tabla
        hdr = ["Semana", "Mes", "Desde", "Hasta", "Etapa", "% Volumen"] + [f"{e}\n(volumen)" for e, _ in ejer] + \
              ["% Intensidad", "VPE\n(ppm)"]
        for i, h in enumerate(hdr):
            celda(ws, f"{LET[i]}5", h, F_HDR, FILL_H)
        ws.row_dimensions[5].height = 32
        cvol, cint, cvpe = LET[5], LET[6 + len(ejer)], LET[7 + len(ejer)]
        for k in range(1, 49):
            r = k + 5
            celda(ws, f"A{r}", f'=IF(DATOS!$B$20>={k},{k},"")', Font(bold=True))
            celda(ws, f"B{r}", f'=IF($A{r}="","",INT(($A{r}-1)/4)+1)')
            celda(ws, f"C{r}", f'=IF($A{r}="","",DATOS!$B$5+7*($A{r}-1))', fmt=FMT_FECHA)
            celda(ws, f"D{r}", f'=IF($A{r}="","",C{r}+6)', fmt=FMT_FECHA)
            celda(ws, f"E{r}", f'=IF($A{r}="","",PERIODIZACION!$I{r})', Font(size=10))
            for col, bcol in ((cvol, "E"), (cint, "F")):
                celda(ws, f"{col}{r}",
                      f'=IF($A{r}="","",IF(PERIODIZACION!$N{r}=0,0,IFERROR(AVERAGEIF(OFFSET({base}!${bcol}$5,'
                      f'PERIODIZACION!$N{r},0,PERIODIZACION!$O{r}-PERIODIZACION!$N{r}+1,1),">0"),0)))', fmt="0")
            for i, (e, row) in enumerate(ejer):
                celda(ws, f"{LET[6 + i]}{r}", f'=IF($A{r}="","",DATOS!$B${row}*${cvol}{r}/100)', fmt="0.0")
            celda(ws, f"{cvpe}{r}", f'=IF($A{r}="","",IF(${cint}{r}=0,"-",(DATOS!$B$21-DATOS!$B$9)*${cint}{r}/100+DATOS!$B$9))',
                  Font(bold=True, color="C00000"), fmt="0")
        for etapa, color in zip(ETAPAS, COLOR_ETAPA):
            ws.conditional_formatting.add(f"A6:{last}53", FormulaRule(formula=[f'$E6="{etapa}"'],
                                          fill=PatternFill("solid", fgColor=color, bgColor=color)))
        ws.freeze_panes = "A6"
        pagina(ws, f"A1:{last}53")

    # ================================================================== MICROCICLOS
    DIAS = ["LUNES", "MARTES", "MIERCOLES", "JUEVES", "VIERNES", "SABADO", "DOMINGO"]
    DIAS_OK = ["LUNES", "MARTES", "MIÉRCOLES", "JUEVES", "VIERNES", "SÁBADO", "DOMINGO"]


    def limpia(t):
        if not t: return ""
        t = " ".join(str(t).split())
        return (t.replace("descanzar", "descansar").replace("Descanzar", "Descansar")
                 .replace("sitemas", "sistemas"))


    def leer_micro(ws):
        filas = {}
        rows = list(ws.iter_rows(min_row=1, max_row=ws.max_row))
        dia_rows = [(r[0].row, str(r[0].value).strip().upper()) for r in rows if str(r[0].value).strip().upper() in DIAS]
        for idx, (fr, dia) in enumerate(dia_rows):
            fin = dia_rows[idx + 1][0] if idx + 1 < len(dia_rows) else ws.max_row + 1
            man, noc = [], []
            destino = man
            for rr in range(fr, fin):
                if str(ws.cell(rr, 2).value).strip().upper() == "NOCHE":
                    destino = noc
                for cc in range(3, 12):
                    v = ws.cell(rr, cc).value
                    if v:
                        # nota de la columna G que continúa el texto de la mañana (EPE, domingo)
                        (man if cc > 3 and man else destino).append(limpia(v))
            filas[dia] = (" ".join(man) or "Libre", " ".join(noc) or "Libre")
        return filas, limpia(ws["A2"].value)


    micro_info = {
        "MICROCICLO_EPG": ("ETAPA PREPARATORIA GENERAL", "Preparatoria General"),
        "MICROCICLO_EPE": ("ETAPA PREPARATORIA ESPECIAL", "Preparatoria Especial"),
        "MICROCICLO_EPC": ("ETAPA PRECOMPETITIVA", "Precompetitiva"),
    }
    # contenido de los microciclos: datos/microciclos.json (editable)
    contenidos = {m["hoja"]: ({DIAS[i]: (d["manana"], d["noche"]) for i, d in enumerate(m["dias"])},
                              m["correlacion"])
                  for m in MICROCICLOS.values() if isinstance(m, dict)}
    for n in micro_info:
        del wb[n]

    for n, (tit, etapa) in micro_info.items():
        filas, correl = contenidos[n]
        ws = wb.create_sheet(n)
        ws.sheet_properties.tabColor = "C55A11"
        ws.column_dimensions["A"].width = 16; ws.column_dimensions["B"].width = 58; ws.column_dimensions["C"].width = 58
        titulo(ws, "A1:C1", "MICROCICLO SEMANAL - " + tit, 30, Font(size=15, bold=True, color="FFFFFF"))
        ref_et = f'"{etapa}"'
        nsem, p1, p2, d1, d2 = f_desde_hasta(ref_et)
        celda(ws, "A2", "Deportista:", F_LBL, None, RIGHT, borde=False)
        celda(ws, "B2", '=DATOS!B3&""', Font(size=12, bold=True), None, LEFT, borde=False)
        celda(ws, "C2", '="Ciclo "&DATOS!B4&"  -  Edad: "&DATOS!B6&" años  -  VPR: "&DATOS!B9&" ppm  -  FC máx: "&DATOS!B21&" ppm"',
              F_TXT, None, LEFT, borde=False)
        celda(ws, "A3", "Duración:", F_LBL, None, RIGHT, borde=False)
        celda(ws, "B3", f'=IF({nsem[1:]}=0,"Esta etapa no forma parte del ciclo elegido",{nsem[1:]}&" semanas (de la semana "&{p1[1:]}&" a la "&{p2[1:]}&" del ciclo)")',
              F_TXT, None, LEFT, borde=False)
        celda(ws, "A4", "Desde:", F_LBL, None, RIGHT, borde=False)
        celda(ws, "B4", d1, F_TXT, None, LEFT, "dd/mm/yyyy", borde=False)
        celda(ws, "A5", "Hasta:", F_LBL, None, RIGHT, borde=False)
        celda(ws, "B5", d2, F_TXT, None, LEFT, "dd/mm/yyyy", borde=False)
        ws.merge_cells("A6:C6")
        celda(ws, "A6", correl, Font(italic=True, bold=True, color="1F3864"), FILL_H, LEFT)
        ws.row_dimensions[6].height = 22
        for ref, t in [("A8", "DÍA"), ("B8", "MAÑANA"), ("C8", "NOCHE")]:
            encab(ws, ref, t)
        for i, (dia, dok) in enumerate(zip(DIAS, DIAS_OK)):
            r = 9 + i
            man, noc = filas.get(dia, ("Libre", "Libre"))
            celda(ws, f"A{r}", dok, F_LBL, FILL_G)
            celda(ws, f"B{r}", man, F_TXT, None, LEFT)
            celda(ws, f"C{r}", noc, F_TXT, None, LEFT)
            ws.row_dimensions[r].height = 48 if max(len(man), len(noc)) > 60 else 34
        r = 17
        if n == "MICROCICLO_EPG":
            ws.merge_cells(f"A{r}:C{r}"); encab(ws, f"A{r}", "PESAS - FUERZA RESISTENCIA: MÉTODO DE REPETICIONES")
            ws.merge_cells(f"A{r+1}:C{r+1}")
            celda(ws, f"A{r+1}", "Trabajar con un peso con el que se puedan realizar alrededor de 30 repeticiones "
                  "(según Dimitar Gurkov, la repetición número 30 es la que entrena la fuerza resistencia).",
                  F_TXT, None, LEFT)
            ws.row_dimensions[r + 1].height = 32
            r += 3
        if n == "MICROCICLO_EPE":
            ws.merge_cells(f"A{r}:C{r}"); encab(ws, f"A{r}", "PESAS - MÉTODO DE FUERZA MÁXIMA (una serie por paso)")
            pasos = [9, 6, 3, 1, 3, 6, 9]
            for i, rep in enumerate(pasos):
                ws.merge_cells(f"B{r+1+i}:C{r+1+i}")
                celda(ws, f"A{r+1+i}", f"Paso {i+1}", F_LBL, FILL_G)
                celda(ws, f"B{r+1+i}", f"Un peso tal que permita {rep} repetici{'ón' if rep == 1 else 'ones'}", F_TXT, None, LEFT)
            r += 8
            ws.merge_cells(f"A{r}:C{r}")
            celda(ws, f"A{r}", "NOTA: no se puede hacer fuerza máxima si no se ha hecho el método de repeticiones "
                  "por no menos de 2 meses.", Font(bold=True, color="C00000"), None, LEFT)
            r += 2
        if n in ("MICROCICLO_EPE", "MICROCICLO_EPC"):
            ws.merge_cells(f"A{r}:C{r}")
            encab(ws, f"A{r}", "PESO DE CADA CANILLERA (gramos por pierna, según el peso corporal)")
            celda(ws, f"A{r+1}", "=DATOS!B8&\" kg\"", F_LBL, FILL_G)
            celda(ws, f"B{r+1}", '="2 %: "&TEXT(DATOS!B24,"0")&" g          3 %: "&TEXT(DATOS!B25,"0")&" g"', Font(size=12, bold=True), None, CENTER)
            celda(ws, f"C{r+1}", '="4 %: "&TEXT(DATOS!B26,"0")&" g          5 %: "&TEXT(DATOS!B27,"0")&" g"', Font(size=12), None, CENTER)
            ws.row_dimensions[r + 1].height = 24
            r += 3
        ws.merge_cells(f"A{r}:C{r}")
        celda(ws, f"A{r}", "Pulso de entrenamiento (VPE) de cada semana: ver hojas PLAN_EFG, PLAN_EFE y PLAN_EFC.   "
              "Índice de sobreentrenamiento: tomar el pulso acostado y luego de pie; si la diferencia supera 15 "
              "pulsaciones es señal de alerta.", Font(italic=True, size=9), None, LEFT, borde=False)
        ws.row_dimensions[r].height = 28
        pagina(ws, f"A1:C{r}", horizontal=True)

    # ================================================================== RESUMEN
    R = wb.create_sheet("RESUMEN")
    R.sheet_properties.tabColor = "2F5597"
    for col, w in zip("ABCDEFG", [24, 13, 13, 13, 13, 13, 22]):
        R.column_dimensions[col].width = w
    titulo(R, "A1:G1", "PLAN DE ENTRENAMIENTO PERSONALIZADO - TAEKWONDO", 32)
    R.merge_cells("A2:G2")
    celda(R, "A2", "Volúmenes e intensidades semanales según las curvas de Matveyev", F_SUB, FILL_H, CENTER, borde=False)
    encab(R, "A4:G4", "DATOS DEL DEPORTISTA")
    pares = [  # fila, (etq, fx, fmt) izquierda, (etq, fx, fmt) derecha
        (5, ("Nombre:", '=DATOS!B3&""', "@"), ("Tipo de ciclo:", '=DATOS!B4&" ("&DATOS!B20&" semanas)"', "@")),
        (6, ("Edad:", '=DATOS!B6&" años"', "@"), ("Fecha de inicio:", "=DATOS!B5", FMT_FECHA)),
        (7, ("Sexo:", '=IF(DATOS!B7="F","Femenino","Masculino")', "@"), ("Fecha final:", "=DATOS!B23", FMT_FECHA)),
        (8, ("Peso corporal:", '=DATOS!B8&" kg"', "@"), ("FC máxima estimada:", '=DATOS!B21&" ppm"', "@")),
        (9, ("VPR (pulso en reposo):", '=DATOS!B9&" ppm"', "@"), ("100 % de esfuerzo:", '=DATOS!B22&" ppm"', "@")),
        (10, ("VPB (pulso basal):", '=DATOS!B10&" ppm"', "@"), ("", "", "@")),
    ]
    for f, izq, der in pares:
        celda(R, f"A{f}", izq[0], F_LBL, None, RIGHT, borde=False)
        R.merge_cells(f"B{f}:C{f}"); celda(R, f"B{f}", izq[1], Font(bold=f == 5, size=11), FILL_G, LEFT, izq[2])
        R.merge_cells(f"D{f}:E{f}"); celda(R, f"D{f}", der[0], F_LBL, None, RIGHT, borde=False)
        R.merge_cells(f"F{f}:G{f}"); celda(R, f"F{f}", der[1], F_TXT, FILL_G if der[0] else None, LEFT, der[2], borde=bool(der[0]))
    encab(R, "A12:G12", "ETAPAS DEL CICLO")
    for i, h in enumerate(["Etapa", "Semanas", "Semana inicial", "Semana final", "Desde", "Hasta", "Microciclo (hoja)"]):
        celda(R, f"{'ABCDEFG'[i]}13", h, F_HDR, FILL_H)
    R.row_dimensions[13].height = 30
    for i, (et, color) in enumerate(zip(ETAPAS, COLOR_ETAPA)):
        f = 14 + i
        fl = PatternFill("solid", fgColor=color)
        celda(R, f"A{f}", et, F_LBL, fl, LEFT)
        fx = f_desde_hasta(f"$A{f}")
        for j, col in enumerate("BCDEF"):
            celda(R, f"{col}{f}", fx[j], F_TXT, fl, CENTER, FMT_FECHA if col in "EF" else "0")
        celda(R, f"G{f}", MICRO.get(et, "(según calendario)"), Font(size=9), fl)
    encab(R, "A20:G20", "CANTIDADES AL 100 % DE ESFUERZO")
    cant = [("Cross (EFG)", 12), ("Escalera (EFG)", 13), ("Patadas (EFE)", 14), ("Saco (EFE)", 15),
            ("Canillera (EFE)", 16), ("Combate condicionado (EFC)", 17), ("Combate libre (EFC)", 18)]
    for i, (e, row) in enumerate(cant):
        f = 21 + (i // 2)
        cl, cv = ("A", "B") if i % 2 == 0 else ("D", "F")
        if cl == "D":
            R.merge_cells(f"D{f}:E{f}")
        celda(R, f"{cl}{f}", e + ":", F_LBL, None, RIGHT, borde=False)
        celda(R, f"{cv}{f}", f"=DATOS!B{row}", F_TXT, FILL_G, CENTER, "0.0")
    encab(R, "A26:G26", "PESO DE CADA CANILLERA (gramos por pierna)")
    for i, (p, row) in enumerate([("2 %", 24), ("3 %", 25), ("4 %", 26), ("5 %", 27)]):
        celda(R, f"{'BCDE'[i]}27", p, F_HDR, FILL_H)
        celda(R, f"{'BCDE'[i]}28", f"=DATOS!B{row}", Font(size=12, bold=True), FILL_G, CENTER, '0" g"')
    celda(R, "A28", '="Peso: "&DATOS!B8&" kg"', F_LBL, None, RIGHT, borde=False)
    encab(R, "A30:G30", "FÓRMULAS UTILIZADAS")
    notas = [
        "FC máxima = 220 - edad (hombres)  ·  226 - edad (mujeres)",
        "100 % de esfuerzo = FC máxima - VPR",
        "VPE (pulso de entrenamiento) = % Intensidad × (FC máxima - VPR) / 100 + VPR",
        "Volumen de cada ejercicio = cantidad al 100 % × % Volumen / 100",
        "Ciclo semestral: cada etapa comprime proporcionalmente la misma etapa del ciclo anual.",
    ]
    for i, t in enumerate(notas):
        R.merge_cells(f"A{31+i}:G{31+i}"); celda(R, f"A{31+i}", t, F_TXT, None, LEFT, borde=False)
    encab(R, "A37:G37", "ÍNDICE DE SOBREENTRENAMIENTO (control diario)")
    R.merge_cells("A38:G40")
    celda(R, "A38", "Cada mañana, tomar la frecuencia cardiaca ACOSTADO y luego inmediatamente al LEVANTARSE (de pie). "
          "Si la diferencia supera las 15 pulsaciones es señal de alerta: revisar descanso, trasnochos, consumo de "
          "alcohol, cargas de entrenamiento y, si se repite, consultar al médico. Registrar los valores en la hoja "
          "INDICE_SOBREENTRENAMIENTO.", F_TXT, None, LEFT, borde=False)
    for rr in (38, 39, 40):
        R.row_dimensions[rr].height = 18
    pagina(R, "A1:G40")

    # ================================================================== ÍNDICE DE SOBREENTRENAMIENTO
    del wb["INDICE_SOBREENTRENAMIENTO"]
    I = wb.create_sheet("INDICE_SOBREENTRENAMIENTO")
    I.sheet_properties.tabColor = "C00000"
    for col, w in zip("ABCDE", [14, 16, 16, 14, 40]):
        I.column_dimensions[col].width = w
    titulo(I, "A1:E1", "ÍNDICE DE SOBREENTRENAMIENTO", 28, Font(size=14, bold=True, color="FFFFFF"))
    celda(I, "A2", "Deportista:", F_LBL, None, RIGHT, borde=False)
    I.merge_cells("B2:E2"); celda(I, "B2", '=DATOS!B3&""', Font(size=12, bold=True), None, LEFT, borde=False)
    I.merge_cells("A3:E4")
    celda(I, "A3", "Tomar el pulso ACOSTADO al despertar y luego inmediatamente DE PIE. Si la diferencia es mayor "
          "de 15 pulsaciones es señal de alerta: revisar descanso, trasnochos, alcohol y cargas; si se repite, "
          "consultar al médico.", Font(size=10, italic=True), None, LEFT, borde=False)
    I.row_dimensions[3].height = 20; I.row_dimensions[4].height = 20
    for i, h in enumerate(["Fecha", "Pulso acostado (ppm)", "Pulso de pie (ppm)", "Diferencia", "Resultado"]):
        celda(I, f"{'ABCDE'[i]}6", h, F_HDR, FILL_H)
    I.row_dimensions[6].height = 30
    for r in range(7, 373):
        for col in "ABC":
            c = celda(I, f"{col}{r}", None, F_TXT, None, CENTER, FMT_FECHA if col == "A" else "0")
            c.protection = UNLOCK
        celda(I, f"D{r}", f'=IF(OR(B{r}="",C{r}=""),"",C{r}-B{r})', F_LBL, None, CENTER, "0")
        celda(I, f"E{r}", f'=IF(D{r}="","",IF(D{r}>15,"ALERTA: posible sobreentrenamiento","Normal"))', F_TXT, None, LEFT)
    I.conditional_formatting.add("A7:E372", FormulaRule(formula=['AND($D7<>"",$D7>15)'],
                                 fill=PatternFill("solid", fgColor="F8CBAD", bgColor="F8CBAD"),
                                 font=Font(bold=True, color="C00000")))
    I.freeze_panes = "A7"
    pagina(I, "A1:E34", alto_una=False)
    I.print_title_rows = "6:6"

    # ================================================================== MENÚ
    M = wb.create_sheet("MENÚ")
    M.sheet_properties.tabColor = "1F3864"
    M.sheet_view.showGridLines = False
    M.sheet_view.showRowColHeaders = False
    for col, w in zip("ABCDEFGH", [3, 30, 2, 28, 4, 2, 52, 3]):
        M.column_dimensions[col].width = w
    titulo(M, "B2:H2", "PLANIFICACIÓN DEL ENTRENAMIENTO DE TAEKWONDO", 36, Font(size=18, bold=True, color="FFFFFF"))
    M.merge_cells("B3:H3")
    celda(M, "B3", "Volúmenes, intensidades y microciclos semanales - ciclo ANUAL o SEMESTRAL - curvas de Matveyev",
          F_SUB, FILL_H, CENTER, borde=False)
    M.merge_cells("B4:H4")
    celda(M, "B4", "Si no ve los botones de colores, haga clic en «Habilitar contenido» (barra amarilla) para activar el menú.",
          Font(bold=True, color="C00000"), None, CENTER, borde=False)
    encab(M, "B5:D5", "DATOS DEL DEPORTISTA")
    panel = [("Nombre:", '=DATOS!B3&""', "@"), ("Ciclo:", '=DATOS!B4&" ("&DATOS!B20&" semanas)"', "@"),
             ("Inicio del ciclo:", "=DATOS!B5", FMT_FECHA), ("Fin del ciclo:", "=DATOS!B23", FMT_FECHA),
             ("Edad:", '=DATOS!B6&" años"', "@"), ("Sexo:", "=DATOS!B7", "@"), ("Peso:", '=DATOS!B8&" kg"', "@"),
             ("VPR - pulso en reposo:", '=DATOS!B9&" ppm"', "@"), ("VPB - pulso basal:", '=DATOS!B10&" ppm"', "@"),
             ("FC máxima estimada:", '=DATOS!B21&" ppm"', "@")]
    for i, (lbl, fx, fmt) in enumerate(panel):
        f = 6 + i
        celda(M, f"B{f}", lbl, F_LBL, None, RIGHT, borde=False)
        celda(M, f"D{f}", fx, Font(size=12, bold=(i == 0)), FILL_G, CENTER, fmt)
    encab(M, "B17:D17", "CANTIDADES AL 100 % DE ESFUERZO")
    for i, (e, row) in enumerate(cant):
        f = 18 + i
        celda(M, f"B{f}", e + ":", F_LBL, None, RIGHT, borde=False)
        celda(M, f"D{f}", f"=DATOS!B{row}", Font(size=12), FILL_G, CENTER, "0.0")
    encab(M, "G5", "MENÚ DE OPCIONES  (haga clic en una opción)")
    for f in range(6, 26):
        M.row_dimensions[f].height = 21
    encab(M, "B27:H27", "CÓMO USAR ESTA PLANTILLA")
    pasos = ["1.  Clic en «1. Nuevo plan» y responda las preguntas: primero si el ciclo es ANUAL o SEMESTRAL.",
             "2.  El libro calcula volúmenes, intensidades (VPE) y las fechas de cada semana y de cada microciclo.",
             "3.  Revise el plan con las opciones 2, 3 y 4, o con la «Vista previa».",
             "4.  Imprima el plan completo (opción 7) o guárdelo en PDF (opción 8) para entregarlo al deportista.",
             "5.  Cada mañana registre el índice de sobreentrenamiento (opción 5): el libro avisa si la diferencia supera 15."]
    for i, t in enumerate(pasos):
        M.merge_cells(f"B{28+i}:H{28+i}"); celda(M, f"B{28+i}", t, F_TXT, None, LEFT, borde=False)
    pagina(M, "B2:H32", horizontal=True)

    # ================================================================== base EFG/EFE/EFC enlazadas a DATOS
    link = {"EFG": {"B3": "B6", "D3": "B9", "F3": "B10", "H3": "B7", "J1": "B12", "J2": "B13"},
            "EFE": {"B3": "B6", "D3": "B9", "F3": "B10", "J3": "B7", "J1": "B14", "J2": "B15", "L1": "B16"},
            "EFC": {"B3": "B6", "D3": "B9", "F3": "B10", "H3": "B7", "J1": "B17", "J2": "B18"}}
    for sh, mp in link.items():
        ws = wb[sh]
        for dst, src in mp.items():
            ws[dst].value = f"=DATOS!{src}"
        ws["A2"].value = "CICLO ANUAL DE ENTRENAMIENTO (base de las curvas de Matveyev)"
    ws = wb["EFG"]
    for r in range(22, 54):
        ws[f"G{r}"] = f"=$J$1*E{r}/100"
        ws[f"I{r}"] = f"=$J$2*E{r}/100"
        for col in "HJ":
            ws[f"{col}{r}"] = (f'=IF($H$3="M",(220-$B$3-$D$3)*F{r}/100+$D$3,(226-$B$3-$D$3)*F{r}/100+$D$3)')
        for col in "GHIJ":
            ws[f"{col}{r}"].number_format = ws[f"{col}21"].number_format
            ws[f"{col}{r}"]._style = ws[f"{col}21"]._style
    wb["EFE"]["AE1"] = "M"; wb["EFE"]["AE2"] = "F"
    for c in ("AE1", "AE2"):
        wb["EFE"][c].font = Font(color="FFFFFF")
    wb["FORMULAS"]["D22"] = "Canillera 3%"

    # ================================================================== orden, codeNames
    orden = ["MENÚ", "RESUMEN", "PLAN_EFG", "PLAN_EFE", "PLAN_EFC", "MICROCICLO_EPG", "MICROCICLO_EPE",
             "MICROCICLO_EPC", "INDICE_SOBREENTRENAMIENTO", "DATOS", "PERIODIZACION", "FORMULAS", "EFG", "EFE",
             "EFC", "Gráfico_EFG", "Gráfico_EFE", "Gráfico_EFC", "GRAFICA_CORRELACION_GE"]
    assert sorted(orden) == sorted(wb.sheetnames), wb.sheetnames
    wb._sheets = [wb[n] for n in orden]
    CODES = ["HojaMenu", "HojaResumen", "HojaPlanEFG", "HojaPlanEFE", "HojaPlanEFC", "HojaMicroEPG", "HojaMicroEPE",
             "HojaMicroEPC", "HojaIndice", "HojaDatos", "HojaPeriodos", "HojaFormulas", "HojaEFG", "HojaEFE",
             "HojaEFC", "HojaGrafEFG", "HojaGrafEFE", "HojaGrafEFC", "HojaCorrelacion"]
    for n, cn in zip(orden, CODES):
        wb[n].sheet_properties.codeName = cn
        wb[n].sheet_view.tabSelected = (n == "MENÚ")
    wb.active = 0
    wb.code_name = "ThisWorkbook"
    wb.calculation.fullCalcOnLoad = True
    return wb


def _ocultar_semanas(wb, ciclo: str) -> None:
    n = SEMANAS[ciclo]
    for hoja in ("PLAN_EFG", "PLAN_EFE", "PLAN_EFC"):
        ws = wb[hoja]
        for r in range(6, 54):
            ws.row_dimensions[r].hidden = r >= 6 + n


def _a_bytes(wb) -> bytes:
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def libro_xlsx(p: Parametros) -> bytes:
    """Libro SIN macros, listo para usar (no necesita habilitar contenido)."""
    wb = construir_workbook(p)
    del wb["MENÚ"]                       # el menú solo funciona con la macro
    wb.active = 0
    for ws in wb.worksheets:
        ws.sheet_view.tabSelected = ws.title == "RESUMEN"
    _ocultar_semanas(wb, p.ciclo)
    return _a_bytes(wb)


def libro_xlsm(p: Parametros) -> bytes:
    """Libro CON la macro: al abrirlo muestra el MENÚ (nuevo plan, índice, imprimir, PDF)."""
    wb = construir_workbook(p)
    _ocultar_semanas(wb, p.ciclo)
    return inyectar_vba(_a_bytes(wb))
