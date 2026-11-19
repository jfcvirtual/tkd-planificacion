"""Planificación del entrenamiento de Taekwondo - curvas de Matveyev.

App de Streamlit: el entrenador ingresa los datos del deportista, elige ciclo ANUAL o
SEMESTRAL y descarga el plan en PDF, en Excel (.xlsx) o en Excel con macro (.xlsm).

Ejecutar:  streamlit run app.py
"""
import datetime as dt

import pandas as pd
import streamlit as st

from motor.calculos import (EJERCICIOS, MICROCICLOS, NOMBRE_FORMACION, SEMANAS, Deportista,
                            Parametros, calcular_plan, indice_sobreentrenamiento, lunes_de)
from motor.excel import libro_xlsm, libro_xlsx
from motor.pdf import generar_pdf

st.set_page_config(page_title="Plan de entrenamiento TKD", page_icon="🥋", layout="wide")

COLOR_ETAPA = {"Preparatoria General": "#E2EFDA", "Preparatoria Especial": "#FFF2CC",
               "Precompetitiva": "#FCE4D6", "Competitiva": "#F8CBAD", "Transición": "#EDEDED"}
TITULO_FORMACION = {"EFG": "Formación General (EFG)", "EFE": "Formación Especial (EFE)",
                    "EFC": "Formación Competitiva (EFC)"}
CIEN_DEFECTO = {"cross": 90.0, "escalera": 70.0, "patadas": 200.0, "saco": 150.0, "canillera": 120.0,
                "combate_condicionado": 16.0, "combate_libre": 12.0}


# ---------------------------------------------------------------- utilidades
def fecha(d):
    return d.strftime("%d/%m/%Y") if d else ""


def nombre_archivo(texto: str) -> str:
    limpio = "".join(c if c.isalnum() else "_" for c in texto.strip())
    return "_".join(p for p in limpio.split("_") if p) or "deportista"


def parametros_desde_clave(clave):
    ciclo, inicio, nombre, edad, sexo, peso, vpr, vpb, cien = clave
    return Parametros(ciclo, inicio, Deportista(nombre, edad, sexo, peso, vpr, vpb, dict(cien)))


@st.cache_data(show_spinner=False)
def excel_xlsx(clave) -> bytes:
    return libro_xlsx(parametros_desde_clave(clave))


@st.cache_data(show_spinner=False)
def excel_xlsm(clave) -> bytes:
    return libro_xlsm(parametros_desde_clave(clave))


def pdf_plan(clave, registros) -> bytes:
    return generar_pdf(calcular_plan(parametros_desde_clave(clave)), registros)


# ---------------------------------------------------------------- encabezado
st.title("🥋 Planificación del entrenamiento de Taekwondo")
st.caption("Volúmenes, intensidades y microciclos semanales según las curvas de Matveyev · "
           "Plantilla original de Justo Fuentes")

with st.sidebar:
    st.header("¿Cómo funciona?")
    st.markdown(
        "1. Elija **ciclo anual** (48 semanas) o **semestral** (24 semanas).\n"
        "2. Escriba los datos del deportista y sus **cantidades al 100 %**.\n"
        "3. Presione **Calcular plan**.\n"
        "4. Revise las pestañas y **descargue** el PDF o el Excel.")
    st.divider()
    st.markdown("**Fórmulas**")
    st.markdown("- FC máx = 220 − edad (H) · 226 − edad (M)\n"
                "- VPE = % Intensidad × (FC máx − VPR) / 100 + VPR\n"
                "- Volumen = cantidad al 100 % × % Volumen / 100")
    st.divider()
    st.caption("Esta herramienta es una guía aproximada para personalizar el pulso y las cargas de "
               "entrenamiento; no reemplaza la valoración médica.")

# ---------------------------------------------------------------- formulario
with st.form("datos"):
    st.subheader("1. Tipo de ciclo")
    c1, c2 = st.columns([1, 1])
    ciclo = c1.radio("¿Qué desea planificar?", ["ANUAL", "SEMESTRAL"], horizontal=True,
                     format_func=lambda c: f"Ciclo {c.lower()} ({SEMANAS[c]} semanas)")
    inicio = c2.date_input("Fecha de inicio del ciclo (se recomienda un lunes)",
                           value=lunes_de(dt.date.today() + dt.timedelta(days=7)), format="DD/MM/YYYY")
    ajustar = c2.checkbox("Si la fecha no es lunes, empezar el lunes de esa semana", value=True)

    st.subheader("2. Datos del deportista")
    c1, c2, c3, c4 = st.columns([2.2, 1, 1, 1])
    nombre = c1.text_input("Nombre completo")
    edad = c2.number_input("Edad (años)", 5, 90, 20)
    sexo = c3.selectbox("Sexo", ["M", "F"], format_func=lambda s: "Masculino" if s == "M" else "Femenino")
    peso = c4.number_input("Peso corporal (kg)", 15.0, 200.0, 60.0, step=0.5)
    c1, c2, _ = st.columns([1, 1, 2])
    vpr = c1.number_input("VPR - pulso en reposo (ppm)", 30, 120, 64,
                          help="Pulsaciones por minuto, sentado y tranquilo.")
    vpb = c2.number_input("VPB - pulso basal (ppm)", 30, 120, 60,
                          help="Al despertar, antes de levantarse.")

    st.subheader("3. Cantidades al 100 % de esfuerzo (máximo estimado del deportista)")
    cien = {}
    cols = st.columns(3)
    for col, (form, ejer) in zip(cols, EJERCICIOS.items()):
        col.markdown(f"**{TITULO_FORMACION[form]}**")
        for clave, etiqueta in ejer:
            cien[clave] = col.number_input(etiqueta, 0.1, 100000.0, CIEN_DEFECTO[clave], step=1.0, key=clave)

    enviado = st.form_submit_button("Calcular plan", type="primary", width="stretch")

if enviado:
    if not nombre.strip():
        st.error("Escriba el nombre del deportista.")
        st.stop()
    if ajustar:
        inicio = lunes_de(inicio)
    st.session_state["clave"] = (ciclo, inicio, nombre.strip(), float(edad), sexo, float(peso), float(vpr),
                                 float(vpb), tuple(sorted(cien.items())))

if "clave" not in st.session_state:
    st.info("Complete el formulario y presione **Calcular plan**.")
    st.stop()

# ---------------------------------------------------------------- resultados
clave = st.session_state["clave"]
p = parametros_desde_clave(clave)
plan = calcular_plan(p)
d = p.deportista

st.divider()
st.subheader(f"Plan de {d.nombre} · ciclo {p.ciclo.lower()} ({SEMANAS[p.ciclo]} semanas)")
st.markdown(f"**Del {fecha(p.inicio)} al {fecha(plan['fin'])}** · Edad: {d.edad:.0f} años · "
            f"Sexo: {'Femenino' if d.sexo == 'F' else 'Masculino'} · Peso: {d.peso:g} kg · "
            f"VPR: {d.vpr:.0f} ppm · VPB: {d.vpb:.0f} ppm")
m = st.columns(3)
m[0].metric("FC máxima estimada", f"{plan['fc_maxima']:.0f} ppm")
m[1].metric("100 % de esfuerzo (FC máx − VPR)", f"{plan['esfuerzo_100']:.0f} ppm")
m[2].metric("Canilleras 2 % / 3 % (por pierna)", f"{plan['canilleras'][2]:.0f} / {plan['canilleras'][3]:.0f} g")


def tabla_formacion(form):
    filas = []
    for s in plan["semanas"]:
        f = s[form]
        fila = {"Semana": s["semana"], "Mes": s["mes"], "Desde": fecha(s["desde"]), "Hasta": fecha(s["hasta"]),
                "Etapa": s["etapa"], "% Volumen": round(f["pct_volumen"], 1)}
        for (c, n), v in zip(EJERCICIOS[form], f["volumenes"].values()):
            fila[n] = round(v, 1)
        fila["% Intensidad"] = round(f["pct_intensidad"], 1)
        fila["VPE (ppm)"] = round(f["vpe"]) if f["vpe"] is not None else None
        filas.append(fila)
    df = pd.DataFrame(filas)
    return df.style.apply(lambda r: [f"background-color: {COLOR_ETAPA[r['Etapa']]}"] * len(r), axis=1) \
                   .format(precision=1, na_rep="-").format({"VPE (ppm)": "{:.0f}"}, na_rep="-")


tabs = st.tabs(["📋 Resumen", "EFG", "EFE", "EFC", "📅 Microciclos", "📈 Gráficas",
                "❤️ Índice de sobreentrenamiento", "⬇️ Descargas"])

with tabs[0]:
    st.markdown("**Etapas del ciclo**")
    et = pd.DataFrame([{
        "Etapa": e["etapa"], "Semanas": e["semanas"],
        "Semana inicial": e["primera"] or "", "Semana final": e["ultima"] or "",
        "Desde": fecha(e["desde"]), "Hasta": fecha(e["hasta"]),
        "Microciclo": MICROCICLOS[e["microciclo"]]["hoja"] if e["microciclo"] else "(según calendario)",
    } for e in plan["etapas"]])
    st.dataframe(et.style.apply(lambda r: [f"background-color: {COLOR_ETAPA[r['Etapa']]}"] * len(r), axis=1),
                 hide_index=True, width="stretch")
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("**Cantidades al 100 %**")
        st.dataframe(pd.DataFrame([{"Formación": f, "Ejercicio": n, "Cantidad": d.cien[c]}
                                   for f, ej in EJERCICIOS.items() for c, n in ej]),
                     hide_index=True, width="stretch")
    with c2:
        st.markdown("**Peso de cada canillera (gramos por pierna)**")
        st.dataframe(pd.DataFrame([{"%": f"{k} %", "Gramos": round(v)} for k, v in plan["canilleras"].items()]),
                     hide_index=True, width="stretch")

for i, form in enumerate(EJERCICIOS, start=1):
    with tabs[i]:
        st.markdown(f"**{NOMBRE_FORMACION[form]}**")
        st.dataframe(tabla_formacion(form), hide_index=True, width="stretch", height=560)

with tabs[4]:
    for codigo in ("EPG", "EPE", "EPC"):
        mc = MICROCICLOS[codigo]
        e = next(x for x in plan["etapas"] if x["etapa"] == mc["etapa"])
        rango = (f"{e['semanas']} semanas · del {fecha(e['desde'])} al {fecha(e['hasta'])}"
                 if e["semanas"] else "no forma parte de este ciclo")
        with st.expander(f"{mc['titulo']} — {rango}", expanded=(codigo == "EPG")):
            st.caption(mc["correlacion"])
            st.table(pd.DataFrame([{"Día": x["dia"], "Mañana": x["manana"], "Noche": x["noche"]}
                                   for x in mc["dias"]]).set_index("Día"))
            pesas = mc.get("pesas")
            if pesas and "pasos" in pesas:
                st.markdown(f"**{pesas['titulo']}:** " + " → ".join(f"{r} rep." for r in pesas["pasos"]))
                st.warning(pesas["nota"])
            elif pesas:
                st.markdown(f"**{pesas['titulo']}:** {pesas['texto']}")
            if mc.get("canilleras"):
                st.markdown("**Canilleras (gramos por pierna):** " + " · ".join(
                    f"{k} % = {v:.0f} g" for k, v in plan["canilleras"].items()))
    st.caption("Las etapas Competitiva y Transición se planifican según el calendario de competencias.")

with tabs[5]:
    for form in EJERCICIOS:
        st.markdown(f"**{NOMBRE_FORMACION[form]}** - % de volumen e intensidad por semana")
        st.line_chart(pd.DataFrame({
            "% Volumen": [s[form]["pct_volumen"] for s in plan["semanas"]],
            "% Intensidad": [s[form]["pct_intensidad"] for s in plan["semanas"]],
        }, index=pd.Index([s["semana"] for s in plan["semanas"]], name="Semana")), height=240)

with tabs[6]:
    st.markdown("Tome el pulso **acostado** al despertar y luego inmediatamente **de pie**. "
                "Si la diferencia supera **15 pulsaciones**, es señal de alerta.")
    if "indice" not in st.session_state:
        st.session_state["indice"] = pd.DataFrame(
            {"Fecha": pd.Series(dtype="datetime64[ns]"), "Pulso acostado": pd.Series(dtype="float"),
             "Pulso de pie": pd.Series(dtype="float")})
    reg = st.data_editor(st.session_state["indice"], num_rows="dynamic", width="stretch",
                         column_config={"Fecha": st.column_config.DateColumn(format="DD/MM/YYYY"),
                                        "Pulso acostado": st.column_config.NumberColumn(min_value=25, max_value=200),
                                        "Pulso de pie": st.column_config.NumberColumn(min_value=25, max_value=220)},
                         key="editor_indice")
    registros = []
    for _, r in reg.dropna().iterrows():
        dif, alerta = indice_sobreentrenamiento(r["Pulso acostado"], r["Pulso de pie"])
        registros.append({"fecha": pd.Timestamp(r["Fecha"]).date(), "acostado": r["Pulso acostado"],
                          "de_pie": r["Pulso de pie"]})
        if alerta:
            st.error(f"{fecha(pd.Timestamp(r['Fecha']).date())}: diferencia de {dif:.0f} pulsaciones. "
                     "Posible riesgo de sobreentrenamiento: revisar sueño y trasnochos, consumo de alcohol, "
                     "cargas recientes o malestar. Si se repite, consultar al médico.")
    if registros and not any(r["de_pie"] - r["acostado"] > 15 for r in registros):
        st.success("Todas las mediciones están dentro de lo normal.")
    st.caption("Los registros no se guardan en el servidor: se incluyen en el PDF que descargue. "
               "Para un registro permanente use el libro de Excel.")

with tabs[7]:
    base = f"Plan_{nombre_archivo(d.nombre)}_{p.ciclo.lower()}"
    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown("**PDF del plan completo**")
        st.caption("Resumen, EFG, EFE, EFC, microciclos y planilla del índice de sobreentrenamiento.")
        with st.spinner("Preparando PDF..."):
            st.download_button("⬇️ Descargar PDF", pdf_plan(clave, registros), f"{base}.pdf",
                               "application/pdf", type="primary", width="stretch")
    with c2:
        st.markdown("**Excel (.xlsx), sin macros**")
        st.caption("Ya trae los datos del deportista. Se abre sin habilitar contenido.")
        with st.spinner("Preparando Excel..."):
            st.download_button("⬇️ Descargar Excel", excel_xlsx(clave), f"{base}.xlsx",
                               "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               width="stretch")
    with c3:
        st.markdown("**Excel con macro (.xlsm)**")
        st.caption("Incluye el MENÚ para crear nuevos planes, registrar el índice e imprimir. "
                   "Al abrirlo: «Habilitar contenido».")
        with st.spinner("Preparando Excel con macro..."):
            st.download_button("⬇️ Descargar Excel con macro", excel_xlsm(clave), f"{base}.xlsm",
                               "application/vnd.ms-excel.sheet.macroEnabled.12", width="stretch")
    st.info("Si el archivo .xlsm se descarga de Internet, Windows puede bloquear la macro: clic derecho "
            "sobre el archivo → Propiedades → marcar «Desbloquear» → Aceptar.")
