# 🥋 Planificación del entrenamiento de Taekwondo

Aplicación web (Streamlit + Python) que genera el **plan de entrenamiento personalizado** de un
deportista de taekwondo a partir de las **curvas de Matveyev**:

- Ciclo **ANUAL** (48 semanas) o **SEMESTRAL** (24 semanas).
- **Volúmenes** e **intensidades** semanales de los ejercicios de formación General (EFG),
  Especial (EFE) y Competitiva (EFC), con el **pulso de entrenamiento (VPE)** de cada semana.
- **Microciclos** semanales de cada etapa con sus fechas, el método de pesas y el **peso de las
  canilleras** según el peso corporal.
- **Índice de sobreentrenamiento** (pulso acostado vs. de pie; alerta si la diferencia supera 15).
- Descargas: **PDF** del plan, **Excel (.xlsx)** sin macros y **Excel con macro (.xlsm)** con menú.

> Plantilla original y metodología: **Justo Fuentes** (Universidad de Sucre).

---

## Fórmulas

| Cálculo | Fórmula |
|---|---|
| FC máxima | 220 − edad (hombres) · 226 − edad (mujeres) |
| 100 % de esfuerzo | FC máxima − VPR |
| VPE (pulso de entrenamiento) | % Intensidad × (FC máxima − VPR) / 100 + VPR |
| Volumen de un ejercicio | Cantidad al 100 % × % Volumen / 100 |
| Canillera (g por pierna) | (peso × 0,25 / 2) × % × 1000 |

**Ciclo semestral:** cada semana de una etapa corresponde, en proporción, a un grupo de semanas de
la misma etapa del ciclo anual, y se toma el promedio (sin contar ceros) de sus porcentajes.
Ejemplo: la semana 1 del semestre promedia las semanas 1 y 2 del año.

## Estructura del repositorio

```
tkd-planificacion/
├── app.py                    # interfaz Streamlit
├── motor/
│   ├── calculos.py           # fórmulas, ciclo anual/semestral, etapas y fechas
│   ├── pdf.py                # PDF del plan (reportlab)
│   ├── excel.py              # libro .xlsx / .xlsm (openpyxl)
│   └── vba.py                # crea el proyecto VBA (macro) sin necesidad de Excel
├── datos/
│   ├── curvas.json           # % de volumen e intensidad de las 48 semanas (Matveyev)
│   ├── periodizacion.json    # etapa de cada semana (anual y semestral)
│   ├── microciclos.json      # microciclos EPG, EPE y EPC
│   └── plantilla_base.xlsx   # libro original (fórmulas, curvas, gráficos)
├── vba/                      # código de la macro del menú (modMenu.bas, etc.)
├── tests/test_calculos.py    # pruebas automáticas
├── .devcontainer/            # configuración para GitHub Codespaces
└── requirements.txt
```

## Cómo usarla

### Opción A: en GitHub Codespaces (recomendada, sin instalar nada)
1. Suba este repositorio a GitHub.
2. Botón verde **Code → Codespaces → Create codespace on main**.
3. Espere a que termine la instalación: la app se abre sola en el puerto **8501**.
   Si no se abre, en la terminal escriba: `streamlit run app.py`

### Opción B: en su computador
```bash
python -m venv .venv
.venv\Scripts\activate          # Windows   (en Linux/Mac: source .venv/bin/activate)
pip install -r requirements.txt
streamlit run app.py
```

### Opción C: publicarla en Internet gratis (Streamlit Community Cloud)
1. Entre a <https://share.streamlit.io> con su cuenta de GitHub.
2. **Create app → Deploy a public app from GitHub**.
3. Repositorio: el suyo · Rama: `main` · Archivo principal: `app.py` → **Deploy**.
4. Comparta el enlace con los entrenadores: funciona en computador, tableta y celular.

## Cómo modificar el contenido (sin tocar el código)

- **Microciclos:** edite `datos/microciclos.json` (texto de mañana y noche de cada día).
- **Etapas de cada semana:** edite `datos/periodizacion.json`. Las semanas de una misma etapa
  deben ir seguidas. Las etapas válidas son: Preparatoria General, Preparatoria Especial,
  Precompetitiva, Competitiva y Transición.
- **Curvas de Matveyev:** edite `datos/curvas.json` (48 valores de volumen y 48 de intensidad
  para EFG, EFE y EFC).

Después de un cambio, ejecute las pruebas: `pytest -q`

## Notas

- Los registros del índice de sobreentrenamiento **no se guardan en el servidor**: se incluyen en
  el PDF descargado. Para un registro permanente, use el libro de Excel.
- Si el archivo **.xlsm** llega por Internet o por correo, Windows puede bloquear la macro:
  clic derecho → **Propiedades** → marcar **Desbloquear** → Aceptar; al abrirlo, **Habilitar contenido**.
- Las etapas **Competitiva** y **Transición** aún no tienen microciclo propio (se planifican según el
  calendario de competencias).
- Esta herramienta es una guía aproximada para personalizar las cargas y el pulso de entrenamiento;
  no reemplaza la valoración médica.
