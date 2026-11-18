"""Motor Start-up Studio | Streamlit adaptation of the original Tkinter calculator.

The calculation equations intentionally follow the source application. Values and
units are kept together; spreadsheet curves are never fabricated.
"""
from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from io import BytesIO
from pathlib import Path
from typing import Any
import html
import math
import os
import re
import urllib.request

import pandas as pd
import streamlit as st
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, KeepTogether

ROOT = Path(__file__).parent
DEFAULT_V47 = "PMD NEMA FS 56 to FS 500 - V47-R1.xlsx"
DEFAULT_BASE = "base_motores.xlsx"
COLUMN_DEFAULTS = {"MLFB": "A", "HP": "C", "INERCIA": "AS", "CURVA": "BT", "HOT": "AN", "COLD": "AO"}
MLFB_DESCRIPTIONS = {
    1: {"1": "Motor Series: Standard or Definite Purpose"},
    2: {"L": "Standard GP/SD", "P": "Definite Purpose"},
    3: {"E": "Standard GP/SD", "C": "Definite Purpose"},
    4: {"4": "SD200", "6": "Next Gen NEMA"},
    5: {"3": "SD200", "5": "DP200 HPS"},
    6: {"2": "Eficiencia estándar"},
    7: {"1": "Tipo básico", "2": "Tipo mejorado"},
    8: {"4": "Frame 444-449", "5": "Frame 5000s"},
    9: {"B": "Shaft 444-445T", "C": "Shaft 447-449T", "D": "Shaft L449T", "A": "Shaft 509-5011", "F": "Shaft 444-445TS", "G": "Shaft 447-449TS"},
    10: {"A": "2 polos", "B": "4 polos", "C": "6 polos", "D": "8 polos"},
}


# Los archivos originales se descargan automáticamente desde el repositorio público
# de origen cuando aún no están presentes en el repositorio de despliegue.
SOURCE_ASSETS = (
    DEFAULT_V47, DEFAULT_BASE, "promedios.xlsx",
    "logo.png", "motor_header.png", "fondo_motor.png", "logo.ico",
)
SOURCE_REPO_RAW = (
    "https://raw.githubusercontent.com/yahir21547/"
    "CALCULO-DE-TIEMPO-DE-ARRANQUE-STREAMLIT/main/"
)

@st.cache_resource(show_spinner=False)
def hydrate_assets() -> tuple[str, ...]:
    failed = []
    for name in SOURCE_ASSETS:
        path = ROOT / name
        if path.is_file() and path.stat().st_size:
            continue
        from urllib.parse import quote
        url = SOURCE_REPO_RAW + quote(name)
        try:
            req = urllib.request.Request(url, headers={"User-Agent": "motor-startup-streamlit/2"})
            with urllib.request.urlopen(req, timeout=30) as response:
                payload = response.read()
            if not payload:
                raise ValueError("Archivo vacío")
            path.write_bytes(payload)
        except Exception:
            failed.append(name)
    return tuple(failed)


def clean_mlfb(v: Any) -> str:
    return re.sub(r"[^A-Z0-9]", "", str(v).upper().strip()) if pd.notna(v) else ""


def col_index(letter: str) -> int:
    s = letter.strip().upper()
    if not re.fullmatch(r"[A-Z]+", s):
        raise ValueError(f"Columna Excel inválida: {letter!r}")
    i = 0
    for c in s:
        i = i * 26 + ord(c) - ord("A") + 1
    return i - 1


def parse_num(value: Any) -> float | None:
    if value is None or pd.isna(value):
        return None
    try:
        if isinstance(value, str):
            value = value.strip().replace(" ", "").replace(",", "")
            if not value or value in ("-", "—", "N/A", "#N/A"):
                return None
        f = float(value)
        return f if math.isfinite(f) else None
    except (ValueError, TypeError, OverflowError):
        return None


def display(v: float | None, decimals: int = 5) -> str:
    if v is None or not math.isfinite(v):
        return "No disponible"
    return f"{v:,.{decimals}f}".rstrip("0").rstrip(".")


@st.cache_data(show_spinner=False)
def workbook_sheets(data: bytes) -> list[str]:
    return pd.ExcelFile(BytesIO(data), engine="openpyxl").sheet_names


@st.cache_data(show_spinner=False)
def read_sheet(data: bytes, sheet: str, header_row: int) -> pd.DataFrame:
    return pd.read_excel(BytesIO(data), sheet_name=sheet, header=header_row - 1, engine="openpyxl", dtype=object)


def read_local(name: str) -> bytes | None:
    fp = ROOT / name
    try:
        return fp.read_bytes() if fp.is_file() else None
    except OSError:
        return None


def lookup(df: pd.DataFrame, code: str, mapping: dict[str, str]) -> dict[str, Any] | None:
    ix = col_index(mapping["MLFB"])
    if ix >= df.shape[1]:
        raise ValueError(f"La columna MLFB {mapping['MLFB']} no existe en esa hoja ({df.shape[1]} columnas)")
    found = df.loc[df.iloc[:, ix].map(clean_mlfb) == clean_mlfb(code)]
    if found.empty:
        return None
    row = found.iloc[0]
    result = {"_row": int(found.index[0])}
    for key, letter in mapping.items():
        idx = col_index(letter)
        result[key] = row.iloc[idx] if idx < len(row) else None
    return result


def infer_rpm(code: str) -> float | None:
    # 60 Hz synchronous reference RPM, as in the desktop source code.
    return {"A": 3600., "B": 1800., "C": 1200., "D": 900.}.get(code[9]) if len(code) >= 10 else None


@dataclass
class Calculation:
    values: dict[str, float]
    trace: list[tuple[str, str, float]]
    status: str
    status_detail: str
    raw_wr2: float
    wr2_unit: str
    raw_mnv: float
    mnv_unit: str


def calculate(*, hp: float, motor_rpm: float, motor_inertia: float, motor_factor: float,
              eq_rpm: float, load_inertia: float, load_torque: float, load_factor: float,
              hot: float | None, cold: float | None, raw_wr2: float, wr2_unit: str,
              raw_mnv: float, mnv_unit: str) -> Calculation:
    numeric = (hp, motor_rpm, motor_inertia, motor_factor, eq_rpm, load_inertia, load_torque, load_factor)
    if not all(math.isfinite(v) for v in numeric):
        raise ValueError("Los datos del cálculo deben ser numéricos y finitos.")
    if hp <= 0 or motor_rpm <= 0 or eq_rpm <= 0:
        raise ValueError("HP y ambas velocidades deben ser mayores que cero.")
    if motor_inertia < 0 or load_inertia < 0 or load_torque < 0:
        raise ValueError("Las inercias y el torque de la carga no pueden ser negativos.")
    if motor_factor <= 0 or load_factor < 0:
        raise ValueError("Los factores de torque deben ser válidos (motor > 0, carga >= 0).")
    mnm = hp * 726.15 / motor_rpm
    mmm = mnm * motor_factor
    mmv = load_torque * load_factor
    jv = load_inertia
    gd2m = 4 * motor_inertia
    gd2v = 4 * jv
    ratio = eq_rpm / motor_rpm
    mmv_motor = round(mmv * ratio, 2)  # original application rounding
    gd2v_motor = round(gd2v * ratio ** 2, 2)  # original application rounding
    ngd2 = gd2m + gd2v_motor
    accelerating_torque = mmm - mmv_motor
    if accelerating_torque <= 0:
        raise ValueError("El torque medio del motor no supera al torque de carga referido. No existe un tiempo de arranque positivo válido con esta aproximación.")
    ta_raw = ngd2 * motor_rpm / (375 * accelerating_torque)
    if not math.isfinite(ta_raw) or ta_raw < 0:
        raise ValueError("Tiempo de arranque inválido; comprueba los valores de entrada.")
    # The original UI compared the 2-decimal display ta to thermal limits.
    ta = round(ta_raw, 2)
    ceiling = math.ceil(ta_raw)
    vals = dict(HP=hp, RPM_MOTOR=motor_rpm, RPM_EQ=eq_rpm, JM=motor_inertia,
                MNM=mnm, MMM=mmm, MNV=load_torque, MMV=mmv, JV=jv,
                GD2M=gd2m, GD2V=gd2v, MMV_NMOT=mmv_motor,
                GD2V_NMOT=gd2v_motor, NGD2=ngd2, TA_RAW=ta_raw,
                TA=ta, TA_CEIL=ceiling, HOT=hot, COLD=cold,
                FACTOR_MMM=motor_factor, FACTOR_MMV=load_factor)
    if hot is None or cold is None or hot <= 0 or cold <= 0:
        status, detail = "Sin evaluación térmica", "Se calculó el tiempo, pero faltan límites HOT y/o COLD válidos."
    elif ta_raw < hot and ta_raw < cold:
        status, detail = "Apto HOT y COLD", "El tiempo calculado es menor a ambos límites térmicos."
    elif ta_raw < cold:
        status, detail = "Solo apto COLD", "El tiempo calculado es menor al límite COLD, pero no al HOT."
    elif ta_raw < hot:
        status, detail = "Solo apto HOT", "El tiempo calculado es menor al límite HOT, pero no al COLD."
    else:
        status, detail = "No apto", "El tiempo calculado no es menor a ninguno de los límites térmicos."
    trace = [
        ("MNM", "HP × 726.15 / RPM motor", mnm),
        ("MMM", "MNM × factor de curva", mmm),
        ("MMV", "MNV × factor de carga", mmv),
        ("JM", "Inercia de motor [kg·m²]", motor_inertia),
        ("GD²M", "4 × JM", gd2m),
        ("JV", "Inercia equivalente de carga [kg·m²]", jv),
        ("GD²V", "4 × JV", gd2v),
        ("MMV referido", "round(MMV × RPM equipo / RPM motor, 2)", mmv_motor),
        ("GD²V referido", "round(GD²V × (RPM equipo / RPM motor)², 2)", gd2v_motor),
        ("nGD²", "GD²M + GD²V referido", ngd2),
        ("ta", "nGD² × RPM motor / [375 × (MMM − MMV referido)]", ta_raw),
        ("ta CEILING", "ceil(ta)", float(ceiling)),
    ]
    return Calculation(vals, trace, status, detail, raw_wr2, wr2_unit, raw_mnv, mnv_unit)


def build_pdf(code: str, result: Calculation, catalog: str, curve_source: str) -> bytes:
    out = BytesIO()
    pdf = SimpleDocTemplate(out, pagesize=A4, leftMargin=35, rightMargin=35, topMargin=40, bottomMargin=38,
                            title="Cálculo de tiempo de arranque", author="Motor Start-up Studio")
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name="HeroMotor", parent=styles["Title"], fontSize=17, leading=21, textColor=colors.HexColor("#E12B32"), spaceAfter=9))
    styles.add(ParagraphStyle(name="SubMotor", parent=styles["Normal"], fontSize=9, leading=14, textColor=colors.HexColor("#475569")))
    styles.add(ParagraphStyle(name="SectionMotor", parent=styles["Heading2"], fontSize=11, textColor=colors.HexColor("#202938"), spaceBefore=16, spaceAfter=7))
    story = [Paragraph("MOTOR START-UP | REPORTE TÉCNICO", styles["HeroMotor"]),
             Paragraph(f"MLFB: <b>{html.escape(code or 'Manual')}</b> &nbsp; | &nbsp; Fecha: {datetime.now():%d/%m/%Y %H:%M}", styles["SubMotor"]),
             Paragraph(f"Catálogo: {html.escape(catalog)}", styles["SubMotor"]),
             Paragraph(f"Origen del factor de curva: {html.escape(curve_source)}", styles["SubMotor"]),
             Spacer(1, 12)]
    v = result.values
    sections = [
        ("Datos del motor", [
            ("Potencia", v["HP"], "HP"), ("RPM nominal de referencia", v["RPM_MOTOR"], "rpm"),
            ("Torque nominal MNM", v["MNM"], "kgf·m"), ("Inercia rotor JM", v["JM"], "kg·m²"),
            ("GD² motor", v["GD2M"], "kg·m² (convención del modelo)"),
            ("Factor de curva", v["FACTOR_MMM"], "fracción"), ("Torque medio MMM", v["MMM"], "kgf·m"),
            ("Límite HOT", v["HOT"], "s"), ("Límite COLD", v["COLD"], "s")]),
        ("Equipo accionado y conversiones", [
            ("Velocidad equipo", v["RPM_EQ"], "rpm"),
            ("Inercia introducida", result.raw_wr2, result.wr2_unit),
            ("Inercia normalizada JV", v["JV"], "kg·m²"),
            ("Torque introducido", result.raw_mnv, result.mnv_unit),
            ("Torque normalizado MNV", v["MNV"], "kgf·m"),
            ("Factor MMV", v["FACTOR_MMV"], "fracción"),
            ("Torque medio MMV", v["MMV"], "kgf·m"),
            ("GD² carga", v["GD2V"], "kg·m² (convención del modelo)"),
            ("MMV referido al motor", v["MMV_NMOT"], "kgf·m"),
            ("GD² carga referido", v["GD2V_NMOT"], "kg·m²"),
            ("nGD² total", v["NGD2"], "kg·m²")]),
        ("Resultado", [
            ("Tiempo sin redondear", v["TA_RAW"], "s"),
            ("Tiempo presentado", v["TA"], "s"),
            ("Tiempo hacia arriba (CEILING)", v["TA_CEIL"], "s")])]
    for title, rows in sections:
        story.append(Paragraph(title, styles["SectionMotor"]))
        data = [["Parámetro", "Valor", "Unidad"]] + [[str(l), display(val, 6), unit] for l, val, unit in rows]
        table = Table(data, colWidths=[230, 93, 170], repeatRows=1, hAlign="LEFT")
        table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F2937")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F3F6FA")]),
            ("LINEBELOW", (0, -1), (-1, -1), 0.5, colors.HexColor("#E2E8F0")),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("FONTSIZE", (0, 0), (-1, -1), 8), ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
            ("TOPPADDING", (0, 0), (-1, -1), 7),
        ]))
        story.append(table)
    story.append(Paragraph("Evaluación de arranque", styles["SectionMotor"]))
    story.append(Paragraph(f"<b>{html.escape(result.status)}</b>. {html.escape(result.status_detail)}", styles["Normal"]))
    story.append(Paragraph("Operaciones utilizadas", styles["SectionMotor"]))
    for label, formula, value in result.trace:
        story.append(Paragraph(f"<b>{html.escape(label)}:</b> {html.escape(formula)} = {display(value, 6)}", styles["SubMotor"]))
    story.append(Spacer(1, 12))
    story.append(Paragraph("Nota: estimación mediante la fórmula y criterios del aplicativo original. Verificar unidades, curvas y datos de placa antes de decisiones de ingeniería.", styles["SubMotor"]))
    def decorate(canvas, doc):
        canvas.saveState()
        # Recursos gráficos del aplicativo original, cuando están en GitHub.
        for filename, x, w, h in (("logo.png", 35, 63, 28), ("motor_header.png", 492, 62, 28)):
            image_path = ROOT / filename
            if image_path.is_file():
                try:
                    canvas.drawImage(str(image_path), x, 798, width=w, height=h,
                                     preserveAspectRatio=True, mask="auto", anchor="c")
                except Exception:
                    pass
        canvas.setStrokeColor(colors.HexColor("#E12B32"))
        canvas.setLineWidth(2)
        canvas.line(35, 790, 560, 790)
        canvas.setFont("Helvetica", 7)
        canvas.setFillColor(colors.HexColor("#64748B"))
        canvas.drawCentredString(297.5, 22, f"Reporte de tiempo de arranque · Página {doc.page}")
        canvas.restoreState()
    pdf.build(story, onFirstPage=decorate, onLaterPages=decorate)
    return out.getvalue()



# -----------------------------------------------------------------------------
# Original Tkinter layout adapted to Streamlit, without changing calculation core.
# -----------------------------------------------------------------------------

def theme() -> None:
    st.markdown("""<style>
    :root { color-scheme: dark; }
    [data-testid="stAppViewContainer"], .stApp {background:#10151d;color:#f0f4fa;}
    [data-testid="stHeader"] {background:#10151d;}
    [data-testid="stSidebar"] {background:#19212b;border-right:1px solid #344150;}
    .block-container {max-width:1580px;padding-top:1.1rem;padding-bottom:3rem;}
    h1,h2,h3,p,label,small {color:#eef3fa;}
    .app-banner {background:linear-gradient(115deg,#202c3b,#141d29 67%,#311c25); border:1px solid #3d4a5c;
                 border-left:6px solid #ec1e24;border-radius:13px;padding:17px 22px;margin-bottom:12px;}
    .app-banner h1 {font-size:1.75rem!important;color:#fff;margin:0;font-weight:850;}
    .app-banner p {color:#cbd5e1;margin:6px 0 0;font-size:.93rem;}
    .section-bar {background:#273343;border:1px solid #415165;border-left:5px solid #ee262e;
        border-radius:9px;padding:8px 13px;margin:16px 0 10px;font-weight:800;color:#fff;letter-spacing:.02em;}
    .small-note {color:#bdc9d8;font-size:.83rem;line-height:1.45;}
    .step-line {background:#1d2936;border-radius:8px;padding:9px 12px;color:#e0eaf5;border:1px solid #394a5c;margin:6px 0;}
    .step-ok {border-color:#36775e;background:#163128;color:#baffd5;}
    .unit-pill {display:inline-block;background:#35465b;color:#fff;padding:3px 9px;border-radius:6px;font-size:.8rem;}
    div[data-testid="stVerticalBlockBorderWrapper"] > div {background:#19232f;border-color:#394b60 !important;}
    div[data-testid="stTextInput"] input, div[data-testid="stNumberInput"] input,
    div[data-baseweb="select"] > div {background:#253243 !important;color:#f5f8ff !important;border:1px solid #51647a!important;}
    div[data-testid="stTextInput"] input:disabled, div[data-testid="stNumberInput"] input:disabled {opacity:.78;color:#bacada!important;background:#26313e!important;}
    div[data-testid="stNumberInput"] input:focus, div[data-testid="stTextInput"] input:focus {border-color:#ff6066!important;}
    [data-testid="stWidgetLabel"] p {color:#edf4ff!important;font-weight:670;font-size:.89rem;}
    [data-testid="stMetric"] {background:#253242;border:1px solid #4d6076;border-radius:11px;padding:14px;}
    [data-testid="stMetricLabel"] p {color:#d7e2ee!important;}
    [data-testid="stMetricValue"] {color:#fff;}
    div.stButton > button {background:#2b3b4e;color:#fff;border:1px solid #6b7c8d;font-weight:700;}
    div.stButton > button:hover {border-color:#ff545c;color:#fff;background:#384b60;}
    div.stButton > button[kind="primary"], div.stDownloadButton > button[kind="primary"] {
      background:#dd262c;color:#fff;border:1px solid #ff585e;font-weight:850;}
    div.stButton > button:disabled {background:#27313e!important;color:#8a9bad!important;border-color:#344252!important;}
    [data-testid="stAlert"] {border-radius:10px;}
    [data-testid="stExpander"] {border:1px solid #475c75;border-radius:10px;background:#1b2836;}
    [data-testid="stTabs"] button {color:#d4dce7;font-weight:700;}
    .field-group {color:#edf6ff;font-size:1.06rem;font-weight:800;border-bottom:1px solid #46586a;padding:5px 0 10px;margin:0 0 9px;}
    </style>""", unsafe_allow_html=True)


def section(label: str) -> None:
    st.markdown(f'<div class="section-bar">{html.escape(label)}</div>', unsafe_allow_html=True)


def positive_state(key: str) -> bool:
    try:
        return float(st.session_state.get(key, 0)) > 0
    except (TypeError, ValueError):
        return False


def jump_step(n: int) -> None:
    st.session_state.stage = max(int(st.session_state.get('stage', 0)), n)


def clear_form() -> None:
    # The widget values are cleared before the next Streamlit rerun.
    for key in ('mlfb', 'hp', 'motor_rpm', 'jm', 'factor_mmm', 'hot', 'cold',
                'eq_rpm', 'load_wr2', 'load_mnv', 'factor_mmv', 'load_type'):
        st.session_state.pop(key, None)
    for key in ('motor', 'calc', 'calc_code', 'calc_curve', 'calc_catalog'):
        st.session_state.pop(key, None)
    st.session_state.stage = 0


def set_motor(rec: dict[str, Any], code: str, catalog: str, curve_source: str) -> None:
    st.session_state.motor = {'code':code, 'data':rec, 'curve_source':curve_source,'catalog':catalog}
    st.session_state.hp = float(rec.get('HP') or 0)
    st.session_state.motor_rpm = float(infer_rpm(code) or 0)
    st.session_state.jm = float((rec.get('INERCIA') or 0) / 23.73)
    st.session_state.factor_mmm = float((rec.get('CURVA') or 0) / 100)
    st.session_state.hot = float(rec.get('HOT') or 0)
    st.session_state.cold = float(rec.get('COLD') or 0)
    for k in ('eq_rpm', 'load_wr2', 'load_mnv'):
        st.session_state[k] = 0.0
    st.session_state.factor_mmv = 0.33
    st.session_state.load_type = 'Bomba o ventilador'
    st.session_state.unit_wr2 = 'kgm2'
    st.session_state.unit_mnv = 'kgm'
    st.session_state.stage = 1
    st.session_state.pop('calc',None)


def default_value(key: str, value: Any) -> None:
    if key not in st.session_state:
        st.session_state[key] = value


def main() -> None:
    st.set_page_config(page_title='ABB | Cálculo de tiempo de arranque', page_icon='⚙️', layout='wide', initial_sidebar_state='collapsed')
    theme()
    unavailable_assets = hydrate_assets()
    if unavailable_assets:
        st.warning('No se pudieron descargar automáticamente estos archivos del repositorio original: ' + ', '.join(unavailable_assets) + '. Puedes agregarlos manualmente en la raíz del nuevo repositorio.')
    default_value('stage',0)
    default_value('unit_wr2','kgm2')
    default_value('unit_mnv','kgm')
    default_value('load_type','Bomba o ventilador')
    default_value('factor_mmv',0.33)
    for key in ('eq_rpm','load_wr2','load_mnv','hp','jm','motor_rpm','factor_mmm','hot','cold'):
        default_value(key,0.0)
    st.markdown('<div class="app-banner"><h1>⚙️ Plataforma de Cálculos de Motores</h1><p>Motor Calculations Platform &nbsp; · &nbsp; Cálculo de Tiempo de Arranque / Start-up Time Calculation</p></div>',unsafe_allow_html=True)
    section('🔍 BUSCAR POR CÓDIGO MLFB / SEARCH BY MLFB CODE')
    aa,bb,cc=st.columns([2.2,1.2,1.4],vertical_alignment='bottom')
    with aa:
        code=clean_mlfb(st.text_input('Código MLFB (11 caracteres)', key='mlfb', placeholder='Ej. 1LE...'))
    with bb:
        search=st.button('🔎 BUSCAR MOTOR',use_container_width=True,type='primary')
    with cc:
        st.button('↺ REINICIAR TODO',on_click=clear_form,use_container_width=True)

    section('📁 CONFIGURACIÓN DE ARCHIVO / FILE CONFIGURATION')
    available=[n for n in (DEFAULT_V47,DEFAULT_BASE) if (ROOT/n).is_file()]
    uploaded=None
    file_cols=st.columns([2.0,1.3,1.3])
    with file_cols[0]:
        chosen=st.selectbox('Archivo Excel incorporado',available or ['Sin Excel incorporado'],help='Selecciona el mismo archivo de catálogo utilizado por la aplicación original; puedes sustituirlo por otro')
    with file_cols[1]:
        uploaded=st.file_uploader('Cambiar / Actualizar archivo Excel',type=['xlsx'],key='file_change')
    with file_cols[2]:
        st.markdown('<div class="small-note" style="padding-top:9px">El Excel incorporado se carga automáticamente. Si subes una versión nueva, se utiliza sin modificar los archivos de GitHub.</div>',unsafe_allow_html=True)
    catalog_bytes=uploaded.getvalue() if uploaded is not None else read_local(chosen)
    catalog_name=uploaded.name if uploaded is not None else chosen
    if catalog_bytes is None:
        st.error('No se encontró un catálogo incorporado. Mantén base_motores.xlsx junto a streamlit_app.py o carga el Excel aquí.')
    sheets=[]; df=None
    with st.expander('⚙️ Columnas, hoja y respaldo de curvas  ·  Mostrar / editar',expanded=False):
        cfg1,cfg2=st.columns([2,1])
        with cfg1:
            if catalog_bytes:
                try:
                    sheets=workbook_sheets(catalog_bytes)
                except Exception as exc:
                    st.error(f'Error leyendo las hojas del catálogo: {exc}')
            sheet_name=st.selectbox('Hoja de datos',sheets if sheets else ['Sin hojas'], index=sheets.index('Core Level Data 60 Hz') if 'Core Level Data 60 Hz' in sheets else 0)
        with cfg2:
            headers=st.number_input('Fila de encabezados',min_value=1,max_value=100,value=3,step=1)
        grid=st.columns(6)
        mapping={}
        label={'MLFB':'MLFB','HP':'HP','INERCIA':'Inercia','CURVA':'Curvas','HOT':'Stall HOT','COLD':'Stall COLD'}
        for i,(k,v) in enumerate(COLUMN_DEFAULTS.items()):
            with grid[i]:
                mapping[k]=st.text_input(f'Columna {label[k]}',value=v,key=f'column_{k}').strip().upper()
        st.caption('Por defecto: A (MLFB), C (HP), AS (inercia), BT (curva), AN (HOT), AO (COLD). Puedes editar las seis columnas.')
        st.markdown('**Curvas de respaldo**  ·  Si el Excel nuevo no contiene el factor, consultar base_motores.xlsx por código MLFB')
        uploaded_curves=st.file_uploader('Cambiar archivo de respaldo de curvas',type=['xlsx'],key='backup_file')
        backup_bytes=uploaded_curves.getvalue() if uploaded_curves else read_local(DEFAULT_BASE)
        if backup_bytes:
            try:
                backup_sheets=workbook_sheets(backup_bytes)
            except Exception:
                backup_sheets=[]
        else:
            backup_sheets=[]
        bs1,bs2,bs3,bs4=st.columns(4)
        with bs1:
            backup_sheet=st.selectbox('Hoja base de curvas',backup_sheets or ['Sin respaldo'], index=backup_sheets.index('Core Level Data 60 Hz') if 'Core Level Data 60 Hz' in backup_sheets else 0)
        with bs2:
            backup_header=st.number_input('Fila encabezado respaldo',min_value=1,max_value=100,value=3)
        with bs3:
            backup_code_col=st.text_input('Columna MLFB respaldo',value='A')
        with bs4:
            backup_curve_col=st.text_input('Columna curva respaldo',value='BT')
    # Config widgets are always rendered, even while the expander is collapsed.
    if catalog_bytes and sheets:
        try:
            df=read_sheet(catalog_bytes,sheet_name,int(headers))
        except Exception as exc:
            st.error(f'No se pudo leer la tabla: {exc}')
    if catalog_bytes:
        st.caption(f'📊 Archivo en uso: **{catalog_name}**  ·  Hoja: {sheet_name}  ·  Registros: {len(df) if df is not None else "sin lectura"}')
    if search:
        if len(code)!=11:
            st.error('El código MLFB debe tener exactamente 11 caracteres.')
        elif df is None:
            st.error('Primero carga un catálogo válido.')
        else:
            try:
                record=lookup(df,code,mapping)
                if record is None:
                    st.warning(f'No se encontró el MLFB {code} en {catalog_name}.')
                else:
                    vals={k:parse_num(record.get(k)) for k in ('HP','INERCIA','CURVA','HOT','COLD')}
                    curve_source=catalog_name
                    if vals['CURVA'] is None and backup_bytes and backup_sheets:
                        try:
                            base=read_sheet(backup_bytes,backup_sheet,int(backup_header))
                            curve_rec=lookup(base,code,{'MLFB':backup_code_col,'CURVA':backup_curve_col})
                            if curve_rec:
                                val=parse_num(curve_rec.get('CURVA'))
                                if val is not None:
                                    vals['CURVA']=val
                                    curve_source=(uploaded_curves.name if uploaded_curves else DEFAULT_BASE)
                        except Exception as exc:
                            st.warning(f'No fue posible leer la curva de respaldo: {exc}')
                    set_motor(vals,code,catalog_name,curve_source)
                    st.success(f'✔ Motor encontrado: {code} · Datos cargados · Factor curva: {curve_source}')
            except Exception as exc:
                st.error(f'Error buscando motor: {exc}')
    motor=st.session_state.get('motor',{})
    if motor:
        source=motor.get('curve_source','')
        if st.session_state.get('factor_mmm',0)==0:
            st.warning('Este MLFB no tiene un factor de curva numérico. Puedes capturarlo manualmente en Datos del Motor.')
        st.markdown(f'<div class="step-line step-ok">✔ MLFB <b>{html.escape(motor["code"])}</b> cargado. Archivo: {html.escape(motor["catalog"])} · Curva: {html.escape(source)}</div>',unsafe_allow_html=True)
    else:
        st.info('Empieza buscando un código MLFB. Después se habilitará el ingreso paso a paso, como en el programa de escritorio.')

    section('⚙️ DATOS DEL MOTOR Y DEL EQUIPO ACCIONADO')
    frame_eq,frame_motor=st.columns(2,gap='large')
    stage=int(st.session_state.stage)
    active=bool(motor)
    with frame_eq:
        with st.container(border=True):
            st.markdown('<div class="field-group">📦 Datos del equipo accionado / Driven Equipment Data</div>',unsafe_allow_html=True)
            st.markdown('**Paso 1. Tipo de equipo**')
            load_type=st.radio('¿Es bomba o ventilador?', ['Bomba o ventilador','Otro equipo'],horizontal=True,key='load_type',disabled=not active)
            if load_type=='Bomba o ventilador':
                st.session_state.factor_mmv=0.33
            colfac,facbtn=st.columns([3,1],vertical_alignment='bottom')
            with colfac:
                factor_mmv=st.number_input('Factor MMV',min_value=0.0,format='%.6f',key='factor_mmv',disabled=not active or load_type=='Bomba o ventilador',help='Para bomba o ventilador el factor es 0.33; otros equipos utilizan el factor manual confirmado.')
            with facbtn:
                st.button('✔ Validar',key='next_type',disabled=not active or (load_type!='Bomba o ventilador' and factor_mmv<=0),on_click=jump_step,args=(2,),use_container_width=True)
            st.markdown('**Paso 2. Velocidad del equipo**')
            rpm_col,rpm_btn=st.columns([3,1],vertical_alignment='bottom')
            with rpm_col:
                eq_rpm=st.number_input('RPM (Equipo)',min_value=0.0,step=1.0,format='%.2f',key='eq_rpm',disabled=stage<2)
            with rpm_btn:
                st.button('✔ Siguiente',key='next_rpm',disabled=stage<2 or not positive_state('eq_rpm'),on_click=jump_step,args=(3,),use_container_width=True)
            st.markdown('**Paso 3. Momento de inercia**')
            wi,wunit,wbtn=st.columns([2.8,1.3,1.3],vertical_alignment='bottom')
            with wi:
                wr2=st.number_input('WR²',min_value=0.0,format='%.6f',key='load_wr2',disabled=stage<3)
            with wunit:
                unit_wr2=st.selectbox('Unidad WR²',['kgm2','lbft2'],key='unit_wr2',disabled=stage<3)
            with wbtn:
                st.button('✔ Siguiente',key='next_wr2',disabled=stage<3 or not positive_state('load_wr2'),on_click=jump_step,args=(4,),use_container_width=True)
            st.markdown('**Paso 4. Momento nominal de la carga**')
            mi,munit,mbtn=st.columns([2.8,1.3,1.3],vertical_alignment='bottom')
            with mi:
                mnv_raw=st.number_input('MNV',min_value=0.0,format='%.6f',key='load_mnv',disabled=stage<4)
            with munit:
                unit_mnv=st.selectbox('Unidad MNV',['kgm','lbft'],key='unit_mnv',disabled=stage<4)
            with mbtn:
                st.button('✔ Siguiente',key='next_mnv',disabled=stage<4 or not positive_state('load_mnv'),on_click=jump_step,args=(5,),use_container_width=True)
            st.markdown('**Adaptaciones de la carga**')
            with st.container(border=True):
                if st.session_state.get('calc'):
                    cv=st.session_state.calc.values
                    for left,right,unit in [('MMV','MMV','kgf·m'),('JV','JV','kg·m²')]:
                        a,b=st.columns([1,2]);a.markdown(f'**{left}:**');b.markdown(f'`{display(cv[right],7)} {unit}`')
                else:
                    st.caption('MMV y JV se rellenan automáticamente al calcular, como en la versión Tkinter.')
    with frame_motor:
        with st.container(border=True):
            st.markdown('<div class="field-group">⚙️ Datos del motor / Motor Data</div>',unsafe_allow_html=True)
            st.number_input('HP (Motor)',min_value=0.0,key='hp',format='%.4f',disabled=not active)
            st.number_input('RPM (Motor)',min_value=0.0,key='motor_rpm',format='%.2f',disabled=not active,help='Se precarga la RPM inferida desde el MLFB; es editable para usar RPM nominal de placa.')
            st.number_input('JM · Inercia del rotor (kg·m²)',min_value=0.0,key='jm',format='%.7f',disabled=not active)
            st.number_input('Factor MMM · Curva del motor (fracción)',min_value=0.0,key='factor_mmm',format='%.7f',disabled=not active,help='Valor BT del Excel dividido entre 100, igual que en el original.')
            hot_a,cold_a=st.columns(2)
            with hot_a:
                st.number_input('Stall Time HOT (seg)',min_value=0.0,key='hot',format='%.3f',disabled=not active)
            with cold_a:
                st.number_input('Stall Time COLD (seg)',min_value=0.0,key='cold',format='%.3f',disabled=not active)
            st.markdown('**Valores del motor calculados automáticamente**')
            with st.container(border=True):
                if active and positive_state('motor_rpm') and positive_state('hp'):
                    mnm=float(st.session_state.hp)*726.15/float(st.session_state.motor_rpm)
                    gd2m=float(st.session_state.jm)*4
                    mmm=mnm*float(st.session_state.factor_mmm)
                    for label,value,unit in [('MNM',mnm,'kgf·m'),('GD²M',gd2m,'kg·m²'),('MMM',mmm,'kgf·m')]:
                        a,b=st.columns([1,2]);a.markdown(f'**{label}:**');b.markdown(f'`{display(value,7)} {unit}`')
                else:
                    st.caption('MNM, GD²M y MMM estarán disponibles después de buscar un motor.')

    section('📈 ADAPTACIONES / ADAPTATIONS')
    v=st.session_state.get('calc')
    labels=[('MMV nMot.','MMV_NMOT','kgf·m'),('GD²V','GD2V','kg·m²'),('GD²V nMot.','GD2V_NMOT','kg·m²'),('nGD²','NGD2','kg·m²'),('ta','TA','seg'),('ta CEILING','TA_CEIL','seg')]
    cols=st.columns(6)
    for col,(title,k,u) in zip(cols,labels):
        col.metric(title, f'{display(v.values[k],3)} {u}' if v else '—')
    btn1,btn2,btn3=st.columns([2,1,1],vertical_alignment='bottom')
    with btn1:
        do_calc=st.button('⚡ CALCULAR TIEMPO DE ARRANQUE',type='primary',use_container_width=True,disabled=stage<5 or not active)
    with btn2:
        if st.button('↻ Volver a paso 1',use_container_width=True,disabled=not active):
            st.session_state.stage=1;st.session_state.pop('calc',None);st.rerun()
    with btn3:
        st.caption(f'Paso {min(stage,5)} de 5 · Secuencia guiada')
    if do_calc:
        try:
            raw_wr2=float(st.session_state.load_wr2)
            raw_mnv=float(st.session_state.load_mnv)
            ui=st.session_state.unit_wr2
            ut=st.session_state.unit_mnv
            calc=calculate(hp=float(st.session_state.hp),motor_rpm=float(st.session_state.motor_rpm),
                motor_inertia=float(st.session_state.jm),motor_factor=float(st.session_state.factor_mmm),
                eq_rpm=float(st.session_state.eq_rpm),
                load_inertia=raw_wr2/23.731 if ui=='lbft2' else raw_wr2,
                load_torque=raw_mnv/7.233 if ut=='lbft' else raw_mnv,
                load_factor=0.33 if st.session_state.load_type=='Bomba o ventilador' else float(st.session_state.factor_mmv),
                hot=float(st.session_state.hot) or None,cold=float(st.session_state.cold) or None,
                raw_wr2=raw_wr2,wr2_unit='lb·ft²' if ui=='lbft2' else 'kg·m²',
                raw_mnv=raw_mnv,mnv_unit='lb·ft' if ut=='lbft' else 'kgf·m')
            st.session_state.calc=calc
            st.session_state.calc_code=motor.get('code',code)
            st.session_state.calc_curve=motor.get('curve_source','Manual')
            st.session_state.calc_catalog=catalog_name
            st.rerun()
        except Exception as exc:
            st.session_state.pop('calc',None)
            st.error(f'No se pudo realizar el cálculo: {exc}')
    if v:
        if v.status=='Apto HOT y COLD':st.success(f'✔ {v.status}. {v.status_detail}')
        elif v.status=='No apto':st.error(f'✖ {v.status}. {v.status_detail}')
        else:st.warning(f'{v.status}. {v.status_detail}')
        st.markdown(f'<div class="step-line step-ok">Conversión comprobable: WR² ingresado = <b>{display(v.raw_wr2,6)} {html.escape(v.wr2_unit)}</b> → JV calculado = <b>{display(v.values["JV"],7)} kg·m²</b>. El PDF muestra ambos valores con sus respectivas unidades.</div>',unsafe_allow_html=True)
        op,report=st.columns([2,1])
        with op:
            with st.expander('📋 Operaciones utilizadas / Debug',expanded=False):
                st.dataframe(pd.DataFrame(v.trace,columns=['Variable','Operación','Resultado']),hide_index=True,use_container_width=True)
        with report:
            try:
                data_pdf=build_pdf(st.session_state.calc_code,v,st.session_state.calc_catalog,st.session_state.calc_curve)
                st.download_button('📄 EXPORTAR PDF',data=data_pdf,file_name=f'{st.session_state.calc_code} - {datetime.now():%d-%m-%Y}.pdf',mime='application/pdf',type='primary',use_container_width=True)
            except Exception as exc:
                st.error(f'No se pudo generar el PDF: {exc}')
    section('🔎 HERRAMIENTAS ADICIONALES')
    with st.expander('Comparador de MLFB / MLFB Comparator',expanded=False):
        co1,co2=st.columns(2)
        with co1:m1=st.text_input('MLFB 1',key='compare_1')
        with co2:m2=st.text_input('MLFB 2',key='compare_2')
        if st.button('Comparar MLFB',key='compare_btn'):
            a,b=clean_mlfb(m1),clean_mlfb(m2)
            if not a or not b:st.error('Ambos códigos deben estar llenos.')
            else:
                rows=[]
                for i in range(max(len(a),len(b))):
                    x=a[i] if i<len(a) else '(vacío)';y=b[i] if i<len(b) else '(vacío)'
                    if x!=y:
                        rows.append({'Posición':i+1,'MLFB 1':x,'Significado 1':MLFB_DESCRIPTIONS.get(i+1,{}).get(x,'Sin descripción'),
                            'MLFB 2':y,'Significado 2':MLFB_DESCRIPTIONS.get(i+1,{}).get(y,'Sin descripción')})
                if rows:st.dataframe(pd.DataFrame(rows),hide_index=True,use_container_width=True)
                else:st.success('Los códigos son iguales.')
    st.caption('Cálculos basados en la lógica del formulario Tkinter original. Confirma las unidades físicas y el catálogo correspondiente antes de usar los resultados para selección de motores.')

if __name__ == '__main__':
    main()
