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
    pdf.build(story)
    return out.getvalue()


def theme() -> None:
    st.markdown('''<style>
    .stApp {background: #f6f8fb; color: #1b2838;}
    .block-container {max-width: 1260px; padding-top: 2rem; padding-bottom: 3rem;}
    [data-testid="stMetric"] {background: white; border: 1px solid #e1e7ee; border-radius: 14px; padding: 18px;}
    h1,h2,h3 {letter-spacing: -.035em;}
    h1 {font-size: 2.3rem !important;}
    [data-testid="stSidebar"] {background: #ffffff;}
    .hero {border-radius:18px; padding:25px 29px; background:linear-gradient(110deg,#152337,#263c54); color:#fff; margin-bottom: 24px;}
    .hero h1 {color: #fff !important; margin: 0 0 7px;}
    .hero p {color:#cbd7e4; margin:0;}
    .tag {display:inline-block; color:#fecaca; font-size:12px; font-weight:700; margin-bottom:9px; letter-spacing:.07em;}
    div.stButton > button[kind="primary"] {background:#d92332; border-color:#d92332;}
    </style>''', unsafe_allow_html=True)


def main() -> None:
    st.set_page_config(page_title="Motor Start-up Studio", page_icon="⚙️", layout="wide", initial_sidebar_state="expanded")
    theme()
    st.markdown('<section class="hero"><span class="tag">MOTOR ENGINEERING TOOLS</span><h1>⚙️ Motor Start-up Studio</h1><p>Cálculo de arranque NEMA · Catálogos MLFB actualizables · Reportes técnicos</p></section>', unsafe_allow_html=True)

    with st.sidebar:
        st.header("📁 Catálogo de motores")
        available = [n for n in (DEFAULT_V47, DEFAULT_BASE) if (ROOT / n).is_file()]
        if available:
            default_name = available[0]
            selected = st.selectbox("Archivo base", available, help="V47-R1 tiene prioridad cuando existe en el repositorio.")
            catalog_bytes = read_local(selected)
            catalog_name = selected
        else:
            catalog_bytes, catalog_name = None, "Sin catálogo integrado"
        upload = st.file_uploader("Actualizar catálogo Excel", type=["xlsx"], key="catalog_upload", help="El archivo subido tiene prioridad durante esta sesión. No altera el repositorio.")
        if upload is not None:
            catalog_bytes, catalog_name = upload.getvalue(), upload.name
        if catalog_bytes is None:
            st.warning("Sube el catálogo Excel para habilitar la búsqueda MLFB.")
        else:
            st.caption(f"En uso: {catalog_name}")
        st.markdown("#### Configuración de lectura")
        headers = st.number_input("Fila de encabezados", min_value=1, max_value=30, value=3, step=1)
        sheet_name = None
        data = None
        if catalog_bytes:
            try:
                sheets = workbook_sheets(catalog_bytes)
                initial = sheets.index("Core Level Data 60 Hz") if "Core Level Data 60 Hz" in sheets else 0
                sheet_name = st.selectbox("Hoja Excel", sheets, index=initial)
                data = read_sheet(catalog_bytes, sheet_name, int(headers))
                st.caption(f"{len(data):,} filas · {data.shape[1]} columnas")
            except Exception as exc:
                st.error(f"No se pudo abrir este Excel: {exc}")
        with st.expander("Columnas (letras Excel)", expanded=False):
            mapping = {k: st.text_input(k, value=v, max_chars=3, key=f"map_{k}").strip().upper() for k, v in COLUMN_DEFAULTS.items()}
        st.markdown("#### Factores de curvas")
        st.caption("Si el catálogo actual no tiene curva, se busca en base_motores.xlsx por MLFB. Nunca se generan promedios ficticios.")
        backup_upload = st.file_uploader("Base de curvas adicional (opcional)", type=["xlsx"], key="curves_upload")
        backup_bytes = backup_upload.getvalue() if backup_upload else read_local(DEFAULT_BASE)
        backup_sheet = "Core Level Data 60 Hz"
        backup_map = dict(COLUMN_DEFAULTS)
        with st.expander("Configuración de la base de curvas"):
            if backup_bytes:
                try:
                    bsheets = workbook_sheets(backup_bytes)
                    backup_sheet = st.selectbox("Hoja de curvas", bsheets, index=bsheets.index(backup_sheet) if backup_sheet in bsheets else 0)
                except Exception as exc:
                    st.warning(f"No se pudo leer la base de respaldo: {exc}")
            backup_header = st.number_input("Fila de encabezados de curvas", min_value=1, max_value=30, value=3)
            backup_map["MLFB"] = st.text_input("Columna MLFB de respaldo", value="A").strip().upper()
            backup_map["CURVA"] = st.text_input("Columna factor de curva de respaldo", value="BT").strip().upper()

    tab_calc, tab_mlfb, tab_guide = st.tabs(["🧮 Cálculo de arranque", "🔎 Comparador MLFB", "📘 Guía y fuentes"])
    with tab_mlfb:
        st.subheader("Comparación de códigos MLFB")
        c1, c2 = st.columns(2)
        with c1: m1 = st.text_input("MLFB 1", key="compare1")
        with c2: m2 = st.text_input("MLFB 2", key="compare2")
        if m1.strip() and m2.strip():
            a, b = clean_mlfb(m1), clean_mlfb(m2)
            changes = []
            for i in range(max(len(a), len(b))):
                aa = a[i] if i < len(a) else "∅"
                bb = b[i] if i < len(b) else "∅"
                if aa != bb:
                    changes.append({"Posición": i + 1, "MLFB 1": aa, "Descripción 1": MLFB_DESCRIPTIONS.get(i + 1, {}).get(aa, "Sin descripción"),
                                    "MLFB 2": bb, "Descripción 2": MLFB_DESCRIPTIONS.get(i + 1, {}).get(bb, "Sin descripción")})
            if changes: st.dataframe(pd.DataFrame(changes), hide_index=True, use_container_width=True)
            else: st.success("Ambos códigos son iguales.")
        st.caption("Decodificador posicional heredado del comparador original. Las descripciones dependen de la familia de producto y deben contrastarse con el catálogo.")
    with tab_guide:
        st.subheader("Cómo actualizar el Excel")
        st.markdown("""1. En la barra izquierda, **Actualizar catálogo Excel** permite subir una versión nueva (incluida V47-R1). Se utiliza de inmediato durante la sesión actual.
2. Ajusta la **hoja**, **fila de encabezados** y las **letras de columna** si la estructura del libro cambió.
3. Si la nueva versión no trae factores de curva, sube **base_motores.xlsx** como respaldo y configura sus columnas. Solo se reutilizan curvas con coincidencia exacta de MLFB.
4. Para que el catálogo sea permanente para todos los usuarios, sube el `.xlsx` al repositorio de GitHub. El aplicativo detecta automáticamente V47-R1 cuando está junto a `streamlit_app.py`.

**Importante:** la carga temporal desde Streamlit no escribe cambios en GitHub. La aplicación nunca usa valores de curva inventados, y cuando falta el factor exige capturarlo manualmente.

**Convenciones:** se reproduce la fórmula original de arranque y el redondeo intermedio original. La conversión de entrada es `lb·ft² / 23.731 = kg·m²` y `lb·ft / 7.233 = kgf·m` según el programa de escritorio. Las etiquetas del PDF reflejan el valor original y el normalizado por separado. Verifica la convención física exacta `WR²/GD²` con tu documentación técnica.""")
    with tab_calc:
        st.subheader("01 · Selección del motor")
        col1, col2 = st.columns([2, 1])
        with col1:
            code = clean_mlfb(st.text_input("Código MLFB", placeholder="Introduce los 11 caracteres del motor"))
        with col2:
            search = st.button("Buscar motor en catálogo", type="primary", use_container_width=True)
        if search:
            if len(code) != 11:
                st.error("El MLFB debe tener 11 caracteres alfanuméricos.")
            elif data is None:
                st.error("No hay un catálogo Excel válido cargado.")
            else:
                try:
                    rec = lookup(data, code, mapping)
                    if rec is None:
                        st.error(f"No se encontró {code} en {catalog_name}.")
                    else:
                        vals = {k: parse_num(rec.get(k)) for k in ("HP", "INERCIA", "CURVA", "HOT", "COLD")}
                        source = "Catálogo activo"
                        if vals["CURVA"] is None and backup_bytes:
                            try:
                                back = read_sheet(backup_bytes, backup_sheet, int(backup_header))
                                match = lookup(back, code, backup_map)
                                if match is not None and parse_num(match.get("CURVA")) is not None:
                                    vals["CURVA"] = parse_num(match["CURVA"])
                                    source = f"Respaldo: {backup_upload.name if backup_upload else DEFAULT_BASE}"
                            except Exception as exc:
                                st.warning(f"No fue posible consultar las curvas de respaldo: {exc}")
                        st.session_state["motor"] = {"code": code, "data": vals, "curve_source": source, "catalog": catalog_name}
                        st.session_state.pop("calc", None)
                        st.success(f"Motor {code} localizado. Datos y curva: {source}.")
                except Exception as exc:
                    st.error(f"No se pudo buscar el motor: {exc}")
        motor = st.session_state.get("motor", {})
        mv = motor.get("data", {})
        if motor:
            if motor["code"] != code and code:
                st.info(f"Datos cargados para {motor['code']}. Pulsa Buscar para sustituirlos.")
            if mv.get("CURVA") is None:
                st.warning("El catálogo no tiene un factor de curva numérico para este MLFB y no se encontró en el respaldo. Captúralo manualmente en el formulario.")
            st.caption(f"Fuente: {motor.get('catalog')} · Curva: {motor.get('curve_source')}")
        st.subheader("02 · Datos de operación")
        with st.form("motor_form"):
            left, right = st.columns(2, gap="large")
            with left:
                st.markdown("#### Motor seleccionado")
                hp = st.number_input("Potencia del motor (HP)", min_value=0.0, value=float(mv.get("HP") or 0.0), step=1.0, format="%.4f")
                rpm_default = infer_rpm(motor.get("code", "")) or 0.0
                motor_rpm = st.number_input("Velocidad del motor (rpm)", min_value=0.0, value=float(rpm_default), step=1.0, format="%.2f", help="La velocidad sugerida desde el MLFB es sincrónica de referencia, no necesariamente la velocidad nominal de placa.")
                inertia_lbft2 = mv.get("INERCIA")
                inertia_default = float(inertia_lbft2 / 23.73) if inertia_lbft2 is not None else 0.0
                motor_inertia = st.number_input("Inercia del rotor JM (kg·m²)", min_value=0.0, value=inertia_default, format="%.7f", help="Conversión del valor AS del catálogo: / 23.73, igual que la app original.")
                curve_default = (float(mv["CURVA"]) / 100.0) if mv.get("CURVA") is not None else 0.0
                motor_factor = st.number_input("Factor de torque medio del motor (fracción)", min_value=0.0, value=curve_default, format="%.6f", help="Valor Excel BT / 100. Si la curva faltaba, introduce aquí el factor confirmado.")
                hot = st.number_input("Stall Time HOT (segundos; 0 = no disponible)", min_value=0.0, value=float(mv.get("HOT") or 0), format="%.3f")
                cold = st.number_input("Stall Time COLD (segundos; 0 = no disponible)", min_value=0.0, value=float(mv.get("COLD") or 0), format="%.3f")
            with right:
                st.markdown("#### Equipo accionado")
                eq_rpm = st.number_input("RPM del equipo accionado", min_value=0.0, value=0.0, step=1.0)
                u_wr2 = st.selectbox("Unidad de inercia introducida", ["lb·ft²", "kg·m²"], index=0)
                raw_wr2 = st.number_input(f"Inercia del equipo ({u_wr2})", min_value=0.0, value=0.0, format="%.6f")
                u_mnv = st.selectbox("Unidad de torque introducido", ["kgf·m", "lb·ft"], index=0)
                raw_mnv = st.number_input(f"Torque nominal de la carga ({u_mnv})", min_value=0.0, value=0.0, format="%.6f")
                is_fan = st.checkbox("Es bomba o ventilador (factor MMV = 0.33)", value=True)
                manual_factor = st.number_input("Factor MMV (fracción)", min_value=0.0, value=0.33, format="%.6f", disabled=is_fan)
                st.caption("La interfaz conserva tus valores originales. El cálculo opera con unidades normalizadas.")
            submitted = st.form_submit_button("⚡ Calcular tiempo de arranque", type="primary", use_container_width=True)
        if submitted:
            load_inertia = raw_wr2 / 23.731 if u_wr2 == "lb·ft²" else raw_wr2
            load_torque = raw_mnv / 7.233 if u_mnv == "lb·ft" else raw_mnv
            try:
                calc = calculate(hp=hp, motor_rpm=motor_rpm, motor_inertia=motor_inertia,
                                 motor_factor=motor_factor, eq_rpm=eq_rpm, load_inertia=load_inertia,
                                 load_torque=load_torque, load_factor=0.33 if is_fan else manual_factor,
                                 hot=hot or None, cold=cold or None,
                                 raw_wr2=raw_wr2, wr2_unit=u_wr2, raw_mnv=raw_mnv, mnv_unit=u_mnv)
                st.session_state["calc"] = calc
                st.session_state["calc_code"] = motor.get("code", code)
                st.session_state["calc_catalog"] = motor.get("catalog", catalog_name)
                st.session_state["calc_curve"] = motor.get("curve_source", "Introducido manualmente")
            except ValueError as exc:
                st.session_state.pop("calc", None)
                st.error(str(exc))
        result = st.session_state.get("calc")
        if result:
            v = result.values
            st.subheader("03 · Resultado de ingeniería")
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Tiempo de arranque", f"{v['TA']:.2f} s")
            c2.metric("Tiempo CEILING", f"{v['TA_CEIL']:.0f} s")
            c3.metric("Torque medio motor", f"{display(v['MMM'], 2)} kgf·m")
            c4.metric("Inercia equivalente", f"{display(v['NGD2'], 2)} kg·m²")
            if result.status == "Apto HOT y COLD": st.success(f"✅ {result.status}. {result.status_detail}")
            elif result.status in ("No apto",): st.error(f"⛔ {result.status}. {result.status_detail}")
            else: st.warning(f"ℹ️ {result.status}. {result.status_detail}")
            with st.expander("Ver conversiones, variables y operaciones", expanded=False):
                st.info(f"Inercia: {display(result.raw_wr2, 6)} {result.wr2_unit} → {display(v['JV'], 7)} kg·m².  Torque: {display(result.raw_mnv, 6)} {result.mnv_unit} → {display(v['MNV'], 7)} kgf·m.")
                st.dataframe(pd.DataFrame(result.trace, columns=["Variable", "Operación", "Valor"]), use_container_width=True, hide_index=True)
            try:
                pdf_bytes = build_pdf(st.session_state.get("calc_code", ""), result, st.session_state.get("calc_catalog", "Manual"), st.session_state.get("calc_curve", "Manual"))
                st.download_button("📄 Descargar reporte PDF", data=pdf_bytes,
                                   file_name=f"arranque_{st.session_state.get('calc_code','manual')}_{datetime.now():%Y%m%d}.pdf",
                                   mime="application/pdf", type="primary")
            except Exception as exc:
                st.error(f"No se pudo construir el reporte PDF: {exc}")
            if st.button("Limpiar resultados"):
                st.session_state.pop("calc", None)
                st.rerun()


if __name__ == "__main__":
    main()
