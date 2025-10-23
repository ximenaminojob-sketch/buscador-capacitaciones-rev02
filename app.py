import os
import hashlib
from pathlib import Path
from datetime import datetime, date
import unicodedata

import pandas as pd
import streamlit as st

# ================== CONFIG BÁSICA ==================
st.set_page_config(page_title="Seguimiento de Competencias", page_icon="✅", layout="wide")

BASE_DIR   = Path(__file__).resolve().parent
XLSX_PATH  = BASE_DIR / "Nómina de Capacitación - QR - REV. 02.xlsx"
SHEET_NAME = "TECHINT"

# Estructura (0-based): ajustá si cambia tu archivo
ROW_HEADER = 5   # fila 6 en Excel (encabezados de temas)
ROW_START  = 6   # fila 7 en Excel (comienzo de datos)
COL_DNI    = 2   # columna C (DNI)
COL_START  = 6   # columna G (primer tema)

# ================== THEME (toggle claro/oscuro con paleta Techint) ==================
def apply_theme(mode="light"):
    # Paleta Techint
    GREEN = "#009900"   # Pantone 355
    NAVY  = "#002B5C"   # Pantone 289
    GRAY  = "#6E6E6E"

    # ----- Paletas completamente invertidas -----
    if mode == "dark":
        BG      = NAVY          # fondo principal
        BG2     = "#0F2F55"     # paneles/cards/inputs
        TEXT    = "#FFFFFF"      # texto
        MUTED   = "rgba(255,255,255,.65)"
        BORDER  = "#1f3a64"
        HOVER   = "rgba(255,255,255,.12)"
        TABLEZ  = "rgba(255,255,255,.03)"   # zebra
        LINK    = "#7ec8ff"
        CHIP_BG = "#0F3F1F"
        CHIP_BD = GREEN
    else:
        BG      = "#FFFFFF"
        BG2     = "#F2F4F8"
        TEXT    = NAVY
        MUTED   = "#5b6b7c"
        BORDER  = "#d5dde7"
        HOVER   = "rgba(0,0,0,.06)"
        TABLEZ  = "#fafcff"
        LINK    = "#0b6bbf"
        CHIP_BG = "#E9F7EA"
        CHIP_BD = GREEN

    st.markdown(f"""
    <style>
      /* Base */
      .stApp, html, body, [data-testid="stAppViewContainer"] {{
        background:{BG} !important; color:{TEXT} !important;
      }}
      p,span,div,li,label,small,code,pre,
      h1,h2,h3,h4,h5,h6,
      .stMarkdown, .stText, .stTooltip, [data-testid="stMetric"] * {{
        color:{TEXT} !important;
      }}
      a, a:visited {{ color:{LINK} !important; }}

      /* Contenedores secundarios / cards */
      [data-testid="stSidebar"], .stTabs, .stAlert, .css-12w0qpk, .css-uvzsq5 {{
        background:{BG2} !important; color:{TEXT} !important;
      }}

      /* Inputs, textareas, sliders */
      input, textarea {{ background:{BG2} !important; color:{TEXT} !important; border-color:{BORDER} !important; }}
      .stSlider > div > div > div > div {{ background:{GREEN} !important; }}

      /* Botones */
      .stButton > button {{
        background:{GREEN} !important; color:#fff !important;
        border:0 !important; border-radius:10px !important; padding:.5rem 1rem !important;
        box-shadow:none !important;
      }}
      .stDownloadButton > button {{ background:{GREEN} !important; color:#fff !important; }}

      /* Radio/checkbox */
      div[role="radiogroup"] label span, label[for*="checkbox"] span {{ color:{TEXT} !important; }}

      /* Selectbox/Multiselect (BaseWeb) */
      div[data-baseweb="select"] > div {{
        background:{BG2} !important; color:{TEXT} !important; border-color:{BORDER} !important;
      }}
      div[data-baseweb="select"] * {{ color:{TEXT} !important; fill:{TEXT} !important; }}
      div[data-baseweb="popover"] {{ background:{BG2} !important; color:{TEXT} !important; border:1px solid {BORDER} !important; }}
      div[data-baseweb="popover"] [role="listbox"] {{ background:{BG2} !important; }}
      div[data-baseweb="popover"] [role="option"] {{ color:{TEXT} !important; }}
      div[data-baseweb="popover"] [role="option"]:hover {{ background:{HOVER} !important; }}

      /* Tablas */
      table {{ border-color:{BORDER} !important; }}
      thead tr th {{ background:{BG2} !important; color:{TEXT} !important; border-bottom:1px solid {BORDER} !important; }}
      tbody tr:nth-child(odd) td {{ background:{TABLEZ} !important; }}
      tbody tr:hover td {{ background:{HOVER} !important; }}

      /* Separadores / líneas */
      hr, [role="separator"] {{ border-color:{BORDER} !important; }}

      /* Chips (temas realizados) */
      .tag {{
        display:inline-block; padding:6px 10px; border-radius:14px;
        background:{CHIP_BG}; border:1px solid {CHIP_BD}; color:{TEXT};
        margin:4px 6px 8px 0; font-size:14px
      }}

      /* Scrollbar (sutil) */
      ::-webkit-scrollbar {{ height:12px; width:12px; }}
      ::-webkit-scrollbar-thumb {{ background:{BORDER}; border-radius:10px; }}
      ::-webkit-scrollbar-track {{ background:transparent; }}
    </style>
    """, unsafe_allow_html=True)
    
col_modo, _ = st.columns([1, 5])
with col_modo:
    oscuro = st.toggle("Modo oscuro", value=False)
apply_theme("dark" if oscuro else "light")

# ================== HELPERS ==================
def _file_version(path: Path) -> str:
    """Hash corto del archivo para invalidar caché cuando cambia."""
    h = hashlib.md5()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(8192), b""):
            h.update(chunk)
    return h.hexdigest()[:12]

@st.cache_data  # podés usar ttl=300 si querés auto-refresh cada 5 min
def load_data(xlsx_path: Path, sheet_name: str, _ver: str):
    # _ver no se usa: solo fuerza a recachear si cambia el archivo
    return pd.read_excel(xlsx_path, sheet_name=sheet_name, header=None, engine="openpyxl")

def excel_serial_to_datetime(val):
    try:
        return pd.to_datetime(val, unit="d", origin="1899-12-30")
    except Exception:
        return None

def parse_fecha(v):
    """Devuelve datetime.date si la celda representa una fecha válida; si no, None."""
    if pd.isna(v):
        return None
    if isinstance(v, pd.Timestamp):
        return v.date()
    if isinstance(v, (datetime, date)):
        return v if isinstance(v, date) and not isinstance(v, datetime) else v.date()
    if isinstance(v, (int, float)) and v > 0:
        dt = excel_serial_to_datetime(v)
        return dt.date() if dt is not None else None
    if isinstance(v, str):
        dt = pd.to_datetime(v, dayfirst=True, errors="coerce")
        return dt.date() if pd.notna(dt) else None
    return None

def normalize_text(s: str) -> str:
    if s is None:
        return ""
    s = str(s)
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
    return s.lower().strip()

def find_col_by_keywords(headers_row, keywords):
    for idx, val in enumerate(headers_row):
        if isinstance(val, str):
            t = normalize_text(val)
            for kw in keywords:
                if kw in t:
                    return idx
    return None

# ================== CARGA DE DATOS ==================
if not XLSX_PATH.exists():
    st.error(f"No se encontró el Excel: {XLSX_PATH.name}")
    st.stop()

df = load_data(XLSX_PATH, SHEET_NAME, _file_version(XLSX_PATH))
if df is None or not hasattr(df, "iloc"):
    st.error("No se pudo cargar el Excel. Verificá el nombre de la hoja y la función load_data().")
    st.stop()

# Botón para refrescar manualmente
col_refresh, _ = st.columns([1, 5])
with col_refresh:
    if st.button("🔁 Actualizar datos"):
        load_data.clear()
        st.rerun()
# Detecto última columna de temas (desde G hacia la derecha)
headers_row = df.iloc[ROW_HEADER, :].tolist()
last_col = COL_START
for c in range(COL_START, df.shape[1]):
    val = df.iat[ROW_HEADER, c]
    if (isinstance(val, str) and val.strip() != "") or (pd.notna(val) and str(val).strip() != ""):
        last_col = c
COL_END = last_col

# Detecto columnas Nombre, Puesto, Especialidad por encabezado
COL_NOMBRE       = find_col_by_keywords(headers_row, ["nombre y apellido", "nombre", "apellido", "apellidos"])
COL_PUESTO       = find_col_by_keywords(headers_row, ["puesto"])
COL_ESPECIALIDAD = find_col_by_keywords(headers_row, ["especialidad"])

# Series base (DNI normalizado para números/espacios)
dni_col_raw = df.iloc[ROW_START:, COL_DNI]
def _norm_dni(v):
    if pd.isna(v): return ""
    if isinstance(v, (int,)) or (isinstance(v, float) and v.is_integer()):
        return str(int(v))
    return str(v).strip()
dni_series = dni_col_raw.map(_norm_dni)
dni_unicos = sorted(set([d for d in dni_series.tolist() if d and d.lower() != "nan"]))

nombres_series = None
if COL_NOMBRE is not None:
    nombres_series = df.iloc[ROW_START:, COL_NOMBRE].astype(str).str.strip()

temas = (
    df.iloc[ROW_HEADER, COL_START:COL_END+1]
      .fillna("")
      .astype(str).str.strip()
      .tolist()
)

# ================== HEADER CON LOGO ==================
logo_path = BASE_DIR / "logo_techint.png"
col_logo, col_title = st.columns([2, 6])
with col_logo:
    if logo_path.exists():
        st.image(str(logo_path), width=250)
with col_title:
    st.title("Seguimiento de Competencias")
st.caption("Elegí una persona por **DNI** o **Nombre y Apellido**. Se listan solo los temas con **fecha de realización**.")

# ================== CONTROLES DE BÚSQUEDA ==================
modo = st.radio("Buscar por", ["DNI", "Nombre y Apellido"], horizontal=True)

row_idx = None
dni_sel = None
nombre_sel = None

if modo == "DNI":
    dni_sel = st.selectbox("DNI", options=["— Seleccioná —"] + dni_unicos, index=0)
    if dni_sel and dni_sel != "— Seleccioná —":
        mask = (dni_series == str(dni_sel).strip())
        if mask.any():
            row_idx = mask[mask].index[0]
            if nombres_series is not None:
                nombre_sel = str(df.iat[row_idx, COL_NOMBRE]) if COL_NOMBRE is not None else None
else:
    if nombres_series is None:
        st.warning("No se encontró la columna de 'Nombre' en la fila de encabezados (fila 6).")
    else:
        opciones = sorted(set([n for n in nombres_series.tolist() if n and n.lower() != "nan"]))
        nombre_sel = st.selectbox("Nombre y apellido", options=["— Seleccioná —"] + opciones, index=0)
        if nombre_sel and nombre_sel != "— Seleccioná —":
            mask = (nombres_series == nombre_sel)
            if mask.sum() > 1:
                dnis_coinc = dni_series[mask].tolist()
                dni_sel = st.selectbox("Coinciden varios, elegí el DNI", options=dnis_coinc)
                mask = mask & (dni_series == dni_sel)
            if mask.any():
                row_idx = mask[mask].index[0]
                dni_sel = str(df.iat[row_idx, COL_DNI])

# ================== DATOS Y CAPACITACIONES REALIZADAS ==================
if row_idx is not None:
    # Cabecera
    nombre = str(df.iat[row_idx, COL_NOMBRE]) if COL_NOMBRE is not None else "-"
    puesto = str(df.iat[row_idx, COL_PUESTO]) if COL_PUESTO is not None else "-"
    espec  = str(df.iat[row_idx, COL_ESPECIALIDAD]) if COL_ESPECIALIDAD is not None else "-"

    cA, cB, cC, cD = st.columns([2, 3, 2, 2])
    with cA: st.write("**DNI**");               st.write(dni_sel or "-")
    with cB: st.write("**Nombre y Apellido**"); st.write(nombre)
    with cC: st.write("**Puesto**");            st.write(puesto)
    with cD: st.write("**Especialidad**");      st.write(espec)

    st.divider()

    # Registros (solo con fecha válida)
    valores = df.iloc[row_idx, COL_START:COL_END+1].tolist()
    registros = []
    for h, v in zip(temas, valores):
        if not h:
            continue
        f = parse_fecha(v)
        if f is not None:
            registros.append({"Tema": h, "Fecha": f.strftime("%d/%m/%Y")})

    total_realizadas = len(registros)
    st.subheader(f"✅ Capacitaciones realizadas ({total_realizadas})")

    if total_realizadas == 0:
        st.info("No hay capacitaciones realizadas registradas para esta persona.")
    else:
        df_out = pd.DataFrame(registros)
        df_out["__orden"] = pd.to_datetime(df_out["Fecha"], dayfirst=True, errors="coerce")
        df_out = df_out.sort_values("__orden", ascending=False).drop(columns="__orden")

        st.markdown("**Temas realizados:**")
        st.markdown("".join([f"<span class='tag'>{t}</span>" for t in df_out["Tema"].tolist()]),
                    unsafe_allow_html=True)

        st.dataframe(df_out, use_container_width=True)

        csv = df_out.to_csv(index=False).encode("utf-8-sig")
        st.download_button("⬇️ Descargar CSV",
                           data=csv,
                           file_name=f"capacitaciones_realizadas_{dni_sel or 'persona'}.csv",
                           mime="text/csv")
else:
    st.info("Elegí un DNI o un Nombre para comenzar.")
