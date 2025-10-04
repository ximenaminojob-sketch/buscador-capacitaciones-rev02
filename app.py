import pandas as pd
import streamlit as st
from datetime import datetime, date
from pathlib import Path
import unicodedata

# ================== CONFIG ==================
st.set_page_config(page_title="Reporte de Formación por Persona", page_icon="✅", layout="wide")

BASE_DIR   = Path(__file__).resolve().parent
XLSX_PATH  = BASE_DIR / "Nómina de Capacitación - QR - REV. 02.xlsx"   # <-- tu archivo real
SHEET_NAME = "TECHINT"

# Estructura base (0-based). Si cambia, ajustá acá:
ROW_HEADER = 5   # encabezados en la fila 6
ROW_START  = 6   # datos desde la fila 7
COL_DNI    = 2   # columna C (DNI)
COL_START  = 6   # columna G (primer tema)

# ================== HELPERS ==================
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

@st.cache_data
def load_data(xlsx_path: Path, sheet_name: str):
    df = pd.read_excel(xlsx_path, sheet_name=sheet_name, header=None, engine="openpyxl")
    return df

def excel_serial_to_datetime(val):
    try:
        return pd.to_datetime(val, unit="d", origin="1899-12-30")
    except Exception:
        return None

def parse_fecha(v):
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

# ================== LOAD ==================
if not XLSX_PATH.exists():
    st.error(f"No se encontró el Excel: {XLSX_PATH.name}")
    st.stop()

df = load_data(XLSX_PATH, SHEET_NAME)

# Detecto última columna de temas (COL_END)
headers_row = df.iloc[ROW_HEADER, :].tolist()
last_col = COL_START
for c in range(COL_START, df.shape[1]):
    val = df.iat[ROW_HEADER, c]
    if (isinstance(val, str) and val.strip() != "") or (pd.notna(val) and str(val).strip() != ""):
        last_col = c
COL_END = last_col

# Detecto columnas de Nombre, Puesto, Especialidad por encabezado
COL_NOMBRE       = find_col_by_keywords(headers_row, ["nombre y apellido", "nombre", "apellido", "apellidos"])
COL_PUESTO       = find_col_by_keywords(headers_row, ["puesto"])
COL_ESPECIALIDAD = find_col_by_keywords(headers_row, ["especialidad"])

# Series base
dni_series = df.iloc[ROW_START:, COL_DNI].astype(str).str.strip()

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
from pathlib import Path

BASE_DIR  = Path(__file__).resolve().parent
logo_path = BASE_DIR / "logo_techint.png"  # agregalo al repo con este nombre

col_logo, col_title = st.columns([1, 6])
with col_logo:
    if logo_path.exists():
        st.image(str(logo_path), width=120)
with col_title:
    st.title("Buscador de Capacitaciones (solo realizadas)")

st.caption(" Elegí una persona por **DNI** o **Nombre y Apellido**. Se listan solo los temas con **fecha de realización**.")

# ================== CONTROLES DE BÚSQUEDA ==================
modo = st.radio("Buscar por", ["DNI", "Nombre y Apellido"], horizontal=True)

row_idx = None
dni_sel = None
nombre_sel = None

if modo == "DNI":
    dni_unicos = sorted(set([d for d in dni_series.tolist() if d and d.lower() != "nan"]))
    dni_sel = st.selectbox("DNI", options=["— Seleccioná —"] + dni_unicos, index=0)
    if dni_sel and dni_sel != "— Seleccioná —":
        mask = (dni_series == str(dni_sel).strip())
        if mask.any():
            row_idx = mask[mask].index[0]
            if nombres_series is not None:
                nombre_sel = str(df.iat[row_idx, COL_NOMBRE]) if COL_NOMBRE is not None else None

else:  # Buscar por nombre
    if nombres_series is None:
        st.warning("No se encontró la columna de 'Nombre' en la fila de encabezados (fila 6).")
    else:
        # Normalizo para buscar ignorando acentos y may/min
        nombres_norm = nombres_series.map(normalize_text)
        opciones = sorted(set([n for n in nombres_series.tolist() if n and n.lower() != "nan"]))
        nombre_sel = st.selectbox("Nombre y apellido", options=["— Seleccioná —"] + opciones, index=0)
        if nombre_sel and nombre_sel != "— Seleccioná —":
            mask = (nombres_series == nombre_sel)
            if mask.sum() > 1:
                # Si hay duplicados, pedimos el DNI para desambiguar
                dnis_coinc = dni_series[mask].tolist()
                dni_sel = st.selectbox("Coinciden varios, elegí el DNI", options=dnis_coinc)
                mask = mask & (dni_series == dni_sel)
            if mask.any():
                row_idx = mask[mask].index[0]
                dni_sel = str(df.iat[row_idx, COL_DNI])

# ================== DATOS DE LA PERSONA + LISTA REALIZADAS ==================
# ================== DATOS DE LA PERSONA + LISTA REALIZADAS ==================
if row_idx is not None:
    # --- Datos de cabecera ---
    nombre = str(df.iat[row_idx, COL_NOMBRE]) if COL_NOMBRE is not None else "-"
    puesto = str(df.iat[row_idx, COL_PUESTO]) if COL_PUESTO is not None else "-"
    espec  = str(df.iat[row_idx, COL_ESPECIALIDAD]) if COL_ESPECIALIDAD is not None else "-"

    cA, cB, cC, cD = st.columns([2, 3, 2, 2])
    with cA: st.write("**DNI**");               st.write(dni_sel or "-")
    with cB: st.write("**Nombre y Apellido**"); st.write(nombre)
    with cC: st.write("**Puesto**");            st.write(puesto)
    with cD: st.write("**Especialidad**");      st.write(espec)

    st.divider()

    # --- Construyo 'registros' (solo si hay FECHA) ---
    valores = df.iloc[row_idx, COL_START:COL_END+1].tolist()
    registros = []
    for h, v in zip(temas, valores):
        if not h:
            continue
        f = parse_fecha(v)            # usa tu helper; devuelve date o None
        if f is not None:
            registros.append({"Tema": h, "Fecha": f.strftime("%d/%m/%Y")})

    # --- Métricas ---
    total_realizadas = len(registros)
    total_temarios   = len(temas)
    c1, c2, c3 = st.columns(3)
    with c1: st.metric("Capacitaciones realizadas", total_realizadas)
    with c2: st.metric("Total de temas", total_temarios)
    with c3:
        pct = 0 if total_temarios == 0 else round(100*total_realizadas/total_temarios, 1)
        st.metric("% de avance", f"{pct}%")

    st.subheader("✅ Capacitaciones realizadas")

    if total_realizadas == 0:
        st.info("No hay capacitaciones realizadas registradas para esta persona.")
    else:
        # Ordeno por fecha DESC y muestro
        import pandas as pd
        df_out = pd.DataFrame(registros)
        df_out["__orden"] = pd.to_datetime(df_out["Fecha"], dayfirst=True, errors="coerce")
        df_out = df_out.sort_values("__orden", ascending=False).drop(columns="__orden")

        # Chips de temas (lista visual)
        st.markdown("""
        <style>
        .tag {display:inline-block; padding:6px 10px; border-radius:14px;
              background:#eef2ff; margin:4px 6px 8px 0; border:1px solid #c7d2fe; font-size:14px}
        </style>
        """, unsafe_allow_html=True)
        st.markdown("**Temas realizados:**", unsafe_allow_html=True)
        st.markdown("".join([f"<span class='tag'>{t}</span>" for t in df_out["Tema"].tolist()]),
                    unsafe_allow_html=True)

        # Tabla Tema–Fecha
        st.dataframe(df_out, use_container_width=True)

        # Descargar
        csv = df_out.to_csv(index=False).encode("utf-8-sig")
        st.download_button("⬇️ Descargar CSV", data=csv,
                           file_name=f"capacitaciones_realizadas_{(dni_sel or 'persona')}.csv",
                           mime="text/csv")
else:
    st.info("Elegí un DNI o un Nombre para comenzar.")
