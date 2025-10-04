import pandas as pd
import streamlit as st
from datetime import datetime, date
from pathlib import Path

st.set_page_config(page_title="Buscador de Capacitaciones", page_icon="✅", layout="wide")

BASE_DIR   = Path(__file__).resolve().parent
XLSX_PATH  = BASE_DIR / "Nómina de Capacitación - QR - REV. 02.xlsx"
SHEET_NAME = "TECHINT"

ROW_HEADER = 5
ROW_START  = 6
COL_DNI    = 2
COL_START  = 6

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
    if pd.isna(v): return None
    if isinstance(v, pd.Timestamp): return v.date()
    if isinstance(v, (datetime, date)): return v if isinstance(v, date) and not isinstance(v, datetime) else v.date()
    if isinstance(v, (int, float)) and v > 0:
        dt = excel_serial_to_datetime(v); return dt.date() if dt is not None else None
    if isinstance(v, str):
        dt = pd.to_datetime(v, dayfirst=True, errors="coerce"); return dt.date() if pd.notna(dt) else None
    return None

if not XLSX_PATH.exists():
    st.error(f"No se encontró el archivo Excel esperado: {XLSX_PATH.name}")
    st.stop()

df = load_data(XLSX_PATH, SHEET_NAME)

headers_row = df.iloc[ROW_HEADER, :]
last_col = COL_START
for c in range(COL_START, df.shape[1]):
    val = headers_row.iat[c]
    if isinstance(val, str) and val.strip() != "":
        last_col = c
    elif not pd.isna(val):
        last_col = c
COL_END = last_col

dni_series = df.iloc[ROW_START:, COL_DNI].astype(str).str.strip()
dni_unicos = sorted(set([d for d in dni_series.tolist() if d and d.lower() != "nan"]))

temas = df.iloc[ROW_HEADER, COL_START:COL_END+1].fillna("").astype(str).str.strip().tolist()

st.title("🔎 Buscador de Capacitaciones (solo realizadas)")
st.caption("Elegí un DNI para ver **solo** los temas con **fecha de realización**.")

c1, c2 = st.columns([2,1])
with c1:
    dni_sel = st.selectbox("DNI", options=["— Seleccioná —"] + dni_unicos, index=0)
with c2:
    st.write(""); st.write("")
    if st.button("🔄 Limpiar selección"): st.experimental_rerun()

if dni_sel and dni_sel != "— Seleccioná —":
    mask = (dni_series == str(dni_sel).strip())
    if not mask.any():
        st.info("No se encontró ese DNI en la base.")
    else:
        row_idx = mask[mask].index[0]
        valores = df.iloc[row_idx, COL_START:COL_END+1].tolist()
        registros = []
        for h, v in zip(temas, valores):
            if not h: continue
            f = parse_fecha(v)
            if f is not None:
                registros.append({"Tema": h, "Fecha": f.strftime("%d/%m/%Y")})

        total_realizadas = len(registros)
        total_temarios = len(temas)
        colA, colB, colC = st.columns(3)
        with colA: st.metric("Capacitaciones realizadas", total_realizadas)
        with colB: st.metric("Total de temas", total_temarios)
        with colC:
            pct = 0 if total_temarios == 0 else round(100*total_realizadas/total_temarios, 1)
            st.metric("% de avance", f"{pct}%")

        st.divider()
        if total_realizadas == 0:
            st.warning("No se registran capacitaciones realizadas para este DNI.")
        else:
            import pandas as pd
            df_out = pd.DataFrame(registros)
            st.subheader("✅ Capacitaciones realizadas")
            st.dataframe(df_out, use_container_width=True)
            csv = df_out.to_csv(index=False).encode("utf-8-sig")
            st.download_button("⬇️ Descargar CSV", data=csv, file_name=f"capacitaciones_realizadas_{dni_sel}.csv", mime="text/csv")
else:
    st.info("Elegí un DNI para comenzar.")