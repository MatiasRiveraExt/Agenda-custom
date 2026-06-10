import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime
from io import BytesIO
import re

# =====================================================
# CONFIG
# =====================================================

st.set_page_config(page_title="Agenda PRO DB", layout="wide")
st.title("📊 Agenda Automática PRO - FIX FINAL GSheets")

# =====================================================
# GOOGLE SHEETS
# =====================================================

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

creds = Credentials.from_service_account_info(
    st.secrets["gcp_service_account"],
    scopes=SCOPES
)

client = gspread.authorize(creds)

SHEET_ID = "1vOcVAGUzQjVGMKnFZUTQ0SLM7lBGBgAhK6RHGKJyBpk"
sheet = client.open_by_key(SHEET_ID)

st.success("✅ Conectado a Google Sheets")

# =====================================================
# HELPERS
# =====================================================

def normalize_columns(df):
    df = df.copy()
    df.columns = df.columns.astype(str).str.strip()
    return df


def normalize_order(x):
    if pd.isna(x):
        return ""
    return str(x).strip().replace(".0", "").lstrip("0")


def extract_hour(text):
    if pd.isna(text):
        return ""
    m = re.search(r'(\d{1,2}:\d{2})', str(text))
    return m.group(1) if m else ""


def format_money(v):
    try:
        v = float(str(v).replace(".", "").replace(",", ""))
        return f"{int(v):,}".replace(",", ".")
    except:
        return v


def to_datetime(series):
    return pd.to_datetime(series, errors="coerce")

# =====================================================
# 🔥 CLEAN FINAL PARA GOOGLE SHEETS (FIX CLAVE)
# =====================================================

def clean_for_gsheet(df):
    df = df.copy()

    # reemplazar NaT / NaN primero
    df = df.fillna("")

    # convertir todo a string puro (evita Timestamp/numpy issues)
    for col in df.columns:
        df[col] = df[col].astype(str)

    # limpiar basura común
    df = df.replace({
        "nan": "",
        "NaT": "",
        "None": ""
    })

    return df


def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()

# =====================================================
# LOAD SHEET
# =====================================================

def load_latest():
    ws = sheet.worksheet("Agenda Final")
    data = ws.get_all_records()
    return pd.DataFrame(data) if data else pd.DataFrame()

# =====================================================
# UI
# =====================================================

st.subheader("📤 Generar agenda")

ordenes_file = st.file_uploader("Ordenes", type=["xlsx"])
track_file = st.file_uploader("Track", type=["xlsx"])
maestro_file = st.file_uploader("Maestro", type=["xlsx"])

# =====================================================
# GENERAR
# =====================================================

if st.button("🚀 Generar"):

    if not ordenes_file or not track_file or not maestro_file:
        st.error("Faltan archivos")
        st.stop()

    ordenes = normalize_columns(pd.read_excel(ordenes_file))
    track = normalize_columns(pd.read_excel(track_file))
    maestro = normalize_columns(pd.read_excel(maestro_file))

    st.success("📥 Archivos cargados")

    # =================================================
    # TRACK
    # =================================================

    track_final = track[[
        "Created On",
        "PO Number",
        "Sold to Name",
        "Delivered Quantity",
        "Delivered Amount"
    ]].rename(columns={
        "Created On": "Fecha Creación",
        "PO Number": "Num Order",
        "Sold to Name": "Cliente",
        "Delivered Quantity": "Unidades",
        "Delivered Amount": "Monto"
    })

    maestro_final = maestro.reindex(columns=["Num Order", "Departamento", "PD"])

    ordenes_final = ordenes.reindex(columns=["O/C Cliente", "Fecha Entrega", "Instrucciones"])

    ordenes_final = ordenes_final.rename(columns={
        "O/C Cliente": "Num Order",
        "Fecha Entrega": "Fecha de entrega"
    })

    ordenes_final["Hora"] = ordenes_final.get("Instrucciones", "").apply(extract_hour)

    # =================================================
    # NORMALIZAR KEY
    # =================================================

    for df in [track_final, maestro_final, ordenes_final]:
        df["Num Order"] = df["Num Order"].apply(normalize_order)

    # =================================================
    # MERGE
    # =================================================

    df = track_final.merge(maestro_final, on="Num Order", how="left")
    df = df.merge(ordenes_final, on="Num Order", how="left")

    df["Monto"] = df["Monto"].apply(format_money)

    # =================================================
    # FECHAS SEGURAS
    # =================================================

    df["Fecha Creación"] = to_datetime(df["Fecha Creación"])
    df["Fecha de entrega"] = to_datetime(df["Fecha de entrega"])

    # =================================================
    # SUBIDA SEGURA A SHEETS (FIX REAL)
    # =================================================

    ws = sheet.worksheet("Agenda Final")
    ws.clear()

    clean_df = clean_for_gsheet(df)

    ws.update(
        [clean_df.columns.tolist()] +
        clean_df.values.tolist()
    )

    st.success("☁️ Base actualizada correctamente")

# =====================================================
# LOAD + FILTERS
# =====================================================

st.divider()
st.subheader("📦 Agenda")

latest_df = load_latest()

if not latest_df.empty:

    latest_df = normalize_columns(latest_df)

    df_filtered = latest_df.copy()

    # =================================================
    # CLIENTES
    # =================================================

    clientes = df_filtered["Cliente"].dropna().unique().tolist()
    clientes_sel = st.multiselect("Cliente", clientes)

    if clientes_sel:
        df_filtered = df_filtered[df_filtered["Cliente"].isin(clientes_sel)]

    # =================================================
    # FECHA CREACIÓN (FILTRO SIMPLE Y ESTABLE)
    # =================================================

    fechas = pd.to_datetime(df_filtered["Fecha Creación"], errors="coerce").dropna()

    if fechas.empty:
        st.warning("No hay fechas válidas")
        st.stop()

    min_c = fechas.min()
    max_c = fechas.max()

    fecha_creacion = st.date_input(
        "Fecha creación",
        value=(min_c.date(), max_c.date())
    )

    start_c = pd.to_datetime(fecha_creacion[0])
    end_c = pd.to_datetime(fecha_creacion[1])

    df_filtered["Fecha Creación"] = pd.to_datetime(df_filtered["Fecha Creación"], errors="coerce")

    df_filtered = df_filtered[
        df_filtered["Fecha Creación"].between(start_c, end_c)
    ]

    # =================================================
    # OUTPUT
    # =================================================

    cols = [
        "Num Order",
        "Fecha Creación",
        "Cliente",
        "Departamento",
        "PD",
        "Unidades",
        "Fecha de entrega",
        "Hora",
        "Monto"
    ]

    df_filtered = df_filtered.reindex(columns=cols)

    st.metric("Resultados", len(df_filtered))
    st.dataframe(df_filtered, use_container_width=True)

    st.download_button(
        "Descargar agenda",
        to_excel(df_filtered),
        file_name=f"Agenda_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.warning("Sin datos")
