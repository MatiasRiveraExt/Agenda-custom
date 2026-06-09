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
st.title("📊 Agenda Automática PRO - Base de Datos")

# =====================================================
# GOOGLE SHEETS AUTH
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
# UTILIDADES
# =====================================================

def clean_dataframe(df):
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    df = df.dropna(axis=1, how="all")
    df = df.loc[:, ~pd.Index(df.columns).duplicated(keep="first")]
    return df

def extract_hour(text):
    if pd.isna(text):
        return ""
    match = re.search(r'(\d{1,2}:\d{2})', str(text))
    return match.group(1) if match else ""

def normalize_order(x):
    if pd.isna(x):
        return ""
    x = str(x).strip()
    x = x.replace("\u00a0", "")
    x = re.sub(r"\.0$", "", x)
    x = x.lstrip("0")
    return x

def format_chilean_money(value):
    try:
        value = float(value)
        return f"{int(value):,}".replace(",", ".")
    except:
        return value

def format_date(value):
    try:
        return pd.to_datetime(value).strftime("%d-%m-%Y")
    except:
        return value

# =====================================================
# EXPORT EXCEL
# =====================================================

def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Agenda")
    return output.getvalue()

# =====================================================
# UPSERT (BASE DE DATOS REAL)
# =====================================================

def upsert_to_sheet(df, sheet_name):

    ws = sheet.worksheet(sheet_name)

    existing = ws.get_all_records()

    if len(existing) == 0:
        ws.clear()
        df = df.fillna("").astype(str)

        values = [df.columns.tolist()]
        for row in df.values:
            values.append([str(x) for x in row])

        ws.update(values)
        return

    existing_df = pd.DataFrame(existing)

    if "Num Order" not in existing_df.columns:
        ws.clear()
        existing_df = pd.DataFrame()

    df["Num Order"] = df["Num Order"].astype(str)
    existing_df["Num Order"] = existing_df["Num Order"].astype(str)

    existing_df = existing_df.set_index("Num Order")
    df = df.set_index("Num Order")

    combined = df.combine_first(existing_df)
    combined.update(df)

    combined = combined.reset_index()

    ws.clear()

    combined = combined.fillna("").astype(str)

    values = [combined.columns.tolist()]
    for row in combined.values:
        values.append([str(x) for x in row])

    ws.update(values)

# =====================================================
# LOAD DESDE SHEETS (DESCARGA SIN EXCEL)
# =====================================================

def load_latest_from_sheet():

    ws = sheet.worksheet("Agenda Final")
    data = ws.get_all_records()

    if not data:
        return pd.DataFrame()

    return pd.DataFrame(data)

# =====================================================
# UI - INPUTS
# =====================================================

st.subheader("📤 Generar nueva agenda")

ordenes_file = st.file_uploader("Ordenes", type=["xlsx"])
track_file = st.file_uploader("Track", type=["xlsx"])
maestro_file = st.file_uploader("Maestro", type=["xlsx"])

# =====================================================
# GENERAR
# =====================================================

if st.button("🚀 Generar Agenda"):

    if not ordenes_file or not track_file or not maestro_file:
        st.error("❌ Faltan archivos")
        st.stop()

    ordenes = clean_dataframe(pd.read_excel(ordenes_file))
    track = clean_dataframe(pd.read_excel(track_file))
    maestro = clean_dataframe(pd.read_excel(maestro_file))

    ordenes["Hora"] = ordenes["Instrucciones"].apply(extract_hour)

    track_final = track[[
        "PO Number",
        "Sold to Name",
        "Delivered Quantity",
        "Delivered Amount"
    ]].copy()

    track_final = track_final.rename(columns={
        "PO Number": "Num Order",
        "Sold to Name": "Cliente",
        "Delivered Quantity": "Suma de Unidades",
        "Delivered Amount": "Monto"
    })

    maestro_final = maestro[[
        "Num Order",
        "Departamento",
        "PD"
    ]].copy()

    ordenes_final = ordenes[[
        "O/C Cliente",
        "Fecha Entrega",
        "Hora"
    ]].copy()

    ordenes_final = ordenes_final.rename(columns={
        "O/C Cliente": "Num Order",
        "Fecha Entrega": "Fecha de entrega"
    })

    track_final["Num Order"] = track_final["Num Order"].apply(normalize_order)
    maestro_final["Num Order"] = maestro_final["Num Order"].apply(normalize_order)
    ordenes_final["Num Order"] = ordenes_final["Num Order"].apply(normalize_order)

    df = track_final.merge(maestro_final, on="Num Order", how="left")
    df = df.merge(ordenes_final, on="Num Order", how="left")

    df["Monto"] = df["Monto"].apply(format_chilean_money)
    df["Fecha de entrega"] = df["Fecha de entrega"].apply(format_date)

    df = df[[
        "Num Order",
        "Cliente",
        "Departamento",
        "PD",
        "Suma de Unidades",
        "Fecha de entrega",
        "Hora",
        "Monto"
    ]]

    df = df.fillna("").drop_duplicates()

    upsert_to_sheet(df, "Agenda Final")

    st.success("☁️ Agenda actualizada en base de datos")

    st.dataframe(df)

    st.download_button(
        "📥 Descargar Excel generado ahora",
        to_excel(df),
        file_name="Agenda_Nueva.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# =====================================================
# 🔥 DESCARGA DESDE BASE DE DATOS (SIN ARCHIVOS)
# =====================================================

st.divider()
st.subheader("📦 Descargar última agenda (Base de Datos)")

latest_df = load_latest_from_sheet()

if latest_df.empty:
    st.warning("No hay datos aún en la base")
else:

    st.success("📊 Última versión cargada desde Google Sheets")

    st.dataframe(latest_df, use_container_width=True)

    st.download_button(
        "📥 Descargar última agenda",
        to_excel(latest_df),
        file_name=f"Agenda_DB_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
