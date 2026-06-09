import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime
import re

# =====================================================
# CONFIG
# =====================================================
st.set_page_config(page_title="Agenda PRO", layout="wide")

st.title("📊 Sistema Agenda Automática PRO")

# =====================================================
# GOOGLE AUTH
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
# FUNCIONES
# =====================================================

def clean_dataframe(df):

    df = df.copy()

    df.columns = [str(c).strip() for c in df.columns]

    df = df.dropna(axis=1, how="all")

    df = df.loc[:, ~pd.Index(df.columns).duplicated(keep="first")]

    return df

# =====================================================

def extract_hour(text):

    if pd.isna(text):
        return ""

    text = str(text)

    match = re.search(r'(\d{1,2}:\d{2})', text)

    if match:
        return match.group(1)

    return ""

# =====================================================

def upload_to_sheet(df, sheet_name):

    ws = sheet.worksheet(sheet_name)

    ws.clear()

    df = df.fillna("").astype(str)

    values = [df.columns.tolist()]

    for row in df.values:
        values.append([str(x) for x in row])

    ws.update(values)

# =====================================================
# UI
# =====================================================

st.subheader("📤 Subir archivos Excel")

ordenes_file = st.file_uploader(
    "Archivo Ordenes",
    type=["xlsx"]
)

track_file = st.file_uploader(
    "Archivo Track",
    type=["xlsx"]
)

maestro_file = st.file_uploader(
    "Archivo Maestro",
    type=["xlsx"]
)

# =====================================================
# MAIN
# =====================================================

if st.button("🚀 Generar Agenda"):

    if not ordenes_file or not track_file or not maestro_file:

        st.error("❌ Debes subir los 3 archivos")

        st.stop()

    # =================================================
    # LEER EXCEL
    # =================================================

    ordenes = pd.read_excel(ordenes_file)

    track = pd.read_excel(track_file)

    maestro = pd.read_excel(maestro_file)

    st.success("📥 Archivos cargados")

    # =================================================
    # LIMPIEZA
    # =================================================

    ordenes = clean_dataframe(ordenes)

    track = clean_dataframe(track)

    maestro = clean_dataframe(maestro)

    # =================================================
    # DEBUG
    # =================================================

    st.write("📌 Ordenes cols:")
    st.write(ordenes.columns.tolist())

    st.write("📌 Track cols:")
    st.write(track.columns.tolist())

    st.write("📌 Maestro cols:")
    st.write(maestro.columns.tolist())

    # =================================================
    # EXTRAER HORA
    # =================================================

    ordenes["Hora"] = ordenes["Instrucciones"].apply(extract_hour)

    # =================================================
    # CREAR DATAFRAMES NECESARIOS
    # =================================================

    # TRACK
    track_final = track[[
        "PO Number",
        "Delivered Quantity",
        "Delivered Amount"
    ]].copy()

    track_final = track_final.rename(columns={
        "PO Number": "Num Order",
        "Delivered Quantity": "Suma de Unidades",
        "Delivered Amount": "Monto"
    })

    # MAESTRO
    maestro_final = maestro[[
        "Order Number",
        "Departamento",
        "PD"
    ]].copy()

    maestro_final = maestro_final.rename(columns={
        "Order Number": "Num Order"
    })

    # ORDENES
    ordenes_final = ordenes[[
        "Order Number",
        "Fecha Entrega",
        "Hora"
    ]].copy()

    ordenes_final = ordenes_final.rename(columns={
        "Order Number": "Num Order",
        "Fecha Entrega": "Fecha de entrega"
    })

    # =================================================
    # MERGES
    # =================================================

    df = track_final.merge(
        maestro_final,
        on="Num Order",
        how="left"
    )

    df = df.merge(
        ordenes_final,
        on="Num Order",
        how="left"
    )

    # =================================================
    # ORDEN FINAL
    # =================================================

    df = df[[
        "Num Order",
        "Departamento",
        "PD",
        "Suma de Unidades",
        "Fecha de entrega",
        "Hora",
        "Monto"
    ]]

    # =================================================
    # LIMPIEZA FINAL
    # =================================================

    df = df.fillna("")

    df = df.drop_duplicates()

    # =================================================
    # SUBIR A GOOGLE SHEETS
    # =================================================

    upload_to_sheet(df, "Agenda Final")

    # =================================================
    # TIMESTAMP
    # =================================================

    ws = sheet.worksheet("Agenda Final")

    timestamp = datetime.now().strftime(
        "%Y-%m-%d %H:%M:%S"
    )

    ws.update(
        "J1",
        [[f"Última actualización: {timestamp}"]]
    )

    # =================================================
    # SUCCESS
    # =================================================

    st.success("🎯 Agenda generada correctamente")

    st.dataframe(df)

# =====================================================
# MOSTRAR TIMESTAMP
# =====================================================

try:

    ws = sheet.worksheet("Agenda Final")

    last_update = ws.acell("J1").value

    st.info(last_update)

except:

    st.warning("Sin actualizaciones aún")
