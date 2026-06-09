import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime
import re

# =====================================================
# CONFIG
# =====================================================

st.set_page_config(
    page_title="Agenda PRO",
    layout="wide"
)

st.title("📊 Sistema Agenda Automática PRO")

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

# 🔥 TU SHEET ID
SHEET_ID = "1vOcVAGUzQjVGMKnFZUTQ0SLM7lBGBgAhK6RHGKJyBpk"

sheet = client.open_by_key(SHEET_ID)

st.success("✅ Conectado a Google Sheets")

# =====================================================
# MOSTRAR HOJAS
# =====================================================

st.write(
    "📄 Hojas disponibles:",
    [ws.title for ws in sheet.worksheets()]
)

# =====================================================
# FUNCIONES
# =====================================================

def clean_dataframe(df):

    df = df.copy()

    # limpiar nombres
    df.columns = [str(c).strip() for c in df.columns]

    # eliminar columnas vacías
    df = df.dropna(axis=1, how="all")

    # eliminar duplicadas
    df = df.loc[:, ~pd.Index(df.columns).duplicated(keep="first")]

    return df

# =====================================================

def extract_hour(text):

    if pd.isna(text):
        return ""

    text = str(text)

    # buscar HH:MM
    match = re.search(r'(\d{1,2}:\d{2})', text)

    if match:
        return match.group(1)

    return ""

# =====================================================

def format_chilean_money(value):

    try:

        value = float(value)

        return f"{int(value):,}".replace(",", ".")

    except:

        return value

# =====================================================

def format_date(value):

    try:

        return pd.to_datetime(value).strftime("%d-%m-%Y")

    except:

        return value

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

    # =================================================
    # VALIDACIÓN
    # =================================================

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
    # DEBUG COLUMNAS
    # =================================================

    st.write("📌 Ordenes cols:")
    st.write(ordenes.columns.tolist())

    st.write("📌 Track cols:")
    st.write(track.columns.tolist())

    st.write("📌 Maestro cols:")
    st.write(maestro.columns.tolist())

    # =================================================
    # VALIDAR COLUMNAS NECESARIAS
    # =================================================

    required_ordenes = [
        "O/C Cliente",
        "Fecha Entrega",
        "Instrucciones"
    ]

    required_track = [
        "PO Number",
        "Sold to Name",
        "Delivered Quantity",
        "Delivered Amount"
    ]

    required_maestro = [
        "Num Order",
        "Departamento",
        "PD"
    ]

    for col in required_ordenes:

        if col not in ordenes.columns:
            st.error(f"❌ Falta '{col}' en Ordenes")
            st.stop()

    for col in required_track:

        if col not in track.columns:
            st.error(f"❌ Falta '{col}' en Track")
            st.stop()

    for col in required_maestro:

        if col not in maestro.columns:
            st.error(f"❌ Falta '{col}' en Maestro")
            st.stop()

    # =================================================
    # EXTRAER HORA
    # =================================================

    ordenes["Hora"] = ordenes["Instrucciones"].apply(
        extract_hour
    )

    # =================================================
    # TRACK
    # =================================================

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

    # =================================================
    # MAESTRO
    # =================================================

    maestro_final = maestro[[
        "Num Order",
        "Departamento",
        "PD"
    ]].copy()

    # =================================================
    # ORDENES
    # =================================================

    ordenes_final = ordenes[[
        "O/C Cliente",
        "Fecha Entrega",
        "Hora"
    ]].copy()

    ordenes_final = ordenes_final.rename(columns={
        "O/C Cliente": "Num Order",
        "Fecha Entrega": "Fecha de entrega"
    })

    # =================================================
    # CONVERTIR KEYS A STRING
    # =================================================

    track_final["Num Order"] = (
        track_final["Num Order"]
        .astype(str)
        .str.strip()
    )

    maestro_final["Num Order"] = (
        maestro_final["Num Order"]
        .astype(str)
        .str.strip()
    )

    ordenes_final["Num Order"] = (
        ordenes_final["Num Order"]
        .astype(str)
        .str.strip()
    )

    # =================================================
    # DEBUG MERGE
    # =================================================

    st.write("📌 Track Final")
    st.dataframe(track_final.head())

    st.write("📌 Maestro Final")
    st.dataframe(maestro_final.head())

    st.write("📌 Ordenes Final")
    st.dataframe(ordenes_final.head())

    # =================================================
    # MERGES
    # =================================================

    try:

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

    except Exception as e:

        st.error(f"❌ Error merge: {e}")

        st.stop()

    # =================================================
    # FORMATO
    # =================================================

    df["Monto"] = df["Monto"].apply(
        format_chilean_money
    )

    df["Fecha de entrega"] = df[
        "Fecha de entrega"
    ].apply(format_date)

    # =================================================
    # ORDEN FINAL
    # =================================================

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

    # =================================================
    # LIMPIEZA FINAL
    # =================================================

    df = df.fillna("")

    df = df.drop_duplicates()

    # =================================================
    # SUBIR A GOOGLE SHEETS
    # =================================================

    upload_to_sheet(df, "Agenda Final")

    st.success("☁️ Agenda subida a Google Sheets")

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
# MOSTRAR ÚLTIMA ACTUALIZACIÓN
# =====================================================

try:

    ws = sheet.worksheet("Agenda Final")

    last_update = ws.acell("J1").value

    st.info(last_update)

except:

    st.warning("Sin actualizaciones aún")
