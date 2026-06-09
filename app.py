import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime

# =========================
# CONFIG
# =========================
st.set_page_config(page_title="Agenda PRO", layout="wide")
st.title("📊 Sistema Agenda Automática PRO")

# =========================
# GOOGLE AUTH
# =========================
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

# =========================
# HOJAS DISPONIBLES
# =========================
st.write("📄 Hojas:", [ws.title for ws in sheet.worksheets()])

# =========================
# NORMALIZADOR DE COLUMNAS (CLAVE)
# =========================
def normalize_columns(df):
    df = df.copy()
    df.columns = df.columns.str.strip()

    mapping = {
        "PO Number": "Order Number",
        "Num Order": "Order Number",
        "O/C Cliente": "Order Number",
        "Order": "Order Number",
        "Order ID": "Order Number"
    }

    df = df.rename(columns=mapping)
    return df

# =========================
# UPLOAD SAFE A SHEETS
# =========================
def upload_to_sheet(df, sheet_name):
    ws = sheet.worksheet(sheet_name)
    ws.clear()

    df = df.fillna("").astype(str)

    values = [df.columns.tolist()]
    for row in df.values:
        values.append([str(x) for x in row])

    ws.update(values)

# =========================
# UI UPLOAD
# =========================
st.subheader("📤 Subir archivos Excel")

ordenes_file = st.file_uploader("Ordenes", type=["xlsx"])
track_file = st.file_uploader("Track", type=["xlsx"])
maestro_file = st.file_uploader("Maestro", type=["xlsx"])

# =========================
# PROCESO
# =========================
if st.button("🚀 Generar Agenda PRO"):

    if not ordenes_file or not track_file or not maestro_file:
        st.error("❌ Debes subir los 3 archivos")
        st.stop()

    # =========================
    # LECTURA
    # =========================
    ordenes = pd.read_excel(ordenes_file)
    track = pd.read_excel(track_file)
    maestro = pd.read_excel(maestro_file)

    st.success("📥 Archivos cargados")

    # =========================
    # LIMPIEZA
    # =========================
    ordenes = normalize_columns(ordenes)
    track = normalize_columns(track)
    maestro = normalize_columns(maestro)

    # =========================
    # VALIDACIÓN
    # =========================
    required = "Order Number"

    for name, df in [("Ordenes", ordenes), ("Track", track), ("Maestro", maestro)]:
        if required not in df.columns:
            st.error(f"❌ Falta 'Order Number' en {name}")
            st.write(df.columns)
            st.stop()

    # =========================
    # SUBIR A SHEETS
    # =========================
    upload_to_sheet(ordenes, "Ordenes")
    upload_to_sheet(track, "Track")
    upload_to_sheet(maestro, "Maestro")

    st.success("☁️ Datos subidos a Google Sheets")

    # =========================
    # MERGE SEGURO
    # =========================
    df = ordenes.merge(track, on="Order Number", how="left")
    df = df.merge(maestro, on="Order Number", how="left")

    df = df.fillna("")
    df = df.drop_duplicates()

    # =========================
    # GUARDAR AGENDA FINAL
    # =========================
    upload_to_sheet(df, "Agenda Final")

    # timestamp
    ws = sheet.worksheet("Agenda Final")
    ws.update("H1", [[
        f"Última actualización: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
    ]])

    st.success("🎯 Agenda generada correctamente")

    st.dataframe(df)

# =========================
# TIMESTAMP VIEW
# =========================
try:
    ws = sheet.worksheet("Agenda Final")
    st.info(ws.acell("H1").value)
except:
    st.warning("Sin actualizaciones aún")
