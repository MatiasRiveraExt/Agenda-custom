import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime

# =========================
# CONFIG
# =========================
st.set_page_config(page_title="Agenda PRO", layout="wide")
st.title("📊 Agenda Automática PRO")

# =========================
# AUTH GOOGLE
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

SHEET_ID = "PEGA_AQUI_TU_SHEET_ID"
sheet = client.open_by_key(SHEET_ID)

st.success("✅ Conectado a Google Sheets")

# =========================
# TIMESTAMP
# =========================
def set_last_update():
    ws = sheet.worksheet("Agenda Final")
    ws.update("H1", [["Última actualización: " + datetime.now().strftime("%Y-%m-%d %H:%M:%S")]])

# =========================
# SUBIDA DE ARCHIVOS
# =========================
st.subheader("📤 Cargar archivos Excel")

ordenes_file = st.file_uploader("Ordenes", type=["xlsx"])
track_file = st.file_uploader("Track", type=["xlsx"])
maestro_file = st.file_uploader("Maestro", type=["xlsx"])

# =========================
# FUNCIÓN EXCEL → SHEETS
# =========================
def upload_to_sheet(df, sheet_name):
    ws = sheet.worksheet(sheet_name)
    ws.clear()
    ws.update([df.columns.tolist()] + df.fillna("").values.tolist())

# =========================
# PROCESO PRINCIPAL
# =========================
if st.button("🚀 Actualizar Base y Generar Agenda"):

    if not ordenes_file or not track_file or not maestro_file:
        st.error("❌ Debes subir los 3 archivos")
    else:

        # Leer Excel
        ordenes = pd.read_excel(ordenes_file)
        track = pd.read_excel(track_file)
        maestro = pd.read_excel(maestro_file)

        st.success("📥 Archivos cargados correctamente")

        # =========================
        # NORMALIZAR COLUMNAS
        # =========================
        ordenes.columns = ordenes.columns.str.strip()
        track.columns = track.columns.str.strip()
        maestro.columns = maestro.columns.str.strip()

        # =========================
        # SUBIR A SHEETS
        # =========================
        upload_to_sheet(ordenes, "Ordenes")
        upload_to_sheet(track, "Track")
        upload_to_sheet(maestro, "Maestro")

        st.success("☁️ Datos actualizados en Google Sheets")

        # =========================
        # GENERAR AGENDA
        # =========================
        df = ordenes.merge(track, on="Order Number", how="left")
        df = df.merge(maestro, on="Order Number", how="left")

        df = df.fillna("")
        df = df.drop_duplicates()

        upload_to_sheet(df, "Agenda Final")

        set_last_update()

        st.success("🎯 Agenda generada correctamente")

        st.dataframe(df)

# =========================
# MOSTRAR ÚLTIMA ACTUALIZACIÓN
# =========================
try:
    ws = sheet.worksheet("Agenda Final")
    last_update = ws.acell("H1").value
    st.info(last_update)
except:
    st.warning("Sin registro de actualización aún")
