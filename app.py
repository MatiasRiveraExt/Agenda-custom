import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime

# =========================
# CONFIG
# =========================
st.set_page_config(page_title="Agenda PRO", layout="wide")
st.title("📊 Sistema de Agenda Automática PRO")

# =========================
# AUTH GOOGLE SHEETS
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
# MOSTRAR HOJAS
# =========================
worksheets = sheet.worksheets()
st.write("📄 Hojas disponibles:", [ws.title for ws in worksheets])

# =========================
# FUNCIÓN DE LECTURA
# =========================
def load_sheet(name):
    ws = sheet.worksheet(name)
    return pd.DataFrame(ws.get_all_records())

# =========================
# 🔥 FUNCIÓN FIX DEFINITIVA (ANTI ERROR JSON)
# =========================
def upload_to_sheet(df, sheet_name):
    ws = sheet.worksheet(sheet_name)
    ws.clear()

    df_clean = df.copy()

    # 🔥 eliminar NaN
    df_clean = df_clean.fillna("")

    # 🔥 convertir TODO a string seguro
    df_clean = df_clean.astype(str)

    # 🔥 construir matriz 2D limpia
    values = []
    values.append([str(col) for col in df_clean.columns])

    for row in df_clean.values:
        clean_row = [str(cell) for cell in row]
        values.append(clean_row)

    ws.update(values)

# =========================
# SUBIDA DE ARCHIVOS
# =========================
st.subheader("📤 Subir archivos Excel")

ordenes_file = st.file_uploader("Ordenes", type=["xlsx"])
track_file = st.file_uploader("Track", type=["xlsx"])
maestro_file = st.file_uploader("Maestro", type=["xlsx"])

# =========================
# PROCESO PRINCIPAL
# =========================
if st.button("🚀 Generar Agenda PRO"):

    if not ordenes_file or not track_file or not maestro_file:
        st.error("❌ Debes subir los 3 archivos")
        st.stop()

    # =========================
    # LEER EXCEL
    # =========================
    ordenes = pd.read_excel(ordenes_file)
    track = pd.read_excel(track_file)
    maestro = pd.read_excel(maestro_file)

    st.success("📥 Archivos cargados correctamente")

    # =========================
    # LIMPIEZA COLUMNAS
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

    st.success("☁️ Datos subidos a Google Sheets")

    # =========================
    # MERGE
    # =========================
    df = ordenes.merge(track, on="Order Number", how="left")
    df = df.merge(maestro, on="Order Number", how="left")

    df = df.fillna("")
    df = df.drop_duplicates()

    # =========================
    # GUARDAR AGENDA
    # =========================
    upload_to_sheet(df, "Agenda Final")

    # timestamp
    ws = sheet.worksheet("Agenda Final")
    ws.update("H1", [[f"Última actualización: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"]])

    st.success("🎯 Agenda generada correctamente")

    st.dataframe(df)

# =========================
# TIMESTAMP DISPLAY
# =========================
try:
    ws = sheet.worksheet("Agenda Final")
    st.info(ws.acell("H1").value)
except:
    st.warning("Sin actualizaciones aún")
