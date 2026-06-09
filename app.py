import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

# =========================
# CONFIGURACIÓN GENERAL
# =========================
st.set_page_config(page_title="Agenda System", layout="wide")
st.title("📊 Sistema de Agenda con Google Sheets")

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

# =========================
# CONECTAR SHEET
# =========================
SHEET_ID = "PEGA_AQUI_TU_SHEET_ID"

sheet = client.open_by_key(SHEET_ID)

st.success("✅ Conectado a Google Sheets")

# =========================
# MOSTRAR HOJAS
# =========================
worksheets = sheet.worksheets()

st.subheader("📄 Hojas disponibles")
st.write([ws.title for ws in worksheets])

# =========================
# FUNCIÓN PARA LEER HOJAS
# =========================
def load_sheet(name):
    ws = sheet.worksheet(name)
    data = ws.get_all_records()
    return pd.DataFrame(data)

# =========================
# CARGA DE DATOS
# =========================
try:
    ordenes = load_sheet("ordenes")
    track = load_sheet("track")
    maestro = load_sheet("maestro")

    st.subheader("📌 Ordenes")
    st.dataframe(ordenes)

    st.subheader("📌 Track")
    st.dataframe(track)

    st.subheader("📌 Maestro")
    st.dataframe(maestro)

except Exception as e:
    st.error(f"❌ Error cargando datos: {e}")

# =========================
# GENERAR AGENDA
# =========================
st.divider()

if st.button("🚀 Generar Agenda"):

    try:
        df = ordenes.merge(track, on="Order Number", how="left")
        df = df.merge(maestro, on="Order Number", how="left")

        df = df.drop_duplicates()
        df = df.fillna("")

        st.success("✅ Agenda generada")

        st.dataframe(df)

        # =========================
        # GUARDAR EN GOOGLE SHEETS
        # =========================
        ws = sheet.worksheet("agenda_final")
        ws.clear()

        ws.update(
            [df.columns.tolist()] + df.values.tolist()
        )

        st.success("💾 Agenda guardada en Google Sheets")

    except Exception as e:
        st.error(f"❌ Error generando agenda: {e}")
