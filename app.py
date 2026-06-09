import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

st.set_page_config(page_title="Agenda App", layout="wide")
st.title("📊 Agenda con Google Sheets")

# =========================
# AUTH
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
# ABRIR SHEET
# =========================
SHEET_ID = "PEGA_AQUI_TU_SHEET_ID"

sheet = client.open_by_key(SHEET_ID)

st.success("✅ Conectado a Google Sheets")
st.write("Hojas disponibles:", sheet.sheetnames)
