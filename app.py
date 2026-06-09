import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime

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

# 🔥 PEGA TU ID
SHEET_ID = "1vOcVAGUzQjVGMKnFZUTQ0SLM7lBGBgAhK6RHGKJyBpk"

sheet = client.open_by_key(SHEET_ID)

st.success("✅ Conectado a Google Sheets")

# =====================================================
# DEBUG HOJAS
# =====================================================
st.write(
    "📄 Hojas disponibles:",
    [ws.title for ws in sheet.worksheets()]
)

# =====================================================
# LIMPIEZA GENERAL
# =====================================================
def clean_dataframe(df):

    df = df.copy()

    # convertir nombres a string
    df.columns = [str(c).strip() for c in df.columns]

    # eliminar columnas vacías
    df = df.dropna(axis=1, how="all")

    # eliminar duplicadas exactas
    df = df.loc[:, ~pd.Index(df.columns).duplicated(keep="first")]

    return df

# =====================================================
# HACER COLUMNAS ÚNICAS
# =====================================================
def ensure_unique_columns(df):

    cols = []
    seen = {}

    for col in df.columns:

        if col not in seen:
            seen[col] = 0
            cols.append(col)

        else:
            seen[col] += 1
            cols.append(f"{col}_{seen[col]}")

    df.columns = cols

    return df

# =====================================================
# NORMALIZADOR
# =====================================================
def normalize_columns(df):

    mapping = {
        "PO Number": "Order Number",
        "Num Order": "Order Number",
        "O/C Cliente": "Order Number",
        "Order": "Order Number",
        "Order ID": "Order Number"
    }

    df = df.rename(columns=mapping)

    return df

# =====================================================
# UPLOAD GOOGLE SHEETS
# =====================================================
def upload_to_sheet(df, sheet_name):

    ws = sheet.worksheet(sheet_name)

    ws.clear()

    df = df.fillna("")

    # convertir TODO a string
    df = df.astype(str)

    values = [df.columns.tolist()]

    for row in df.values:
        values.append([str(x) for x in row])

    ws.update(values)

# =====================================================
# UI
# =====================================================
st.subheader("📤 Subir archivos Excel")

ordenes_file = st.file_uploader(
    "Ordenes",
    type=["xlsx"]
)

track_file = st.file_uploader(
    "Track",
    type=["xlsx"]
)

maestro_file = st.file_uploader(
    "Maestro",
    type=["xlsx"]
)

# =====================================================
# MAIN
# =====================================================
if st.button("🚀 Generar Agenda PRO"):

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
    # NORMALIZAR
    # =================================================
    ordenes = normalize_columns(ordenes)
    track = normalize_columns(track)
    maestro = normalize_columns(maestro)

    # =================================================
    # HACER COLUMNAS ÚNICAS
    # =================================================
    ordenes = ensure_unique_columns(ordenes)
    track = ensure_unique_columns(track)
    maestro = ensure_unique_columns(maestro)

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
    # VALIDAR KEY
    # =================================================
    required = "Order Number"

    for name, df in [
        ("Ordenes", ordenes),
        ("Track", track),
        ("Maestro", maestro)
    ]:

        if required not in df.columns:

            st.error(f"❌ No existe '{required}' en {name}")

            st.write(df.columns.tolist())

            st.stop()

    # =================================================
    # SUBIR A SHEETS
    # =================================================
    upload_to_sheet(ordenes, "Ordenes")

    upload_to_sheet(track, "Track")

    upload_to_sheet(maestro, "Maestro")

    st.success("☁️ Datos subidos a Google Sheets")

    # =================================================
    # MERGE
    # =================================================
    try:

        df = ordenes.merge(
            track,
            on="Order Number",
            how="left"
        )

        df = df.merge(
            maestro,
            on="Order Number",
            how="left"
        )

    except Exception as e:

        st.error(f"❌ Error en merge: {e}")

        st.stop()

    # =================================================
    # LIMPIEZA FINAL
    # =================================================
    df = df.fillna("")

    df = df.drop_duplicates()

    # =================================================
    # GUARDAR AGENDA
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
        "H1",
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

    last_update = ws.acell("H1").value

    st.info(last_update)

except:

    st.warning("Sin actualizaciones aún")
