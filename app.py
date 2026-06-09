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

st.set_page_config(
    page_title="Agenda PRO DB",
    layout="wide"
)

st.title("📊 Agenda Automática PRO (Base de Datos)")

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

# 🔥 FIX CLAVE: NORMALIZAR OC
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

def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Agenda")
    return output.getvalue()

# =====================================================
# 🔥 UPSERT REAL (BASE DE DATOS)
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
# UI
# =====================================================

st.subheader("📤 Subir archivos Excel")

ordenes_file = st.file_uploader("Ordenes", type=["xlsx"])
track_file = st.file_uploader("Track", type=["xlsx"])
maestro_file = st.file_uploader("Maestro", type=["xlsx"])

# =====================================================
# MAIN
# =====================================================

if st.button("🚀 Generar Agenda"):

    if not ordenes_file or not track_file or not maestro_file:
        st.error("❌ Faltan archivos")
        st.stop()

    # =================================================
    # LOAD
    # =================================================

    ordenes = clean_dataframe(pd.read_excel(ordenes_file))
    track = clean_dataframe(pd.read_excel(track_file))
    maestro = clean_dataframe(pd.read_excel(maestro_file))

    st.success("📥 Archivos cargados")

    # =================================================
    # VALIDACIÓN
    # =================================================

    required_ordenes = ["O/C Cliente", "Fecha Entrega", "Instrucciones"]
    required_track = ["PO Number", "Sold to Name", "Delivered Quantity", "Delivered Amount"]
    required_maestro = ["Num Order", "Departamento", "PD"]

    for col in required_ordenes:
        if col not in ordenes.columns:
            st.error(f"❌ Falta {col}")
            st.stop()

    for col in required_track:
        if col not in track.columns:
            st.error(f"❌ Falta {col}")
            st.stop()

    for col in required_maestro:
        if col not in maestro.columns:
            st.error(f"❌ Falta {col}")
            st.stop()

    # =================================================
    # HORA
    # =================================================

    ordenes["Hora"] = ordenes["Instrucciones"].apply(extract_hour)

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
    # NORMALIZAR KEYS
    # =================================================

    track_final["Num Order"] = track_final["Num Order"].apply(normalize_order)
    maestro_final["Num Order"] = maestro_final["Num Order"].apply(normalize_order)
    ordenes_final["Num Order"] = ordenes_final["Num Order"].apply(normalize_order)

    # =================================================
    # MERGE
    # =================================================

    df = track_final.merge(maestro_final, on="Num Order", how="left")
    df = df.merge(ordenes_final, on="Num Order", how="left")

    # =================================================
    # FORMATO
    # =================================================

    df["Monto"] = df["Monto"].apply(format_chilean_money)
    df["Fecha de entrega"] = df["Fecha de entrega"].apply(format_date)

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

    df = df.fillna("")
    df = df.drop_duplicates()

    # =================================================
    # UPSERT A SHEETS (🔥 NUEVO SISTEMA)
    # =================================================

    upsert_to_sheet(df, "Agenda Final")

    st.success("☁️ Datos sincronizados (UPSERT)")

    # =================================================
    # TIMESTAMP SOLO VISUAL (NO BD)
    # =================================================

    ws = sheet.worksheet("Agenda Final")
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    ws.update("J1", [[f"Última actualización: {timestamp}"]])

    # =================================================
    # OUTPUT
    # =================================================

    st.success("🎯 Agenda generada correctamente")
    st.dataframe(df)

    excel_data = to_excel(df)

    st.download_button(
        "📥 Descargar Agenda Excel",
        excel_data,
        file_name=f"Agenda_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# =====================================================
# STATUS
# =====================================================

try:
    ws = sheet.worksheet("Agenda Final")
    st.info(ws.acell("J1").value)
except:
    st.warning("Sin datos aún")
