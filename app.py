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
st.title("📊 Agenda Automática PRO - FIX FILTROS")

# =====================================================
# GOOGLE SHEETS
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
# HELPERS
# =====================================================

def normalize_columns(df):
    df = df.copy()
    df.columns = df.columns.astype(str).str.strip()
    return df


def safe_datetime(series):
    return pd.to_datetime(series, errors="coerce")


def extract_hour(text):
    if pd.isna(text):
        return ""
    match = re.search(r'(\d{1,2}:\d{2})', str(text))
    return match.group(1) if match else ""


def normalize_order(x):
    if pd.isna(x):
        return ""
    x = str(x).strip()
    x = re.sub(r"\.0$", "", x)
    return x.lstrip("0")


def format_money(v):
    try:
        v = float(str(v).replace(".", "").replace(",", ""))
        return f"{int(v):,}".replace(",", ".")
    except:
        return v


def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()

# =====================================================
# LOAD SHEET
# =====================================================

def load_latest():
    ws = sheet.worksheet("Agenda Final")
    data = ws.get_all_records()
    return pd.DataFrame(data) if data else pd.DataFrame()

# =====================================================
# UI
# =====================================================

st.subheader("📤 Generar agenda")

ordenes_file = st.file_uploader("Ordenes", type=["xlsx"])
track_file = st.file_uploader("Track", type=["xlsx"])
maestro_file = st.file_uploader("Maestro", type=["xlsx"])

# =====================================================
# GENERAR
# =====================================================

if st.button("🚀 Generar"):

    if not ordenes_file or not track_file or not maestro_file:
        st.error("Faltan archivos")
        st.stop()

    ordenes = normalize_columns(pd.read_excel(ordenes_file))
    track = normalize_columns(pd.read_excel(track_file))
    maestro = normalize_columns(pd.read_excel(maestro_file))

    st.success("📥 Archivos cargados")

    # =================================================
    # TRACK
    # =================================================

    track_final = track[[
        "Created On",
        "PO Number",
        "Sold to Name",
        "Delivered Quantity",
        "Delivered Amount"
    ]].rename(columns={
        "Created On": "Fecha Creación",
        "PO Number": "Num Order",
        "Sold to Name": "Cliente",
        "Delivered Quantity": "Unidades",
        "Delivered Amount": "Monto"
    })

    # =================================================
    # MAESTRO + ORDENES
    # =================================================

    maestro_final = maestro[["Num Order", "Departamento", "PD"]]

    ordenes_final = ordenes[["O/C Cliente", "Fecha Entrega", "Instrucciones"]].rename(columns={
        "O/C Cliente": "Num Order",
        "Fecha Entrega": "Fecha de entrega"
    })

    ordenes_final["Hora"] = ordenes_final["Instrucciones"].apply(extract_hour)

    # =================================================
    # NORMALIZAR KEY
    # =================================================

    for df in [track_final, maestro_final, ordenes_final]:
        df["Num Order"] = df["Num Order"].apply(normalize_order)

    # =================================================
    # MERGE
    # =================================================

    df = track_final.merge(maestro_final, on="Num Order", how="left")
    df = df.merge(ordenes_final, on="Num Order", how="left")

    df["Monto"] = df["Monto"].apply(format_money)

    # =================================================
    # 🔥 FIX CRÍTICO FECHAS
    # =================================================

    df["Fecha Creación"] = safe_datetime(df["Fecha Creación"])
    df["Fecha de entrega"] = safe_datetime(df["Fecha de entrega"])

    ws = sheet.worksheet("Agenda Final")
    ws.clear()
    ws.update([df.columns.tolist()] + df.fillna("").values.tolist())

    st.success("☁️ Base actualizada")

# =====================================================
# LOAD + FILTER
# =====================================================

st.divider()
st.subheader("📦 Agenda")

latest_df = load_latest()

if not latest_df.empty:

    latest_df = normalize_columns(latest_df)

    # 🔥 FORZAR TIPO REAL (CLAVE)
    latest_df["Fecha Creación"] = pd.to_datetime(latest_df["Fecha Creación"], errors="coerce")
    latest_df["Fecha de entrega"] = pd.to_datetime(latest_df["Fecha de entrega"], errors="coerce")

    df_filtered = latest_df.copy()

    # =================================================
    # FILTROS
    # =================================================

    clientes = df_filtered["Cliente"].dropna().unique().tolist()
    clientes_sel = st.multiselect("Cliente", clientes)

    min_c, max_c = df_filtered["Fecha Creación"].min(), df_filtered["Fecha Creación"].max()
    min_e, max_e = df_filtered["Fecha de entrega"].min(), df_filtered["Fecha de entrega"].max()

    fecha_creacion = st.date_input("Fecha creación", value=(min_c.date(), max_c.date()))
    fecha_entrega = st.date_input("Fecha entrega", value=(min_e.date(), max_e.date()))

    incluir_null = st.checkbox("Incluir sin fecha entrega", True)

    # =================================================
    # CLIENTE
    # =================================================

    if clientes_sel:
        df_filtered = df_filtered[df_filtered["Cliente"].isin(clientes_sel)]

    # =================================================
    # 🔥 FIX FECHA CREACIÓN
    # =================================================

    start_c = pd.to_datetime(fecha_creacion[0])
    end_c = pd.to_datetime(fecha_creacion[1])

    df_filtered = df_filtered[
        (df_filtered["Fecha Creación"] >= start_c) &
        (df_filtered["Fecha Creación"] <= end_c)
    ]

    # =================================================
    # 🔥 FIX FECHA ENTREGA (ESTE ES EL PROBLEMA REAL)
    # =================================================

    start_e = pd.to_datetime(fecha_entrega[0])
    end_e = pd.to_datetime(fecha_entrega[1])

    entrega = df_filtered["Fecha de entrega"]

    mask = (entrega >= start_e) & (entrega <= end_e)

    if incluir_null:
        df_filtered = df_filtered[mask | entrega.isna()]
    else:
        df_filtered = df_filtered[mask]

    # =================================================
    # OUTPUT FINAL
    # =================================================

    df_filtered = df_filtered[[
        "Num Order",
        "Fecha Creación",
        "Cliente",
        "Departamento",
        "PD",
        "Unidades",
        "Fecha de entrega",
        "Hora",
        "Monto"
    ]]

    st.metric("Resultados", len(df_filtered))
    st.dataframe(df_filtered, use_container_width=True)

    st.download_button(
        "Descargar agenda",
        to_excel(df_filtered),
        file_name=f"Agenda_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.warning("Sin datos")
