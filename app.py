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
st.title("📊 Agenda Automática PRO - FIX DEFINITIVO")

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
# 🔥 NORMALIZADOR REAL DE COLUMNAS (CLAVE)
# =====================================================

def normalize_columns(df):
    df = df.copy()
    df.columns = (
        df.columns
        .astype(str)
        .str.strip()
        .str.replace("\n", "", regex=False)
        .str.replace("\t", "", regex=False)
    )
    return df

# =====================================================
# FECHAS ROBUSTAS
# =====================================================

def safe_date(series):
    s = series.replace(["", " ", "N/A", "nan", "None"], pd.NaT)
    return pd.to_datetime(s, errors="coerce")

# =====================================================
# HELPERS
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


def normalize_order(x):
    if pd.isna(x):
        return ""
    x = str(x).strip()
    x = re.sub(r"\.0$", "", x)
    return x.lstrip("0")


def format_money(value):
    try:
        value = float(str(value).replace(".", "").replace(",", ""))
        return f"{int(value):,}".replace(",", ".")
    except:
        return value


def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Agenda")
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

st.subheader("📤 Generar Agenda")

ordenes_file = st.file_uploader("Ordenes", type=["xlsx"])
track_file = st.file_uploader("Track", type=["xlsx"])
maestro_file = st.file_uploader("Maestro", type=["xlsx"])

# =====================================================
# GENERAR
# =====================================================

if st.button("🚀 Generar Agenda"):

    if not ordenes_file or not track_file or not maestro_file:
        st.error("Faltan archivos")
        st.stop()

    ordenes = clean_dataframe(pd.read_excel(ordenes_file))
    track = clean_dataframe(pd.read_excel(track_file))
    maestro = clean_dataframe(pd.read_excel(maestro_file))

    st.success("📥 Archivos cargados")

    # =================================================
    # VALIDACIÓN
    # =================================================

    required = [
        "Created On",
        "PO Number",
        "Sold to Name",
        "Delivered Quantity",
        "Delivered Amount"
    ]

    for c in required:
        if c not in track.columns:
            st.error(f"Falta en Track: {c}")
            st.stop()

    # =================================================
    # TRANSFORMACIÓN
    # =================================================

    track_final = track[required].rename(columns={
        "Created On": "Fecha Creación",
        "PO Number": "Num Order",
        "Sold to Name": "Cliente",
        "Delivered Quantity": "Suma de Unidades",
        "Delivered Amount": "Monto"
    })

    maestro_final = maestro[["Num Order", "Departamento", "PD"]]

    ordenes_final = ordenes[["O/C Cliente", "Fecha Entrega", "Instrucciones"]].rename(columns={
        "O/C Cliente": "Num Order",
        "Fecha Entrega": "Fecha de entrega"
    })

    ordenes_final["Hora"] = ordenes_final["Instrucciones"].apply(extract_hour)

    # normalización clave
    for df_ in [track_final, maestro_final, ordenes_final]:
        df_["Num Order"] = df_["Num Order"].apply(normalize_order)

    # merge
    df = track_final.merge(maestro_final, on="Num Order", how="left")
    df = df.merge(ordenes_final, on="Num Order", how="left")

    df["Monto"] = df["Monto"].apply(format_money)

    df["Fecha Creación"] = safe_date(df["Fecha Creación"])
    df["Fecha de entrega"] = safe_date(df["Fecha de entrega"])

    ws = sheet.worksheet("Agenda Final")
    ws.clear()
    ws.update([df.columns.tolist()] + df.fillna("").values.tolist())

    st.success("☁️ Base actualizada")

# =====================================================
# LOAD + FILTERS
# =====================================================

st.divider()
st.subheader("📦 Base histórica")

latest_df = load_latest()

if not latest_df.empty:

    # =================================================
    # 🔥 FIX CRÍTICO: COLUMNAS LIMPIAS
    # =================================================

    latest_df = normalize_columns(latest_df)

    latest_df["Fecha Creación"] = safe_date(latest_df["Fecha Creación"])
    latest_df["Fecha de entrega"] = safe_date(latest_df["Fecha de entrega"])

    df_filtered = latest_df.copy()

    st.markdown("## 🎯 Filtros")

    clientes = df_filtered["Cliente"].dropna().unique().tolist()
    clientes_sel = st.multiselect("Cliente", clientes)

    min_c, max_c = df_filtered["Fecha Creación"].min(), df_filtered["Fecha Creación"].max()
    min_e, max_e = df_filtered["Fecha de entrega"].min(), df_filtered["Fecha de entrega"].max()

    fecha_creacion = st.date_input("Fecha creación", value=(min_c.date(), max_c.date()))
    fecha_entrega = st.date_input("Fecha entrega", value=(min_e.date(), max_e.date()))

    incluir_null = st.checkbox("Incluir sin fecha entrega", True)

    # filtros cliente
    if clientes_sel:
        df_filtered = df_filtered[df_filtered["Cliente"].isin(clientes_sel)]

    # filtro creación
    start_c = pd.to_datetime(fecha_creacion[0])
    end_c = pd.to_datetime(fecha_creacion[1])

    df_filtered = df_filtered[
        df_filtered["Fecha Creación"].between(start_c, end_c)
    ]

    # filtro entrega
    start_e = pd.to_datetime(fecha_entrega[0])
    end_e = pd.to_datetime(fecha_entrega[1])

    entrega = df_filtered["Fecha de entrega"]

    mask = entrega.notna() & entrega.between(start_e, end_e)

    if incluir_null:
        df_filtered = df_filtered[entrega.isna() | mask]
    else:
        df_filtered = df_filtered[mask]

    # =================================================
    # 🔥 FIX FINAL ANTI-KEYERROR
    # =================================================

    cols = [
        "Num Order",
        "Fecha Creación",
        "Cliente",
        "Departamento",
        "PD",
        "Suma de Unidades",
        "Fecha de entrega",
        "Hora",
        "Monto"
    ]

    # DEBUG PROTEGIDO
    missing = [c for c in cols if c not in df_filtered.columns]

    if missing:
        st.error(f"Faltan columnas: {missing}")
        st.write("Columnas reales:", df_filtered.columns.tolist())
        st.stop()

    df_filtered = df_filtered[cols]

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
