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
st.title("📊 Agenda Automática PRO - Base Histórica")

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
# FUNCIONES
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
    x = x.replace("\u00a0", "")
    x = re.sub(r"\.0$", "", x)
    x = x.lstrip("0")
    return x


def format_chilean_money(value):
    try:
        value = float(str(value).replace(".", "").replace(",", ""))
        return f"{int(value):,}".replace(",", ".")
    except:
        return value


def parse_date(value):
    try:
        return pd.to_datetime(value).date()
    except:
        return pd.NaT


def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Agenda")
    return output.getvalue()

# =====================================================
# LOAD DB
# =====================================================

def load_latest_from_sheet():
    ws = sheet.worksheet("Agenda Final")
    data = ws.get_all_records()
    if not data:
        return pd.DataFrame()
    return pd.DataFrame(data)

# =====================================================
# UI INPUT
# =====================================================

st.subheader("📤 Actualizar Base Histórica")

ordenes_file = st.file_uploader("Ordenes", type=["xlsx"])
track_file = st.file_uploader("Track", type=["xlsx"])
maestro_file = st.file_uploader("Maestro", type=["xlsx"])

# =====================================================
# GENERAR AGENDA
# =====================================================

if st.button("🚀 Generar Agenda"):

    if not ordenes_file or not track_file or not maestro_file:
        st.error("❌ Debes subir los 3 archivos")
        st.stop()

    ordenes = clean_dataframe(pd.read_excel(ordenes_file))
    track = clean_dataframe(pd.read_excel(track_file))
    maestro = clean_dataframe(pd.read_excel(maestro_file))

    st.success("📥 Archivos cargados")

    required = ["Created On", "PO Number", "Sold to Name", "Delivered Quantity", "Delivered Amount"]

    for c in required:
        if c not in track.columns:
            st.error(f"Falta {c}")
            st.stop()

    # =================================================
    # TRANSFORMACIÓN
    # =================================================

    ordenes["Hora"] = ordenes["Instrucciones"].apply(extract_hour)

    track_final = track[required].rename(columns={
        "Created On": "Fecha Creación",
        "PO Number": "Num Order",
        "Sold to Name": "Cliente",
        "Delivered Quantity": "Suma de Unidades",
        "Delivered Amount": "Monto"
    })

    maestro_final = maestro[["Num Order", "Departamento", "PD"]]

    ordenes_final = ordenes[["O/C Cliente", "Fecha Entrega", "Hora"]].rename(columns={
        "O/C Cliente": "Num Order",
        "Fecha Entrega": "Fecha de entrega"
    })

    track_final["Num Order"] = track_final["Num Order"].apply(normalize_order)
    maestro_final["Num Order"] = maestro_final["Num Order"].apply(normalize_order)
    ordenes_final["Num Order"] = ordenes_final["Num Order"].apply(normalize_order)

    df = track_final.merge(maestro_final, on="Num Order", how="left")
    df = df.merge(ordenes_final, on="Num Order", how="left")

    df["Monto"] = df["Monto"].apply(format_chilean_money)

    df["Fecha Creación"] = df["Fecha Creación"].apply(parse_date)
    df["Fecha de entrega"] = df["Fecha de entrega"].apply(parse_date)

    df = df.fillna("")

    ws = sheet.worksheet("Agenda Final")
    ws.clear()
    ws.update([df.columns.tolist()] + df.values.tolist())

    st.success("☁️ Base actualizada")

# =====================================================
# LOAD + FILTERS
# =====================================================

st.divider()
st.subheader("📦 Base histórica")

latest_df = load_latest_from_sheet()

if not latest_df.empty:

    # =================================================
    # FIX CRÍTICO FECHAS (SEPARADO PARA FILTRO)
    # =================================================

    latest_df["Fecha Creación"] = pd.to_datetime(
        latest_df["Fecha Creación"],
        errors="coerce"
    ).dt.date

    latest_df["Fecha de entrega"] = pd.to_datetime(
        latest_df["Fecha de entrega"],
        errors="coerce"
    ).dt.date

    # copia segura para filtros
    df_filtered = latest_df.copy()

    st.markdown("## 🎯 Filtros")

    # CLIENTE
    clientes = df_filtered["Cliente"].dropna().unique().tolist()
    clientes_sel = st.multiselect("Cliente", clientes)

    # FECHA CREACIÓN (SAFE)
    fechas_c = df_filtered["Fecha Creación"].dropna()

    if len(fechas_c) > 0:
        min_c = fechas_c.min()
        max_c = fechas_c.max()
    else:
        min_c = datetime.today().date()
        max_c = datetime.today().date()

    fecha_creacion = st.date_input("Fecha creación", value=(min_c, max_c))

    # FECHA ENTREGA (SAFE FIX REAL)
    fechas_e = df_filtered["Fecha de entrega"].dropna()

    if len(fechas_e) > 0:
        min_e = fechas_e.min()
        max_e = fechas_e.max()
    else:
        min_e = datetime.today().date()
        max_e = datetime.today().date()

    fecha_entrega = st.date_input("Fecha entrega", value=(min_e, max_e))

    incluir_null = st.checkbox("Incluir sin fecha de entrega", True)

    # =================================================
    # FILTRO CLIENTE
    # =================================================

    if clientes_sel:
        df_filtered = df_filtered[df_filtered["Cliente"].isin(clientes_sel)]

    # =================================================
    # FILTRO FECHA CREACIÓN
    # =================================================

    if isinstance(fecha_creacion, tuple):
        start, end = fecha_creacion

        df_filtered = df_filtered[
            (df_filtered["Fecha Creación"] >= start) &
            (df_filtered["Fecha Creación"] <= end)
        ]

    # =================================================
    # 🔥 FIX CRÍTICO FECHA ENTREGA
    # =================================================

    if isinstance(fecha_entrega, tuple):
        start, end = fecha_entrega

        entrega = df_filtered["Fecha de entrega"]

        mask_rango = (
            entrega.notna() &
            (entrega >= start) &
            (entrega <= end)
        )

        if incluir_null:
            df_filtered = df_filtered[
                entrega.isna() | mask_rango
            ]
        else:
            df_filtered = df_filtered[mask_rango]

    # =================================================
    # OUTPUT
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

    df_filtered = df_filtered[cols]

    st.metric("Resultados", len(df_filtered))
    st.dataframe(df_filtered, use_container_width=True)

    st.download_button(
        "Descargar agenda filtrada",
        to_excel(df_filtered),
        file_name=f"Agenda_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.warning("Sin datos")
