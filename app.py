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
    """SIEMPRE retorna date (no datetime)"""
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
# UPSERT HISTÓRICO
# =====================================================

def upsert_to_sheet(df_new, sheet_name):

    ws = sheet.worksheet(sheet_name)
    existing = ws.get_all_records()

    if len(existing) == 0:
        final_df = df_new.copy()
    else:
        existing_df = pd.DataFrame(existing)
        existing_df = clean_dataframe(existing_df)

        existing_df["Num Order"] = existing_df["Num Order"].apply(normalize_order)
        df_new["Num Order"] = df_new["Num Order"].apply(normalize_order)

        final_df = pd.concat([existing_df, df_new], ignore_index=True)

        final_df = final_df.drop_duplicates(subset="Num Order", keep="last")

    final_df = final_df.fillna("").astype(str)

    desired_order = [
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

    final_df = final_df[[c for c in desired_order if c in final_df.columns]]

    ws.clear()

    values = [final_df.columns.tolist()] + final_df.values.tolist()
    ws.update(values)

# =====================================================
# UI INPUT
# =====================================================

st.subheader("📤 Actualizar Base Histórica")

ordenes_file = st.file_uploader("Archivo Ordenes", type=["xlsx"])
track_file = st.file_uploader("Archivo Track", type=["xlsx"])
maestro_file = st.file_uploader("Archivo Maestro", type=["xlsx"])

# =====================================================
# GENERAR
# =====================================================

if st.button("🚀 Actualizar y Generar Agenda"):

    if not ordenes_file or not track_file or not maestro_file:
        st.error("❌ Debes subir los 3 archivos")
        st.stop()

    ordenes = clean_dataframe(pd.read_excel(ordenes_file))
    track = clean_dataframe(pd.read_excel(track_file))
    maestro = clean_dataframe(pd.read_excel(maestro_file))

    st.success("📥 Archivos cargados correctamente")

    # =================================================
    # VALIDACIÓN
    # =================================================

    required_track = [
        "Created On",
        "PO Number",
        "Sold to Name",
        "Delivered Quantity",
        "Delivered Amount"
    ]

    for col in required_track:
        if col not in track.columns:
            st.error(f"❌ Falta '{col}' en Track")
            st.stop()

    # =================================================
    # TRANSFORMACIÓN
    # =================================================

    ordenes["Hora"] = ordenes["Instrucciones"].apply(extract_hour)

    track_final = track[[
        "Created On",
        "PO Number",
        "Sold to Name",
        "Delivered Quantity",
        "Delivered Amount"
    ]].copy()

    track_final = track_final.rename(columns={
        "Created On": "Fecha Creación",
        "PO Number": "Num Order",
        "Sold to Name": "Cliente",
        "Delivered Quantity": "Suma de Unidades",
        "Delivered Amount": "Monto"
    })

    maestro_final = maestro[["Num Order", "Departamento", "PD"]].copy()

    ordenes_final = ordenes[["O/C Cliente", "Fecha Entrega", "Hora"]].copy()

    ordenes_final = ordenes_final.rename(columns={
        "O/C Cliente": "Num Order",
        "Fecha Entrega": "Fecha de entrega"
    })

    # normalizar OC
    track_final["Num Order"] = track_final["Num Order"].apply(normalize_order)
    maestro_final["Num Order"] = maestro_final["Num Order"].apply(normalize_order)
    ordenes_final["Num Order"] = ordenes_final["Num Order"].apply(normalize_order)

    # merge
    df = track_final.merge(maestro_final, on="Num Order", how="left")
    df = df.merge(ordenes_final, on="Num Order", how="left")

    df["Monto"] = df["Monto"].apply(format_chilean_money)

    # fechas como DATE real (clave del fix)
    df["Fecha Creación"] = df["Fecha Creación"].apply(parse_date)
    df["Fecha de entrega"] = df["Fecha de entrega"].apply(parse_date)

    df = df.fillna("")

    upsert_to_sheet(df, "Agenda Final")

    st.success("☁️ Base histórica actualizada")

# =====================================================
# LOAD DB
# =====================================================

st.divider()
st.subheader("📦 Base histórica")

latest_df = load_latest_from_sheet()

if not latest_df.empty:

    # FIX CRÍTICO: fechas consistentes tipo DATE
    latest_df["Fecha Creación"] = pd.to_datetime(latest_df["Fecha Creación"], errors="coerce").dt.date
    latest_df["Fecha de entrega"] = pd.to_datetime(latest_df["Fecha de entrega"], errors="coerce").dt.date

    st.markdown("## 🎯 Filtros de descarga")

    # CLIENTE
    clientes = sorted(latest_df["Cliente"].dropna().unique().tolist())
    clientes_sel = st.multiselect("👤 Cliente", clientes)

    # FECHAS
    min_c = latest_df["Fecha Creación"].min()
    max_c = latest_df["Fecha Creación"].max()

    fecha_creacion = st.date_input(
        "📅 Fecha creación",
        value=(min_c, max_c)
    )

    min_e = latest_df["Fecha de entrega"].min()
    max_e = latest_df["Fecha de entrega"].max()

    fecha_entrega = st.date_input(
        "🚚 Fecha entrega",
        value=(min_e, max_e)
    )

    incluir_null = st.checkbox("📌 Incluir sin fecha de entrega", True)

    df_filtered = latest_df.copy()

    # filtro cliente
    if clientes_sel:
        df_filtered = df_filtered[df_filtered["Cliente"].isin(clientes_sel)]

    # filtro creación
    if isinstance(fecha_creacion, tuple):
        start, end = fecha_creacion
        df_filtered = df_filtered[
            (df_filtered["Fecha Creación"] >= start) &
            (df_filtered["Fecha Creación"] <= end)
        ]

    # filtro entrega
    if isinstance(fecha_entrega, tuple):
        start, end = fecha_entrega

        if incluir_null:
            df_filtered = df_filtered[
                df_filtered["Fecha de entrega"].isna() |
                (
                    (df_filtered["Fecha de entrega"] >= start) &
                    (df_filtered["Fecha de entrega"] <= end)
                )
            ]
        else:
            df_filtered = df_filtered[
                (df_filtered["Fecha de entrega"] >= start) &
                (df_filtered["Fecha de entrega"] <= end)
            ]

    # orden final
    df_filtered = df_filtered[
        [
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
    ]

    st.metric("📊 Resultados", len(df_filtered))

    st.dataframe(df_filtered, use_container_width=True)

    st.download_button(
        "📥 Descargar agenda filtrada",
        to_excel(df_filtered),
        file_name=f"Agenda_FILTRADA_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.warning("No hay datos aún")
