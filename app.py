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

st.title("📊 Agenda Automática PRO - Base de Datos")

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
# FUNCIONES
# =====================================================

def clean_dataframe(df):

    df = df.copy()

    df.columns = [str(c).strip() for c in df.columns]

    df = df.dropna(axis=1, how="all")

    df = df.loc[:, ~pd.Index(df.columns).duplicated(keep="first")]

    return df

# =====================================================

def extract_hour(text):

    if pd.isna(text):
        return ""

    text = str(text)

    match = re.search(r'(\d{1,2}:\d{2})', text)

    if match:
        return match.group(1)

    return ""

# =====================================================

def normalize_order(x):

    if pd.isna(x):
        return ""

    x = str(x).strip()

    x = x.replace("\u00a0", "")

    x = re.sub(r"\.0$", "", x)

    x = x.lstrip("0")

    return x

# =====================================================

def format_chilean_money(value):

    try:
        value = float(value)
        return f"{int(value):,}".replace(",", ".")
    except:
        return value

# =====================================================

def format_date(value):

    try:
        return pd.to_datetime(value).strftime("%d-%m-%Y")
    except:
        return value

# =====================================================

def to_excel(df):

    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Agenda")

    return output.getvalue()

# =====================================================
# UPSERT
# =====================================================

def upsert_to_sheet(df, sheet_name):

    ws = sheet.worksheet(sheet_name)

    existing = ws.get_all_records()

    # =================================================
    # SI NO HAY DATA PREVIA
    # =================================================

    if len(existing) == 0:

        ws.clear()

        df = df.fillna("").astype(str)

        values = [df.columns.tolist()]

        for row in df.values:
            values.append([str(x) for x in row])

        ws.update(values)

        return

    # =================================================
    # EXISTING DF
    # =================================================

    existing_df = pd.DataFrame(existing)

    if "Num Order" not in existing_df.columns:

        ws.clear()

        existing_df = pd.DataFrame()

    # =================================================
    # NORMALIZAR KEYS
    # =================================================

    df["Num Order"] = df["Num Order"].astype(str)
    existing_df["Num Order"] = existing_df["Num Order"].astype(str)

    # =================================================
    # 🔥 ELIMINAR DUPLICADOS
    # =================================================

    existing_df = existing_df.drop_duplicates(
        subset="Num Order",
        keep="last"
    )

    df = df.drop_duplicates(
        subset="Num Order",
        keep="last"
    )

    # =================================================
    # DEBUG DUPLICADOS
    # =================================================

    duplicates = df[df.duplicated("Num Order", keep=False)]

    if not duplicates.empty:

        st.warning("⚠️ Se encontraron órdenes duplicadas")

        st.dataframe(duplicates)

    # =================================================
    # INDEX
    # =================================================

    existing_df = existing_df.set_index("Num Order")
    df = df.set_index("Num Order")

    # =================================================
    # UPSERT
    # =================================================

    combined = df.combine_first(existing_df)

    combined.update(df)

    combined = combined.reset_index()

    # =================================================
    # WRITE BACK
    # =================================================

    ws.clear()

    combined = combined.fillna("").astype(str)

    values = [combined.columns.tolist()]

    for row in combined.values:
        values.append([str(x) for x in row])

    ws.update(values)

# =====================================================
# LOAD FROM SHEETS
# =====================================================

def load_latest_from_sheet():

    ws = sheet.worksheet("Agenda Final")

    data = ws.get_all_records()

    if not data:
        return pd.DataFrame()

    return pd.DataFrame(data)

# =====================================================
# UI
# =====================================================

st.subheader("📤 Generar nueva agenda")

ordenes_file = st.file_uploader(
    "Archivo Ordenes",
    type=["xlsx"]
)

track_file = st.file_uploader(
    "Archivo Track",
    type=["xlsx"]
)

maestro_file = st.file_uploader(
    "Archivo Maestro",
    type=["xlsx"]
)

# =====================================================
# GENERAR AGENDA
# =====================================================

if st.button("🚀 Generar Agenda"):

    if not ordenes_file or not track_file or not maestro_file:

        st.error("❌ Debes subir los 3 archivos")

        st.stop()

    # =================================================
    # READ FILES
    # =================================================

    ordenes = clean_dataframe(
        pd.read_excel(ordenes_file)
    )

    track = clean_dataframe(
        pd.read_excel(track_file)
    )

    maestro = clean_dataframe(
        pd.read_excel(maestro_file)
    )

    st.success("📥 Archivos cargados correctamente")

    # =================================================
    # VALIDACIONES
    # =================================================

    required_ordenes = [
        "O/C Cliente",
        "Fecha Entrega",
        "Instrucciones"
    ]

    required_track = [
        "PO Number",
        "Sold to Name",
        "Delivered Quantity",
        "Delivered Amount"
    ]

    required_maestro = [
        "Num Order",
        "Departamento",
        "PD"
    ]

    for col in required_ordenes:

        if col not in ordenes.columns:

            st.error(f"❌ Falta '{col}' en Ordenes")

            st.stop()

    for col in required_track:

        if col not in track.columns:

            st.error(f"❌ Falta '{col}' en Track")

            st.stop()

    for col in required_maestro:

        if col not in maestro.columns:

            st.error(f"❌ Falta '{col}' en Maestro")

            st.stop()

    # =================================================
    # HORA
    # =================================================

    ordenes["Hora"] = ordenes["Instrucciones"].apply(
        extract_hour
    )

    # =================================================
    # TRACK FINAL
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
    # MAESTRO FINAL
    # =================================================

    maestro_final = maestro[[
        "Num Order",
        "Departamento",
        "PD"
    ]].copy()

    # =================================================
    # ORDENES FINAL
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

    track_final["Num Order"] = track_final["Num Order"].apply(
        normalize_order
    )

    maestro_final["Num Order"] = maestro_final["Num Order"].apply(
        normalize_order
    )

    ordenes_final["Num Order"] = ordenes_final["Num Order"].apply(
        normalize_order
    )

    # =================================================
    # ELIMINAR DUPLICADOS
    # =================================================

    track_final = track_final.drop_duplicates(
        subset="Num Order",
        keep="last"
    )

    maestro_final = maestro_final.drop_duplicates(
        subset="Num Order",
        keep="last"
    )

    ordenes_final = ordenes_final.drop_duplicates(
        subset="Num Order",
        keep="last"
    )

    # =================================================
    # MERGE
    # =================================================

    df = track_final.merge(
        maestro_final,
        on="Num Order",
        how="left"
    )

    df = df.merge(
        ordenes_final,
        on="Num Order",
        how="left"
    )

    # =================================================
    # FORMATO
    # =================================================

    df["Monto"] = df["Monto"].apply(
        format_chilean_money
    )

    df["Fecha de entrega"] = df["Fecha de entrega"].apply(
        format_date
    )

    # =================================================
    # COLUMNAS FINALES
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
    # UPSERT TO SHEETS
    # =================================================

    upsert_to_sheet(df, "Agenda Final")

    st.success("☁️ Base de datos actualizada correctamente")

    # =================================================
    # TIMESTAMP
    # =================================================

    ws = sheet.worksheet("Agenda Final")

    timestamp = datetime.now().strftime(
        "%Y-%m-%d %H:%M:%S"
    )

    ws.update(
        "J1",
        [[f"Última actualización: {timestamp}"]]
    )

    # =================================================
    # OUTPUT
    # =================================================

    st.success("🎯 Agenda generada correctamente")

    st.dataframe(
        df,
        use_container_width=True
    )

    # =================================================
    # DOWNLOAD GENERATED
    # =================================================

    st.download_button(
        "📥 Descargar Excel generado",
        to_excel(df),
        file_name=f"Agenda_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# =====================================================
# DOWNLOAD FROM DATABASE
# =====================================================

st.divider()

st.subheader("📦 Descargar última agenda desde base de datos")

latest_df = load_latest_from_sheet()

if latest_df.empty:

    st.warning("No hay datos aún")

else:

    st.success("📊 Última agenda cargada desde Google Sheets")

    st.dataframe(
        latest_df,
        use_container_width=True
    )

    st.download_button(
        "📥 Descargar última agenda",
        to_excel(latest_df),
        file_name=f"Agenda_DB_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# =====================================================
# STATUS
# =====================================================

try:

    ws = sheet.worksheet("Agenda Final")

    status = ws.acell("J1").value

    if status:
        st.info(status)

except:
    pass
