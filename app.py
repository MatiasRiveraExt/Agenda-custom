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
        value = float(str(value).replace(".", "").replace(",", ""))
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
        df.to_excel(
            writer,
            index=False,
            sheet_name="Agenda"
        )

    return output.getvalue()

# =====================================================
# CARGAR BASE HISTÓRICA
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

    # =================================================
    # SI NO EXISTE BASE
    # =================================================

    if len(existing) == 0:

        final_df = df_new.copy()

    else:

        existing_df = pd.DataFrame(existing)

        existing_df = clean_dataframe(existing_df)

        # =============================================
        # NORMALIZAR
        # =============================================

        existing_df["Num Order"] = existing_df["Num Order"].apply(
            normalize_order
        )

        df_new["Num Order"] = df_new["Num Order"].apply(
            normalize_order
        )

        # =============================================
        # CONCAT
        # =============================================

        final_df = pd.concat(
            [existing_df, df_new],
            ignore_index=True
        )

        # =============================================
        # ELIMINAR DUPLICADOS
        # =============================================

        final_df = final_df.drop_duplicates(
            subset="Num Order",
            keep="last"
        )

    # =================================================
    # LIMPIEZA FINAL
    # =================================================

    final_df = final_df.fillna("").astype(str)

    # =================================================
    # WRITE
    # =================================================

    ws.clear()

    values = [final_df.columns.tolist()]

    for row in final_df.values:
        values.append([str(x) for x in row])

    ws.update(values)

# =====================================================
# UI
# =====================================================

st.subheader("📤 Actualizar Base Histórica")

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
# GENERAR
# =====================================================

if st.button("🚀 Actualizar y Generar Agenda"):

    if not ordenes_file or not track_file or not maestro_file:

        st.error("❌ Debes subir los 3 archivos")

        st.stop()

    # =================================================
    # LEER ARCHIVOS
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
        "Created On",
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
    # EXTRAER HORA
    # =================================================

    ordenes["Hora"] = ordenes["Instrucciones"].apply(
        extract_hour
    )

    # =================================================
    # TRACK
    # =================================================

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
    # NORMALIZAR
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

    df["Fecha Creación"] = df["Fecha Creación"].apply(
        format_date
    )

    # =================================================
    # COLUMNAS FINALES
    # =================================================

    df = df[[
        "Fecha Creación",
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

    # =================================================
    # ACTUALIZAR BASE HISTÓRICA
    # =================================================

    upsert_to_sheet(df, "Agenda Final")

    st.success("☁️ Base histórica actualizada")

    # =================================================
    # RECARGAR TODA LA BASE
    # =================================================

    df = load_latest_from_sheet()

    st.success("📦 Agenda reconstruida desde TODA la base histórica")

    # =================================================
    # TIMESTAMP
    # =================================================

    ws = sheet.worksheet("Agenda Final")

    timestamp = datetime.now().strftime(
        "%d-%m-%Y %H:%M:%S"
    )

    ws.update(
        "K1",
        [[f"Última actualización: {timestamp}"]]
    )

    # =================================================
    # MÉTRICAS
    # =================================================

    st.metric(
        "📊 OC históricas",
        len(df)
    )

    # =================================================
    # OUTPUT
    # =================================================

    st.dataframe(
        df,
        use_container_width=True
    )

    # =================================================
    # DESCARGA
    # =================================================

    st.download_button(
        "📥 Descargar agenda histórica",
        to_excel(df),
        file_name=f"Agenda_Historica_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# =====================================================
# DESCARGA DIRECTA DESDE BD
# =====================================================

st.divider()

st.subheader("📦 Última agenda desde Base Histórica")

latest_df = load_latest_from_sheet()

if latest_df.empty:

    st.warning("No hay datos aún")

else:

    st.success("📊 Base histórica cargada")

    st.metric(
        "📈 Total OC históricas",
        len(latest_df)
    )

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

    status = ws.acell("K1").value

    if status:
        st.info(status)

except:
    pass
