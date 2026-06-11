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

st.set_page_config(page_title="Agenda PRO", layout="wide")
st.title("📦 Agenda Automática PRO")


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
spreadsheet = client.open_by_key(SHEET_ID)
ws = spreadsheet.worksheet("Agenda Final")


# =====================================================
# HELPERS
# =====================================================

def clean_columns(df):
    df = df.copy()
    df.columns = df.columns.astype(str).str.strip()
    return df


def clean_order(x):
    if pd.isna(x):
        return ""
    x = str(x).strip()
    x = x.replace(".0", "")
    x = re.sub(r"\s+", "", x)
    return x.lstrip("0")


def get_hour(x):
    if pd.isna(x):
        return ""
    r = re.search(r"\d{1,2}:\d{2}", str(x))
    return r.group(0) if r else ""


# =====================================================
# 🔥 FIX REAL NUMÉRICO (CLAVE)
# =====================================================

def clean_number(x):
    if pd.isna(x):
        return 0

    x = str(x).strip()

    # caso 1: 1.234,56 → 1234.56
    if "." in x and "," in x:
        x = x.replace(".", "")
        x = x.replace(",", ".")

    # caso 2: 1,234 → 1234
    elif "," in x:
        x = x.replace(",", "")

    # caso 3: 1234.56 OK
    try:
        return float(x)
    except:
        return 0


def money(x):
    try:
        return f"{int(x):,}".replace(",", ".")
    except:
        return x


def read_database():
    data = ws.get_all_records()
    if not data:
        return pd.DataFrame()
    return pd.DataFrame(data)


def sheet_values(df):
    values = []
    values.append([str(c) for c in df.columns])

    for row in df.itertuples(index=False):
        values.append(["" if pd.isna(v) else str(v) for v in row])

    return values


def excel_download(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()


# =====================================================
# UI
# =====================================================

st.subheader("📤 Actualizar base")

ordenes_file = st.file_uploader("Ordenes", type=["xlsx"])
track_file = st.file_uploader("Track", type=["xlsx"])
maestro_file = st.file_uploader("Maestro", type=["xlsx"])


# =====================================================
# GENERAR
# =====================================================

if st.button("🚀 Actualizar"):

    if not all([ordenes_file, track_file, maestro_file]):
        st.error("Faltan archivos")
        st.stop()

    ordenes = clean_columns(pd.read_excel(ordenes_file))
    track = clean_columns(pd.read_excel(track_file))
    maestro = clean_columns(pd.read_excel(maestro_file))

    # =================================================
    # TRACK
    # =================================================

    track = track[[
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

    track["Num Order"] = track["Num Order"].apply(clean_order)

    # 🔥 FIX NUMÉRICO REAL
    track["Unidades"] = track["Unidades"].apply(clean_number)
    track["Monto"] = track["Monto"].apply(clean_number)

    # =================================================
    # 🔥 AGRUPACIÓN CORRECTA
    # =================================================

    track = track.groupby(
        "Num Order",
        as_index=False
    ).agg({
        "Fecha Creación": "max",
        "Cliente": "first",
        "Unidades": "sum",
        "Monto": "sum"
    })

    # =================================================
    # MAESTRO
    # =================================================

    maestro = maestro.reindex(columns=[
        "Num Order",
        "Departamento",
        "PD"
    ])

    maestro["Num Order"] = maestro["Num Order"].apply(clean_order)

    # =================================================
    # ORDENES
    # =================================================

    ordenes = ordenes.reindex(columns=[
        "O/C Cliente",
        "Fecha Entrega",
        "Instrucciones"
    ])

    ordenes = ordenes.rename(columns={
        "O/C Cliente": "Num Order",
        "Fecha Entrega": "Fecha de entrega"
    })

    ordenes["Hora"] = ordenes["Instrucciones"].apply(get_hour)
    ordenes["Num Order"] = ordenes["Num Order"].apply(clean_order)

    # =================================================
    # MERGE
    # =================================================

    nuevo = track.merge(maestro, on="Num Order", how="left")
    nuevo = nuevo.merge(ordenes, on="Num Order", how="left")

    nuevo["Fecha Creación"] = pd.to_datetime(nuevo["Fecha Creación"], errors="coerce")
    nuevo["Fecha de entrega"] = pd.to_datetime(nuevo["Fecha de entrega"], errors="coerce")

    nuevo["Monto"] = nuevo["Monto"].apply(money)

    # =================================================
    # HISTORICO
    # =================================================

    viejo = read_database()

    if not viejo.empty:
        viejo["Num Order"] = viejo["Num Order"].apply(clean_order)
        final = pd.concat([viejo, nuevo], ignore_index=True)
    else:
        final = nuevo

    final["Num Order"] = final["Num Order"].apply(clean_order)
    final["Fecha Creación"] = pd.to_datetime(final["Fecha Creación"], errors="coerce")

    final = final.sort_values(by="Fecha Creación", na_position="first")

    final = final.drop_duplicates(subset=["Num Order"], keep="last")

    final = final.reset_index(drop=True)

    # =================================================
    # GUARDAR
    # =================================================

    ws.clear()
    ws.update(sheet_values(final))

    st.success(f"Base actualizada: {len(final)} OC")


# =====================================================
# AGENDA
# =====================================================

st.divider()
st.subheader("📋 Agenda")

df = read_database()

if not df.empty:

    df = clean_columns(df)

    clientes = df["Cliente"].dropna().unique().tolist()
    filtro = st.multiselect("Cliente", clientes)

    if filtro:
        df = df[df["Cliente"].isin(filtro)]

    df["Fecha Creación"] = pd.to_datetime(df["Fecha Creación"], errors="coerce")
    fechas = df["Fecha Creación"].dropna()

    if not fechas.empty:
        rango = st.date_input(
            "Fecha creación",
            (fechas.min().date(), fechas.max().date())
        )

        df = df[df["Fecha Creación"].between(
            pd.to_datetime(rango[0]),
            pd.to_datetime(rango[1])
        )]

    columnas = [
        "Num Order",
        "Fecha Creación",
        "Cliente",
        "Departamento",
        "PD",
        "Unidades",
        "Fecha de entrega",
        "Hora",
        "Monto"
    ]

    df = df.reindex(columns=columnas)

    st.metric("Resultados", len(df))
    st.dataframe(df, use_container_width=True)

    st.download_button(
        "⬇ Descargar agenda",
        excel_download(df),
        file_name=f"Agenda_{datetime.now().strftime('%Y%m%d')}.xlsx"
    )

else:
    st.warning("Sin datos")
