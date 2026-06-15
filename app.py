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
ws = client.open_by_key(SHEET_ID).worksheet("Agenda Final")


# =====================================================
# HELPERS
# =====================================================

def clean_columns(df):
    df = df.copy()
    df.columns = df.columns.astype(str).str.strip()
    return df


# =====================================================
# 🔥 FIX DEFINITIVO PO NUMBER
# =====================================================

def clean_order(x):
    if pd.isna(x):
        return ""

    x = str(x)

    # eliminar espacios invisibles
    x = x.replace("\u00a0", "").replace(" ", "")
    x = x.replace("\n", "").replace("\t", "")

    # eliminar formato Excel float
    if x.endswith(".0"):
        x = x[:-2]

    # dejar solo números
    x = re.sub(r"[^0-9]", "", x)

    # quitar ceros izquierda
    x = x.lstrip("0")

    return x


def get_hour(x):
    if pd.isna(x):
        return ""
    r = re.search(r"\d{1,2}:\d{2}", str(x))
    return r.group(0) if r else ""


def clean_number(x):
    if pd.isna(x):
        return 0

    x = str(x).strip()

    if "." in x and "," in x:
        x = x.replace(".", "")
        x = x.replace(",", ".")

    elif "," in x:
        x = x.replace(",", "")

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
# GENERAR BASE
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
        "Order Number",
        "Delivery Number",
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
    st.write(
    "Duplicados por Order Number:",
    track["Order Number"].duplicated().sum()
    )

    st.dataframe(
    track[track["Order Number"].duplicated(keep=False)]
    )

    st.write("Filas Track:", len(track))

    st.write(
        "Order Number únicos:",
        track["Order Number"].nunique()
    )
    
    st.write(
        "Duplicados detectados:",
        len(track) - track["Order Number"].nunique()
    )

    # 🔥 NORMALIZACIÓN PO DEFINITIVA
    track["Num Order"] = track["Num Order"].apply(clean_order)
    track["Delivery Number"] = pd.to_numeric(
    track["Delivery Number"],
    errors="coerce"
    )

    track["Unidades"] = track["Unidades"].apply(clean_number)
    track["Monto"] = track["Monto"].apply(clean_number)

    # =================================================
    # ELIMINAR DUPLICADOS DE ORDER NUMBER
    # CONSERVANDO EL MENOR DELIVERY NUMBER
    # =================================================

    track = track.sort_values(
        ["Order Number", "Delivery Number"],
        ascending=[True, True]
    )

    track = track.drop_duplicates(
        subset=["Order Number"],
        keep="first"
    )
    
    st.write(
        "Registros después de eliminar duplicados:",
        len(track)
    )

    # =================================================
    # GROUPBY CORRECTO
    # =================================================

    track = track.groupby(
        "Num Order",
        as_index=False,
        dropna=False
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
    # MERGE FINAL
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
