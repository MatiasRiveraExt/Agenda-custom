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

st.title("📊 Agenda Automática PRO")


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

ws = sheet.worksheet("Agenda Final")


st.success("✅ Conectado a Google Sheets")


# =====================================================
# HELPERS
# =====================================================

def normalize_columns(df):
    df = df.copy()
    df.columns = (
        df.columns
        .astype(str)
        .str.strip()
    )
    return df



def normalize_order(x):

    if pd.isna(x):
        return ""

    x = str(x)
    x = x.strip()
    x = x.replace(".0","")

    # elimina espacios invisibles
    x = re.sub(r"\s+", "", x)

    # elimina ceros iniciales
    x = x.lstrip("0")

    return x



def extract_hour(text):

    if pd.isna(text):
        return ""

    m = re.search(
        r"(\d{1,2}:\d{2})",
        str(text)
    )

    return m.group(1) if m else ""



def format_money(v):

    try:

        v = float(
            str(v)
            .replace(".","")
            .replace(",","")
        )

        return f"{int(v):,}".replace(",", ".")

    except:
        return v



def safe_gsheets_values(df):

    values = []

    values.append(
        [str(c) for c in df.columns]
    )


    for row in df.itertuples(index=False):

        clean=[]

        for v in row:

            if pd.isna(v):
                clean.append("")
            else:
                clean.append(str(v))

        values.append(clean)


    return values



def to_excel(df):

    output = BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

        df.to_excel(
            writer,
            index=False
        )


    return output.getvalue()



def load_database():

    data = ws.get_all_records()

    if len(data)==0:
        return pd.DataFrame()

    return pd.DataFrame(data)



# =====================================================
# ARCHIVOS
# =====================================================

st.subheader("📤 Nueva carga")


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
# GENERAR NUEVA CARGA
# =====================================================


if st.button("🚀 Actualizar base"):


    if not ordenes_file or not track_file or not maestro_file:

        st.error(
            "Faltan archivos"
        )

        st.stop()



    ordenes = normalize_columns(
        pd.read_excel(ordenes_file)
    )

    track = normalize_columns(
        pd.read_excel(track_file)
    )

    maestro = normalize_columns(
        pd.read_excel(maestro_file)
    )


    st.success(
        "📥 Archivos cargados"
    )


    # -------------------------
    # TRACK
    # -------------------------

    track_final = track[[
        "Created On",
        "PO Number",
        "Sold to Name",
        "Delivered Quantity",
        "Delivered Amount"
    ]].rename(
        columns={
            "Created On":"Fecha Creación",
            "PO Number":"Num Order",
            "Sold to Name":"Cliente",
            "Delivered Quantity":"Unidades",
            "Delivered Amount":"Monto"
        }
    )



    maestro_final = maestro.reindex(
        columns=[
            "Num Order",
            "Departamento",
            "PD"
        ]
    )



    ordenes_final = ordenes.reindex(
        columns=[
            "O/C Cliente",
            "Fecha Entrega",
            "Instrucciones"
        ]
    )



    ordenes_final = ordenes_final.rename(
        columns={
            "O/C Cliente":"Num Order",
            "Fecha Entrega":"Fecha de entrega"
        }
    )



    ordenes_final["Hora"] = (
        ordenes_final["Instrucciones"]
        .apply(extract_hour)
    )


    # -------------------------
    # NORMALIZAR
    # -------------------------

    for d in [
        track_final,
        maestro_final,
        ordenes_final
    ]:

        d["Num Order"] = (
            d["Num Order"]
            .apply(normalize_order)
        )



    # -------------------------
    # MERGE
    # -------------------------

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


    df["Monto"] = (
        df["Monto"]
        .apply(format_money)
    )


    df["Fecha Creación"] = pd.to_datetime(
        df["Fecha Creación"],
        errors="coerce"
    )


    df["Fecha de entrega"] = pd.to_datetime(
        df["Fecha de entrega"],
        errors="coerce"
    )



    # =================================================
    # 🔥 UPSERT REAL
    # =================================================


    old = load_database()



    if not old.empty:


        old["Num Order"] = (
            old["Num Order"]
            .apply(normalize_order)
        )


        final_df = pd.concat(
            [
                old,
                df
            ],
            ignore_index=True
        )


    else:

        final_df = df.copy()



    final_df["Num Order"] = (
        final_df["Num Order"]
        .apply(normalize_order)
    )



    # elimina duplicados conservando último
    final_df = final_df.drop_duplicates(
        subset=["Num Order"],
        keep="last"
    )



    # guardar

    ws.clear()

    ws.update(
        safe_gsheets_values(final_df)
    )


    st.success(
        f"☁️ Base actualizada. Registros: {len(final_df)}"
    )



# =====================================================
# AGENDA
# =====================================================

st.divider()

st.subheader("📦 Agenda")


agenda = load_database()



if not agenda.empty:


    df = agenda.copy()


    # -------------------------
    # CLIENTE
    # -------------------------

    clientes = (
        df["Cliente"]
        .dropna()
        .unique()
        .tolist()
    )


    seleccion = st.multiselect(
        "Cliente",
        clientes
    )


    if seleccion:

        df = df[
            df["Cliente"]
            .isin(seleccion)
        ]



    # -------------------------
    # FECHA CREACION
    # -------------------------

    fechas = pd.to_datetime(
        df["Fecha Creación"],
        errors="coerce"
    ).dropna()



    if not fechas.empty:


        rango = st.date_input(
            "Fecha creación",
            value=(
                fechas.min().date(),
                fechas.max().date()
            )
        )


        inicio = pd.to_datetime(rango[0])
        fin = pd.to_datetime(rango[1])


        df["Fecha Creación"] = pd.to_datetime(
            df["Fecha Creación"],
            errors="coerce"
        )


        df = df[
            df["Fecha Creación"]
            .between(inicio, fin)
        ]



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


    df = df.reindex(
        columns=columnas
    )


    st.metric(
        "Resultados",
        len(df)
    )


    st.dataframe(
        df,
        use_container_width=True
    )


    st.download_button(
        "⬇️ Descargar agenda",
        to_excel(df),
        file_name=
        f"Agenda_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime=
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


else:

    st.warning(
        "No hay datos"
    )
