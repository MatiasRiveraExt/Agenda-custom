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
    page_title="Agenda PRO",
    layout="wide"
)

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
# FUNCIONES
# =====================================================

def clean_columns(df):
    df = df.copy()
    df.columns = df.columns.astype(str).str.strip()
    return df



def clean_order(x):

    if pd.isna(x):
        return ""

    x = str(x)
    x = x.strip()
    x = x.replace(".0", "")
    x = re.sub(r"\s+", "", x)

    return x.lstrip("0")



def get_hour(x):

    if pd.isna(x):
        return ""

    result = re.search(
        r"\d{1,2}:\d{2}",
        str(x)
    )

    return result.group(0) if result else ""



def money(x):

    try:
        x = float(
            str(x)
            .replace(".","")
            .replace(",","")
        )

        return f"{int(x):,}".replace(",", ".")

    except:
        return x



def read_database():

    data = ws.get_all_records()

    if not data:
        return pd.DataFrame()

    return pd.DataFrame(data)



def sheet_values(df):

    result = []

    result.append(
        [str(c) for c in df.columns]
    )

    for row in df.itertuples(index=False):

        new_row=[]

        for value in row:

            if pd.isna(value):
                new_row.append("")
            else:
                new_row.append(str(value))

        result.append(new_row)


    return result



def excel_download(df):

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



# =====================================================
# CARGA
# =====================================================

st.subheader("📤 Actualizar base")

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



if st.button("🚀 Actualizar"):


    if not all(
        [
            ordenes_file,
            track_file,
            maestro_file
        ]
    ):

        st.error(
            "Faltan archivos"
        )

        st.stop()



    ordenes = clean_columns(
        pd.read_excel(ordenes_file)
    )

    track = clean_columns(
        pd.read_excel(track_file)
    )

    maestro = clean_columns(
        pd.read_excel(maestro_file)
    )


    # ------------------------------
    # TRACK
    # ------------------------------

    track = track[
        [
            "Created On",
            "PO Number",
            "Sold to Name",
            "Delivered Quantity",
            "Delivered Amount"
        ]
    ]


    track = track.rename(
        columns={
            "Created On":"Fecha Creación",
            "PO Number":"Num Order",
            "Sold to Name":"Cliente",
            "Delivered Quantity":"Unidades",
            "Delivered Amount":"Monto"
        }
    )


    # ------------------------------
    # MAESTRO
    # ------------------------------

    maestro = maestro.reindex(
        columns=[
            "Num Order",
            "Departamento",
            "PD"
        ]
    )


    # ------------------------------
    # ORDENES
    # ------------------------------

    ordenes = ordenes.reindex(
        columns=[
            "O/C Cliente",
            "Fecha Entrega",
            "Instrucciones"
        ]
    )


    ordenes = ordenes.rename(
        columns={
            "O/C Cliente":"Num Order",
            "Fecha Entrega":"Fecha de entrega"
        }
    )


    ordenes["Hora"] = (
        ordenes["Instrucciones"]
        .apply(get_hour)
    )


    # ------------------------------
    # LIMPIAR KEYS
    # ------------------------------

    for x in [
        track,
        maestro,
        ordenes
    ]:

        x["Num Order"] = (
            x["Num Order"]
            .apply(clean_order)
        )



    # ------------------------------
    # MERGES
    # ------------------------------

    nuevo = track.merge(
        maestro,
        on="Num Order",
        how="left"
    )


    nuevo = nuevo.merge(
        ordenes,
        on="Num Order",
        how="left"
    )


    nuevo["Monto"] = (
        nuevo["Monto"]
        .apply(money)
    )


    nuevo["Fecha Creación"] = pd.to_datetime(
        nuevo["Fecha Creación"],
        errors="coerce"
    )


    nuevo["Fecha de entrega"] = pd.to_datetime(
        nuevo["Fecha de entrega"],
        errors="coerce"
    )



    # =================================================
    # UPSERT HISTORICO
    # =================================================

    viejo = read_database()


    if not viejo.empty:

        viejo = clean_columns(viejo)

        viejo["Num Order"] = (
            viejo["Num Order"]
            .apply(clean_order)
        )


        viejo["Fecha Creación"] = pd.to_datetime(
            viejo["Fecha Creación"],
            errors="coerce"
        )


        final = pd.concat(
            [
                viejo,
                nuevo
            ],
            ignore_index=True
        )

    else:

        final = nuevo



    final["Num Order"] = (
        final["Num Order"]
        .apply(clean_order)
    )


    final = final.sort_values(
        "Fecha Creación"
    )


    final = final.drop_duplicates(
        subset=["Num Order"],
        keep="last"
    )


    final = final.reset_index(drop=True)



    # guardar

    ws.clear()

    ws.update(
        sheet_values(final)
    )


    st.success(
        f"Base actualizada: {len(final)} registros"
    )



# =====================================================
# AGENDA
# =====================================================

st.divider()

st.subheader("📋 Agenda")


df = read_database()


if not df.empty:


    df = clean_columns(df)


    filtro = st.multiselect(
        "Cliente",
        df["Cliente"].dropna().unique()
    )


    if filtro:

        df = df[
            df["Cliente"].isin(filtro)
        ]



    fechas = pd.to_datetime(
        df["Fecha Creación"],
        errors="coerce"
    ).dropna()


    if not fechas.empty:


        rango = st.date_input(
            "Fecha creación",
            (
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
            df["Fecha Creación"].between(
                inicio,
                fin
            )
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
        "⬇ Descargar agenda",
        excel_download(df),
        file_name=f"Agenda_{datetime.now().strftime('%Y%m%d')}.xlsx"
    )

else:

    st.warning(
        "Base vacía"
    )
