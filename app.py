import streamlit as st
import pandas as pd
import re

st.set_page_config(page_title="Consolidador Excel PRO", layout="wide")
st.title("📊 Consolidador Automático de Excel (PRO)")

st.write("Sube los 3 archivos: SO Track, Pronóstico, Maestro")

# =========================
# UPLOAD
# =========================
track_file = st.file_uploader("SO Track", type=["xlsx"])
plan_file = st.file_uploader("Pronóstico", type=["xlsx"])
maestro_file = st.file_uploader("Maestro", type=["xlsx"])

# =========================
# HELPERS
# =========================
def clean_cols(df):
    df.columns = (
        df.columns.astype(str)
        .str.replace("\n", " ")
        .str.replace("\xa0", " ")
        .str.strip()
    )
    return df

def find_col(df, keywords):
    for col in df.columns:
        c = str(col).lower().replace(" ", "").replace("/", "").replace("_", "")
        for k in keywords:
            k2 = k.lower().replace(" ", "").replace("/", "").replace("_", "")
            if k2 in c:
                return col
    return None

def detect_header_row(df_raw):
    for i in range(len(df_raw)):
        row = df_raw.iloc[i].astype(str).str.lower()
        if any(
            ("cliente" in x or "order" in x or "oc" in x or "po" in x)
            for x in row
        ):
            return i
    return 0

# =========================
# MAIN
# =========================
if st.button("🚀 Generar reporte"):

    if not track_file or not plan_file or not maestro_file:
        st.error("❌ Debes subir los 3 archivos")
        st.stop()

    # =========================
    # READ TRACK
    # =========================
    track = pd.read_excel(track_file)
    track = clean_cols(track)

    # =========================
    # READ MAESTRO
    # =========================
    maestro = pd.read_excel(maestro_file)
    maestro = clean_cols(maestro)

    # =========================
    # READ PLAN (RAW)
    # =========================
    plan_raw = pd.read_excel(plan_file, sheet_name="Pronostico JUNIO", header=None)

    st.write("📌 Vista cruda del pronóstico:")
    st.dataframe(plan_raw.head(10))

    # =========================
    # DETECT HEADER ROW
    # =========================
    header_row = detect_header_row(plan_raw)
    st.write("📌 Header detectado en fila:", header_row)

    # =========================
    # RELOAD PLAN CORRECTLY
    # =========================
    plan = pd.read_excel(plan_file, sheet_name="Pronostico JUNIO", header=header_row)
    plan = clean_cols(plan)

    st.write("📌 Columnas reales del plan:")
    st.write(plan.columns.tolist())

    # =========================
    # DETECT COLUMNS TRACK
    # =========================
    track_order = find_col(track, ["po", "order", "orden"])
    track_qty = find_col(track, ["quantity", "delivered", "qty"])
    track_amt = find_col(track, ["amount", "monto"])

    # =========================
    # DETECT COLUMNS PLAN
    # =========================
    plan_order = find_col(plan, ["occliente", "cliente", "order", "po", "oc"])
    plan_instr = find_col(plan, ["instru"])
    plan_date = find_col(plan, ["fecha"])

    # =========================
    # DETECT COLUMNS MAESTRO
    # =========================
    maestro_order = find_col(maestro, ["numorder", "order", "num"])
    maestro_dep = find_col(maestro, ["depart"])
    maestro_pd = find_col(maestro, ["pd"])

    # =========================
    # VALIDATION
    # =========================
    if not track_order:
        st.error("❌ No se encontró Order en SO Track")
        st.stop()

    if not plan_order:
        st.error("❌ No se encontró Order en Pronóstico")
        st.stop()

    if not maestro_order:
        st.error("❌ No se encontró Order en Maestro")
        st.stop()

    # =========================
    # NORMALIZE TRACK
    # =========================
    track = track.rename(columns={
        track_order: "Order Number",
        track_qty: "Suma de Unidades",
        track_amt: "Monto"
    })

    # =========================
    # NORMALIZE PLAN
    # =========================
    plan = plan.rename(columns={
        plan_order: "Order Number",
        plan_instr: "Instrucciones" if plan_instr else "Instrucciones",
        plan_date: "Fecha de entrega" if plan_date else "Fecha de entrega"
    })

    # =========================
    # NORMALIZE MAESTRO
    # =========================
    maestro = maestro.rename(columns={
        maestro_order: "Order Number",
        maestro_dep: "Departamento",
        maestro_pd: "PD"
    })

    # =========================
    # EXTRACT HOUR
    # =========================
    if "Instrucciones" in plan.columns:
        plan["Hora"] = plan["Instrucciones"].astype(str).str.extract(r'(\d{1,2}:\d{2})')
    else:
        plan["Hora"] = ""

    # =========================
    # MERGE
    # =========================
    df = track.merge(plan, on="Order Number", how="left")
    df = df.merge(maestro, on="Order Number", how="left")

    # =========================
    # FINAL OUTPUT
    # =========================
    final = df[[
        "Order Number",
        "Departamento",
        "PD",
        "Suma de Unidades",
        "Fecha de entrega",
        "Hora",
        "Monto"
    ]].drop_duplicates()

    final = final.fillna("")

    st.success("✅ Reporte generado correctamente")

    st.dataframe(final.head(50))

    # =========================
    # DOWNLOAD
    # =========================
    csv = final.to_csv(index=False).encode("utf-8")

    st.download_button(
        "⬇ Descargar reporte",
        csv,
        "reporte_final.csv",
        "text/csv"
    )
