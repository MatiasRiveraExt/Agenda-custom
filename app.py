import streamlit as st
import pandas as pd
import re

st.set_page_config(page_title="Consolidador Excel", layout="wide")
st.title("📊 Consolidador Automático de Excel")

st.write("Sube tus 3 archivos: SO Track, Pronóstico, Maestro")

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
        .str.strip()
        .str.replace("\n", " ")
        .str.replace("\xa0", " ")
    )
    return df

def find_col(df, keywords):
    for col in df.columns:
        c = col.lower().replace(" ", "").replace("/", "")
        for k in keywords:
            k2 = k.lower().replace(" ", "").replace("/", "")
            if k2 in c:
                return col
    return None

# =========================
# MAIN
# =========================
if st.button("🚀 Generar reporte"):

    if not track_file or not plan_file or not maestro_file:
        st.error("❌ Debes subir los 3 archivos")
        st.stop()

    # =========================
    # READ FILES
    # =========================
    track = pd.read_excel(track_file)
    plan_raw = pd.read_excel(plan_file, sheet_name=None)  # 🔥 TODAS LAS HOJAS
    maestro = pd.read_excel(maestro_file)

    track = clean_cols(track)
    maestro = clean_cols(maestro)

    # =========================
    # 🔥 ELEGIR HOJA MÁS GRANDE DEL PLAN
    # =========================
    plan = max(plan_raw.values(), key=lambda x: x.shape[0])
    plan = clean_cols(plan)

    st.write("📌 Hoja seleccionada automáticamente:", plan.shape)

    # =========================
    # DETECTAR COLUMNAS TRACK
    # =========================
    track_order = find_col(track, ["po number", "order", "po"])
    track_qty = find_col(track, ["quantity", "delivered", "qty"])
    track_amt = find_col(track, ["amount", "monto"])

    # =========================
    # DETECTAR PLAN
    # =========================
    plan_order = find_col(plan, ["o/c", "cliente", "order", "po"])
    plan_instr = find_col(plan, ["instru"])
    plan_date = find_col(plan, ["fecha"])

    # =========================
    # DETECTAR MAESTRO
    # =========================
    maestro_order = find_col(maestro, ["num order", "order"])
    maestro_dep = find_col(maestro, ["depart"])
    maestro_pd = find_col(maestro, ["pd"])

    # =========================
    # VALIDACIÓN
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
    # NORMALIZAR
    # =========================
    track = track.rename(columns={
        track_order: "Order Number",
        track_qty: "Suma de Unidades",
        track_amt: "Monto"
    })

    plan = plan.rename(columns={
        plan_order: "Order Number",
        plan_instr: "Instrucciones" if plan_instr else "Instrucciones",
        plan_date: "Fecha de entrega" if plan_date else "Fecha de entrega"
    })

    maestro = maestro.rename(columns={
        maestro_order: "Order Number",
        maestro_dep: "Departamento",
        maestro_pd: "PD"
    })

    # =========================
    # HORA
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
    # FINAL
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

    st.success("✅ Reporte generado")

    st.dataframe(final.head(50))

    csv = final.to_csv(index=False).encode("utf-8")

    st.download_button(
        "⬇ Descargar reporte",
        csv,
        "reporte_final.csv",
        "text/csv"
    )
