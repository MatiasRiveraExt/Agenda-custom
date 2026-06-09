import streamlit as st
import pandas as pd
import re

st.set_page_config(page_title="Consolidador Excel", layout="wide")

st.title("📊 Consolidador Automático de Excel")
st.write("Sube los 3 archivos para generar el reporte final.")

# =========================
# UPLOAD FILES
# =========================
track_file = st.file_uploader("📁 SO Track", type=["xlsx"])
plan_file = st.file_uploader("📁 Pronóstico / Planificación", type=["xlsx"])
maestro_file = st.file_uploader("📁 Maestro (Departamento / PD)", type=["xlsx"])

# =========================
# FUNCTION: CLEAN COLUMNS
# =========================
def clean_columns(df):
    df.columns = df.columns.str.strip()
    return df

# =========================
# PROCESS BUTTON
# =========================
if st.button("🚀 Generar reporte"):

    if not track_file or not plan_file or not maestro_file:
        st.error("❌ Debes subir los 3 archivos")
        st.stop()

    # =========================
    # READ FILES
    # =========================
    track = pd.read_excel(track_file)
    plan = pd.read_excel(plan_file)
    maestro = pd.read_excel(maestro_file)

    track = clean_columns(track)
    plan = clean_columns(plan)
    maestro = clean_columns(maestro)

    # =========================
    # NORMALIZAR COLUMNAS CLAVE
    # =========================
    def find_col(df, keywords):
        for col in df.columns:
            for k in keywords:
                if k.lower() in col.lower():
                    return col
        return None

    track_order = find_col(track, ["po number"])
    plan_order = find_col(plan, ["o/c cliente"])
    maestro_order = find_col(maestro, ["num order"])

    if not track_order:
    st.error("❌ No se encontró 'PO Number' en SO Track")

    if not plan_order:
    st.error("❌ No se encontró 'O/C Cliente' en Pronóstico")

    if not maestro_order:
    st.error("❌ No se encontró 'Num Order' en Órdenes Liberadas")

    if not track_order or not plan_order or not maestro_order:
    st.stop()

    # =========================
    # STANDARDIZE
    # =========================
    track = track.rename(columns={
        track_order: "Order Number",
        find_col(track, ["delivered quantity"]): "Suma de Unidades",
        find_col(track, ["delivered amount"]): "Monto"
    })

    plan = plan.rename(columns={
        plan_order: "Order Number",
        find_col(plan, ["instru"]): "Instrucciones",
        find_col(plan, ["fecha"]): "Fecha de entrega"
    })

    maestro = maestro.rename(columns={
        maestro_order: "Order Number",
        find_col(maestro, ["depart"]): "Departamento",
        find_col(maestro, ["pd"]): "PD"
    })

    # =========================
    # EXTRAER HORA
    # =========================
    if "Instrucciones" in plan.columns:
        plan["Hora"] = plan["Instrucciones"].astype(str).str.extract(r'(\d{1,2}:\d{2})')
    else:
        plan["Hora"] = ""

    # =========================
    # MERGES
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

    st.success("✅ Reporte generado correctamente")

    st.dataframe(final.head(30))

    # =========================
    # DOWNLOAD
    # =========================
    csv = final.to_csv(index=False).encode("utf-8")

    st.download_button(
        "⬇ Descargar reporte (CSV)",
        csv,
        "reporte_final.csv",
        "text/csv"
    )
