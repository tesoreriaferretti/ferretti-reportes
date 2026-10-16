from reportes import conciliador_odoo, cxc_pendientes, no_identificado_bcp
import streamlit as st

st.set_page_config(
    page_title="Ferretti - Sistema de Reportes", layout="wide", page_icon="📊"
)

st.title("📊 Ferretti - Generador Centralizado de Reportes")

REPORTES = {
    "Conciliador Semanal: Odoo vs Bancos": conciliador_odoo.render,
    "Depósitos y Tarjetas No Identificados (BCP)": no_identificado_bcp.render,
    "Cuentas por Cobrar (CxC) Pendientes": cxc_pendientes.render,
}

st.sidebar.title("Navegación")
opcion = st.sidebar.selectbox(
    "Selecciona el reporte a generar:", list(REPORTES.keys())
)
st.sidebar.divider()
st.sidebar.caption(
    "💡 Tus archivos se procesan en memoria RAM sin guardarse en la nube."
)

REPORTES[opcion]()