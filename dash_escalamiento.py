import streamlit as st 
import pandas as pd
import plotly.express as px
from io import BytesIO
from streamlit_autorefresh import st_autorefresh
from datetime import datetime
import numpy as np
import requests


st.set_page_config(page_title="Dashboard SAC", layout="wide")

# -----------------------
# MENU DASHBOARD
# -----------------------

dashboard = st.selectbox(
    "Seleccionar Dashboard",
    ["Gestión Escalamiento", "Gestión Casos Cerrados"]
)

# -----------------------
# PALETA DE COLORES
# -----------------------

PALETA_SAC = [
    "#8f5cda",
    "#7069d8",
    "#3a81d5",
    "#38a9d2",
    "#4cb2ca",
    "#ffffff"
]

# -----------------------
# ESTILO BOTÓN VERDE
# -----------------------

st.markdown("""
<style>
div.stButton > button {
    background-color: #28a745;
    color: white;
    border-radius: 8px;
    height: 40px;
    width: 180px;
    font-weight: bold;
}
div.stButton > button:hover {
    background-color: #218838;
    color: white;
}
</style>
""", unsafe_allow_html=True)

# -----------------------
# ESTILO CORPORATIVO DASHBOARD
# -----------------------

st.markdown("""
<style>

.main {
    background-color: #0f1116;
}

section[data-testid="stSidebar"] {
    background-color: #161a24;
}

h1, h2, h3 {
    color: #8f5cda;
}

div[data-testid="metric-container"] {
    background-color: #161a24;
    border: 1px solid #2a2f3a;
    padding: 10px;
    border-radius: 10px;
}

[data-testid="stDataFrame"] {
    border-radius: 10px;
}

div.stButton > button:hover {
    transform: scale(1.02);
    transition: 0.2s;
}

label {
    color: #cfd8ff !important;
}

</style>
""", unsafe_allow_html=True)

# -----------------------
# TÍTULO
# -----------------------

col_title, col_button = st.columns([9,1])

with col_title:
    
    if dashboard == "Gestión Escalamiento":
        st.title("Dashboard Gestión Escalamiento")
    else:
        st.title("Dashboard Casos Cerrados")

with col_button:
    if st.button("🔄 Actualizar"):
        st.cache_data.clear()
        st.rerun()

# -----------------------
# AUTO ACTUALIZACIÓN DIARIA
# -----------------------

st_autorefresh(
    interval=24 * 60 * 60 * 1000,
    key="auto_refresh"
)

# -----------------------
# CARGAR DATOS
# -----------------------

url_excel = "https://suncompanycol-my.sharepoint.com/personal/sac_dispower_co/_layouts/15/download.aspx?share=IQD470ahen_KQa7HKVzXNh_EAa_2TzrHt1M9hOShRFYbwaM&e=mDU1fU"


@st.cache_data(ttl=3600)
def cargar_datos():

    response = requests.get(url_excel)

    archivo = BytesIO(response.content)

    df_abiertos = pd.read_excel(
        archivo,
        sheet_name="Consolidado",
        engine="openpyxl"
    )

    archivo.seek(0)

    df_cerrados = pd.read_excel(
        archivo,
        sheet_name="Casos_Cerrados",
        engine="openpyxl"
    )

    return df_abiertos, df_cerrados


df_abiertos, df_cerrados = cargar_datos()

# -----------------------
# SELECCIÓN DATAFRAME
# -----------------------

if dashboard == "Gestión Escalamiento":
    df = df_abiertos.copy()
else:
    df = df_cerrados.copy()


df["FechaCreacion"] = pd.to_datetime(df["FechaCreacion"], errors="coerce")

# -----------------------
# CALCULO DE DIAS PARA CIERRE
# -----------------------

hoy = np.datetime64(pd.Timestamp.today().normalize(), 'D')

df["Dias para Cierre"] = np.busday_count(
    df["FechaCreacion"].values.astype("datetime64[D]"),
    hoy
)

# -----------------------
# SEMAFORO POR DIAS
# -----------------------

df["Semaforo_Dias"] = pd.cut(
    df["Dias para Cierre"],
    bins=[-999,5,10,999],
    labels=["🟢 En tiempo","🟡 En riesgo","🔴 Vencido"]
)

# -----------------------

st.caption(f"Última actualización del dashboard: {datetime.now().strftime('%d/%m/%Y %H:%M')}")

# -----------------------
# FILTROS
# -----------------------

st.sidebar.header("Filtros")

responsable = st.sidebar.multiselect(
    "Responsable",
    df["Responsable"].dropna().unique()
)

seccional = st.sidebar.multiselect(
    "NombreSeccionales",
    df["NombreSeccionales"].dropna().unique()
)

menu = st.sidebar.multiselect(
    "Menu",
    df["Menu"].dropna().unique()
)

submenu1 = st.sidebar.multiselect(
    "SubMenu1",
    df["SubMenu1"].dropna().unique()
)

estado = st.sidebar.multiselect(
    "Semaforo",
    df["Semaforo"].dropna().unique()
)

fecha = st.sidebar.date_input(
    "Fecha de Creación",
    []
)

canal = st.sidebar.multiselect(
    "Canal",
    df["canal"].dropna().unique()

)
# aplicar filtros

if seccional:
    df = df[df["NombreSeccionales"].isin(seccional)]

if canal:
    df = df[df["canal"].isin(canal)]

if estado:
    df = df[df["Semaforo"].isin(estado)]

if menu:
    df = df[df["Menu"].isin(menu)]

if submenu1:
    df = df[df["SubMenu1"].isin(submenu1)]

if responsable:
    df = df[df["Responsable"].isin(responsable)]

if len(fecha) == 2:
    df = df[(df["FechaCreacion"] >= pd.to_datetime(fecha[0])) &
            (df["FechaCreacion"] <= pd.to_datetime(fecha[1]))]

# -----------------------
# INDICADORES
# -----------------------

col1, col2, col3, col4 = st.columns(4)

col1.metric("Total Tickets", len(df))
col2.metric("Seccionales", df["NombreSeccionales"].nunique())
col3.metric("NUI", df["NUI"].nunique())

promedio_cierre = round(df["Dias para Cierre"].mean(),2)

col4.metric("Promedio días cierre", promedio_cierre)

st.divider()

# -----------------------
# INDICADORES SEMAFORO
# -----------------------

verdes = len(df[df["Semaforo_Dias"]=="🟢 En tiempo"])
amarillos = len(df[df["Semaforo_Dias"]=="🟡 En riesgo"])
rojos = len(df[df["Semaforo_Dias"]=="🔴 Vencido"])

colA, colB, colC = st.columns(3)

colA.metric("🟢 Tickets en tiempo", verdes)
colB.metric("🟡 Tickets en riesgo", amarillos)
colC.metric("🔴 Tickets vencidos", rojos)

st.divider()

# -----------------------
# TABLA DETALLE
# -----------------------

st.subheader("Detalle de Tickets")

columnas_tabla = [
    "NUI",
    "NombreSeccionales",
    "Id_Tickets",
    "Semaforo",
    "Menu",
    "SubMenu1",
    "FechaCreacion",
    "Dias para Cierre",
    "Creador_gestion",
    "Responsable",
    "Fecha Asignación",
    "Descripción"
]

tabla = df[columnas_tabla]

tabla["FechaCreacion"] = pd.to_datetime(
    tabla["FechaCreacion"], errors="coerce"
).dt.strftime("%d-%m-%Y")

tabla["Fecha Asignación"] = pd.to_datetime(
    tabla["Fecha Asignación"], errors="coerce"
).dt.strftime("%d-%m-%Y")

st.dataframe(tabla, use_container_width=True, height=400)

# -----------------------
# DESCARGAR EXCEL
# -----------------------

def convertir_excel(df):

    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Tickets")

    return output.getvalue()

excel_data = convertir_excel(tabla)

st.download_button(
    label="📥 Descargar tabla en Excel",
    data=excel_data,
    file_name="reporte_tickets_filtrado.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

st.divider()

# -----------------------
# GRAFICOS PRINCIPALES
# -----------------------

col1, col2 = st.columns(2)

tickets_seccional = df.groupby("NombreSeccionales").size().reset_index(name="Tickets")

fig1 = px.bar(
    tickets_seccional,
    x="NombreSeccionales",
    y="Tickets",
    title="Tickets por Seccional",
    color="NombreSeccionales",
    color_discrete_sequence=PALETA_SAC
)

col1.plotly_chart(fig1, use_container_width=True)

tickets_canal = df.groupby("canal").size().reset_index(name="Tickets")

fig2 = px.pie(
    tickets_canal,
    names="canal",
    values="Tickets",
    title="Distribución por Canal",
    color_discrete_sequence=PALETA_SAC
)

col2.plotly_chart(fig2, use_container_width=True)

# -----------------------
# SEMAFORO
# -----------------------

st.subheader("Estado de Tickets (Semáforo)")

tickets_semaforo = df.groupby("Semaforo").size().reset_index(name="Tickets")

fig3 = px.bar(
    tickets_semaforo,
    x="Semaforo",
    y="Tickets",
    color="Semaforo",
    color_discrete_sequence=PALETA_SAC
)

st.plotly_chart(fig3, use_container_width=True)

# -----------------------
# RANKING SECCIONALES
# -----------------------

st.subheader("Ranking de Seccionales")

ranking_seccionales = df.groupby("NombreSeccionales").size().reset_index(name="Tickets")

ranking_seccionales = ranking_seccionales.sort_values("Tickets", ascending=False)

fig4 = px.bar(
    ranking_seccionales,
    x="Tickets",
    y="NombreSeccionales",
    orientation="h",
    color="NombreSeccionales",
    color_discrete_sequence=PALETA_SAC
)

st.plotly_chart(fig4, use_container_width=True)

# -----------------------
# EVOLUCION EN EL TIEMPO
# -----------------------

st.subheader("Evolución de Tickets")

tickets_fecha = df.groupby(df["FechaCreacion"].dt.date).size().reset_index(name="Tickets")

fig5 = px.line(
    tickets_fecha,
    x="FechaCreacion",
    y="Tickets",
    color_discrete_sequence=PALETA_SAC
)

st.plotly_chart(fig5, use_container_width=True)

# -----------------------
# TOP RESPONSABLES
# -----------------------

st.subheader("Top 10 Responsables con más Tickets")

top_responsables = df.groupby("Responsable").size().reset_index(name="Tickets")

top_responsables = top_responsables.sort_values("Tickets", ascending=False).head(10)

fig6 = px.bar(
    top_responsables,
    x="Tickets",
    y="Responsable",
    orientation="h",
    color="Responsable",
    color_discrete_sequence=PALETA_SAC
)

st.plotly_chart(fig6, use_container_width=True)

# -----------------------
# PROMEDIO CIERRE RESPONSABLE
# -----------------------

st.subheader("Tiempo Promedio de Cierre por Responsable")

promedio_responsable = df.groupby("Responsable")["Dias para Cierre"].mean().reset_index()

promedio_responsable = promedio_responsable.sort_values("Dias para Cierre", ascending=False)

fig8 = px.bar(
    promedio_responsable,
    x="Dias para Cierre",
    y="Responsable",
    orientation="h",
    color="Responsable",
    color_discrete_sequence=PALETA_SAC
)

st.plotly_chart(fig8, use_container_width=True)
