import streamlit as st
import pandas as pd
import os

# 1. Configuración de la página
st.set_page_config(
    page_title="Control Planta Tandil", 
    page_icon="🏭", 
    layout="wide", 
    initial_sidebar_state="expanded"
)

st.markdown("""
    <style>
    .main-header { font-size: 24px; font-weight: 800; color: #1f2937; }
    </style>
""", unsafe_allow_html=True)

# 2. Definir la ruta base de Google Drive en tu PC
# (Ajustá esta ruta si tus archivos están en otra subcarpeta dentro de G:\Mi unidad)
DRIVE_PATH = r"G:\Mi unidad"

# 3. Función robusta para leer Excel localmente
@st.cache_data(ttl=600)
def load_local_excel(folder_name, file_name, sheet_name, header_row=0):
    try:
        file_path = os.path.join(DRIVE_PATH, folder_name, file_name)
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=header_row)
        df.columns = df.columns.str.strip()
        return df
    except Exception as e:
        st.error(f"No se pudo leer el archivo '{file_name}': {e}")
        return pd.DataFrame()

# Menú lateral
with st.sidebar:
    st.image("https://cdn-icons-png.flaticon.com/512/2830/2830305.png", width=50)
    st.markdown("### Control Planta Tandil")
    st.markdown("---")
    
    modulo = st.radio(
        "Seleccionar Módulo:", 
        [
            "🥛 Recepción Coopagro", 
            "🚛 Recepción Mastellone (Fasón)", 
            "🧀 Producción y Rendimiento", 
            "📦 Insumos y Costos"
        ]
    )

# =====================================================================
# MÓDULO 1: RECEPCIÓN COOPAGRO
# =====================================================================
if modulo == "🥛 Recepción Coopagro":
    st.markdown('<p class="main-header">Recepción de Materia Prima — Coopagro</p>', unsafe_allow_html=True)
    st.markdown("Lectura directa desde el archivo sincronizado en Google Drive.")
    
    # Nombre de tu carpeta en Drive y nombre exacto del archivo
    CARPETA_DRIVE = "App Muzzarella Coopagro" # Cambialo por el nombre real de tu carpeta si es distinto
    NOMBRE_ARCHIVO = "Resumen planilla Recibo  OD-PRO-03.xlsx"
    
    # Cargamos la solapa 'Résumen OD-PRO-03' (el encabezado real está en la fila 4, índice 4)
    df_recibo = load_local_excel(CARPETA_DRIVE, NOMBRE_ARCHIVO, sheet_name="Résumen OD-PRO-03", header_row=4)
    
    if not df_recibo.empty:
        # Limpiamos filas vacías si las hubiera
        df_recibo = df_recibo.dropna(subset=['Fecha'])
        
        # Filtro por Tambo
        col1, col2 = st.columns(2)
        with col1:
            tambos = ["Todos"] + sorted(df_recibo['Tambo'].dropna().unique().tolist()) if 'Tambo' in df_recibo.columns else ["Todos"]
            tambo_sel = st.selectbox("Filtrar por Tambo:", tambos)
            
        df_view = df_recibo.copy()
        if tambo_sel != "Todos":
            df_view = df_view[df_view['Tambo'] == tambo_sel]
            
        st.markdown("<br>", unsafe_allow_html=True)
        
        # Mostramos métricas rápidas arriba
        if 'Litros\n(Ticket)' in df_view.columns:
            total_litros = df_view['Litros\n(Ticket)'].sum()
            col_m1, col_m2 = st.columns(2)
            col_m1.metric("Litros Totales (Ticket)", f"{total_litros:,.0f} L")
            col_m2.metric("Total Registros", len(df_view))
            st.markdown("<br>", unsafe_allow_html=True)

        # Tabla interactiva
        st.dataframe(df_view, use_container_width=True, hide_index=True)
    else:
        st.warning("Verificá que el nombre de la carpeta y del archivo de recepción coincidan exactamente en tu Google Drive.")

elif modulo != "🥛 Recepción Coopagro":
    st.info("Módulo en desarrollo...")