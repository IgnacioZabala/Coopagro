import streamlit as st
import pandas as pd

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

# 2. Función para leer el Excel directamente desde tu Google Drive en la nube
@st.cache_data(ttl=600)
def load_excel_from_drive(file_id, sheet_name, header_row=0):
    try:
        url = f"https://docs.google.com/spreadsheets/d/19OVD6xBeK08o4cW1XrdMr54L1nciAJC2/export?format=xlsx"
        df = pd.read_excel(url, sheet_name=sheet_name, header=header_row)
        df.columns = df.columns.str.strip()
        return df
    except Exception as e:
        st.error(f"Error al leer el archivo desde Google Drive: {e}")
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
    st.markdown("Panel conectado automáticamente a Google Drive.")
    
    # PEGA TU FILE ID ACÁ ENTRE LAS COMILLAS
    FILE_ID_RECIBO = "TU_FILE_ID_AQUI" 
    
    if FILE_ID_RECIBO == "TU_FILE_ID_AQUI":
        st.warning("⚠️ Por favor, configurá tu FILE_ID de Google Drive en el código del `app.py` para ver los datos.")
    else:
        # Cargamos la solapa 'Résumen OD-PRO-03' (encabezado real en la fila 4)
        df_recibo = load_excel_from_drive(FILE_ID_RECIBO, sheet_name="Résumen OD-PRO-03", header_row=4)
        
        if not df_recibo.empty:
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
            
            if 'Litros\n(Ticket)' in df_view.columns:
                total_litros = df_view['Litros\n(Ticket)'].sum()
                col_m1, col_m2 = st.columns(2)
                col_m1.metric("Litros Totales (Ticket)", f"{total_litros:,.0f} L")
                col_m2.metric("Total Registros", len(df_view))
                st.markdown("<br>", unsafe_allow_html=True)

            st.dataframe(df_view, use_container_width=True, hide_index=True)
        else:
            st.error("No se pudieron leer los datos. Verificá que el archivo esté compartido públicamente como 'Lector'.")

elif modulo != "🥛 Recepción Coopagro":
    st.info("Módulo en desarrollo...")
