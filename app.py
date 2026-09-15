import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from fpdf import FPDF
import tempfile
import os

# 1. Configuración de la página
st.set_page_config(
    page_title="Control Planta Tandil | Coopagro", 
    page_icon="🏭", 
    layout="wide", 
    initial_sidebar_state="expanded"
)

st.markdown("""
    <style>
    .main-header { font-size: 24px; font-weight: 800; color: #1f2937; }
    .metric-card { background-color: #f9fafb; padding: 15px; border-radius: 8px; border-left: 4px solid #3b82f6; }
    </style>
""", unsafe_allow_html=True)

# 2. Carga del Excel desde Google Drive (vía File ID)
@st.cache_data(ttl=600)
def load_recibo_data(file_id):
    try:
        url = f"https://docs.google.com/spreadsheets/d/{file_id}/export?format=xlsx"
        df = pd.read_excel(url, sheet_name="Resumen OD-PRO-03", header=4)
        df.columns = df.columns.str.strip()
        
        # Limpieza y formateo de fechas
        df['Fecha'] = pd.to_datetime(df['Fecha'], errors='coerce')
        
        # Renombramos "Litros\n(Ticket)" a "Litros" para estandarizar
        if 'Litros\n(Ticket)' in df.columns:
            df.rename(columns={'Litros\n(Ticket)': 'Litros'}, inplace=True)
            
        return df.dropna(subset=['Fecha'])
    except Exception as e:
        st.error(f"Error al cargar el archivo de recepción: {e}")
        return pd.DataFrame()

# ID de tu archivo en Google Drive (reemplazá con tu ID real)
FILE_ID_RECIBO = "1Zaqtkadw4Mhgcc8WuuFb1YlXvvWbMsM4" 

df_recibo = load_recibo_data(FILE_ID_RECIBO)

# Menú lateral de Navegación
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
# MÓDULO 1: RECEPCIÓN COOPAGRO (Con Gestión Semanal / Mensual y PDF)
# =====================================================================
if modulo == "🥛 Recepción Coopagro":
    st.markdown('<p class="main-header">Gestión y Reportes de Recepción — Coopagro</p>', unsafe_allow_html=True)
    
    if not df_recibo.empty:
        # Controles laterales de período y tambo
        st.sidebar.markdown("---")
        st.sidebar.subheader("Filtros de Reporte")
        
        tipo_periodo = st.sidebar.radio("Modalidad de Período:", ["Semanal (Sáb a Vie)", "Mensual"])
        
        tambos_disponibles = sorted(df_recibo['Tambo'].dropna().unique().tolist())
        tambo_sel = st.sidebar.selectbox("Seleccione el Tambo:", tambos_disponibles)
        
        # Filtrar por tambo seleccionado
        df_tambo = df_recibo[df_recibo['Tambo'] == tambo_sel].copy()
        
        if tipo_periodo == "Semanal (Sáb a Vie)":
            # Creamos la lógica de semanas (sábado a viernes)
            # Desplazamos la fecha para que el sábado sea el inicio de semana (weekday de sabado es 5)
            df_tambo['Inicio_Semana'] = df_tambo['Fecha'].apply(lambda x: x - timedelta(days=(x.weekday() - 5) % 7))
            df_tambo['Fin_Semana'] = df_tambo['Inicio_Semana'] + timedelta(days=6)
            df_tambo['Rango_Semana'] = df_tambo.apply(lambda row: f"Del {row['Inicio_Semana'].strftime('%d/%m/%Y')} al {row['Fin_Semana'].strftime('%d/%m/%Y')}", axis=1)
            
            semanas_disponibles = sorted(df_tambo['Rango_Semana'].unique().tolist(), reverse=True)
            semana_sel = st.sidebar.selectbox("Cierre de Semana:", semanas_disponibles)
            
            df_filtrado = df_tambo[df_tambo['Rango_Semana'] == semana_sel]
            periodo_str = semana_sel
        else:
            df_tambo['Mes_Anio'] = df_tambo['Fecha'].dt.strftime('%B %Y')
            meses_disponibles = sorted(df_tambo['Mes_Anio'].unique().tolist(), reverse=True)
            mes_sel = st.sidebar.selectbox("Seleccione Mes:", meses_disponibles)
            
            df_filtrado = df_tambo[df_tambo['Mes_Anio'] == mes_sel]
            periodo_str = mes_sel

        # --- Métricas Principales (Tarjetas KPI) ---
        total_litros = df_filtrado['Litros'].sum() if 'Litros' in df_filtrado.columns else 0
        temp_prom = df_filtrado['Temperatura'].mean() if 'Temperatura' in df_filtrado.columns else 0
        grasa_prom = df_filtrado['Grasa'].mean() if 'Grasa' in df_filtrado.columns else 0
        prot_prom = df_filtrado['Proteína'].mean() if 'Proteína' in df_filtrado.columns else 0

        st.subheader(f"Resumen — {tambo_sel} ({periodo_str})")
        
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Litros Totales", f"{total_litros:,.0f} L")
        c2.metric("Temp. Promedio", f"{temp_prom:.1f}°C")
        c3.metric("Grasa Promedio", f"{grasa_prom:.2f}%" if grasa_prom > 0 else "S/D")
        c4.metric("Proteína Promedio", f"{prot_prom:.2f}%" if prot_prom > 0 else "S/D")
        
        st.markdown("<br>", unsafe_allow_html=True)

        # Mostrar tabla interactiva
        columnas_mostrar = [c for c in ['Fecha', 'N°Remito', 'Litros', 'Temperatura', 'Grasa', 'Proteína', 'Diferencia'] if c in df_filtrado.columns]
        
        # Ordenar por fecha
        df_filtrado = df_filtrado.sort_values(by='Fecha')
        
        st.dataframe(
            df_filtrado[columnas_mostrar].style.format({
                'Litros': '{:,.0f}',
                'Temperatura': '{:.1f}',
                'Grasa': '{:.2f}%',
                'Proteína': '{:.2f}%'
            }, na_rep="-"), 
            use_container_width=True, 
            hide_index=True
        )

        # --- GENERACIÓN DE PDF ---
        st.markdown("---")
        if st.button("📥 Descargar Reporte en PDF"):
            class PDF(FPDF):
                def header(self):
                    self.set_font('Arial', 'B', 14)
                    self.cell(0, 8, 'TANDIL - COOPAGRO', 0, 1, 'L')
                    self.set_font('Arial', '', 10)
                    self.cell(0, 6, 'Resumen de recolección de leche', 0, 1, 'L')
                    self.ln(5)

                def footer(self):
                    self.set_y(-15)
                    self.set_font('Arial', 'I', 8)
                    self.cell(0, 10, f'Página {self.page_no()}', 0, 0, 'C')

            pdf = PDF()
            pdf.add_page()
            pdf.set_font('Arial', '', 11)
            
            # Datos del encabezado del reporte
            pdf.cell(0, 6, f"Productor: {tambo_sel}", 0, 1)
            pdf.cell(0, 6, f"Período: {periodo_str}", 0, 1)
            pdf.cell(0, 6, f"Total Litros: {total_litros:,.0f} L", 0, 1)
            pdf.cell(0, 6, f"Temperatura Promedio: {temp_prom:.1f}°C", 0, 1)
            pdf.ln(8)
            
            # Tabla de registros
            pdf.set_font('Arial', 'B', 9)
            pdf.cell(30, 7, "Fecha", 1, 0, 'C')
            pdf.cell(40, 7, "N° Remito", 1, 0, 'C')
            pdf.cell(35, 7, "Litros", 1, 0, 'C')
            pdf.cell(30, 7, "Temp (°C)", 1, 0, 'C')
            pdf.cell(30, 7, "Grasa (%)", 1, 0, 'C')
            pdf.cell(30, 7, "Prot (%)", 1, 1, 'C')
            
            pdf.set_font('Arial', '', 9)
            for _, row in df_filtrado.iterrows():
                f_str = row['Fecha'].strftime('%d/%m/%Y') if pd.notnull(row['Fecha']) else ""
                rem_str = str(row.get('N°Remito', ''))
                lit_str = f"{row.get('Litros', 0):,.0f}"
                t_str = f"{row.get('Temperatura', 0):.1f}" if pd.notnull(row.get('Temperatura')) else "-"
                g_str = f"{row.get('Grasa', 0):.2f}" if pd.notnull(row.get('Grasa')) else "-"
                p_str = f"{row.get('Proteína', 0):.2f}" if pd.notnull(row.get('Proteína')) else "-"
                
                pdf.cell(30, 6, f_str, 1, 0, 'C')
                pdf.cell(40, 6, rem_str, 1, 0, 'C')
                pdf.cell(35, 6, lit_str, 1, 0, 'C')
                pdf.cell(30, 6, t_str, 1, 0, 'C')
                pdf.cell(30, 6, g_str, 1, 0, 'C')
                pdf.cell(30, 6, p_str, 1, 1, 'C')
                
            # Guardar PDF temporalmente para descarga
            temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
            pdf.output(temp_file.name)
            
            with open(temp_file.name, "rb") as f:
                st.download_button(
                    label="💾 Hacer clic aquí para descargar el PDF",
                    data=f,
                    file_name=f"Resumen_{tambo_sel.replace(' ', '_')}.pdf",
                    mime="application/pdf"
                )

    else:
        st.warning("No se pudieron cargar los datos de recepción.")

elif modulo != "🥛 Recepción Coopagro":
    st.info("Módulo en desarrollo...")
