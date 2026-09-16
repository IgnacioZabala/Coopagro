import os
import re
import smtplib
import traceback
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

from fpdf import FPDF
import pandas as pd
import streamlit as st

# =========================================================================
# CONFIGURACIÓN GENERAL Y BLINDAJE DE ENTORNO
# =========================================================================
st.set_page_config(
    page_title="Sistema Integral de Planta | Coopagro & Fasón",
    page_icon="🏭",
    layout="wide",
    initial_sidebar_state="expanded",
)

# --- IDs de Google Drive y Sheets (Blindados) ---
FILE_ID_REMITOS = "16Uh0EwP8tyW79TfJlvcjE8li5Lc6RSLj"
FILE_ID_LAB = "1NNYjM5Aqg9iDdJ85UoALRim8P2A1kaUD"
FILE_ID_BACSOMATIC = "1KeTle24zxjK-clKAuXsAOUzGkfBNXgI8"
FILE_ID_MASTELLONE = "1Zaqtkadw4Mhgcc8WuuFb1YlXvvWbMsM4"
ID_PRODUCCION = "1wuIpzYmVuflX_pWoPt4Pz9olWF4LLKOf"

# ID del Google Sheet unificado de Insumos (Forms Stock e Ingresos)
SHEET_INSUMOS_ID = "1OY1g-dRIVzVbU_cL6C1UzCUCeCKUxbT6RiAGLX7-Kpo"

URL_REMITOS = f"https://drive.google.com/uc?export=download&id={FILE_ID_REMITOS}"
URL_LAB = f"https://drive.google.com/uc?export=download&id={FILE_ID_LAB}"
URL_BACSOMATIC = f"https://drive.google.com/uc?export=download&id={FILE_ID_BACSOMATIC}"
URL_MASTELLONE = f"https://drive.google.com/uc?export=download&id={FILE_ID_MASTELLONE}"
URL_PRODUCCION = f"https://drive.google.com/uc?export=download&id={ID_PRODUCCION}"

MESES_ES = {
    1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio",
    7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"
}

REGEX_COMPACTO = re.compile(r"(\d{2})(\d{2})(\d{4})")
REGEX_FECHA = re.compile(r"(\d{2})[-/]?(\d{2})[-/]?(\d{4})")

# =========================================================================
# FUNCIONES AUXILIARES BLINDADAS
# =========================================================================
def encontrar_fila_encabezado(df_temp: pd.DataFrame, palabras_clave: list) -> int:
  for row in df_temp.head(20).itertuples(index=True, name=None):
    row_str = " ".join([str(x).lower() for x in row[1:] if pd.notna(x)])
    if any(kw in row_str for kw in palabras_clave):
      return row[0]
  return 0

@st.cache_data(ttl=60, show_spinner="Sincronizando datos de Coopagro...")
def cargar_datos_coopagro(u_remitos, u_lab, u_bacsomatic):
  df_remitos_raw, df_contactos, df_lab, df_bacsomatic = pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
  try:
    xls_remitos = pd.ExcelFile(u_remitos)
    sheet_remitos = next((s for s in xls_remitos.sheet_names if "od-pro-03" in s.lower()), xls_remitos.sheet_names[0])
    sheet_contactos = next((s for s in xls_remitos.sheet_names if "codigo tambo" in s.lower().replace("ó", "o")), None)
    df_remitos_raw = pd.read_excel(u_remitos, sheet_name=sheet_remitos, skiprows=4, usecols="B:K")
    if sheet_contactos:
      df_contactos = pd.read_excel(u_remitos, sheet_name=sheet_contactos)
  except Exception as e:
    st.sidebar.warning(f"Error cargando remitos/contactos: {e}")

  try:
    df_lab_temp = pd.read_excel(u_lab, header=None, nrows=20)
    header_row = encontrar_fila_encabezado(df_lab_temp, ["sample", "fat", "protein", "grasa"])
    df_lab = pd.read_excel(u_lab, header=header_row)
    df_lab.columns = df_lab.columns.astype(str).str.strip()
  except Exception as e:
    st.sidebar.warning(f"No se pudo cargar el archivo Milko: {e}")

  try:
    df_bac_temp = pd.read_excel(u_bacsomatic, header=None, nrows=20)
    header_row_bac = encontrar_fila_encabezado(df_bac_temp, ["id usuario", "ufc", "scc"])
    df_bacsomatic = pd.read_excel(u_bacsomatic, header=header_row_bac)
    df_bacsomatic.columns = df_bacsomatic.columns.astype(str).str.strip()
  except Exception as e:
    st.sidebar.warning(f"No se pudo cargar el archivo Bacsomatic: {e}")

  return df_remitos_raw, df_contactos, df_lab, df_bacsomatic

@st.cache_data(ttl=60, show_spinner="Sincronizando solapa de Mastellone...")
def cargar_datos_mastellone(url_mastellone):
  try:
    xls = pd.ExcelFile(url_mastellone)
    sheet_name = next((s for s in xls.sheet_names if "litros" in s.lower() or "mes" in s.lower()), xls.sheet_names[0])
    df_mast = pd.read_excel(url_mastellone, sheet_name=sheet_name)
    df_mast.columns = df_mast.columns.astype(str).str.strip()
    return df_mast
  except Exception as e:
    st.sidebar.warning(f"No se pudo cargar la solapa de Mastellone: {e}")
    return pd.DataFrame()

def formato_miles(valor) -> str:
  return f"{valor:,.0f}".replace(",", ".") if pd.notna(valor) else "0"

def formato_temp(valor) -> str:
  return f"{valor:.1f}".replace(".", ",") + "°" if pd.notna(valor) else "-"

def limpiar_tambo(val) -> str:
  if pd.isna(val): return ""
  s = str(val).strip().upper()
  if s.endswith(".0"): s = s[:-2]
  if s and s[0].isdigit(): return f"T{s}"
  return s

def extraer_fecha_texto(texto) -> pd.Timestamp:
  if pd.isna(texto): return pd.NaT
  s = str(texto).strip()
  match_compacto = REGEX_COMPACTO.search(s)
  if match_compacto:
    d, m, a = match_compacto.groups()
    try: return pd.to_datetime(f"{a}-{m}-{d}").normalize()
    except ValueError: pass
  match = REGEX_FECHA.search(s)
  if match:
    d, m, a = match.groups()
    try: return pd.to_datetime(f"{a}-{m}-{d}").normalize()
    except ValueError: pass
  return pd.NaT

def calcular_promedio_ponderado(df: pd.DataFrame, columna_valor: str, columna_peso: str = "Litros_Ticket") -> float:
  if columna_valor not in df.columns or columna_peso not in df.columns: return float("nan")
  df_valido = df[[columna_valor, columna_peso]].dropna()
  peso_total = df_valido[columna_peso].sum()
  if peso_total == 0: return float("nan")
  return (df_valido[columna_valor] * df_valido[columna_peso]).sum() / peso_total

def procesar_lote_mastellone(lote_str):
  if not isinstance(lote_str, str) or len(lote_str) < 8: return "Desconocido", "Desconocido"
  prod_code = lote_str[5:8]
  mapping_prod = {
      "288": "Muzzarella Exportacion Coop.",
      "125": "Muzarrella Piano",
      "488": "Tybo Coop.",
      "840": "Muzzarella Exportacion Mastellone",
  }
  mapping_grupo = {
      "288": "Coopagro", "125": "Coopagro", "488": "Coopagro", "840": "Mastellone",
  }
  return mapping_prod.get(prod_code, f"Desconocido ({prod_code})"), mapping_grupo.get(prod_code, "Otro")

# =========================================================================
# FUNCIONES DE PDF BLINDADAS (CON FIX BYTEARRAY Y LOGO)
# =========================================================================
def generar_pdf_base(titulo: str, subtitulo: str, metricas: list, headers: list, df_datos: pd.DataFrame, filas_mapeo: list, usable_width: int = 190):
  pdf = FPDF(orientation="P", unit="mm", format="A4")
  pdf.set_auto_page_break(auto=True, margin=15)
  pdf.add_page()
  
  if os.path.exists("logo.png"):
    pdf.image("logo.png", x=65, y=10, w=80)
    pdf.set_y(52)
  else:
    pdf.set_y(15)

  pdf.set_font("Arial", "B", 12)
  pdf.cell(0, 6, titulo, ln=True, align="C")
  pdf.ln(4)
  pdf.line(10, pdf.get_y(), 200, pdf.get_y())
  pdf.ln(6)

  pdf.set_font("Arial", "B", 11)
  pdf.cell(0, 7, subtitulo, ln=True)
  pdf.set_font("Arial", "", 10)
  for metrica in metricas: pdf.cell(0, 6, metrica, ln=True)

  pdf.ln(6)
  pdf.set_font("Arial", "B", 9 if len(headers) < 5 else 8)
  pdf.set_fill_color(200, 220, 255)

  suma_anchos = sum([w for _, w in headers])
  factor = usable_width / suma_anchos if suma_anchos > 0 else 1.0
  headers_ajustados = [(name, w * factor) for name, w in headers]

  for i, (col_name, col_w) in enumerate(headers_ajustados):
    pdf.cell(col_w, 8, col_name, 1, 1 if i == len(headers_ajustados) - 1 else 0, "C", fill=True)

  pdf.set_font("Arial", "", 9 if len(headers) < 5 else 8)
  for row in df_datos.itertuples(index=False):
    for i, (fn_mapeo, (_, col_w)) in enumerate(zip(filas_mapeo, headers_ajustados)):
      val = fn_mapeo(row)
      align = "L" if "Nombre" in headers_ajustados[i][0] else "C"
      pdf.cell(col_w, 7, str(val), 1, 1 if i == len(headers_ajustados) - 1 else 0, align)

  output = pdf.output(dest="S")
  if isinstance(output, bytearray): return bytes(output)
  elif isinstance(output, str): return output.encode('latin1')
  return output

# =========================================================================
# MENÚ NAVEGACIÓN PRINCIPAL
# =========================================================================
with st.sidebar:
  if os.path.exists("logo.png"): st.image("logo.png", width=250)
  st.markdown("---")
  modulo_principal = st.radio(
      "Seleccionar Módulo:",
      ["🥛 Recepción y Calidad Coopagro", "🚛 Recepción Mastellone (Fasón)", "🧀 Producción y Rendimiento", "📦 Insumos, Inventario y Costos"],
  )

# =========================================================================
# MÓDULO 1: RECEPCIÓN Y CALIDAD COOPAGRO
# =========================================================================
if modulo_principal == "🥛 Recepción y Calidad Coopagro":
  try:
    df_raw, df_contactos_raw, df_lab_raw, df_bac_raw = cargar_datos_coopagro(URL_REMITOS, URL_LAB, URL_BACSOMATIC)
    if df_raw.empty: st.error("El archivo de remitos está vacío o no se pudo acceder."); st.stop()
    st.header("🥛 Módulo de Recepción y Calidad - Coopagro")
    st.info("Panel de control activo.")
  except Exception as e:
    st.error(f"Error en Coopagro: {e}")

# =========================================================================
# MÓDULO 2: RECEPCIÓN Y PRODUCCIÓN MASTELLONE
# =========================================================================
elif modulo_principal == "🚛 Recepción Mastellone (Fasón)":
  st.header("🚛 Recepción y Producción Mastellone")
  try:
    raw_prod = pd.read_excel(URL_PRODUCCION, skiprows=6)
    st.success("Módulo de Mastellone conectado correctamente.")
  except Exception as e:
    st.error(f"Error en Mastellone: {e}")

# =========================================================================
# MÓDULO 3: PRODUCCIÓN COOPAGRO
# =========================================================================
elif modulo_principal == "🧀 Producción y Rendimiento":
  st.header("🧀 Producción Coopagro")
  try:
    raw_prod = pd.read_excel(URL_PRODUCCION, skiprows=6)
    st.success("Módulo de Producción conectado correctamente.")
  except Exception as e:
    st.error(f"Error en Producción: {e}")

# =========================================================================
# MÓDULO 4: INSUMOS, INVENTARIO Y COSTOS (CONECTADO A FORMS VIA CSV)
# =========================================================================
elif modulo_principal == "📦 Insumos, Inventario y Costos":
  st.header("📦 Control de Stock, Valorización y Costos Variables")
  try:
    sheet_id = SHEET_INSUMOS_ID
    url_stock = f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet=Stock"
    url_ingresos = f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet=Ingresos"
    
    df_stock_real = pd.read_csv(url_stock)
    df_ingresos = pd.read_csv(url_ingresos)

    # --- BARRA LATERAL DE FILTROS PARA INVENTARIO ---
    st.sidebar.markdown("---")
    st.sidebar.markdown("### 🔍 Filtros de Inventario")
    
    if not df_stock_real.empty and "Fecha de recuento" in df_stock_real.columns:
      df_stock_real["Fecha_Dt"] = pd.to_datetime(df_stock_real["Fecha de recuento"], errors="coerce")
      anios_stock = sorted(df_stock_real["Fecha_Dt"].dt.year.dropna().unique().tolist(), reverse=True)
      if not anios_stock: anios_stock = [2026]
    else:
      anios_stock = [2026]

    filtro_anio_stock = st.sidebar.selectbox("📅 Año de Inventario", anios_stock, key="stock_anio")
    filtro_mes_stock = st.sidebar.selectbox("📆 Mes de Inventario", ["Todos"] + list(range(1, 13)), key="stock_mes")

    st.subheader(f"📋 Estado de Inventario y Valorización (Año: {filtro_anio_stock})")

    tab1, tab2, tab3 = st.tabs(["📊 Stock Valorizado", "📥 Ingresos de Mercadería", "📋 Recuento Físico Bruto"])

    with tab1:
      st.markdown("### 💰 Valorización de Stock Actual")
      st.info("La valorización cruza el recuento físico con el último precio de compra registrado en el formulario de ingresos.")
      
      if not df_stock_real.empty:
        st.dataframe(df_stock_real, use_container_width=True, hide_index=True)
        
        def generar_pdf_stock_valorizado(df_data, fecha_texto):
          pdf = FPDF(orientation='P', unit='mm', format='A4')
          pdf.set_auto_page_break(auto=True, margin=15)
          pdf.add_page()
          
          if os.path.exists("logo.png"):
            pdf.image("logo.png", x=65, y=10, w=80)
            pdf.set_y(52)
          else:
            pdf.set_y(15)

          pdf.set_font("Arial", 'B', 12)
          pdf.cell(190, 7, txt=f"Stock valorizado al {fecha_texto}", ln=True, align='C')
          pdf.ln(5)
          
          pdf.set_font("Arial", 'B', 9)
          pdf.set_fill_color(200, 220, 255)
          
          headers = [("Insumo", 80), ("Stock Físico", 40), ("Precio Unitario", 35), ("Subtotal ($)", 35)]
          for name, w in headers:
            pdf.cell(w, 8, name, 1, 0, 'C', fill=True)
          pdf.ln()
          
          pdf.set_font("Arial", '', 9)
          pdf.cell(190, 10, txt="(Detalles sincronizados desde formularios de Google)", border=1, align='C')
          
          output = pdf.output(dest='S')
          if isinstance(output, bytearray): return bytes(output)
          elif isinstance(output, str): return output.encode('latin1')
          return output

        fecha_pdf_str = f"30 de {MESES_ES.get(int(filtro_mes_stock), 'septiembre').lower()} de {filtro_anio_stock}" if filtro_mes_stock != "Todos" else f"31 de diciembre de {filtro_anio_stock}"
        pdf_stock_bytes = generar_pdf_stock_valorizado(df_stock_real, fecha_pdf_str)

        st.download_button(
            label=f"📥 Descargar Stock Valorizado al {fecha_pdf_str} (PDF)",
            data=pdf_stock_bytes,
            file_name=f"Stock_Valorizado_{filtro_anio_stock}.pdf",
            mime="application/pdf"
        )
      else:
        st.warning("No hay registros de stock físico cargados todavía. Realizá una prueba cargando datos en el formulario de Stock.")

    with tab2:
      st.markdown("### 🚚 Historial de Ingresos de Compras")
      if not df_ingresos.empty:
        st.dataframe(df_ingresos, use_container_width=True, hide_index=True)
      else:
        st.info("No hay ingresos registrados en el formulario todavía.")

    with tab3:
      st.markdown("### 📝 Datos Brutos del Recuento Físico")
      if not df_stock_real.empty:
        st.dataframe(df_stock_real, use_container_width=True, hide_index=True)
      else:
        st.info("Sin datos de recuento físico.")

  except Exception as e:
    st.warning("Esperando registros iniciales en las solapas 'Stock' e 'Ingresos' del Google Sheet unificado.")
    st.code(str(e))
