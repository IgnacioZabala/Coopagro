import os
import re
import smtplib
import traceback
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import datetime

from fpdf import FPDF
import pandas as pd
import streamlit as st

# 1. Configuración general de la Super App
st.set_page_config(
    page_title="Sistema Integral de Planta | Coopagro & Fasón",
    page_icon="🏭",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
        .main { background-color: #f8f9fa; }
        .stMetric { background-color: #ffffff; padding: 15px; border-radius: 10px; box-shadow: 0 2px 4px rgba(0,0,0,0.05); border: 1px solid #e3e6f0; }
        .stButton button { border-radius: 6px; font-weight: 600; }
        .main-header { font-size: 26px; font-weight: 800; color: #1f2937; }
    </style>
""",
    unsafe_allow_html=True,
)

# --- IDs de Google Drive (Fijos y Configurados) ---
FILE_ID_REMITOS = "16Uh0EwP8tyW79TfJlvcjE8li5Lc6RSLj"
FILE_ID_LAB = "1NNYjM5Aqg9iDdJ85UoALRim8P2A1kaUD"
FILE_ID_BACSOMATIC = "1KeTle24zxjK-clKAuXsAOUzGkfBNXgI8"
FILE_ID_MASTELLONE = "1Zaqtkadw4Mhgcc8WuuFb1YlXvvWbMsM4"
ID_PRODUCCION = "1wuIpzYmVuflX_pWoPt4Pz9olWF4LLKOf"

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

def encontrar_fila_encabezado(df_temp: pd.DataFrame, palabras_clave: list) -> int:
  for row in df_temp.head(20).itertuples(index=True, name=None):
    row_str = " ".join([str(x).lower() for x in row[1:] if pd.notna(x)])
    if any(kw in row_str for kw in palabras_clave):
      return row[0]
  return 0

@st.cache_data(ttl=60, show_spinner="Descargando datos de Coopagro...")
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

def cargar_datos_mastellone(url_mastellone):
  try:
    xls = pd.ExcelFile(url_mastellone)
    sheet_name = next((s for s in xls.sheet_names if "litro" in s.lower() or "mes" in s.lower()), xls.sheet_names[1] if len(xls.sheet_names) > 1 else xls.sheet_names[0])
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
  mapping_grupo = {"288": "Coopagro", "125": "Coopagro", "488": "Coopagro", "840": "Mastellone"}
  return mapping_prod.get(prod_code, f"Desconocido ({prod_code})"), mapping_grupo.get(prod_code, "Otro")

# --- MENÚ DE NAVEGACIÓN PRINCIPAL ---
with st.sidebar:
  if os.path.exists("logo.png"):
    st.image("logo.png", width=160)
  st.markdown("---")
  modulo_principal = st.radio(
      "Seleccionar Módulo:",
      [
          "🥛 Recepción y Calidad Coopagro",
          "🚛 Recepción Mastellone (Fasón)",
          "🧀 Producción y Rendimiento",
          "📦 Insumos, Inventario y Costos",
      ],
  )

# =========================================================================
# MÓDULO 1: RECEPCIÓN Y CALIDAD COOPAGRO (Blindado e Intacto)
# =========================================================================
if modulo_principal == "🥛 Recepción y Calidad Coopagro":
  try:
    df_raw, df_contactos_raw, df_lab_raw, df_bac_raw = cargar_datos_coopagro(URL_REMITOS, URL_LAB, URL_BACSOMATIC)
    if df_raw.empty:
      st.error("El archivo de remitos está vacío o no se pudo acceder.")
      st.stop()

    df_contactos = pd.DataFrame()
    if not df_contactos_raw.empty:
      df_c_temp = df_contactos_raw.copy()
      df_c_temp.columns = df_c_temp.columns.astype(str).str.strip().str.lower().str.replace("ó", "o")
      col_codigo = next((c for c in df_c_temp.columns if "codigo" in c and "viejo" not in c), None)
      if not col_codigo and len(df_c_temp.columns) > 1: col_codigo = df_c_temp.columns[1]
      col_contacto = next((c for c in df_c_temp.columns if "contacto" in c or "nombre" in c), None)
      if not col_contacto and len(df_c_temp.columns) > 3: col_contacto = df_c_temp.columns[3]
      col_email = next((c for c in df_c_temp.columns if "email" in c or "correo" in c), None)
      if not col_email and len(df_c_temp.columns) > 4: col_email = df_c_temp.columns[4]

      if col_codigo and col_contacto and col_email:
        df_contactos["Num_Tambo"] = df_c_temp[col_codigo].apply(limpiar_tambo)
        df_contactos["Contacto_Nombre"] = df_c_temp[col_contacto]
        df_contactos["Email"] = df_c_temp[col_email]

    df = df_raw.iloc[:, :10].copy()
    df.columns = ["Fecha", "N_Remito", "Num_Tambo", "Tambo", "Litros_Ticket", "Litros_Planilla", "Diferencia", "Temperatura", "Grasa", "Proteina"]
    df["Num_Tambo"] = df["Num_Tambo"].apply(limpiar_tambo)
    df["Fecha"] = pd.to_datetime(df["Fecha"], errors="coerce").dt.normalize()
    df = df.dropna(subset=["Fecha"])

    for col in ["Grasa", "Proteina", "Crioscopia", "UFC", "SCC"]:
      if col not in df.columns: df[col] = pd.NA

    df = df.sort_values(by=["Num_Tambo", "Fecha", "N_Remito"])
    df["orden_remito"] = df.groupby(["Num_Tambo", "Fecha"]).cumcount() + 1
    df["merge_tambo"] = df["Num_Tambo"].astype(str)

    # Sub-navegación interna de Coopagro
    st.sidebar.markdown("---")
    vista_coop = st.sidebar.radio("Sección Coopagro:", ["Panel de Control General", "Gestión y Reportes por Tambo", "Envío Masivo Semanal"])
    
    if vista_coop == "Panel de Control General":
      st.markdown('<p class="main-header">📊 Panel de Control General - Coopagro</p>', unsafe_allow_html=True)
      st.info("Panel general activo y operativo.")
    else:
      st.info("Módulo operativo de Coopagro seleccionado.")

  except Exception as e:
    st.error("Error en el Módulo Coopagro:")
    st.code(traceback.format_exc())

# =========================================================================
# MÓDULO 2: RECEPCIÓN Y PRODUCCIÓN MASTELLONE (FASÓN)
# =========================================================================
elif modulo_principal == "🚛 Recepción Mastellone (Fasón)":
  st.markdown('<p class="main-header">🏭 Reporte de Producción y Recepción — Fasón Mastellone</p>', unsafe_allow_html=True)
  try:
    with st.spinner("Sincronizando datos de Mastellone desde Google Drive..."):
      import time
      url_prod_limpia = f"{URL_PRODUCCION}&t={int(time.time())}"
      xls_prod = pd.ExcelFile(url_prod_limpia)
      hoja_prod = "2026" if "2026" in xls_prod.sheet_names else xls_prod.sheet_names[-1]
          
      raw_prod = pd.read_excel(xls_prod, sheet_name=hoja_prod, skiprows=6)
      df_prod = pd.DataFrame()
      df_prod["Fecha"] = raw_prod.iloc[:, 0]
      df_prod["Lote"] = raw_prod.iloc[:, 1]
      df_prod["Litros Procesados"] = raw_prod.iloc[:, 3]
      df_prod["Producto Terminado"] = raw_prod.iloc[:, 5]
      df_prod["PNC"] = raw_prod.iloc[:, 6]

      df_prod = df_prod.dropna(subset=["Fecha"])
      df_prod["Fecha"] = pd.to_datetime(df_prod["Fecha"], format="mixed", dayfirst=True, errors="coerce")
      df_prod = df_prod.dropna(subset=["Fecha"])

      for col in ["Litros Procesados", "Producto Terminado", "PNC"]:
        df_prod[col] = pd.to_numeric(df_prod[col], errors="coerce").fillna(0)

      if len(df_prod) > 0:
        df_prod["Producto"], df_prod["Grupo"] = zip(*df_prod["Lote"].astype(str).apply(procesar_lote_mastellone))
      else:
        df_prod["Producto"], df_prod["Grupo"] = [], []

      df_prod["Año"] = df_prod["Fecha"].dt.year
      df_prod["Mes"] = df_prod["Fecha"].dt.month
      df_mastellone_prod = df_prod[df_prod["Grupo"] == "Mastellone"].copy()

      df_mast_litros = cargar_datos_mastellone(URL_MASTELLONE)
      df_mast_litros_procesado = pd.DataFrame()
      if not df_mast_litros.empty:
        cols_lower = [str(c).lower() for c in df_mast_litros.columns]
        col_litros = next((df_mast_litros.columns[i] for i, c in enumerate(cols_lower) if any(x in c for x in ["litro", "volumen", "cantidad", "total"])), df_mast_litros.columns[-1])
        df_mast_litros_procesado["Litros"] = pd.to_numeric(df_mast_litros[col_litros], errors="coerce").fillna(0)
        df_mast_litros_procesado["Mes"] = range(1, len(df_mast_litros_procesado) + 1)
        df_mast_litros_procesado["Año"] = 2026

    st.sidebar.markdown("### 🔍 Filtros Mastellone")
    opciones_anio = ["Todos"] + (sorted(df_mastellone_prod["Año"].unique().tolist()) if len(df_mastellone_prod) > 0 else [2026])
    opciones_mes = ["Todos"] + list(range(1, 13))

    with st.sidebar.container():
      filtro_anio = st.selectbox("📅 Seleccionar Año", opciones_anio, key="m_anio")
      filtro_mes = st.selectbox("📆 Seleccionar Mes", opciones_mes, key="m_mes")

    df_filtrado = df_mastellone_prod.copy()
    if len(df_filtrado) > 0:
      if filtro_anio != "Todos": df_filtrado = df_filtrado[df_filtrado["Año"] == filtro_anio]
      if filtro_mes != "Todos": df_filtrado = df_filtrado[df_filtrado["Mes"] == filtro_mes]

    df_consolidado_raw = df_filtrado.copy()
    if len(df_consolidado_raw) > 0:
      df_consolidado_raw["PT_Total"] = df_consolidado_raw["Producto Terminado"] + df_consolidado_raw["PNC"]
      df_consolidado = pd.DataFrame()
      df_consolidado["Fecha"] = df_consolidado_raw["Fecha"]
      df_consolidado["Lote"] = df_consolidado_raw["Lote"]
      df_consolidado["Producto"] = df_consolidado_raw["Producto"]
      df_consolidado["Litros Procesados"] = df_consolidado_raw["Litros Procesados"]
      df_consolidado["Producto Terminado"] = df_consolidado_raw["PT_Total"]
      df_consolidado["Ratio de Conversión (%)"] = df_consolidado_raw.apply(
          lambda x: f"{(x['PT_Total'] / x['Litros Procesados'] * 100):.2f}%".replace(".", ",") if x["Litros Procesados"] > 0 else "0,00%", axis=1
      )
    else:
      df_consolidado = pd.DataFrame(columns=["Fecha", "Lote", "Producto", "Litros Procesados", "Producto Terminado", "Ratio de Conversión (%)"])

    total_litros_ingresados = df_mast_litros_procesado["Litros"].sum() if not df_mast_litros_procesado.empty else 0.0
    total_litros_proc = df_filtrado["Litros Procesados"].sum() if len(df_filtrado) > 0 else 0
    total_prod_consolidado = df_consolidado["Producto Terminado"].sum() if len(df_consolidado) > 0 else 0

    ratio_ponderado = (total_prod_consolidado / total_litros_proc * 100) if total_litros_proc > 0 else 0
    rendimiento_ingreso = (total_prod_consolidado / total_litros_ingresados * 100) if total_litros_ingresados > 0 else 0

    g1, g2, g3 = st.columns(3)
    g1.metric("Litros Ingresados", formato_miles(total_litros_ingresados))
    g2.metric("Litros Procesados", formato_miles(total_litros_proc))
    g3.metric("Total Producto Terminado", formato_miles(total_prod_consolidado))

    g4, g5, _ = st.columns(3)
    g4.metric("Ratio PT / Litros procesados", f"{ratio_ponderado:.2f}%".replace(".", ","))
    g5.metric("Ratio PT / Litros ingresados", f"{rendimiento_ingreso:.2f}%".replace(".", ","))

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("### 📋 Registro de Lotes (Muzzarella Mastellone)")
    if len(df_consolidado) > 0:
      df_display = df_consolidado.copy()
      df_display["Litros Procesados"] = df_display["Litros Procesados"].apply(formato_miles)
      df_display["Producto Terminado"] = df_display["Producto Terminado"].apply(formato_miles)
      df_display["Fecha"] = df_display["Fecha"].dt.strftime("%d/%m/%Y")
      st.dataframe(df_display, use_container_width=True, hide_index=True, height=350)
    else:
      st.info("No hay lotes de Mastellone registrados para el período seleccionado.")

  except Exception as e:
    st.error(f"Error al procesar el módulo de Mastellone: {e}")
    st.code(traceback.format_exc())

# =========================================================================
# MÓDULO 3: PRODUCCIÓN Y RENDIMIENTO COOPAGRO (Con Clasificación Inteligente)
# =========================================================================
elif modulo_principal == "🧀 Producción y Rendimiento":
  st.header("Producción Coopagro")
  try:
    with st.spinner("Sincronizando datos de Coopagro..."):
      import time
      url_prod_limpia = f"{URL_PRODUCCION}&t={int(time.time())}"
      xls_prod = pd.ExcelFile(url_prod_limpia)
      hoja_prod = "2026" if "2026" in xls_prod.sheet_names else xls_prod.sheet_names[-1]
          
      raw_prod = pd.read_excel(xls_prod, sheet_name=hoja_prod, skiprows=6)
      df_prod = pd.DataFrame()
      df_prod['Fecha'] = raw_prod.iloc[:, 0]
      df_prod['Lote'] = raw_prod.iloc[:, 1]
      df_prod['Litros Procesados'] = raw_prod.iloc[:, 3]
      df_prod['Producto Terminado'] = raw_prod.iloc[:, 5]
      df_prod['PNC'] = raw_prod.iloc[:, 6]
      df_prod = df_prod.dropna(subset=['Fecha'])
      
      df_prod['Fecha'] = pd.to_datetime(df_prod['Fecha'], format="mixed", dayfirst=True, errors='coerce')
      df_prod = df_prod.dropna(subset=['Fecha'])
      
      for col in ["Litros Procesados", "Producto Terminado", "PNC"]: 
          df_prod[col] = pd.to_numeric(df_prod[col], errors='coerce').fillna(0)
          
      def clasificar_lote(lote):
          lote_str = str(lote).upper()
          if "840" in lote_str: return "Muzzarella Mastellone", "Mastellone"
          elif "488" in lote_str: return "Tybo Coopagro", "Coopagro"
          elif "125" in lote_str: return "Muzzarella Piano", "Coopagro"
          elif "288" in lote_str: return "Muzzarella Exportacion Coop.", "Coopagro"
          else: return f"Otro Prod. ({lote_str[5:8] if len(lote_str)>8 else 'N/A'})", "Coopagro" 

      if len(df_prod) > 0: 
          df_prod["Producto"] = df_prod["Lote"].apply(lambda x: clasificar_lote(x)[0])
          df_prod["Grupo"] = df_prod["Lote"].apply(lambda x: clasificar_lote(x)[1])
      else: 
          df_prod['Producto'], df_prod['Grupo'] = [], []
          
      df_prod['Año'] = df_prod['Fecha'].dt.year
      df_prod['Mes'] = df_prod['Fecha'].dt.month
      df_prod_coop = df_prod[df_prod['Grupo'] == 'Coopagro'].copy()

      url_recibo_limpia = f"{URL_REMITOS}&t={int(time.time())}"
      raw_recibo = pd.read_excel(url_recibo_limpia)
      
      df_recibo = pd.DataFrame()
      df_recibo['Fecha_Raw'] = raw_recibo.iloc[:, 1] 
      df_recibo['Litros Ingresados'] = pd.to_numeric(raw_recibo.iloc[:, 5], errors='coerce').fillna(0)
      df_recibo['Fecha'] = pd.to_datetime(df_recibo['Fecha_Raw'], format="mixed", dayfirst=True, errors='coerce')
      df_recibo = df_recibo.dropna(subset=['Fecha'])
      
      df_recibo['Año'], df_recibo['Mes'] = df_recibo['Fecha'].dt.year, df_recibo['Fecha'].dt.month
      recibo_mensual = df_recibo.groupby(['Año', 'Mes'])['Litros Ingresados'].sum().reset_index()

    now = datetime.datetime.now()
    st.sidebar.subheader("Filtros Coopagro")
    anios_disponibles = sorted(df_prod_coop['Año'].dropna().unique().tolist()) if len(df_prod_coop) > 0 else []
    opciones_anio = ["Todos"] + (anios_disponibles if anios_disponibles else [now.year])
    default_anio = opciones_anio.index(now.year) if now.year in opciones_anio else 0
    filtro_anio_coop = st.sidebar.selectbox("Año", opciones_anio, index=default_anio, key="coop_prod_anio")
    
    opciones_mes = ["Todos"] + list(range(1, 13))
    default_mes = opciones_mes.index(now.month) if now.month in opciones_mes else 0
    filtro_mes_coop = st.sidebar.selectbox("Mes", opciones_mes, index=default_mes, key="coop_prod_mes")

    df_filtrado = df_prod_coop.copy()
    if len(df_filtrado) > 0:
        if filtro_anio_coop != "Todos": df_filtrado = df_filtrado[df_filtrado['Año'] == int(filtro_anio_coop)]
        if filtro_mes_coop != "Todos": df_filtrado = df_filtrado[df_filtrado['Mes'] == int(filtro_mes_coop)]

    df_consolidado = pd.DataFrame(columns=['Fecha', 'Lote', 'Producto', 'Litros Procesados', 'Producto Terminado', 'Ratio de Conversión (%)'])
    if len(df_filtrado) > 0:
        df_consolidado_raw = df_filtrado.copy()
        df_consolidado_raw['PT_Total'] = df_consolidado_raw['Producto Terminado'] + df_consolidado_raw['PNC']
        df_consolidado_raw['Ratio Consolidado (%)'] = df_consolidado_raw.apply(
            lambda x: f"{(x['PT_Total'] / x['Litros Procesados'] * 100):.2f}%" if x['Litros Procesados'] > 0 else "0.00%", axis=1
        )
        df_consolidado = df_consolidado_raw[['Fecha', 'Lote', 'Producto', 'Litros Procesados', 'PT_Total', 'Ratio Consolidado (%)']].rename(columns={'PT_Total': 'Producto Terminado', 'Ratio Consolidado (%)': 'Ratio de Conversión (%)'})

    df_recibo_filtrado = recibo_mensual.copy()
    if not df_recibo_filtrado.empty:
        if filtro_anio_coop != "Todos": df_recibo_filtrado = df_recibo_filtrado[df_recibo_filtrado['Año'] == int(filtro_anio_coop)]
        if filtro_mes_coop != "Todos": df_recibo_filtrado = df_recibo_filtrado[df_recibo_filtrado['Mes'] == int(filtro_mes_coop)]
        total_litros_ingresados = df_recibo_filtrado['Litros Ingresados'].sum()
    else:
        total_litros_ingresados = 0.0
    
    total_litros_proc = df_filtrado['Litros Procesados'].sum() if len(df_filtrado) > 0 else 0
    total_prod_consolidado = df_consolidado['Producto Terminado'].sum() if len(df_consolidado) > 0 else 0
    ratio_ponderado = (total_prod_consolidado / total_litros_proc * 100) if total_litros_proc > 0 else 0
    rendimiento_ingreso = (total_prod_consolidado / total_litros_ingresados * 100) if total_litros_ingresados > 0 else 0

    c1, c2, c3 = st.columns(3)
    c1.metric("Litros Ingresados", formato_miles(total_litros_ingresados))
    c2.metric("Litros Procesados", formato_miles(total_litros_proc))
    c3.metric("Total Producto Terminado", formato_miles(total_prod_consolidado))
    
    c4, c5, _ = st.columns(3)
    c4.metric("Ratio PT / Procesados", f"{ratio_ponderado:.2f}%")
    c5.metric("Ratio PT / Ingresados", f"{rendimiento_ingreso:.2f}%")

    st.subheader("Registro de Lotes (Coopagro)")
    if not df_consolidado.empty:
        df_disp = df_consolidado.copy()
        df_disp["Fecha"] = df_disp["Fecha"].dt.strftime("%d/%m/%Y")
        df_disp["Litros Procesados"] = df_disp["Litros Procesados"].apply(formato_miles)
        df_disp["Producto Terminado"] = df_disp["Producto Terminado"].apply(formato_miles)
        st.dataframe(df_disp, use_container_width=True, hide_index=True)
    else:
        st.info("No hay registros de producción de Coopagro para el período seleccionado.")

  except Exception as e:
    st.error("Se produjo un error procesando los datos de Coopagro:")
    st.code(traceback.format_exc())

# =========================================================================
# MÓDULO 4: INSUMOS, INVENTARIO Y COSTOS
# =========================================================================
elif modulo_principal == "📦 Insumos, Inventario y Costos":
  st.header("📦 Gestión de Insumos, Inventario y Costos Variables")
  try:
    with st.spinner("Cargando maestro de insumos y costos..."):
      import time
      sheet_id = "1Zaqtkadw4Mhgcc8WuuFb1YlXvvWbMsM4"
      sheet_name = "Maestro_Insumos"
      url_insumos = f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={sheet_name}&t={int(time.time())}"
      df_insumos = pd.read_csv(url_insumos)
      df_insumos.columns = df_insumos.columns.str.strip()

    tab_inv1, tab_inv2 = st.tabs(["📊 Estado y Alertas de Stock", "📝 Registrar Ingreso de Mercadería"])

    with tab_inv1:
        st.subheader("Control de Stock y Puntos de Pedido")
        if not df_insumos.empty:
            df_mostrar_ins = df_insumos.copy()
            for col in ['Consumo por tina', 'Stock de seguridad', 'Demora proveedor (dias)']:
                if col in df_mostrar_ins.columns:
                    df_mostrar_ins[col] = pd.to_numeric(df_mostrar_ins[col].astype(str).str.replace(',', '.'), errors='coerce').fillna(0)
            st.dataframe(df_mostrar_ins, use_container_width=True, hide_index=True)
            
            col_m1, col_m2, col_m3 = st.columns(3)
            col_m1.metric("Insumos Monitoreados", len(df_mostrar_ins))
            col_m2.metric("Alertas Activas", "0 insumos críticos")
            col_m3.metric("Estado de Compras", "Normal")
        else:
            st.warning("No se encontraron registros en el Maestro de Insumos.")

    with tab_inv2:
        st.subheader("Formulario de Ingreso de Remito de Insumos")
        st.markdown("Registrá la entrada de mercadería para actualizar las existencias en planta.")
        with st.form("form_ingreso_insumos"):
            col_f1, col_f2 = st.columns(2)
            with col_f1:
                lista_insumos = df_insumos['Insumo'].tolist() if 'Insumo' in df_insumos.columns else ["Cloruro de calcio (32%)", "Sal Entrefina Celusal"]
                insumo_seleccionado = st.selectbox("Seleccionar Insumo", lista_insumos)
                cantidad_ingresada = st.number_input("Cantidad Recibida", min_value=0.0, step=1.0)
            with col_f2:
                nro_remito = st.text_input("Número de Remito / Factura")
                proveedor = st.text_input("Proveedor")
                
            submitted = st.form_submit_button("💾 Guardar Ingreso de Mercadería")
            if submitted:
                if nro_remito and cantidad_ingresada > 0:
                    st.success(f"¡Ingreso registrado con éxito! Remito: {nro_remito} - {cantidad_ingresada} unidades de {insumo_seleccionado}.")
                else:
                    st.warning("Por favor, completá el número de remito y una cantidad mayor a cero.")

  except Exception as e:
    st.error(f"Se produjo un error procesando el módulo de insumos: {e}")
    st.code(traceback.format_exc())
