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

# --- IDs de Google Drive ---
FILE_ID_REMITOS = "16Uh0EwP8tyW79TfJlvcjE8li5Lc6RSLj"
FILE_ID_LAB = "1NNYjM5Aqg9iDdJ85UoALRim8P2A1kaUD"
FILE_ID_BACSOMATIC = "1KeTle24zxjK-clKAuXsAOUzGkfBNXgI8"
FILE_ID_MASTELLONE = "1zNk6whrwaFucv0d7Vkab5rkJduIHzeAg"
ID_PRODUCCION = "1wuIpzYmVuflX_pWoPt4Pz9olWF4LLKOf"

URL_REMITOS = (
    f"https://drive.google.com/uc?export=download&id={FILE_ID_REMITOS}"
)
URL_LAB = f"https://drive.google.com/uc?export=download&id={FILE_ID_LAB}"
URL_BACSOMATIC = (
    f"https://drive.google.com/uc?export=download&id={FILE_ID_BACSOMATIC}"
)
URL_MASTELLONE = (
    f"https://drive.google.com/uc?export=download&id={FILE_ID_MASTELLONE}"
)
URL_PRODUCCION = f"https://drive.google.com/uc?export=download&id={ID_PRODUCCION}"

MESES_ES = {
    1: "Enero",
    2: "Febrero",
    3: "Marzo",
    4: "Abril",
    5: "Mayo",
    6: "Junio",
    7: "Julio",
    8: "Agosto",
    9: "Septiembre",
    10: "Octubre",
    11: "Noviembre",
    12: "Diciembre",
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
  df_remitos_raw, df_contactos, df_lab, df_bacsomatic = (
      pd.DataFrame(),
      pd.DataFrame(),
      pd.DataFrame(),
      pd.DataFrame(),
  )
  try:
    xls_remitos = pd.ExcelFile(u_remitos)
    sheet_remitos = next(
        (s for s in xls_remitos.sheet_names if "od-pro-03" in s.lower()),
        xls_remitos.sheet_names[0],
    )
    sheet_contactos = next(
        (
            s
            for s in xls_remitos.sheet_names
            if "codigo tambo" in s.lower().replace("ó", "o")
        ),
        None,
    )
    df_remitos_raw = pd.read_excel(
        u_remitos, sheet_name=sheet_remitos, skiprows=4, usecols="B:K"
    )
    if sheet_contactos:
      df_contactos = pd.read_excel(u_remitos, sheet_name=sheet_contactos)
  except Exception as e:
    st.sidebar.warning(f"Error cargando remitos/contactos: {e}")

  try:
    df_lab_temp = pd.read_excel(u_lab, header=None, nrows=20)
    header_row = encontrar_fila_encabezado(
        df_lab_temp, ["sample", "fat", "protein", "grasa"]
    )
    df_lab = pd.read_excel(u_lab, header=header_row)
    df_lab.columns = df_lab.columns.astype(str).str.strip()
  except Exception as e:
    st.sidebar.warning(f"No se pudo cargar el archivo Milko: {e}")

  try:
    df_bac_temp = pd.read_excel(u_bacsomatic, header=None, nrows=20)
    header_row_bac = encontrar_fila_encabezado(
        df_bac_temp, ["id usuario", "ufc", "scc"]
    )
    df_bacsomatic = pd.read_excel(u_bacsomatic, header=header_row_bac)
    df_bacsomatic.columns = df_bacsomatic.columns.astype(str).str.strip()
  except Exception as e:
    st.sidebar.warning(f"No se pudo cargar el archivo Bacsomatic: {e}")

  return df_remitos_raw, df_contactos, df_lab, df_bacsomatic


@st.cache_data(ttl=60, show_spinner="Descargando datos de Mastellone...")
def cargar_datos_mastellone(url_mastellone):
  try:
    xls = pd.ExcelFile(url_mastellone)
    sheet_name = next(
        (s for s in xls.sheet_names if "litro" in s.lower() or "mes" in s.lower()),
        xls.sheet_names[1] if len(xls.sheet_names) > 1 else xls.sheet_names[0],
    )
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
  if pd.isna(val):
    return ""
  s = str(val).strip().upper()
  if s.endswith(".0"):
    s = s[:-2]
  if s and s[0].isdigit():
    return f"T{s}"
  return s


def procesar_lote_mastellone(lote_str):
  if not isinstance(lote_str, str) or len(lote_str) < 8:
    return "Desconocido", "Desconocido"
  prod_code = lote_str[5:8]
  mapping_prod = {
      "288": "Muzzarella Exportacion Coop.",
      "125": "Muzarrella Piano",
      "488": "Tybo Coop.",
      "840": "Muzzarella Exportacion Mastellone",
  }
  mapping_grupo = {
      "288": "Coopagro",
      "125": "Coopagro",
      "488": "Coopagro",
      "840": "Mastellone",
  }
  return mapping_prod.get(
      prod_code, f"Desconocido ({prod_code})"
  ), mapping_grupo.get(prod_code, "Otro")


# --- MENÚ DE NAVEGACIÓN PRINCIPAL DE LA SUPER APP ---
with st.sidebar:
  st.image("https://cdn-icons-png.flaticon.com/512/2830/2830305.png", width=60)
  st.title("Gestión Planta Tandil")
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
# MÓDULO 1: RECEPCIÓN Y CALIDAD COOPAGRO
# =========================================================================
if modulo_principal == "🥛 Recepción y Calidad Coopagro":
  st.markdown(
      '<p class="main-header">🥛 Módulo de Recepción y Calidad - Coopagro</p>',
      unsafe_allow_html=True,
  )
  st.info("Sección activa de Coopagro.")


# =========================================================================
# MÓDULO 2: RECEPCIÓN Y PRODUCCIÓN MASTELLONE (FASÓN)
# =========================================================================
elif modulo_principal == "🚛 Recepción Mastellone (Fasón)":
  st.markdown(
      '<p class="main-header">🏭 Reporte de Producción y Recepción — Fasón'
      " Mastellone</p>",
      unsafe_allow_html=True,
  )

  try:
    with st.spinner("Sincronizando datos de Mastellone desde Google Drive..."):
      # 1. Leer Producción General para aislar Mastellone (código 840)[cite: 3]
      raw_prod = pd.read_excel(URL_PRODUCCION, skiprows=6)
      df_prod = pd.DataFrame()
      df_prod["Fecha"] = raw_prod.iloc[:, 0]
      df_prod["Lote"] = raw_prod.iloc[:, 1]
      df_prod["Litros Procesados"] = raw_prod.iloc[:, 3]
      df_prod["Producto Terminado"] = raw_prod.iloc[:, 5]
      df_prod["PNC"] = raw_prod.iloc[:, 6]

      df_prod = df_prod.dropna(subset=["Fecha"])
      df_prod["Fecha"] = pd.to_datetime(
          df_prod["Fecha"], dayfirst=True, errors="coerce"
      )
      df_prod = df_prod.dropna(subset=["Fecha"])

      for col in ["Litros Procesados", "Producto Terminado", "PNC"]:
        df_prod[col] = pd.to_numeric(df_prod[col], errors="coerce").fillna(0)

      if len(df_prod) > 0:
        df_prod["Producto"], df_prod["Grupo"] = zip(
            *df_prod["Lote"].astype(str).apply(procesar_lote_mastellone)
        )
      else:
        df_prod["Producto"] = []
        df_prod["Grupo"] = []

      df_prod["Año"] = df_prod["Fecha"].dt.year
      df_prod["Mes"] = df_prod["Fecha"].dt.month

      # Filtramos estrictamente Mastellone[cite: 3]
      df_mastellone_prod = df_prod[df_prod["Grupo"] == "Mastellone"].copy()

      # 2. Leer Litros Mes desde Datos MHSA.xlsx (Solapa 'litros mes')
      df_mast_litros = cargar_datos_mastellone(URL_MASTELLONE)

      # Procesar fechas y litros del archivo de Mastellone para que cruce por mes y año
      if not df_mast_litros.empty:
        # Detectar columna de fecha automáticamente en la solapa 'litros mes'
        col_fecha_mast = next(
            (
                c
                for c in df_mast_litros.columns
                if any(x in c.lower() for x in ["fecha", "mes", "periodo", "date"])
            ),
            df_mast_litros.columns[0],
        )
        # Detectar columna de litros
        col_litros_mast = next(
            (
                c
                for c in df_mast_litros.columns
                if any(x in c.lower() for x in ["litro", "volumen", "cantidad"])
            ),
            df_mast_litros.columns[-1],
        )

        df_mast_litros["Fecha_Parsed"] = pd.to_datetime(
            df_mast_litros[col_fecha_mast], errors="coerce"
        )
        df_mast_litros["Año"] = df_mast_litros["Fecha_Parsed"].dt.year
        df_mast_litros["Mes"] = df_mast_litros["Fecha_Parsed"].dt.month
        df_mast_litros["Litros_Ingresados"] = pd.to_numeric(
            df_mast_litros[col_litros_mast], errors="coerce"
        ).fillna(0)

    # --- BARRA LATERAL (FILTROS) ---
    st.sidebar.markdown("### 🔍 Filtros Mastellone")
    opciones_anio = ["Todos"] + (
        sorted(df_mastellone_prod["Año"].unique().tolist())
        if len(df_mastellone_prod) > 0
        else []
    )
    opciones_mes = ["Todos"] + (
        sorted(df_mastellone_prod["Mes"].unique().tolist())
        if len(df_mastellone_prod) > 0
        else []
    )

    with st.sidebar.container():
      filtro_anio = st.selectbox("📅 Seleccionar Año", opciones_anio, key="m_anio")
      filtro_mes = st.selectbox("📆 Seleccionar Mes", opciones_mes, key="m_mes")

    # Filtrar datos de producción por año y mes
    df_filtrado = df_mastellone_prod.copy()
    if len(df_filtrado) > 0:
      if filtro_anio != "Todos":
        df_filtrado = df_filtrado[df_filtrado["Año"] == filtro_anio]
      if filtro_mes != "Todos":
        df_filtrado = df_filtrado[df_filtrado["Mes"] == filtro_mes]

    df_consolidado_raw = df_filtrado.copy()
    if len(df_consolidado_raw) > 0:
      df_consolidado_raw["PT_Total"] = (
          df_consolidado_raw["Producto Terminado"]
          + df_consolidado_raw["PNC"]
      )
      df_consolidado_raw["Ratio Consolidado (%)"] = df_consolidado_raw.apply(
          lambda x: (
              f"{(x['PT_Total'] / x['Litros Procesados'] * 100):.2f}%".replace(
                  ".", ","
              )
              if x["Litros Procesados"] > 0
              else "0,00%"
          ),
          axis=1,
      )

      df_consolidado = pd.DataFrame()
      df_consolidado["Fecha"] = df_consolidado_raw["Fecha"]
      df_consolidado["Lote"] = df_consolidado_raw["Lote"]
      df_consolidado["Producto"] = df_consolidado_raw["Producto"]
      df_consolidado["Litros Procesados"] = df_consolidado_raw[
          "Litros Procesados"
      ]
      df_consolidado["Producto Terminado"] = df_consolidado_raw["PT_Total"]
      df_consolidado["Ratio de Conversión (%)"] = df_consolidado_raw[
          "Ratio Consolidado (%)"
      ]
    else:
      df_consolidado = pd.DataFrame(
          columns=[
              "Fecha",
              "Lote",
              "Producto",
              "Litros Procesados",
              "Producto Terminado",
              "Ratio de Conversión (%)",
          ]
      )

    # --- OBTENER LITROS INGRESADOS FILTRADOS POR AÑO Y MES ---
    total_litros_ingresados = 0.0
    if not df_mast_litros.empty and "Litros_Ingresados" in df_mast_litros.columns:
      df_litros_filtrados = df_mast_litros.copy()
      if filtro_anio != "Todos":
        df_litros_filtrados = df_litros_filtrados[
            df_litros_filtrados["Año"] == filtro_anio
        ]
      if filtro_mes != "Todos":
        df_litros_filtrados = df_litros_filtrados[
            df_litros_filtrados["Mes"] == filtro_mes
        ]
      total_litros_ingresados = df_litros_filtrados[
          "Litros_Ingresados"
      ].sum()

    total_litros_proc = (
        df_filtrado["Litros Procesados"].sum()
        if len(df_filtrado) > 0
        else 0
    )
    total_prod_consolidado = (
        df_consolidado["Producto Terminado"].sum()
        if len(df_consolidado) > 0
        else 0
    )

    ratio_ponderado = (
        (total_prod_consolidado / total_litros_proc * 100)
        if total_litros_proc > 0
        else 0
    )
    rendimiento_ingreso = (
        (total_prod_consolidado / total_litros_ingresados * 100)
        if total_litros_ingresados > 0
        else 0
    )

    # --- INTERFAZ PRINCIPAL DE MASTELLONE ---
    st.markdown("### 📈 Indicadores Consolidados - Mastellone")
    g1, g2, g3 = st.columns(3)
    g1.metric("Litros Ingresados", formato_miles(total_litros_ingresados))
    g2.metric("Litros Procesados", formato_miles(total_litros_proc))
    g3.metric(
        "Total Producto Terminado", formato_miles(total_prod_consolidado)
    )

    g4, g5, _ = st.columns(3)
    g4.metric(
        "Ratio PT / Litros procesados", f"{ratio_ponderado:.2f}%".replace(".", ",")
    )
    g5.metric(
        "Ratio PT / Litros ingresados",
        f"{rendimiento_ingreso:.2f}%".replace(".", ","),
    )

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("### 📋 Registro de Lotes (Muzzarella Mastellone)")
    if len(df_consolidado) > 0:
      df_display = df_consolidado.copy()
      df_display["Litros Procesados"] = df_display["Litros Procesados"].apply(
          formato_miles
      )
      df_display["Producto Terminado"] = df_display["Producto Terminado"].apply(
          formato_miles
      )
      df_display["Fecha"] = df_display["Fecha"].dt.strftime("%d/%m/%Y")
      st.dataframe(
          df_display, use_container_width=True, hide_index=True, height=350
      )
    else:
      st.info("No hay lotes de Mastellone registrados para el período seleccionado.")

  except Exception as e:
    st.error(f"Error al procesar el módulo de Mastellone: {e}")
    st.code(traceback.format_exc())


# =========================================================================
# MÓDULOS 3 y 4
# =========================================================================
elif modulo_principal == "🧀 Producción y Rendimiento":
  st.markdown(
      '<p class="main-header">Registro de Producción y Ecuación de Van'
      " Slyke</p>",
      unsafe_allow_html=True,
  )
  st.info("Módulo en desarrollo para análisis global de rendimiento.")

elif modulo_principal == "📦 Insumos, Inventario y Costos":
  st.markdown(
      '<p class="main-header">Gestión de Insumos y Costos Variables</p>",
      unsafe_allow_html=True,
  )
  st.info("Módulo en desarrollo para control de stock y costos.")
