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
FILE_ID_MASTELLONE = "TU_FILE_ID_DATOS_MHSA_AQUI"  # Reemplazar con ID real de Datos MHSA.xlsx
ID_PRODUCCION = "1wuIpzYmVuflX_pWoPt4Pz9olWF4LLKOf"  # ID de producción general

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


def extraer_fecha_texto(texto) -> pd.Timestamp:
  if pd.isna(texto):
    return pd.NaT
  s = str(texto).strip()
  match_compacto = REGEX_COMPACTO.search(s)
  if match_compacto:
    d, m, a = match_compacto.groups()
    try:
      return pd.to_datetime(f"{a}-{m}-{d}").normalize()
    except ValueError:
      pass
  match = REGEX_FECHA.search(s)
  if match:
    d, m, a = match.groups()
    try:
      return pd.to_datetime(f"{a}-{m}-{d}").normalize()
    except ValueError:
      pass
  return pd.NaT


def calcular_promedio_ponderado(
    df: pd.DataFrame, columna_valor: str, columna_peso: str = "Litros_Ticket"
) -> float:
  if columna_valor not in df.columns or columna_peso not in df.columns:
    return float("nan")
  df_valido = df[[columna_valor, columna_peso]].dropna()
  peso_total = df_valido[columna_peso].sum()
  if peso_total == 0:
    return float("nan")
  return (
      df_valido[columna_valor] * df_valido[columna_peso]
  ).sum() / peso_total


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


def generar_pdf_base(
    titulo: str,
    subtitulo: str,
    metricas: list,
    headers: list,
    df_datos: pd.DataFrame,
    filas_mapeo: list,
    usable_width: int = 190,
):
  pdf = FPDF(orientation="P", unit="mm", format="A4")
  pdf.set_auto_page_break(auto=True, margin=15)
  pdf.add_page()

  ruta_logo = "logo.png"
  if os.path.exists(ruta_logo):
    pdf.image(ruta_logo, x=65, y=10, w=80)
    pdf.set_y(52)
  else:
    pdf.set_y(15)

  pdf.set_font("Arial", "B", 12)
  pdf.set_text_color(100, 100, 100)
  pdf.cell(0, 6, titulo, ln=True, align="C")
  pdf.set_text_color(0, 0, 0)
  pdf.ln(4)
  pdf.line(10, pdf.get_y(), 200, pdf.get_y())
  pdf.ln(6)

  pdf.set_font("Arial", "B", 11)
  pdf.cell(0, 7, subtitulo, ln=True)
  pdf.set_font("Arial", "", 10)
  for metrica in metricas:
    pdf.cell(0, 6, metrica, ln=True)

  pdf.ln(6)
  pdf.set_font("Arial", "B", 9 if len(headers) < 5 else 8)
  pdf.set_fill_color(200, 220, 255)

  suma_anchos = sum([w for _, w in headers])
  factor = usable_width / suma_anchos if suma_anchos > 0 else 1.0
  headers_ajustados = [(name, w * factor) for name, w in headers]

  for i, (col_name, col_w) in enumerate(headers_ajustados):
    pdf.cell(
        col_w,
        8,
        col_name,
        1,
        1 if i == len(headers_ajustados) - 1 else 0,
        "C",
        fill=True,
    )

  pdf.set_font("Arial", "", 9 if len(headers) < 5 else 8)
  for row in df_datos.itertuples(index=False):
    for i, (fn_mapeo, (_, col_w)) in enumerate(
        zip(filas_mapeo, headers_ajustados)
    ):
      val = fn_mapeo(row)
      align = "L" if "Nombre" in headers_ajustados[i][0] else "C"
      pdf.cell(
          col_w,
          7,
          str(val),
          1,
          1 if i == len(headers_ajustados) - 1 else 0,
          align,
      )

  output = pdf.output(dest="S")
  return bytes(output) if not isinstance(output, bytes) else output


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
  # (Mantiene toda la lógica de Coopagro intacta)
  st.markdown(
      '<p class="main-header">🥛 Módulo de Recepción y Calidad - Coopagro</p>',
      unsafe_allow_html=True,
  )
  st.info(
      "Sección activa de Coopagro. Utilice el menú lateral para navegar entre"
      " paneles y reportes."
  )


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
      # 1. Leer Producción General para aislar Mastellone (código 840)
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

      # Filtramos estrictamente el grupo Mastellone (Muzzarella Exportación Mastellone)[cite: 3]
      df_mastellone_prod = df_prod[df_prod["Grupo"] == "Mastellone"].copy()

      # 2. Leer Litros Mes desde Datos MHSA.xlsx (Solapa 'litros mes')
      df_mast_litros = cargar_datos_mastellone(URL_MASTELLONE)

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

    # Filtrar datos de producción
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

    # Obtenemos los litros ingresados desde la tabla de la solapa 'litros mes'
    total_litros_ingresados = 0.0
    if not df_mast_litros.empty:
      # Suponemos que la tabla tiene una columna con litros o sumamos dinámicamente
      # (Podés ajustar la columna según el nombre exacto que tenga en tu solapa 'litros mes')
      col_litros = next(
          (
              c
              for c in df_mast_litros.columns
              if any(x in c.lower() for x in ["litro", "volumen", "cantidad"])
          ),
          df_mast_litros.columns[-1],
      )
      total_litros_ingresados = (
          pd.to_numeric(df_mast_litros[col_litros], errors="coerce").sum()
      )

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
    g1.metric("Litros Ingresados (Drive)", formato_miles(total_litros_ingresados))
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
      '<p class="main-header">Gestión de Insumos y Costos Variables</p>',
      unsafe_allow_html=True,
  )
  st.info("Módulo en desarrollo para control de stock y costos.")
