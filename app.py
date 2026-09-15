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

# --- IDs de Google Drive (Fijos y Configurados) ---
FILE_ID_REMITOS = "16Uh0EwP8tyW79TfJlvcjE8li5Lc6RSLj"
FILE_ID_LAB = "1NNYjM5Aqg9iDdJ85UoALRim8P2A1kaUD"
FILE_ID_BACSOMATIC = "1KeTle24zxjK-clKAuXsAOUzGkfBNXgI8"
FILE_ID_MASTELLONE = "1Zaqtkadw4Mhgcc8WuuFb1YlXvvWbMsM4"
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
    # Buscamos de forma exacta o flexible la solapa 'litros mes'
    sheet_name = next(
        (s for s in xls.sheet_names if "litro" in s.lower() or "mes" in s.lower()),
        xls.sheet_names[0],
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


def generar_pdf_panel_general(
    df_macro,
    periodo_titulo,
    total_litros,
    temp_prom,
    grasa_prom,
    prot_prom,
    ratio_gp,
    tambos_activos,
    df_ranking,
):
  ratio_str = f"{ratio_gp:.2f}".replace(".", ",") if pd.notna(ratio_gp) else "S/D"
  grasa_str = (
      f"{grasa_prom:.2f}%".replace(".", ",") if pd.notna(grasa_prom) else "S/D"
  )
  prot_str = (
      f"{prot_prom:.2f}%".replace(".", ",") if pd.notna(prot_prom) else "S/D"
  )

  metricas = [
      (
          f"Tambos Activos: {tambos_activos} | Litros Totales:"
          f" {formato_miles(total_litros)} L"
      ),
      (
          f"Temp. Promedio: {formato_temp(temp_prom)} | Grasa Ponderada:"
          f" {grasa_str} | Prot. Ponderada: {prot_str}"
      ),
      f"Ratio Grasa / Proteína: {ratio_str}",
      "\nRanking de Tambos por Volumen de Litros",
  ]
  headers = [("Código", 30), ("Nombre del Tambo", 100), ("Litros Totales", 60)]
  mapeo = [
      lambda r: str(r.Num_Tambo),
      lambda r: str(r.Tambo),
      lambda r: formato_miles(r.Litros_Ticket),
  ]
  return generar_pdf_base(
      "Informe de Recolección - Cooperativa",
      f"Período Evaluado: {periodo_titulo}",
      metricas,
      headers,
      df_ranking,
      mapeo,
  )


def generar_pdf_bytes(
    df_productor,
    tambo_nombre,
    tambo_id,
    periodo_texto,
    args_visibles,
    es_mensual=False,
):
  titulo = (
      "Resumen mensual de recolección"
      if es_mensual
      else "Resumen semanal de recolección"
  )
  subtitulo = f"Productor: {tambo_nombre} (Código #{tambo_id})"
  temp_prom = (
      df_productor["Temperatura"].mean()
      if "Temperatura" in df_productor
      else float("nan")
  )

  metricas = [
      f"Período: {periodo_texto}",
      f"Total Litros: {formato_miles(df_productor['Litros_Ticket'].sum())} L",
  ]
  if args_visibles["temp"]:
    metricas.append(f"Temperatura Promedio: {formato_temp(temp_prom)}")

  partes_solidos = []
  if (
      args_visibles["grasa"]
      and "Grasa" in df_productor
      and pd.notna(df_productor["Grasa"].mean())
  ):
    partes_solidos.append(f"Grasa: {df_productor['Grasa'].mean():.2f}%".replace(".", ","))
  if (
      args_visibles["prot"]
      and "Proteina" in df_productor
      and pd.notna(df_productor["Proteina"].mean())
  ):
    partes_solidos.append(
        f"Proteína: {df_productor['Proteina'].mean():.2f}%".replace(".", ",")
    )
  if (
      args_visibles["crios"]
      and "Crioscopia" in df_productor
      and pd.notna(df_productor["Crioscopia"].mean())
  ):
    partes_solidos.append(
        f"Crioscopia: {df_productor['Crioscopia'].mean():.3f}".replace(
            ".", ","
        )
    )
  if (
      args_visibles["ufc"]
      and "UFC" in df_productor
      and pd.notna(df_productor["UFC"].mean())
  ):
    partes_solidos.append(f"UFC: {formato_miles(df_productor['UFC'].mean())}")
  if (
      args_visibles["scc"]
      and "SCC" in df_productor
      and pd.notna(df_productor["SCC"].mean())
  ):
    partes_solidos.append(f"SCC: {formato_miles(df_productor['SCC'].mean())}")

  if partes_solidos:
    metricas.append(f"Promedios Lab -> {' | '.join(partes_solidos)}")

  headers = [("Fecha", 26), ("N° Remito", 34), ("Litros", 30)]
  mapeo = [
      (
          lambda r: getattr(r, "Fecha").strftime("%d/%m/%Y")
          if pd.notna(getattr(r, "Fecha"))
          else ""
      ),
      (
          lambda r: str(getattr(r, "N_Remito"))
          if pd.notna(getattr(r, "N_Remito"))
          else "-"
      ),
      (
          lambda r: formato_miles(getattr(r, "Litros_Ticket"))
          if pd.notna(getattr(r, "Litros_Ticket"))
          else "0"
      ),
  ]

  if args_visibles["temp"]:
    headers.append(("Temp", 18))
    mapeo.append(lambda r: formato_temp(getattr(r, "Temperatura", pd.NA)))
  if args_visibles["grasa"]:
    headers.append(("Grasa", 20))
    mapeo.append(
        lambda r: f"{getattr(r, 'Grasa'):.2f}%".replace(".", ",")
        if pd.notna(getattr(r, "Grasa", pd.NA))
        else "-"
    )
  if args_visibles["prot"]:
    headers.append(("Prot", 20))
    mapeo.append(
        lambda r: f"{getattr(r, 'Proteina'):.2f}%".replace(".", ",")
        if pd.notna(getattr(r, "Proteina", pd.NA))
        else "-"
    )
  if args_visibles["crios"]:
    headers.append(("Crios", 22))
    mapeo.append(
        lambda r: f"{getattr(r, 'Crioscopia'):.3f}".replace(".", ",")
        if pd.notna(getattr(r, "Crioscopia", pd.NA))
        else "-"
    )
  if args_visibles["ufc"]:
    headers.append(("UFC", 22))
    mapeo.append(
        lambda r: formato_miles(getattr(r, "UFC", pd.NA))
        if pd.notna(getattr(r, "UFC", pd.NA))
        else "-"
    )
  if args_visibles["scc"]:
    headers.append(("SCC", 24))
    mapeo.append(
        lambda r: formato_miles(getattr(r, "SCC", pd.NA))
        if pd.notna(getattr(r, "SCC", pd.NA))
        else "-"
    )

  return generar_pdf_base(titulo, subtitulo, metricas, headers, df_productor, mapeo)


def enviar_correo_productor(
    destinatario_email,
    nombre_contacto,
    tambo_nombre,
    pdf_bytes,
    nombre_archivo,
    tipo_reporte="semanal",
) -> bool:
  try:
    remitente = st.secrets["email"]["remitente"]
    password = st.secrets["email"]["password"]
    msg = MIMEMultipart()
    msg["From"] = remitente
    destinatarios = [
        e.strip()
        for e in destinatario_email.replace(";", ",").split(",")
        if e.strip()
    ]
    msg["To"] = ", ".join(destinatarios)
    msg["Subject"] = (
        f"Resumen {tipo_reporte.capitalize()} de Recolección - {tambo_nombre}"
    )
    cuerpo_html = (
        f"<html><body><p>Buenas tardes, <b>{nombre_contacto}</b>:</p><p>Le"
        f" adjunto el resumen {tipo_reporte} de recolección y calidad de"
        " leche.</p></body></html>"
    )
    msg.attach(MIMEText(cuerpo_html, "html"))
    part = MIMEBase("application", "octet-stream")
    part.set_payload(pdf_bytes)
    encoders.encode_base64(part)
    part.add_header(
        "Content-Disposition", f'attachment; filename="{nombre_archivo}"'
    )
    msg.attach(part)
    with smtplib.SMTP("smtp.gmail.com", 587) as server:
      server.starttls()
      server.login(remitente, password)
      server.sendmail(remitente, destinatarios, msg.as_string())
    return True
  except Exception as e:
    st.error(f"Error de envío SMTP: {e}")
    return False


# --- MENÚ DE NAVEGACIÓN PRINCIPAL DE LA SUPER APP ---
with st.sidebar:
  if os.path.exists("logo.png"):
    st.image("logo.png", width=300)  # Logo más grande y proporcional
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
    df_raw, df_contactos_raw, df_lab_raw, df_bac_raw = cargar_datos_coopagro(
        URL_REMITOS, URL_LAB, URL_BACSOMATIC
    )

    if df_raw.empty:
      st.error("El archivo de remitos está vacío o no se pudo acceder.")
      st.stop()

    df_contactos = pd.DataFrame()
    if not df_contactos_raw.empty:
      df_c_temp = df_contactos_raw.copy()
      df_c_temp.columns = (
          df_c_temp.columns.astype(str)
          .str.strip()
          .str.lower()
          .str.replace("ó", "o")
      )
      col_codigo = next(
          (c for c in df_c_temp.columns if "codigo" in c and "viejo" not in c),
          None,
      )
      if not col_codigo and len(df_c_temp.columns) > 1:
        col_codigo = df_c_temp.columns[1]
      col_contacto = next(
          (c for c in df_c_temp.columns if "contacto" in c or "nombre" in c),
          None,
      )
      if not col_contacto and len(df_c_temp.columns) > 3:
        col_contacto = df_c_temp.columns[3]
      col_email = next(
          (c for c in df_c_temp.columns if "email" in c or "correo" in c), None
      )
      if not col_email and len(df_c_temp.columns) > 4:
        col_email = df_c_temp.columns[4]

      if (
          col_codigo is not None
          and col_contacto is not None
          and col_email is not None
      ):
        df_contactos["Num_Tambo"] = df_c_temp[col_codigo].apply(limpiar_tambo)
        df_contactos["Contacto_Nombre"] = df_c_temp[col_contacto]
        df_contactos["Email"] = df_c_temp[col_email]

    df = df_raw.iloc[:, :10].copy()
    df.columns = [
        "Fecha",
        "N_Remito",
        "Num_Tambo",
        "Tambo",
        "Litros_Ticket",
        "Litros_Planilla",
        "Diferencia",
        "Temperatura",
        "Grasa",
        "Proteina",
    ]
    df["Num_Tambo"] = df["Num_Tambo"].apply(limpiar_tambo)
    df["Fecha"] = pd.to_datetime(df["Fecha"], errors="coerce").dt.normalize()
    df = df.dropna(subset=["Fecha"])

    for col in ["Grasa", "Proteina", "Crioscopia", "UFC", "SCC"]:
      if col not in df.columns:
        df[col] = pd.NA

    df = df.sort_values(by=["Num_Tambo", "Fecha", "N_Remito"])
    df["orden_remito"] = df.groupby(["Num_Tambo", "Fecha"]).cumcount() + 1
    df["merge_tambo"] = df["Num_Tambo"].astype(str)

    # Procesamiento MilkoScan
    if not df_lab_raw.empty:
      df_lab = df_lab_raw.copy()
      col_sample = next(
          (
              c
              for c in df_lab.columns
              if any(
                  x in c.lower() for x in ["sample", "number", "tambo", "muestra"]
              )
          ),
          df_lab.columns[0],
      )
      col_date = next(
          (
              c
              for c in df_lab.columns
              if any(x in c.lower() for x in ["fecha", "date", "time"])
          ),
          None,
      )
      df_lab["Num_Tambo"] = (
          df_lab[col_sample].astype(str).str.split().str[0].apply(limpiar_tambo)
      )
      if col_date and df_lab[col_date].notna().any():
        df_lab["Fecha"] = (
            pd.to_datetime(df_lab[col_date], errors="coerce").dt.normalize()
        )
      else:
        df_lab["Fecha"] = df_lab[col_sample].apply(
            lambda x: extraer_fecha_texto(x) if pd.notna(x) else pd.NaT
        )
      df_lab = df_lab.dropna(subset=["Fecha", "Num_Tambo"]).sort_values(
          by=["Num_Tambo", "Fecha"]
      )
      df_lab["orden_remito"] = (
          df_lab.groupby(["Num_Tambo", "Fecha"]).cumcount() + 1
      )
      col_fat = next(
          (
              c
              for c in df_lab.columns
              if "fat" in c.lower() or "grasa" in c.lower()
          ),
          None,
      )
      col_prot = next(
          (
              c
              for c in df_lab.columns
              if "protein" in c.lower() or "proteina" in c.lower()
          ),
          None,
      )
      col_fp = next(
          (
              c
              for c in df_lab.columns
              if c.lower() == "fp" or "crios" in c.lower()
          ),
          None,
      )

      map_cols = {}
      if col_fat:
        map_cols[col_fat] = "Grasa_Lab"
      if col_prot:
        map_cols[col_prot] = "Proteina_Lab"
      if col_fp:
        map_cols[col_fp] = "Crioscopia_Lab"

      if map_cols:
        df_milko_clean = df_lab[
            ["Num_Tambo", "Fecha", "orden_remito"] + list(map_cols.keys())
        ].rename(columns=map_cols)
        for c in map_cols.values():
          df_milko_clean[c] = pd.to_numeric(
              df_milko_clean[c].astype(str).str.replace(",", "."),
              errors="coerce",
          )
        df = pd.merge(
            df,
            df_milko_clean,
            on=["Num_Tambo", "Fecha", "orden_remito"],
            how="left",
        )
        if "Grasa_Lab" in df:
          df["Grasa"] = df["Grasa_Lab"].combine_first(df["Grasa"])
        if "Proteina_Lab" in df:
          df["Proteina"] = df["Proteina_Lab"].combine_first(df["Proteina"])
        if "Crioscopia_Lab" in df:
          df["Crioscopia"] = df["Crioscopia_Lab"].combine_first(
              df["Crioscopia"]
          )

    # Procesamiento Bacsomatic
    if not df_bac_raw.empty:
      df_bac = df_bac_raw.copy()
      col_id = next(
          (
              c
              for c in df_bac.columns
              if any(x in c.lower() for x in ["id usuario", "sample", "tambo"])
          ),
          df_bac.columns[0],
      )
      df_bac["Num_Tambo"] = (
          df_bac[col_id].astype(str).str.split().str[0].apply(limpiar_tambo)
      )
      df_bac["Fecha"] = df_bac[col_id].apply(
          lambda x: extraer_fecha_texto(x) if pd.notna(x) else pd.NaT
      )
      df_bac = df_bac.dropna(subset=["Fecha", "Num_Tambo"]).sort_values(
          by=["Num_Tambo", "Fecha"]
      )
      df_bac["orden_remito"] = (
          df_bac.groupby(["Num_Tambo", "Fecha"]).cumcount() + 1
      )
      col_ufc = next(
          (c for c in df_bac.columns if "ufc" in c.lower()),
          None,
      )
      col_scc = next(
          (
              c
              for c in df_bac.columns
              if any(x in c.lower() for x in ["scc", "celulas", "somáticas"])
          ),
          None,
      )
      map_cols_bac = {}
      if col_ufc:
        map_cols_bac[col_ufc] = "UFC_Val"
      if col_scc:
        map_cols_bac[col_scc] = "SCC_Val"

      if map_cols_bac:
        df_bac_clean = df_bac[
            ["Num_Tambo", "Fecha", "orden_remito"] + list(map_cols_bac.keys())
        ].rename(columns=map_cols_bac)
        for c in map_cols_bac.values():
          df_bac_clean[c] = pd.to_numeric(
              df_bac_clean[c].astype(str).str.replace(",", "."),
              errors="coerce",
          )
        df = pd.merge(
            df,
            df_bac_clean,
            on=["Num_Tambo", "Fecha", "orden_remito"],
            how="left",
        )
        if "UFC_Val" in df:
          df["UFC"] = df["UFC_Val"].combine_first(df["UFC"])
        if "SCC_Val" in df:
          df["SCC"] = df["SCC_Val"].combine_first(df["SCC"])

    # Ciclos semanales y mensuales
    df["Fecha_Cierre_Viernes"] = df["Fecha"] + pd.to_timedelta(
        (4 - df["Fecha"].dt.weekday) % 7, unit="D"
    )
    df["Fecha_Inicio_Sabado"] = df["Fecha_Cierre_Viernes"] - pd.Timedelta(days=6)
    df["Ciclo_Semana"] = (
        "Viernes "
        + df["Fecha_Cierre_Viernes"].dt.strftime("%d/%m/%Y")
        + " (Sáb "
        + df["Fecha_Inicio_Sabado"].dt.strftime("%d/%m/%Y")
        + " al Vie "
        + df["Fecha_Cierre_Viernes"].dt.strftime("%d/%m/%Y")
        + ")"
    )
    df["AnioMes"] = df["Fecha"].dt.to_period("M")

    # Sub-navegación interna de Coopagro
    st.sidebar.markdown("---")
    vista_coop = st.sidebar.radio(
        "Sección Coopagro:",
        [
            "Panel de Control General",
            "Gestión y Reportes por Tambo",
            "Envío Masivo Semanal",
        ],
    )

    if vista_coop == "Panel de Control General":
      tipo_reporte_opcion = st.sidebar.radio(
          "Seleccione el período:", ["Semanal", "Mensual"]
      )
      st.markdown(
          '<p class="main-header">📊 Panel de Control General - Coopagro</p>',
          unsafe_allow_html=True,
      )

      if tipo_reporte_opcion == "Semanal":
        ciclos = (
            df[["Fecha_Cierre_Viernes", "Ciclo_Semana"]]
            .drop_duplicates()
            .sort_values("Fecha_Cierre_Viernes", ascending=False)["Ciclo_Semana"]
            .tolist()
        )
        ciclo_gen = (
            st.sidebar.selectbox("Seleccione el Cierre de Semana:", ciclos)
            if ciclos
            else ""
        )
        df_macro = df[df["Ciclo_Semana"] == ciclo_gen] if ciclos else pd.DataFrame()
        periodo_texto = ciclo_gen
      else:
        meses = sorted(df["AnioMes"].unique(), reverse=True)
        mes_gen = (
            st.sidebar.selectbox(
                "Seleccione el Mes:",
                meses,
                format_func=lambda p: f"{MESES_ES.get(p.month)} {p.year}",
            )
            if meses
            else None
        )
        df_macro = df[df["AnioMes"] == mes_gen] if meses else pd.DataFrame()
        periodo_texto = (
            f"{MESES_ES.get(mes_gen.month)} {mes_gen.year}" if mes_gen else ""
        )

      if not df_macro.empty:
        tot_litros = df_macro["Litros_Ticket"].sum()
        grasa_p = calcular_promedio_ponderado(df_macro, "Grasa")
        prot_p = calcular_promedio_ponderado(df_macro, "Proteina")
        ratio_gp = grasa_p / prot_p if (prot_p and prot_p > 0) else pd.NA

        mc1, mc2, mc3, mc4, mc5 = st.columns(5)
        mc1.metric("🥛 Litros Totales", f"{formato_miles(tot_litros)} L")
        mc2.metric(
            "🌡️ Temp. Media", formato_temp(df_macro["Temperatura"].mean())
        )
        mc3.metric(
            "🧈 Grasa Ponderada",
            f"{grasa_p:.2f}%".replace(".", ",") if pd.notna(grasa_p) else "S/D",
        )
        mc4.metric(
            "🧀 Proteína Ponderada",
            f"{prot_p:.2f}%".replace(".", ",") if pd.notna(prot_p) else "S/D",
        )
        mc5.metric(
            "⚖️ Ratio Grasa/Prot.",
            f"{ratio_gp:.2f}".replace(".", ",") if pd.notna(ratio_gp) else "S/D",
        )

        df_ranking = (
            df_macro.groupby(["Tambo", "Num_Tambo"], as_index=False)[
                "Litros_Ticket"
            ]
            .sum()
            .sort_values("Litros_Ticket", ascending=False)
        )
        df_ranking_show = df_ranking.copy()
        df_ranking_show["Litros_Ticket"] = df_ranking_show[
            "Litros_Ticket"
        ].apply(formato_miles)
        st.dataframe(
            df_ranking_show.rename(
                columns={
                    "Tambo": "Nombre del Tambo",
                    "Num_Tambo": "Código",
                    "Litros_Ticket": "Litros Totales",
                }
            ),
            hide_index=True,
            use_container_width=True,
        )

        pdf_bytes = generar_pdf_panel_general(
            df_macro,
            periodo_texto,
            tot_litros,
            df_macro["Temperatura"].mean(),
            grasa_p,
            prot_p,
            ratio_gp,
            df_macro["Num_Tambo"].nunique(),
            df_ranking,
        )
        st.download_button(
            "📥 Descargar Informe en PDF",
            data=pdf_bytes,
            file_name=f"Informe_{periodo_texto.replace(' ', '_')}.pdf",
            mime="application/pdf",
        )
      else:
        st.info("No hay datos para el período seleccionado.")

    elif vista_coop == "Gestión y Reportes por Tambo":
      tipo_reporte_opcion = st.sidebar.radio(
          "Seleccione el período:", ["Semanal", "Mensual"]
      )
      mapeo_tambos = (
          df[["Tambo", "Num_Tambo"]].drop_duplicates().sort_values("Tambo")
      )
      if not mapeo_tambos.empty:
        t_nombre = st.sidebar.selectbox(
            "1. Seleccione el Tambo:", mapeo_tambos["Tambo"].tolist()
        )
        t_id = mapeo_tambos.loc[
            mapeo_tambos["Tambo"] == t_nombre, "Num_Tambo"
        ].values[0]
        df_t = df[df["Num_Tambo"] == str(t_id)]

        if tipo_reporte_opcion == "Semanal":
          ciclos = (
              df_t[["Fecha_Cierre_Viernes", "Ciclo_Semana"]]
              .drop_duplicates()
              .sort_values("Fecha_Cierre_Viernes", ascending=False)[
                  "Ciclo_Semana"
              ]
              .tolist()
          )
          ciclo_sel = (
              st.sidebar.selectbox("2. Cierre de Semana:", ciclos)
              if ciclos
              else ""
          )
          df_per = (
              df_t[df_t["Ciclo_Semana"] == ciclo_sel].sort_values(
                  ["Fecha", "N_Remito"]
              )
              if ciclos
              else pd.DataFrame()
          )
          es_mensual = False
          periodo_pdf = (
              (
                  f"{df_per['Fecha_Inicio_Sabado'].iloc[0]:%d/%m/%Y} al"
                  f" {df_per['Fecha_Cierre_Viernes'].iloc[0]:%d/%m/%Y}"
              )
              if not df_per.empty
              else ""
          )
        else:
          meses = sorted(df_t["AnioMes"].unique(), reverse=True)
          mes_sel = (
              st.sidebar.selectbox(
                  "2. Mes:",
                  meses,
                  format_func=lambda p: f"{MESES_ES.get(p.month)} {p.year}",
              )
              if meses
              else None
          )
          df_per = (
              df_t[df_t["AnioMes"] == mes_sel].sort_values(
                  ["Fecha", "N_Remito"]
              )
              if mes_sel
              else pd.DataFrame()
          )
          es_mensual = True
          periodo_pdf = (
              f"{MESES_ES.get(mes_sel.month)} {mes_sel.year}"
              if mes_sel
              else ""
          )

        st.sidebar.subheader("⚙️ Elementos del Reporte")
        v_temp = st.sidebar.checkbox("Temperatura", True)
        v_grasa = st.sidebar.checkbox("Grasa", True)
        v_prot = st.sidebar.checkbox("Proteína", True)
        v_crios = st.sidebar.checkbox("Crioscopia", True)
        v_ufc = st.sidebar.checkbox("UFC", True)
        v_scc = st.sidebar.checkbox("SCC", True)
        args_vis = {
            "temp": v_temp,
            "grasa": v_grasa,
            "prot": v_prot,
            "crios": v_crios,
            "ufc": v_ufc,
            "scc": v_scc,
        }

        if not df_per.empty:
          st.markdown(
              f'<p class="main-header">Resumen'
              f" {'Mensual' if es_mensual else 'Semanal'} ({periodo_pdf}) -"
              f" {t_nombre} (#{t_id})</p>",
              unsafe_allow_html=True,
          )
          info_c = df_contactos[df_contactos["Num_Tambo"] == str(t_id)]
          email_t = (
              info_c["Email"].values[0]
              if not info_c.empty and pd.notna(info_c["Email"].values[0])
              else ""
          )
          nom_c = (
              info_c["Contacto_Nombre"].values[0]
              if not info_c.empty
              and pd.notna(info_c["Contacto_Nombre"].values[0])
              else "Productor"
          )
          l_act = df_per["Litros_Ticket"].sum()

          cols = st.columns(1 + sum(args_vis.values()))
          cols[0].metric("Litros", f"{formato_miles(l_act)} L")
          idx = 1
          if v_temp:
            cols[idx].metric(
                "Temp. Prom", formato_temp(df_per["Temperatura"].mean())
            )
            idx += 1
          if v_grasa:
            cols[idx].metric(
                "Grasa Prom",
                f"{df_per['Grasa'].mean():.2f}%".replace(".", ",")
                if pd.notna(df_per["Grasa"].mean())
                else "S/D",
            )
            idx += 1
          if v_prot:
            cols[idx].metric(
                "Prot. Prom",
                f"{df_per['Proteina'].mean():.2f}%".replace(".", ",")
                if pd.notna(df_per["Proteina"].mean())
                else "S/D",
            )
            idx += 1
          if v_crios:
            cols[idx].metric(
                "Crios Prom",
                f"{df_per['Crioscopia'].mean():.3f}".replace(".", ",")
                if pd.notna(df_per["Crioscopia"].mean())
                else "S/D",
            )
            idx += 1
          if v_ufc:
            cols[idx].metric(
                "UFC",
                formato_miles(df_per["UFC"].mean())
                if pd.notna(df_per["UFC"].mean())
                else "S/D",
            )
            idx += 1
          if v_scc:
            cols[idx].metric(
                "SCC",
                formato_miles(df_per["SCC"].mean())
                if pd.notna(df_per["SCC"].mean())
                else "S/D",
            )

          cols_show = ["Fecha", "N_Remito", "Litros_Ticket"]
          if v_temp:
            cols_show.append("Temperatura")
          if v_grasa:
            cols_show.append("Grasa")
          if v_prot:
            cols_show.append("Proteina")
          if v_crios:
            cols_show.append("Crioscopia")
          if v_ufc:
            cols_show.append("UFC")
          if v_scc:
            cols_show.append("SCC")

          df_disp = df_per[cols_show].copy()
          df_disp["Fecha"] = df_disp["Fecha"].dt.strftime("%d/%m/%Y")
          df_disp["Litros_Ticket"] = df_disp["Litros_Ticket"].apply(formato_miles)
          if v_temp:
            df_disp["Temperatura"] = df_disp["Temperatura"].apply(formato_temp)
          if v_grasa:
            df_disp["Grasa"] = df_disp["Grasa"].apply(
                lambda x: f"{x:.2f}%".replace(".", ",") if pd.notna(x) else "-"
            )
          if v_prot:
            df_disp["Proteina"] = df_disp["Proteina"].apply(
                lambda x: f"{x:.2f}%".replace(".", ",") if pd.notna(x) else "-"
            )
          if v_crios:
            df_disp["Crioscopia"] = df_disp["Crioscopia"].apply(
                lambda x: f"{x:.3f}".replace(".", ",") if pd.notna(x) else "-"
            )
          if v_ufc:
            df_disp["UFC"] = df_disp["UFC"].apply(
                lambda x: formato_miles(x) if pd.notna(x) else "-"
            )
          if v_scc:
            df_disp["SCC"] = df_disp["SCC"].apply(
                lambda x: formato_miles(x) if pd.notna(x) else "-"
            )

          st.dataframe(
              df_disp.rename(
                  columns={
                      "Litros_Ticket": "Litros",
                      "N_Remito": "N° Remito",
                      "Temperatura": "Temp",
                  }
              ),
              hide_index=True,
              use_container_width=True,
          )

          pdf_b = generar_pdf_bytes(
              df_per, t_nombre, t_id, periodo_pdf, args_vis, es_mensual
          )
          nom_arch = f"Resumen_{'Mensual' if es_mensual else 'Semanal'}_{t_nombre.replace(' ', '_')}.pdf"

          b1, b2 = st.columns(2)
          b1.download_button(
              "📥 Descargar PDF",
              data=pdf_b,
              file_name=nom_arch,
              mime="application/pdf",
              use_container_width=True,
          )
          if email_t:
            if b2.button(f"📧 Enviar Mail a {nom_c}", use_container_width=True):
              if enviar_correo_productor(
                  email_t,
                  nom_c,
                  t_nombre,
                  pdf_b,
                  nom_arch,
                  "mensual" if es_mensual else "semanal",
              ):
                st.success(f"Correo enviado a {email_t}")
          else:
            st.warning("Tambo sin email configurado.")

    elif vista_coop == "Envío Masivo Semanal":
      st.markdown(
          '<p class="main-header">📤 Envío Masivo y Control de Reportes'
          " Semanales</p>",
          unsafe_allow_html=True,
      )
      ciclos_disponibles = (
          df[["Fecha_Cierre_Viernes", "Ciclo_Semana"]]
          .drop_duplicates()
          .sort_values("Fecha_Cierre_Viernes", ascending=False)["Ciclo_Semana"]
          .tolist()
      )
      if ciclos_disponibles:
        ciclo_masivo = st.selectbox(
            "Seleccione el Cierre de Semana a procesar:", ciclos_disponibles
        )
        df_semana_macro = df[df["Ciclo_Semana"] == ciclo_masivo]
        if not df_semana_macro.empty:
          tramos_agrupados = df_semana_macro.groupby(
              ["Num_Tambo", "Tambo"], as_index=False
          ).agg(
              Litros_Totales=("Litros_Ticket", "sum"),
              Remitos_Count=("N_Remito", "count"),
              Temp_Media=("Temperatura", "mean"),
              Grasa_Prom=("Grasa", "mean"),
              Proteina_Prom=("Proteina", "mean"),
          )
          df_preview = pd.merge(
              tramos_agrupados, df_contactos, on="Num_Tambo", how="left"
          )

          def detectar_anomalias(row):
            alertas = []
            if pd.isna(row["Email"]) or not str(row["Email"]).strip():
              alertas.append("⚠️ Falta Email")
            if row["Temp_Media"] > 6.0:
              alertas.append("🚨 Temp Alta (>6°C)")
            if row["Litros_Totales"] <= 0:
              alertas.append("❌ Litros en 0")
            return " | ".join(alertas) if alertas else "✅ OK"

          df_preview["Estado / Alerta"] = df_preview.apply(
              detectar_anomalias, axis=1
          )
          st.dataframe(df_preview, hide_index=True, use_container_width=True)

          if st.button("🚀 Confirmar y Enviar Correos Masivos", type="primary"):
            st.info(
                "Proceso de envío masivo activo para la semana seleccionada."
            )

  except Exception as e:
    st.error("Error en el Módulo Coopagro:")
    st.code(traceback.format_exc())


# =========================================================================
# MÓDULO 2: RECEPCIÓN Y PRODUCCIÓN MASTELLONE (FASÓN) - SOLUCIÓN LITROS MES
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

      df_mastellone_prod = df_prod[df_prod["Grupo"] == "Mastellone"].copy()

      # 2. Leer Litros Mes desde Datos MHSA.xlsx (Solapa 'litros mes')
      df_mast_litros = cargar_datos_mastellone(URL_MASTELLONE)

      df_mast_litros_procesado = pd.DataFrame()
      if not df_mast_litros.empty:
        # Identificamos la columna de litros y la columna de mes/fecha con total precisión
        cols_lower = [str(c).lower() for c in df_mast_litros.columns]
        
        col_litros = next(
            (
                df_mast_litros.columns[i]
                for i, c in enumerate(cols_lower)
                if any(x in c for x in ["litro", "volumen", "cantidad", "total"])
            ),
            df_mast_litros.columns[-1],
        )

        col_fecha = next(
            (
                df_mast_litros.columns[i]
                for i, c in enumerate(cols_lower)
                if any(x in c for x in ["fecha", "mes", "periodo", "date"])
            ),
            df_mast_litros.columns[0],
        )

        df_mast_litros_procesado["Litros"] = pd.to_numeric(
            df_mast_litros[col_litros], errors="coerce"
        ).fillna(0)

        # Intentamos interpretar la columna de mes/fecha
        fechas_parsed = pd.to_datetime(
            df_mast_litros[col_fecha], errors="coerce"
        )
        if fechas_parsed.notna().any():
          df_mast_litros_procesado["Año"] = fechas_parsed.dt.year
          df_mast_litros_procesado["Mes"] = fechas_parsed.dt.month
        else:
          # Si son números de mes (1 al 12) o nombres de meses por fila
          df_mast_litros_procesado["Mes"] = pd.to_numeric(
              df_mast_litros[col_fecha], errors="coerce"
          )
          # Si no hay mes numérico, asumimos orden de filas (Fila 0 = Enero, ..., Fila 7 = Agosto)
          if df_mast_litros_procesado["Mes"].isna().all():
            df_mast_litros_procesado["Mes"] = range(1, len(df_mast_litros_procesado) + 1)
          df_mast_litros_procesado["Año"] = 2026

    # --- BARRA LATERAL (FILTROS) ---
    st.sidebar.markdown("### 🔍 Filtros Mastellone")
    opciones_anio = ["Todos"] + (
        sorted(df_mastellone_prod["Año"].unique().tolist())
        if len(df_mastellone_prod) > 0
        else [2026]
    )
    opciones_mes = ["Todos"] + list(range(1, 13))

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

    # --- OBTENER LITROS INGRESADOS FILTRADOS ESTRICTAMENTE POR AÑO Y MES ---
    total_litros_ingresados = 0.0
    if not df_mast_litros_procesado.empty:
      df_litros_f = df_mast_litros_procesado.copy()
      if filtro_anio != "Todos" and "Año" in df_litros_f.columns:
        df_litros_f = df_litros_f[df_litros_f["Año"] == filtro_anio]
      if filtro_mes != "Todos" and "Mes" in df_litros_f.columns:
        df_litros_f = df_litros_f[df_litros_f["Mes"] == filtro_mes]
      total_litros_ingresados = df_litros_f["Litros"].sum()

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
      '<p class="main-header">Registro de Producción y Ecuación de Van Slyke</p>',
      unsafe_allow_html=True,
  )
  st.info("Módulo en desarrollo para análisis global de rendimiento.")

elif modulo_principal == "📦 Insumos, Inventario y Costos":
  st.markdown(
      '<p class="main-header">Gestión de Insumos y Costos Variables</p>',
      unsafe_allow_html=True,
  )
  st.info("Módulo en desarrollo para control de stock y costos.")
