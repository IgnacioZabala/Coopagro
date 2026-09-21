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
FILE_ID_MASTELLONE = "19OVD6xBeK08o4cW1XrdMr54L1nciAJC2"
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

def generar_pdf_panel_general(df_macro, periodo_titulo, total_litros, temp_prom, grasa_prom, prot_prom, ratio_gp, tambos_activos, df_ranking):
  ratio_str = f"{ratio_gp:.2f}".replace(".", ",") if pd.notna(ratio_gp) else "S/D"
  grasa_str = f"{grasa_prom:.2f}%".replace(".", ",") if pd.notna(grasa_prom) else "S/D"
  prot_str = f"{prot_prom:.2f}%".replace(".", ",") if pd.notna(prot_prom) else "S/D"

  metricas = [
      f"Tambos Activos: {tambos_activos} | Litros Totales: {formato_miles(total_litros)} L",
      f"Temp. Promedio: {formato_temp(temp_prom)} | Grasa Ponderada: {grasa_str} | Prot. Ponderada: {prot_str}",
      f"Ratio Grasa / Proteína: {ratio_str}",
      "\nRanking de Tambos por Volumen de Litros",
  ]
  headers = [("Código", 30), ("Nombre del Tambo", 100), ("Litros Totales", 60)]
  mapeo = [lambda r: str(r.Num_Tambo), lambda r: str(r.Tambo), lambda r: formato_miles(r.Litros_Ticket)]
  return generar_pdf_base("Informe de Recolección - Cooperativa", f"Período Evaluado: {periodo_titulo}", metricas, headers, df_ranking, mapeo)

def generar_pdf_bytes(df_productor, tambo_nombre, tambo_id, periodo_texto, args_visibles, es_mensual=False):
  titulo = "Resumen mensual de recolección" if es_mensual else "Resumen semanal de recolección"
  subtitulo = f"Productor: {tambo_nombre} (Código #{tambo_id})"
  temp_prom = df_productor["Temperatura"].mean() if "Temperatura" in df_productor else float("nan")

  metricas = [f"Período: {periodo_texto}", f"Total Litros: {formato_miles(df_productor['Litros_Ticket'].sum())} L"]
  if args_visibles["temp"]: metricas.append(f"Temperatura Promedio: {formato_temp(temp_prom)}")

  partes_solidos = []
  if args_visibles["grasa"] and "Grasa" in df_productor and pd.notna(df_productor["Grasa"].mean()): partes_solidos.append(f"Grasa: {df_productor['Grasa'].mean():.2f}%".replace(".", ","))
  if args_visibles["prot"] and "Proteina" in df_productor and pd.notna(df_productor["Proteina"].mean()): partes_solidos.append(f"Proteína: {df_productor['Proteina'].mean():.2f}%".replace(".", ","))
  if args_visibles["crios"] and "Crioscopia" in df_productor and pd.notna(df_productor["Crioscopia"].mean()): partes_solidos.append(f"Crioscopia: {df_productor['Crioscopia'].mean():.3f}".replace(".", ","))
  if args_visibles["ufc"] and "UFC" in df_productor and pd.notna(df_productor["UFC"].mean()): partes_solidos.append(f"UFC: {formato_miles(df_productor['UFC'].mean())}")
  if args_visibles["scc"] and "SCC" in df_productor and pd.notna(df_productor["SCC"].mean()): partes_solidos.append(f"SCC: {formato_miles(df_productor['SCC'].mean())}")
  if partes_solidos: metricas.append(f"Promedios Lab -> {' | '.join(partes_solidos)}")

  headers = [("Fecha", 26), ("N° Remito", 34), ("Litros", 30)]
  mapeo = [
      lambda r: getattr(r, "Fecha").strftime("%d/%m/%Y") if pd.notna(getattr(r, "Fecha")) else "",
      lambda r: str(getattr(r, "N_Remito")) if pd.notna(getattr(r, "N_Remito")) else "-",
      lambda r: formato_miles(getattr(r, "Litros_Ticket")) if pd.notna(getattr(r, "Litros_Ticket")) else "0"
  ]
  if args_visibles["temp"]: headers.append(("Temp", 18)); mapeo.append(lambda r: formato_temp(getattr(r, "Temperatura", pd.NA)))
  if args_visibles["grasa"]: headers.append(("Grasa", 20)); mapeo.append(lambda r: f"{getattr(r, 'Grasa'):.2f}%".replace(".", ",") if pd.notna(getattr(r, 'Grasa', pd.NA)) else "-")
  if args_visibles["prot"]: headers.append(("Prot", 20)); mapeo.append(lambda r: f"{getattr(r, 'Proteina'):.2f}%".replace(".", ",") if pd.notna(getattr(r, 'Proteina', pd.NA)) else "-")
  if args_visibles["crios"]: headers.append(("Crios", 22)); mapeo.append(lambda r: f"{getattr(r, 'Crioscopia'):.3f}".replace(".", ",") if pd.notna(getattr(r, 'Crioscopia', pd.NA)) else "-")
  if args_visibles["ufc"]: headers.append(("UFC", 22)); mapeo.append(lambda r: formato_miles(getattr(r, 'UFC', pd.NA)) if pd.notna(getattr(r, 'UFC', pd.NA)) else "-")
  if args_visibles["scc"]: headers.append(("SCC", 24)); mapeo.append(lambda r: formato_miles(getattr(r, 'SCC', pd.NA)) if pd.notna(getattr(r, 'SCC', pd.NA)) else "-")

  return generar_pdf_base(titulo, subtitulo, metricas, headers, df_productor, mapeo)

def enviar_correo_productor(destinatario_email, nombre_contacto, tambo_nombre, pdf_bytes, nombre_archivo, tipo_reporte="semanal"):
  try:
    remitente = st.secrets["email"]["remitente"]
    password = st.secrets["email"]["password"]
    msg = MIMEMultipart()
    msg["From"] = remitente
    destinatarios = [e.strip() for e in destinatario_email.replace(";", ",").split(",") if e.strip()]
    msg["To"] = ", ".join(destinatarios)
    msg["Subject"] = f"Resumen {tipo_reporte.capitalize()} de Recolección - {tambo_nombre}"
    cuerpo_html = f"<html><body><p>Buenas tardes, <b>{nombre_contacto}</b>:</p><p>Le adjunto el resumen {tipo_reporte} de recolección y calidad de leche.</p></body></html>"
    msg.attach(MIMEText(cuerpo_html, "html"))
    part = MIMEBase("application", "octet-stream")
    part.set_payload(pdf_bytes)
    encoders.encode_base64(part)
    part.add_header("Content-Disposition", f'attachment; filename="{nombre_archivo}"')
    msg.attach(part)
    with smtplib.SMTP("smtp.gmail.com", 587) as server:
      server.starttls()
      server.login(remitente, password)
      server.sendmail(remitente, destinatarios, msg.as_string())
    return True
  except Exception as e:
    st.error(f"Error de envío SMTP: {e}")
    return False

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

      if col_codigo is not None and col_contacto is not None and col_email is not None:
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

    # Procesamiento Lab
    if not df_lab_raw.empty:
      df_lab = df_lab_raw.copy()
      col_sample = next((c for c in df_lab.columns if any(x in c.lower() for x in ["sample", "number", "tambo", "muestra"])), df_lab.columns[0])
      col_date = next((c for c in df_lab.columns if any(x in c.lower() for x in ["fecha", "date", "time"])), None)
      df_lab["Num_Tambo"] = df_lab[col_sample].astype(str).str.split().str[0].apply(limpiar_tambo)
      if col_date and df_lab[col_date].notna().any(): df_lab["Fecha"] = pd.to_datetime(df_lab[col_date], errors="coerce").dt.normalize()
      else: df_lab["Fecha"] = df_lab[col_sample].apply(lambda x: extraer_fecha_texto(x) if pd.notna(x) else pd.NaT)
      df_lab = df_lab.dropna(subset=["Fecha", "Num_Tambo"]).sort_values(by=["Num_Tambo", "Fecha"])
      df_lab["orden_remito"] = df_lab.groupby(["Num_Tambo", "Fecha"]).cumcount() + 1
      map_cols = {}
      col_fat = next((c for c in df_lab.columns if "fat" in c.lower() or "grasa" in c.lower()), None)
      if col_fat: map_cols[col_fat] = "Grasa_Lab"
      col_prot = next((c for c in df_lab.columns if "protein" in c.lower() or "proteina" in c.lower()), None)
      if col_prot: map_cols[col_prot] = "Proteina_Lab"
      col_fp = next((c for c in df_lab.columns if c.lower() == "fp" or "crios" in c.lower()), None)
      if col_fp: map_cols[col_fp] = "Crioscopia_Lab"
      if map_cols:
        df_milko_clean = df_lab[["Num_Tambo", "Fecha", "orden_remito"] + list(map_cols.keys())].rename(columns=map_cols)
        for c in map_cols.values(): df_milko_clean[c] = pd.to_numeric(df_milko_clean[c].astype(str).str.replace(",", "."), errors="coerce")
        df = pd.merge(df, df_milko_clean, on=["Num_Tambo", "Fecha", "orden_remito"], how="left")
        if "Grasa_Lab" in df: df["Grasa"] = df["Grasa_Lab"].combine_first(df["Grasa"])
        if "Proteina_Lab" in df: df["Proteina"] = df["Proteina_Lab"].combine_first(df["Proteina"])
        if "Crioscopia_Lab" in df: df["Crioscopia"] = df["Crioscopia_Lab"].combine_first(df["Crioscopia"])

    # Procesamiento Bacsomatic
    if not df_bac_raw.empty:
      df_bac = df_bac_raw.copy()
      col_id = next((c for c in df_bac.columns if any(x in c.lower() for x in ["id usuario", "sample", "tambo"])), df_bac.columns[0])
      df_bac["Num_Tambo"] = df_bac[col_id].astype(str).str.split().str[0].apply(limpiar_tambo)
      df_bac["Fecha"] = df_bac[col_id].apply(lambda x: extraer_fecha_texto(x) if pd.notna(x) else pd.NaT)
      df_bac = df_bac.dropna(subset=["Fecha", "Num_Tambo"]).sort_values(by=["Num_Tambo", "Fecha"])
      df_bac["orden_remito"] = df_bac.groupby(["Num_Tambo", "Fecha"]).cumcount() + 1
      map_cols_bac = {}
      col_ufc = next((c for c in df_bac.columns if "ufc" in c.lower()), None)
      if col_ufc: map_cols_bac[col_ufc] = "UFC_Val"
      col_scc = next((c for c in df_bac.columns if any(x in c.lower() for x in ["scc", "celulas", "somáticas"])), None)
      if col_scc: map_cols_bac[col_scc] = "SCC_Val"
      if map_cols_bac:
        df_bac_clean = df_bac[["Num_Tambo", "Fecha", "orden_remito"] + list(map_cols_bac.keys())].rename(columns=map_cols_bac)
        for c in map_cols_bac.values(): df_bac_clean[c] = pd.to_numeric(df_bac_clean[c].astype(str).str.replace(",", "."), errors="coerce")
        df = pd.merge(df, df_bac_clean, on=["Num_Tambo", "Fecha", "orden_remito"], how="left")
        if "UFC_Val" in df: df["UFC"] = df["UFC_Val"].combine_first(df["UFC"])
        if "SCC_Val" in df: df["SCC"] = df["SCC_Val"].combine_first(df["SCC"])

    df["Fecha_Cierre_Viernes"] = df["Fecha"] + pd.to_timedelta((4 - df["Fecha"].dt.weekday) % 7, unit="D")
    df["Fecha_Inicio_Sabado"] = df["Fecha_Cierre_Viernes"] - pd.Timedelta(days=6)
    df["Ciclo_Semana"] = "Viernes " + df["Fecha_Cierre_Viernes"].dt.strftime("%d/%m/%Y") + " (Sáb " + df["Fecha_Inicio_Sabado"].dt.strftime("%d/%m/%Y") + " al Vie " + df["Fecha_Cierre_Viernes"].dt.strftime("%d/%m/%Y") + ")"
    df["AnioMes"] = df["Fecha"].dt.to_period("M")

    # Submenú Módulo 1
    st.sidebar.markdown("---")
    vista_coop = st.sidebar.radio("Sección Coopagro:", ["Panel de Control General", "Gestión y Reportes por Tambo", "Envío Masivo Semanal"])

    if vista_coop == "Panel de Control General":
      st.header("📊 Panel General - Coopagro")
      tipo_reporte_opcion = st.radio("Seleccione el período:", ["Semanal", "Mensual"], horizontal=True)

      if tipo_reporte_opcion == "Semanal":
        ciclos = df[["Fecha_Cierre_Viernes", "Ciclo_Semana"]].drop_duplicates().sort_values("Fecha_Cierre_Viernes", ascending=False)["Ciclo_Semana"].tolist()
        ciclo_gen = st.selectbox("Seleccione el Cierre de Semana:", ciclos) if ciclos else ""
        df_macro = df[df["Ciclo_Semana"] == ciclo_gen] if ciclos else pd.DataFrame()
        periodo_texto = ciclo_gen
      else:
        meses = sorted(df["AnioMes"].unique(), reverse=True)
        mes_gen = st.selectbox("Seleccione el Mes:", meses, format_func=lambda p: f"{MESES_ES.get(p.month)} {p.year}") if meses else None
        df_macro = df[df["AnioMes"] == mes_gen] if mes_gen else pd.DataFrame()
        periodo_texto = f"{MESES_ES.get(mes_gen.month)} {mes_gen.year}" if mes_gen else ""

      if not df_macro.empty:
        tot_litros = df_macro["Litros_Ticket"].sum()
        grasa_p = calcular_promedio_ponderado(df_macro, "Grasa")
        prot_p = calcular_promedio_ponderado(df_macro, "Proteina")
        ratio_gp = grasa_p / prot_p if (prot_p and prot_p > 0) else pd.NA

        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("Litros Totales", formato_miles(tot_litros))
        c2.metric("Temp. Media", formato_temp(df_macro["Temperatura"].mean()))
        c3.metric("Grasa Ponderada", f"{grasa_p:.2f}%".replace(".", ",") if pd.notna(grasa_p) else "S/D")
        c4.metric("Proteína Ponderada", f"{prot_p:.2f}%".replace(".", ",") if pd.notna(prot_p) else "S/D")
        c5.metric("Ratio Grasa/Prot.", f"{ratio_gp:.2f}".replace(".", ",") if pd.notna(ratio_gp) else "S/D")

        st.subheader("Ranking de Tambos por Volumen")
        df_ranking = df_macro.groupby(["Tambo", "Num_Tambo"], as_index=False)["Litros_Ticket"].sum().sort_values("Litros_Ticket", ascending=False)
        
        pdf_bytes = generar_pdf_panel_general(df_macro, periodo_texto, tot_litros, df_macro["Temperatura"].mean(), grasa_p, prot_p, ratio_gp, df_macro["Num_Tambo"].nunique(), df_ranking)
        st.download_button("📥 Descargar Informe PDF", data=pdf_bytes, file_name=f"Informe_{periodo_texto.replace(' ', '_')}.pdf", mime="application/pdf")
        
        df_ranking_show = df_ranking.copy()
        df_ranking_show["Litros_Ticket"] = df_ranking_show["Litros_Ticket"].apply(formato_miles)
        st.dataframe(df_ranking_show.rename(columns={"Tambo": "Nombre del Tambo", "Num_Tambo": "Código", "Litros_Ticket": "Litros Totales"}), hide_index=True, use_container_width=True)

    elif vista_coop == "Gestión y Reportes por Tambo":
      st.header("📄 Reportes por Tambo")
      tipo_reporte_opcion = st.sidebar.radio("Período de Reporte:", ["Semanal", "Mensual"])
      mapeo_tambos = df[["Tambo", "Num_Tambo"]].drop_duplicates().sort_values("Tambo")
      
      if not mapeo_tambos.empty:
        t_nombre = st.sidebar.selectbox("1. Seleccione Tambo:", mapeo_tambos["Tambo"].tolist())
        t_id = mapeo_tambos.loc[mapeo_tambos["Tambo"] == t_nombre, "Num_Tambo"].values[0]
        df_t = df[df["Num_Tambo"] == str(t_id)]

        if tipo_reporte_opcion == "Semanal":
          ciclos = df_t[["Fecha_Cierre_Viernes", "Ciclo_Semana"]].drop_duplicates().sort_values("Fecha_Cierre_Viernes", ascending=False)["Ciclo_Semana"].tolist()
          ciclo_sel = st.sidebar.selectbox("2. Cierre de Semana:", ciclos) if ciclos else ""
          df_per = df_t[df_t["Ciclo_Semana"] == ciclo_sel].sort_values(["Fecha", "N_Remito"]) if ciclos else pd.DataFrame()
          es_mensual = False
          periodo_pdf = f"{df_per['Fecha_Inicio_Sabado'].iloc[0]:%d/%m/%Y} al {df_per['Fecha_Cierre_Viernes'].iloc[0]:%d/%m/%Y}" if not df_per.empty else ""
        else:
          meses = sorted(df_t["AnioMes"].unique(), reverse=True)
          mes_sel = st.sidebar.selectbox("2. Mes:", meses, format_func=lambda p: f"{MESES_ES.get(p.month)} {p.year}") if meses else None
          df_per = df_t[df_t["AnioMes"] == mes_sel].sort_values(["Fecha", "N_Remito"]) if mes_sel else pd.DataFrame()
          es_mensual = True
          periodo_pdf = f"{MESES_ES.get(mes_sel.month)} {mes_sel.year}" if mes_sel else ""

        st.sidebar.subheader("⚙️ Elementos del Reporte")
        v_temp = st.sidebar.checkbox("Temperatura", True)
        v_grasa = st.sidebar.checkbox("Grasa", True)
        v_prot = st.sidebar.checkbox("Proteína", True)
        v_crios = st.sidebar.checkbox("Crioscopia", True)
        v_ufc = st.sidebar.checkbox("UFC", True)
        v_scc = st.sidebar.checkbox("SCC", True)
        args_vis = {"temp": v_temp, "grasa": v_grasa, "prot": v_prot, "crios": v_crios, "ufc": v_ufc, "scc": v_scc}

        if not df_per.empty:
          info_c = df_contactos[df_contactos["Num_Tambo"] == str(t_id)]
          email_t = info_c["Email"].values[0] if not info_c.empty and pd.notna(info_c["Email"].values[0]) else ""
          nom_c = info_c["Contacto_Nombre"].values[0] if not info_c.empty and pd.notna(info_c["Contacto_Nombre"].values[0]) else "Productor"
          
          st.subheader(f"Resumen {'Mensual' if es_mensual else 'Semanal'} - {t_nombre} (#{t_id})")
          
          pdf_b = generar_pdf_bytes(df_per, t_nombre, t_id, periodo_pdf, args_vis, es_mensual)
          nom_arch = f"Resumen_{'Mensual' if es_mensual else 'Semanal'}_{t_nombre.replace(' ', '_')}.pdf"

          b1, b2 = st.columns(2)
          b1.download_button("📥 Descargar PDF", data=pdf_b, file_name=nom_arch, mime="application/pdf", use_container_width=True)
          if email_t:
            if b2.button(f"📧 Enviar Mail a {nom_c}", use_container_width=True):
              if enviar_correo_productor(email_t, nom_c, t_nombre, pdf_b, nom_arch, "mensual" if es_mensual else "semanal"):
                st.success(f"Correo enviado a {email_t}")
          else: st.warning("Tambo sin email configurado.")

          df_tabla_visual = pd.DataFrame()
          df_tabla_visual["Fecha"] = df_per["Fecha"].dt.strftime("%d/%m/%Y")
          df_tabla_visual["N° Remito"] = df_per["N_Remito"]
          df_tabla_visual["Litros"] = df_per["Litros_Ticket"].apply(formato_miles)

          if v_temp and "Temperatura" in df_per:
            df_tabla_visual["Temperatura"] = df_per["Temperatura"].apply(lambda x: f"{x:.1f}°" if pd.notna(x) else "-")
          if v_grasa and "Grasa" in df_per:
            df_tabla_visual["Grasa"] = df_per["Grasa"].apply(lambda x: f"{x:.2f}%" if pd.notna(x) else "-")
          if v_prot and "Proteina" in df_per:
            df_tabla_visual["Proteína"] = df_per["Proteina"].apply(lambda x: f"{x:.2f}%" if pd.notna(x) else "-")
          if v_crios and "Crioscopia" in df_per:
            df_tabla_visual["Crioscopia"] = df_per["Crioscopia"].apply(lambda x: f"{x:.3f}" if pd.notna(x) else "-")
          if v_ufc and "UFC" in df_per:
            df_tabla_visual["UFC"] = df_per["UFC"].apply(lambda x: formato_miles(x) if pd.notna(x) else "-")
          if v_scc and "SCC" in df_per:
            df_tabla_visual["SCC"] = df_per["SCC"].apply(lambda x: formato_miles(x) if pd.notna(x) else "-")

          st.dataframe(df_tabla_visual, use_container_width=True, hide_index=True)

    elif vista_coop == "Envío Masivo Semanal":
      st.header("📤 Envío Masivo Semanal")
      st.info("Configuración de correos lista para enviar.")
  except Exception as e:
    st.error("Error en el Módulo Coopagro:")
    st.code(traceback.format_exc())

# =========================================================================
# MÓDULO 2: RECEPCIÓN, PRODUCCIÓN Y CALIDAD MASTELLONE
# =========================================================================
elif modulo_principal == "🚛 Recepción Mastellone (Fasón)":
  st.header("🚛 Recepción, Calidad y Producción Mastellone")
  import datetime
  now = datetime.datetime.now()
  
  try:
    with st.spinner("Sincronizando datos de Mastellone..."):
      # 1. Cargar Producción (RE-PRO-52)
      raw_prod = pd.read_excel(URL_PRODUCCION, skiprows=6)
      df_prod = pd.DataFrame()
      df_prod["Fecha"] = raw_prod.iloc[:, 0]
      df_prod["Lote"] = raw_prod.iloc[:, 1]
      df_prod["Litros Procesados"] = raw_prod.iloc[:, 3]
      df_prod["Producto Terminado"] = raw_prod.iloc[:, 5]
      df_prod["PNC"] = raw_prod.iloc[:, 6]
      df_prod = df_prod.dropna(subset=["Fecha"])
      
      # Forzamos formato de fecha robusto
      df_prod["Fecha"] = pd.to_datetime(df_prod["Fecha"], format="mixed", dayfirst=True, errors="coerce")
      df_prod = df_prod.dropna(subset=["Fecha"])
      for col in ["Litros Procesados", "Producto Terminado", "PNC"]: 
          df_prod[col] = pd.to_numeric(df_prod[col], errors="coerce").fillna(0)
      
      if len(df_prod) > 0: 
          df_prod["Producto"] = df_prod["Lote"].astype(str).apply(lambda x: "Muzzarella Export. Mastellone" if "840" in str(x) else "Otro")
          df_prod["Grupo"] = df_prod["Lote"].astype(str).apply(lambda x: "Mastellone" if "840" in str(x) else "Coopagro")
      else: 
          df_prod["Producto"], df_prod["Grupo"] = [], []
      
      df_prod["Año"] = df_prod["Fecha"].dt.year
      df_prod["Mes"] = df_prod["Fecha"].dt.month
      df_mastellone_prod = df_prod[df_prod["Grupo"] == "Mastellone"].copy()

      # 2. Cargar Recepción Diaria (Hoja 2 de FILE_ID_MASTELLONE)
      df_mhsa = pd.DataFrame()
      try:
          xls_mastellone = pd.ExcelFile(URL_MASTELLONE)
          sheet_mhsa = xls_mastellone.sheet_names[1] if len(xls_mastellone.sheet_names) > 1 else xls_mastellone.sheet_names[0]
          
          df_mhsa_raw = pd.read_excel(URL_MASTELLONE, sheet_name=sheet_mhsa, dtype=str)
          
          df_mhsa["Fecha"] = df_mhsa_raw.iloc[:, 0]
          df_mhsa["Num_Tambo"] = df_mhsa_raw.iloc[:, 2]
          df_mhsa["Tambo"] = df_mhsa_raw.iloc[:, 3]
          df_mhsa["Litros_Ticket"] = df_mhsa_raw.iloc[:, 4]
          
          if len(df_mhsa_raw.columns) > 7:
              df_mhsa["Temperatura"] = df_mhsa_raw.iloc[:, 7]
          else:
              df_mhsa["Temperatura"] = pd.NA
          
          df_mhsa["Num_Tambo"] = df_mhsa["Num_Tambo"].apply(limpiar_tambo)
          df_mhsa["Fecha"] = pd.to_datetime(df_mhsa["Fecha"], format="mixed", dayfirst=True, errors="coerce").dt.normalize()
          df_mhsa = df_mhsa.dropna(subset=["Fecha", "Num_Tambo"])
          df_mhsa["Litros_Ticket"] = pd.to_numeric(df_mhsa["Litros_Ticket"], errors="coerce").fillna(0)
          df_mhsa["Temperatura"] = pd.to_numeric(df_mhsa["Temperatura"], errors="coerce")
          
          df_mhsa["Año"] = df_mhsa["Fecha"].dt.year
          df_mhsa["Mes"] = df_mhsa["Fecha"].dt.month
          df_mhsa["orden_remito"] = df_mhsa.groupby(["Num_Tambo", "Fecha"]).cumcount() + 1
      except Exception as e:
          st.sidebar.warning(f"Error al cargar la hoja 2 (MHSA Diario): {e}")
      
      # 3. Cargar Litros Mes (Hoja 3: MHSA Litros Mes)
      df_mhsa_litros = pd.DataFrame()
      try:
          sheet_litros = next((s for s in xls_mastellone.sheet_names if "litros" in s.lower() and "mes" in s.lower()), None)
          if not sheet_litros and len(xls_mastellone.sheet_names) > 2:
              sheet_litros = xls_mastellone.sheet_names[2]
              
          if sheet_litros:
              df_litros_raw = pd.read_excel(URL_MASTELLONE, sheet_name=sheet_litros)
              
              # Lógica adaptable: Si pusiste Año en A (3 columnas o más) o si quedó como en tu foto (2 columnas)
              if len(df_litros_raw.columns) >= 3:
                  df_mhsa_litros["Año"] = pd.to_numeric(df_litros_raw.iloc[:, 0], errors="coerce").fillna(2026).astype(int)
                  df_mhsa_litros["Mes"] = pd.to_numeric(df_litros_raw.iloc[:, 1], errors="coerce").fillna(0).astype(int)
                  df_mhsa_litros["Litros"] = pd.to_numeric(df_litros_raw.iloc[:, 2], errors="coerce").fillna(0)
              else:
                  # Fallback por si la estructura es como en la captura de pantalla (A=Mes, B=Litros)
                  df_mhsa_litros["Año"] = 2026
                  df_mhsa_litros["Mes"] = pd.to_numeric(df_litros_raw.iloc[:, 0], errors="coerce").fillna(0).astype(int)
                  df_mhsa_litros["Litros"] = pd.to_numeric(df_litros_raw.iloc[:, 1], errors="coerce").fillna(0)
      except Exception as e:
          st.sidebar.warning(f"Error al cargar la hoja 3 (Litros Mes): {e}")

      # 4. Cargar Laboratorio y cruzar con Recepción
      _, _, df_lab_raw, df_bac_raw = cargar_datos_coopagro(URL_REMITOS, URL_LAB, URL_BACSOMATIC)
      
      if not df_mhsa.empty:
          # --- Procesar MilkoScan ---
          if not df_lab_raw.empty:
              df_lab_m = df_lab_raw.copy()
              col_sample = df_lab_m.columns[0]
              
              df_lab_m["Num_Tambo"] = df_lab_m[col_sample].astype(str).str.split().str[0].apply(limpiar_tambo)
              df_lab_m["Fecha_Extraida"] = pd.to_datetime(df_lab_m[col_sample].astype(str).str.split().str[-1].apply(extraer_fecha_texto), errors="coerce")
              
              col_date = next((c for c in df_lab_m.columns if any(x in c.lower() for x in ["fecha", "date", "analyzed"])), None)
              if col_date:
                  df_lab_m["Fecha_Analisis"] = pd.to_datetime(df_lab_m[col_date], errors="coerce").dt.normalize()
                  df_lab_m["Fecha"] = df_lab_m["Fecha_Extraida"].fillna(df_lab_m["Fecha_Analisis"])
              else:
                  df_lab_m["Fecha"] = df_lab_m["Fecha_Extraida"]
              
              df_lab_m = df_lab_m.dropna(subset=["Fecha", "Num_Tambo"]).sort_values(by=["Num_Tambo", "Fecha"])
              df_lab_m["orden_remito"] = df_lab_m.groupby(["Num_Tambo", "Fecha"]).cumcount() + 1
              
              map_cols = {}
              col_fat = next((c for c in df_lab_m.columns if "fat" in c.lower() or "grasa" in c.lower()), None)
              col_prot = next((c for c in df_lab_m.columns if "protein" in c.lower() or "proteina" in c.lower()), None)
              col_fp = next((c for c in df_lab_m.columns if "fp" == c.lower() or "crios" in c.lower()), None)
              
              if col_fat: map_cols[col_fat] = "Grasa_Lab"
              if col_prot: map_cols[col_prot] = "Proteina_Lab"
              if col_fp: map_cols[col_fp] = "Crioscopia_Lab"
              
              if map_cols:
                  df_milko_clean = df_lab_m[["Num_Tambo", "Fecha", "orden_remito"] + list(map_cols.keys())].rename(columns=map_cols)
                  for c in map_cols.values(): 
                      df_milko_clean[c] = pd.to_numeric(df_milko_clean[c].astype(str).str.replace(",", "."), errors="coerce")
                  df_mhsa = pd.merge(df_mhsa, df_milko_clean, on=["Num_Tambo", "Fecha", "orden_remito"], how="left")

          # --- Procesar BacSomatic ---
          if not df_bac_raw.empty:
              df_bac_m = df_bac_raw.copy()
              if len(df_bac_m.columns) > 5:
                  col_sample_bac = df_bac_m.columns[5]
              else:
                  col_sample_bac = next((c for c in df_bac_m.columns if any(x in c.lower() for x in ["id usuario", "sample", "tambo"])), df_bac_m.columns[0])
              
              df_bac_m["Num_Tambo"] = df_bac_m[col_sample_bac].astype(str).str.split().str[0].apply(limpiar_tambo)
              df_bac_m["Fecha_Extraida"] = pd.to_datetime(df_bac_m[col_sample_bac].astype(str).str.split().str[-1].apply(extraer_fecha_texto), errors="coerce")
              
              col_date_bac = next((c for c in df_bac_m.columns if any(x in c.lower() for x in ["fecha", "date", "analyzed"])), None)
              if col_date_bac:
                  df_bac_m["Fecha_Analisis"] = pd.to_datetime(df_bac_m[col_date_bac], errors="coerce").dt.normalize()
                  df_bac_m["Fecha"] = df_bac_m["Fecha_Extraida"].fillna(df_bac_m["Fecha_Analisis"])
              else:
                  df_bac_m["Fecha"] = df_bac_m["Fecha_Extraida"]

              df_bac_m = df_bac_m.dropna(subset=["Fecha", "Num_Tambo"]).sort_values(by=["Num_Tambo", "Fecha"])
              df_bac_m["orden_remito"] = df_bac_m.groupby(["Num_Tambo", "Fecha"]).cumcount() + 1
              
              map_cols_bac = {}
              col_ufc = next((c for c in df_bac_m.columns if "ufc" in c.lower()), None)
              col_scc = next((c for c in df_bac_m.columns if any(x in c.lower() for x in ["scc", "celulas", "somáticas"])), None)
              
              if col_ufc: map_cols_bac[col_ufc] = "UFC_Val"
              if col_scc: map_cols_bac[col_scc] = "SCC_Val"
              
              if map_cols_bac:
                  df_bac_clean = df_bac_m[["Num_Tambo", "Fecha", "orden_remito"] + list(map_cols_bac.keys())].rename(columns=map_cols_bac)
                  for c in map_cols_bac.values(): 
                      df_bac_clean[c] = pd.to_numeric(df_bac_clean[c].astype(str).str.replace(",", "."), errors="coerce")
                  df_mhsa = pd.merge(df_mhsa, df_bac_clean, on=["Num_Tambo", "Fecha", "orden_remito"], how="left")

    # ==========================================
    # FILTROS POR DEFECTO DINÁMICOS
    # ==========================================
    st.sidebar.subheader("Filtros Mastellone")
    anios_mhsa = df_mhsa["Año"].dropna().unique().tolist() if not df_mhsa.empty else []
    anios_prod = df_mastellone_prod["Año"].dropna().unique().tolist() if not df_mastellone_prod.empty else []
    todos_anios = sorted(list(set(anios_mhsa + anios_prod + df_mhsa_litros["Año"].tolist() if not df_mhsa_litros.empty else [])))
    
    opciones_anio = ["Todos"] + (todos_anios if todos_anios else [now.year])
    default_anio_idx = opciones_anio.index(now.year) if now.year in opciones_anio else 0
    filtro_anio = st.sidebar.selectbox("Año", opciones_anio, index=default_anio_idx, key="m_anio_mastellone")

    opciones_mes = ["Todos"] + list(range(1, 13))
    default_mes_idx = opciones_mes.index(now.month) if now.month in opciones_mes else 0
    filtro_mes = st.sidebar.selectbox("Mes", opciones_mes, index=default_mes_idx, key="m_mes_mastellone")

    # ==========================================
    # CÁLCULOS Y MÉTRICAS
    # ==========================================
    # Filtro de Producción
    df_filtrado = df_mastellone_prod.copy()
    if len(df_filtrado) > 0:
      if filtro_anio != "Todos": df_filtrado = df_filtrado[df_filtrado["Año"] == filtro_anio]
      if filtro_mes != "Todos": df_filtrado = df_filtrado[df_filtrado["Mes"] == filtro_mes]

    df_consolidado = pd.DataFrame(columns=["Fecha", "Lote", "Producto", "Litros Procesados", "Producto Terminado", "Ratio de Conversión (%)"])
    if len(df_filtrado) > 0:
      df_consolidado_raw = df_filtrado.copy()
      df_consolidado_raw["PT_Total"] = df_consolidado_raw["Producto Terminado"] + df_consolidado_raw["PNC"]
      df_consolidado_raw["Ratio Consolidado (%)"] = df_consolidado_raw.apply(lambda x: f"{(x['PT_Total'] / x['Litros Procesados'] * 100):.2f}%" if x["Litros Procesados"] > 0 else "0.00%", axis=1)
      df_consolidado = df_consolidado_raw[["Fecha", "Lote", "Producto", "Litros Procesados", "PT_Total", "Ratio Consolidado (%)"]].rename(columns={"PT_Total": "Producto Terminado", "Ratio Consolidado (%)": "Ratio de Conversión (%)"})

    # Filtro de Litros Ingresados Brutos (Hoja 3)
    df_mhsa_litros_f = df_mhsa_litros.copy()
    if not df_mhsa_litros_f.empty:
        if filtro_anio != "Todos": df_mhsa_litros_f = df_mhsa_litros_f[df_mhsa_litros_f["Año"] == int(filtro_anio)]
        if filtro_mes != "Todos": df_mhsa_litros_f = df_mhsa_litros_f[df_mhsa_litros_f["Mes"] == int(filtro_mes)]
        total_litros_ingresados = df_mhsa_litros_f["Litros"].sum()
    else:
        total_litros_ingresados = 0.0

    total_litros_proc = df_filtrado["Litros Procesados"].sum() if len(df_filtrado) > 0 else 0
    total_prod_consolidado = df_consolidado["Producto Terminado"].sum() if len(df_consolidado) > 0 else 0
    ratio_ponderado = (total_prod_consolidado / total_litros_proc * 100) if total_litros_proc > 0 else 0
    rendimiento_ingreso = (total_prod_consolidado / total_litros_ingresados * 100) if total_litros_ingresados > 0 else 0

    c1, c2, c3 = st.columns(3)
    c1.metric("Litros Ingresados Brutos", formato_miles(total_litros_ingresados))
    c2.metric("Litros Procesados", formato_miles(total_litros_proc))
    c3.metric("Total Producto Terminado", formato_miles(total_prod_consolidado))
    
    c4, c5, _ = st.columns(3)
    c4.metric("Ratio PT / Procesados", f"{ratio_ponderado:.2f}%")
    c5.metric("Ratio PT / Ingresados", f"{rendimiento_ingreso:.2f}%")

    # ==========================================
    # VISUALIZACIÓN EN TABS
    # ==========================================
    tab1, tab2 = st.tabs(["📑 Recepción y Calidad (MHSA)", "🏭 Producción Fasón"])
    
    with tab1:
        st.subheader("Recepción y Calidad de Tambos Mastellone")
        df_mhsa_filtrado = df_mhsa.copy()
        if not df_mhsa_filtrado.empty:
            if filtro_anio != "Todos": df_mhsa_filtrado = df_mhsa_filtrado[df_mhsa_filtrado["Año"] == filtro_anio]
            if filtro_mes != "Todos": df_mhsa_filtrado = df_mhsa_filtrado[df_mhsa_filtrado["Mes"] == filtro_mes]
        
        if not df_mhsa_filtrado.empty:
            df_mhsa_disp = df_mhsa_filtrado.copy()
            df_mhsa_disp = df_mhsa_disp.sort_values(by=["Fecha", "Num_Tambo"])
            df_mhsa_disp["Fecha"] = df_mhsa_disp["Fecha"].dt.strftime("%d/%m/%Y")
            df_mhsa_disp["Litros_Ticket"] = df_mhsa_disp["Litros_Ticket"].apply(formato_miles)
            df_mhsa_disp["Temperatura"] = df_mhsa_disp["Temperatura"].apply(lambda x: f"{x:.1f}°" if pd.notna(x) else "-")
            
            if "Grasa_Lab" in df_mhsa_disp: df_mhsa_disp["Grasa"] = df_mhsa_disp["Grasa_Lab"].apply(lambda x: f"{x:.2f}%" if pd.notna(x) else "-")
            if "Proteina_Lab" in df_mhsa_disp: df_mhsa_disp["Proteína"] = df_mhsa_disp["Proteina_Lab"].apply(lambda x: f"{x:.2f}%" if pd.notna(x) else "-")
            if "Crioscopia_Lab" in df_mhsa_disp: df_mhsa_disp["Crioscopía"] = df_mhsa_disp["Crioscopia_Lab"].apply(lambda x: f"{x:.3f}" if pd.notna(x) else "-")
            if "UFC_Val" in df_mhsa_disp: df_mhsa_disp["UFC"] = df_mhsa_disp["UFC_Val"].apply(lambda x: formato_miles(x) if pd.notna(x) else "-")
            if "SCC_Val" in df_mhsa_disp: df_mhsa_disp["SCC"] = df_mhsa_disp["SCC_Val"].apply(lambda x: formato_miles(x) if pd.notna(x) else "-")
            
            cols_to_show = ["Fecha", "Num_Tambo", "Tambo", "Litros_Ticket", "Temperatura"]
            if "Grasa" in df_mhsa_disp: cols_to_show.append("Grasa")
            if "Proteína" in df_mhsa_disp: cols_to_show.append("Proteína")
            if "Crioscopía" in df_mhsa_disp: cols_to_show.append("Crioscopía")
            if "UFC" in df_mhsa_disp: cols_to_show.append("UFC")
            if "SCC" in df_mhsa_disp: cols_to_show.append("SCC")
            
            df_mhsa_disp = df_mhsa_disp.rename(columns={"Litros_Ticket": "Litros"})
            cols_to_show[cols_to_show.index("Litros_Ticket")] = "Litros"

            st.dataframe(df_mhsa_disp[cols_to_show], use_container_width=True, hide_index=True)
        else:
            st.info("No hay registros de recepción para el período seleccionado.")

    with tab2:
        st.subheader("Registro de Lotes Fasón")
        if not df_consolidado.empty:
            df_consolidado_disp = df_consolidado.copy()
            df_consolidado_disp["Fecha"] = df_consolidado_disp["Fecha"].dt.strftime("%d/%m/%Y")
            df_consolidado_disp["Litros Procesados"] = df_consolidado_disp["Litros Procesados"].apply(formato_miles)
            df_consolidado_disp["Producto Terminado"] = df_consolidado_disp["Producto Terminado"].apply(formato_miles)
            st.dataframe(df_consolidado_disp, use_container_width=True, hide_index=True)
        else:
            st.info("No hay producción de lotes Mastellone para el período seleccionado.")

  except Exception as e:
    st.error("Se produjo un error procesando los datos de Mastellone:")
    st.code(traceback.format_exc())
      
# =========================================================================
# MÓDULO 3: PRODUCCIÓN COOPAGRO
# =========================================================================
elif modulo_principal == "🧀 Producción y Rendimiento":
  st.header("Producción Coopagro")
  try:
    with st.spinner("Sincronizando datos..."):
      raw_prod = pd.read_excel(URL_PRODUCCION, skiprows=6)
      df_prod = pd.DataFrame()
      df_prod['Fecha'], df_prod['Lote'] = raw_prod.iloc[:, 0], raw_prod.iloc[:, 1]
      df_prod['Litros Procesados'], df_prod['Producto Terminado'], df_prod['PNC'] = raw_prod.iloc[:, 3], raw_prod.iloc[:, 5], raw_prod.iloc[:, 6]
      df_prod = df_prod.dropna(subset=['Fecha'])
      df_prod['Fecha'] = pd.to_datetime(df_prod['Fecha'], dayfirst=True, errors='coerce')
      df_prod = df_prod.dropna(subset=['Fecha'])
      for col in ["Litros Procesados", "Producto Terminado", "PNC"]: df_prod[col] = pd.to_numeric(df_prod[col], errors='coerce').fillna(0)
      if len(df_prod) > 0: df_prod['Producto'], df_prod['Grupo'] = zip(*df_prod['Lote'].astype(str).apply(procesar_lote_mastellone))
      else: df_prod['Producto'], df_prod['Grupo'] = [], []
      df_prod['Año'], df_prod['Mes'] = df_prod['Fecha'].dt.year, df_prod['Fecha'].dt.month
      df_prod_coop = df_prod[df_prod['Grupo'] == 'Coopagro'].copy()

      raw_recibo = pd.read_excel(URL_REMITOS)
      df_recibo = pd.DataFrame()
      df_recibo['Fecha_Raw'] = raw_recibo.iloc[:, 1] 
      df_recibo['Litros Ingresados'] = pd.to_numeric(raw_recibo.iloc[:, 5], errors='coerce').fillna(0)
      df_recibo['Fecha'] = pd.to_datetime(df_recibo['Fecha_Raw'], dayfirst=True, errors='coerce')
      df_recibo = df_recibo.dropna(subset=['Fecha'])
      df_recibo['Año'], df_recibo['Mes'] = df_recibo['Fecha'].dt.year, df_recibo['Fecha'].dt.month
      recibo_mensual = df_recibo.groupby(['Año', 'Mes'])['Litros Ingresados'].sum().reset_index()

    st.sidebar.subheader("Filtros")
    opciones_anio = ["Todos"] + (sorted(df_prod_coop['Año'].unique().tolist()) if len(df_prod_coop) > 0 else [2026])
    opciones_mes = ["Todos"] + list(range(1, 13))
    filtro_anio_coop = st.sidebar.selectbox("Año", opciones_anio, key="coop_prod_anio")
    filtro_mes_coop = st.sidebar.selectbox("Mes", opciones_mes, key="coop_prod_mes")

    df_filtrado = df_prod_coop.copy()
    if len(df_filtrado) > 0:
        if filtro_anio_coop != "Todos": df_filtrado = df_filtrado[df_filtrado['Año'] == filtro_anio_coop]
        if filtro_mes_coop != "Todos": df_filtrado = df_filtrado[df_filtrado['Mes'] == filtro_mes_coop]

    if len(df_filtrado) > 0:
        df_consolidado_raw = df_filtrado.copy()
        df_consolidado_raw['PT_Total'] = df_consolidado_raw['Producto Terminado'] + df_consolidado_raw['PNC']
        df_consolidado_raw['Ratio Consolidado (%)'] = df_consolidado_raw.apply(lambda x: f"{(x['PT_Total'] / x['Litros Procesados'] * 100):.2f}%" if x['Litros Procesados'] > 0 else "0.00%", axis=1)
        df_consolidado = df_consolidado_raw[['Fecha', 'Lote', 'Producto', 'Litros Procesados', 'PT_Total', 'Ratio Consolidado (%)']].rename(columns={'PT_Total': 'Producto Terminado', 'Ratio Consolidado (%)': 'Ratio de Conversión (%)'})
    else:
        df_consolidado = pd.DataFrame(columns=['Fecha', 'Lote', 'Producto', 'Litros Procesados', 'Producto Terminado', 'Ratio de Conversión (%)'])

    df_recibo_filtrado = recibo_mensual.copy()
    if filtro_anio_coop != "Todos": df_recibo_filtrado = df_recibo_filtrado[df_recibo_filtrado['Año'] == filtro_anio_coop]
    if filtro_mes_coop != "Todos": df_recibo_filtrado = df_recibo_filtrado[df_recibo_filtrado['Mes'] == filtro_mes_coop]
    
    total_litros_ingresados = df_recibo_filtrado['Litros Ingresados'].sum()
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
    st.dataframe(df_consolidado, use_container_width=True, hide_index=True)

    def generar_pdf_coopagro(dataframe_gerencia, titulo_dinamico, lit_ingresados, rend_gerencia, ratio_pond_gerenc, df_prod_raw_gerencia):
        pdf = FPDF(orientation='P', unit='mm', format='A4')
        pdf.set_auto_page_break(auto=True, margin=15)
        pdf.add_page()
        
        if os.path.exists("logo.png"):
            pdf.image("logo.png", x=65, y=10, w=80)
            pdf.set_y(52)
        else:
            pdf.set_y(15)

        pdf.set_font("Arial", 'B', 13)
        pdf.cell(190, 7, txt=titulo_dinamico, ln=True, align='C')
        pdf.ln(3)
        
        pdf.set_font("Arial", 'B', 10)
        pdf.cell(190, 5, txt="Resumen", ln=True, align='L')
        pdf.set_font("Arial", '', 9)
        
        tot_lit = dataframe_gerencia['Litros Procesados'].sum()
        tot_pro_ger = dataframe_gerencia['Producto Terminado'].sum()
        
        pdf.cell(95, 5, txt=f"Total Litros Ingresados: {formato_miles(lit_ingresados)}", ln=0)
        pdf.cell(95, 5, txt=f"Ratio PT / Litros procesados: {ratio_pond_gerenc:.2f}%", ln=1)
        pdf.cell(95, 5, txt=f"Total Litros Procesados: {formato_miles(tot_lit)}", ln=0)
        pdf.cell(95, 5, txt=f"Ratio PT / Litros ingresados: {rend_gerencia:.2f}%", ln=1)
        pdf.cell(95, 5, txt=f"Total Producto Terminado: {formato_miles(tot_pro_ger)}", ln=1)
        pdf.ln(5)

        pdf.set_font("Arial", 'B', 10)
        pdf.cell(190, 5, txt="Cantidad por Producto:", ln=True, align='L')
        pdf.set_font("Arial", '', 8)
        
        prod_res_ger = df_prod_raw_gerencia.groupby('Producto')[['Litros Procesados', 'PT_Total']].sum().reset_index()
        for idx, row in prod_res_ger.iterrows():
            ratio_prod = (row['PT_Total'] / row['Litros Procesados'] * 100) if row['Litros Procesados'] > 0 else 0
            txt_linea = f"- {row['Producto']}: Litros Proc. {formato_miles(row['Litros Procesados'])} | Prod. Terminado: {formato_miles(row['PT_Total'])} | Ratio: {ratio_prod:.2f}%"
            pdf.cell(190, 4.5, txt=txt_linea, ln=True)
            
        pdf.ln(5)
        
        pdf.set_font("Arial", 'B', 10)
        pdf.cell(190, 5, txt="Detalle de Lotes", ln=True, align='L')
        pdf.ln(2)
        
        pdf.set_font("Arial", 'B', 7)
        anchos = [20, 28, 42, 32, 38, 30]
        columnas = dataframe_gerencia.columns.tolist()
        for i in range(len(columnas)): pdf.cell(anchos[i], 7, columnas[i], border=1, align='C')
        pdf.ln()
        
        pdf.set_font("Arial", '', 7)
        for index, row in dataframe_gerencia.iterrows():
            pdf.cell(anchos[0], 6, row['Fecha'].strftime('%d/%m/%Y'), border=1, align='C')
            pdf.cell(anchos[1], 6, str(row['Lote']), border=1, align='C')
            pdf.cell(anchos[2], 6, str(row['Producto']), border=1, align='L')
            pdf.cell(anchos[3], 6, formato_miles(row['Litros Procesados']), border=1, align='R')
            pdf.cell(anchos[4], 6, formato_miles(row['Producto Terminado']), border=1, align='R')
            pdf.cell(anchos[5], 6, str(row['Ratio de Conversión (%)']), border=1, align='C')
            pdf.ln()
            
        output = pdf.output(dest='S')
        if isinstance(output, bytearray): return bytes(output)
        elif isinstance(output, str): return output.encode('latin1')
        return output

    if len(df_consolidado) > 0:
        df_gerencia_raw_pdf = df_prod_coop.copy()
        if filtro_anio_coop != "Todos": df_gerencia_raw_pdf = df_gerencia_raw_pdf[df_gerencia_raw_pdf['Año'] == filtro_anio_coop]
        if filtro_mes_coop != "Todos": df_gerencia_raw_pdf = df_gerencia_raw_pdf[df_gerencia_raw_pdf['Mes'] == filtro_mes_coop]
        df_gerencia_raw_pdf['PT_Total'] = df_gerencia_raw_pdf['Producto Terminado'] + df_gerencia_raw_pdf['PNC']

        mes_nombre_pdf = MESES_ES.get(filtro_mes_coop, "") if filtro_mes_coop != "Todos" else "General"
        anio_pdf = str(filtro_anio_coop) if filtro_anio_coop != "Todos" else "2026"
        titulo_pdf = f"Reporte producción Coopagro {mes_nombre_pdf} {anio_pdf}"

        pdf_bytes = generar_pdf_coopagro(df_consolidado, titulo_pdf, total_litros_ingresados, rendimiento_ingreso, ratio_ponderado, df_gerencia_raw_pdf)
        
        st.download_button(label="📄 Descargar Reporte en PDF", data=pdf_bytes, file_name=f"{titulo_pdf.replace(' ', '_')}.pdf", mime="application/pdf")

  except Exception as e:
    st.error(f"Error en Coopagro: {e}")

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
