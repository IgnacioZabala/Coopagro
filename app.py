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

# =========================================================================
# CONFIGURACIÓN GENERAL Y BLINDAJE DE ENTORNO
# =========================================================================
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

# --- IDs de Google Drive y Sheets (Blindados) ---
FILE_ID_REMITOS = "19OVD6xBeK08o4cW1XrdMr54L1nciAJC2"
FILE_ID_MILKO = "1WR3orOFWXyyMqbVrKh792-8VBh2qN68O"
FILE_ID_BACSOMATIC = "1KeTle24zxjK-clKAuXsAOUzGkfBNXgI8"
FILE_ID_MASTELLONE = "19OVD6xBeK08o4cW1XrdMr54L1nciAJC2"
ID_PRODUCCION = "1EH1koI566Bll9b_bqk9Ya4TenOIfczjt"
SHEET_INSUMOS_ID = "1OY1g-dRIVzVbU_cL6C1UzCUCeCKUxbT6RiAGLX7-Kpo"
SHEET_MAESTRO_ID = "1VHJPBN1R2aECfni_5JKHACOZiC6LCFNKyOIPE7ye5yQ"

URL_REMITOS = f"https://drive.google.com/uc?export=download&id={FILE_ID_REMITOS}"
URL_MILKO = f"https://drive.google.com/uc?export=download&id={FILE_ID_MILKO}"
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
    df_lab_temp = pd.read_excel(u_lab, header=None, nrows=20, engine='openpyxl')
    header_row = encontrar_fila_encabezado(df_lab_temp, ["sample", "fat", "protein", "grasa"])
    df_lab = pd.read_excel(u_lab, header=header_row, engine='openpyxl')
    df_lab.columns = df_lab.columns.astype(str).str.strip()
  except Exception as e:
    st.sidebar.warning(f"No se pudo cargar el archivo Milko: {e}")

  try:
    df_bac_temp = pd.read_excel(u_bacsomatic, header=None, nrows=20, engine='openpyxl')
    header_row_bac = encontrar_fila_encabezado(df_bac_temp, ["id usuario", "ufc", "scc"])
    df_bacsomatic = pd.read_excel(u_bacsomatic, header=header_row_bac, engine='openpyxl')
    df_bacsomatic.columns = df_bacsomatic.columns.astype(str).str.strip()
  except Exception as e:
    st.sidebar.warning(f"No se pudo cargar el archivo Bacsomatic: {e}")

  return df_remitos_raw, df_contactos, df_lab, df_bacsomatic

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

def clasificar_lote_general(lote_str):
  if not isinstance(lote_str, str) or len(lote_str) < 8: return "Desconocido", "Otro"
  prod_code = lote_str[5:8]
  mapping_prod = {
      "288": "Muzzarella Exportacion Coop.",
      "125": "Muzarrella Piano",
      "488": "Tybo Coop.",
      "840": "Muzzarella Exportacion Mastellone",
  }
  prod_nombre = mapping_prod.get(prod_code, f"Producto ({prod_code})")
  grupo = "Mastellone" if prod_code == "840" else "Coopagro"
  return prod_nombre, grupo

# =========================================================================
# FUNCIONES DE PDF BLINDADAS (Logo y Espaciado Dinámico)
# =========================================================================
def generar_pdf_base(titulo: str, subtitulo: str, metricas: list, headers: list, df_datos: pd.DataFrame, filas_mapeo: list, usable_width: int = 190):
  pdf = FPDF(orientation="P", unit="mm", format="A4")
  pdf.set_auto_page_break(auto=True, margin=15)
  pdf.add_page()
  
  logo_y = 8
  logo_w = 50
  
  if os.path.exists("logo.png"):
    pdf.image("logo.png", x=(210 - logo_w) / 2, y=logo_y, w=logo_w)
    pdf.set_y(logo_y + 30 + 12)
  else:
    pdf.set_y(20)

  pdf.set_font("Arial", "B", 12)
  pdf.cell(0, 6, titulo, ln=True, align="C")
  pdf.ln(2)
  pdf.line(10, pdf.get_y(), 200, pdf.get_y())
  pdf.ln(4)

  pdf.set_font("Arial", "B", 11)
  pdf.cell(0, 7, subtitulo, ln=True)
  pdf.set_font("Arial", "", 10)
  for metrica in metricas: pdf.cell(0, 6, metrica, ln=True)

  pdf.ln(6)
  pdf.set_font("Arial", "B", 8)
  pdf.set_fill_color(200, 220, 255)

  suma_anchos = sum([w for _, w in headers])
  factor = usable_width / suma_anchos if suma_anchos > 0 else 1.0
  headers_ajustados = [(name, w * factor) for name, w in headers]

  for i, (col_name, col_w) in enumerate(headers_ajustados):
    pdf.cell(col_w, 8, col_name, 1, 1 if i == len(headers_ajustados) - 1 else 0, "C", fill=True)

  pdf.set_font("Arial", "", 8)
  for row in df_datos.itertuples(index=False):
    for i, (fn_mapeo, (_, col_w)) in enumerate(zip(filas_mapeo, headers_ajustados)):
      val = fn_mapeo(row)
      align = "L" if "Nombre" in headers_ajustados[i][0] or "Producto" in headers_ajustados[i][0] or "Insumo" in headers_ajustados[i][0] else "C"
      pdf.cell(col_w, 7, str(val), 1, 1 if i == len(headers_ajustados) - 1 else 0, align)

  output = pdf.output(dest="S")
  if isinstance(output, bytearray): return bytes(output)
  elif isinstance(output, str): return output.encode('latin1')
  return output

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
    df_raw, df_contactos_raw, df_lab_raw, df_bac_raw = cargar_datos_coopagro(URL_REMITOS, URL_MILKO, URL_BACSOMATIC)
    if df_raw.empty: st.error("El archivo de remitos está vacío o no se pudo acceder."); st.stop()

    df_contactos = pd.DataFrame()
    if not df_contactos_raw.empty:
      # Buscar fila de encabezado dinámicamente para la hoja Código Tambos
      header_idx = -1
      for i, row in df_contactos_raw.head(10).iterrows():
          row_str = " ".join([str(x).lower() for x in row if pd.notna(x)])
          if "tambo" in row_str and ("codigo" in row_str or "código" in row_str):
              header_idx = i
              break
      
      if header_idx != -1:
          df_c_temp = df_contactos_raw.iloc[header_idx+1:].copy()
          df_c_temp.columns = df_contactos_raw.iloc[header_idx].astype(str).str.strip().str.lower().str.replace("ó", "o")
      else:
          df_c_temp = df_contactos_raw.copy()
          df_c_temp.columns = df_c_temp.columns.astype(str).str.strip().str.lower().str.replace("ó", "o")

      col_codigo = next((c for c in df_c_temp.columns if "codigo" in c and "viejo" not in c), None)
      if not col_codigo and len(df_c_temp.columns) > 1: col_codigo = df_c_temp.columns[1]
      
      col_tambo_nom = next((c for c in df_c_temp.columns if "tambo" in c), None)
      if not col_tambo_nom and len(df_c_temp.columns) > 2: col_tambo_nom = df_c_temp.columns[2]

      col_contacto = next((c for c in df_c_temp.columns if "contacto" in c or "nombre" in c and "tambo" not in c), None)
      if not col_contacto and len(df_c_temp.columns) > 3: col_contacto = df_c_temp.columns[3]
      
      col_email = next((c for c in df_c_temp.columns if "email" in c or "correo" in c), None)
      if not col_email and len(df_c_temp.columns) > 4: col_email = df_c_temp.columns[4]

      if col_codigo is not None:
        df_contactos["Num_Tambo"] = df_c_temp[col_codigo].apply(limpiar_tambo)
        if col_tambo_nom is not None: df_contactos["Tambo_Maestro"] = df_c_temp[col_tambo_nom]
        if col_contacto is not None: df_contactos["Contacto_Nombre"] = df_c_temp[col_contacto]
        if col_email is not None: df_contactos["Email"] = df_c_temp[col_email]

    df = df_raw.iloc[:, :10].copy()
    df.columns = ["Fecha", "N_Remito", "Num_Tambo", "Tambo", "Litros_Ticket", "Litros_Planilla", "Diferencia", "Temperatura", "Grasa", "Proteina"]
    df["Num_Tambo"] = df["Num_Tambo"].apply(limpiar_tambo)
    df["Fecha"] = pd.to_datetime(df["Fecha"], format="mixed", dayfirst=True, errors="coerce").dt.normalize()
    df = df.dropna(subset=["Fecha"])
    for col in ["Grasa", "Proteina", "Crioscopia", "UFC", "SCC"]:
      if col not in df.columns: df[col] = pd.NA

    # ---> REEMPLAZAR #REF! O NOMBRES ROTOS CON EL MAESTRO DE TAMBOS <---
    if not df_contactos.empty and "Tambo_Maestro" in df_contactos.columns:
        mapeo_nombres = dict(zip(df_contactos["Num_Tambo"], df_contactos["Tambo_Maestro"]))
        # Mapea el nombre desde el maestro, si no lo encuentra deja el original (o el #REF!)
        df["Tambo"] = df["Num_Tambo"].map(mapeo_nombres).combine_first(df["Tambo"])
        
    # Limpieza final por si queda algún #REF! suelto sin mapear en el maestro
    df["Tambo"] = df["Tambo"].replace("#REF!", "Desconocido")

    # Procesamiento Lab (Muestra 1 asegurada)
    if not df_lab_raw.empty:
      df_lab = df_lab_raw.copy()
      col_sample = next((c for c in df_lab.columns if any(x in c.lower() for x in ["sample", "number", "tambo", "muestra"])), df_lab.columns[0])
      df_lab["Num_Tambo"] = df_lab[col_sample].astype(str).str.split().str[0].apply(limpiar_tambo)
      df_lab["Fecha_Extraida"] = pd.to_datetime(df_lab[col_sample].astype(str).str.split().str[-1].apply(extraer_fecha_texto), errors="coerce")
      col_date = next((c for c in df_lab.columns if any(x in c.lower() for x in ["fecha", "date", "time", "analyzed"])), None)
      df_lab["Fecha"] = df_lab["Fecha_Extraida"].fillna(pd.to_datetime(df_lab[col_date], errors="coerce").dt.normalize() if col_date else pd.NaT)
      df_lab = df_lab.dropna(subset=["Fecha", "Num_Tambo"])
      
      df_lab["_sample_str"] = df_lab[col_sample].astype(str)
      df_lab = df_lab.sort_values(by=["_sample_str"]).drop_duplicates(subset=["Num_Tambo", "Fecha"], keep="first")
      
      map_cols = {}
      col_fat = next((c for c in df_lab.columns if "fat" in c.lower() or "grasa" in c.lower()), None)
      col_prot = next((c for c in df_lab.columns if "protein" in c.lower() or "proteina" in c.lower()), None)
      col_fp = next((c for c in df_lab.columns if c.lower() == "fp" or "crios" in c.lower()), None)
      if col_fat: map_cols[col_fat] = "Grasa_Lab"
      if col_prot: map_cols[col_prot] = "Proteina_Lab"
      if col_fp: map_cols[col_fp] = "Crioscopia_Lab"
      
      if map_cols:
        df_milko_clean = df_lab[["Num_Tambo", "Fecha"] + list(map_cols.keys())].rename(columns=map_cols)
        for c in map_cols.values(): df_milko_clean[c] = pd.to_numeric(df_milko_clean[c].astype(str).str.replace(",", "."), errors="coerce")
        df = pd.merge(df, df_milko_clean, on=["Num_Tambo", "Fecha"], how="left")
        if "Grasa_Lab" in df: df["Grasa"] = df["Grasa_Lab"].combine_first(df["Grasa"])
        if "Proteina_Lab" in df: df["Proteina"] = df["Proteina_Lab"].combine_first(df["Proteina"])
        if "Crioscopia_Lab" in df: df["Crioscopia"] = df["Crioscopia_Lab"].combine_first(df["Crioscopia"])

    # Procesamiento Bacsomatic (Muestra 1 asegurada)
    if not df_bac_raw.empty:
      df_bac = df_bac_raw.copy()
      col_id = next((c for c in df_bac.columns if any(x in c.lower() for x in ["id usuario", "sample", "tambo"])), df_bac.columns[0])
      df_bac["Num_Tambo"] = df_bac[col_id].astype(str).str.split().str[0].apply(limpiar_tambo)
      df_bac["Fecha_Extraida"] = pd.to_datetime(df_bac[col_id].astype(str).str.split().str[-1].apply(extraer_fecha_texto), errors="coerce")
      col_date_bac = next((c for c in df_bac.columns if any(x in c.lower() for x in ["fecha", "date", "analyzed"])), None)
      df_bac["Fecha"] = df_bac["Fecha_Extraida"].fillna(pd.to_datetime(df_bac[col_date_bac], errors="coerce").dt.normalize() if col_date_bac else pd.NaT)
      df_bac = df_bac.dropna(subset=["Fecha", "Num_Tambo"])
      
      df_bac["_id_str"] = df_bac[col_id].astype(str)
      df_bac = df_bac.sort_values(by=["_id_str"]).drop_duplicates(subset=["Num_Tambo", "Fecha"], keep="first")
      
      map_cols_bac = {}
      col_ufc = next((c for c in df_bac.columns if "ufc" in c.lower()), None)
      col_scc = next((c for c in df_bac.columns if any(x in c.lower() for x in ["scc", "celulas", "somáticas"])), None)
      if col_ufc: map_cols_bac[col_ufc] = "UFC_Val"
      if col_scc: map_cols_bac[col_scc] = "SCC_Val"
      
      if map_cols_bac:
        df_bac_clean = df_bac[["Num_Tambo", "Fecha"] + list(map_cols_bac.keys())].rename(columns=map_cols_bac)
        for c in map_cols_bac.values(): df_bac_clean[c] = pd.to_numeric(df_bac_clean[c].astype(str).str.replace(",", "."), errors="coerce")
        df = pd.merge(df, df_bac_clean, on=["Num_Tambo", "Fecha"], how="left")
        if "UFC_Val" in df: df["UFC"] = df["UFC_Val"].combine_first(df["UFC"])
        if "SCC_Val" in df: df["SCC"] = df["SCC_Val"].combine_first(df["SCC"])

    df["Fecha_Cierre_Viernes"] = df["Fecha"] + pd.to_timedelta((4 - df["Fecha"].dt.weekday) % 7, unit="D")
    df["Fecha_Inicio_Sabado"] = df["Fecha_Cierre_Viernes"] - pd.Timedelta(days=6)
    df["Ciclo_Semana"] = "Viernes " + df["Fecha_Cierre_Viernes"].dt.strftime("%d/%m/%Y") + " (Sáb " + df["Fecha_Inicio_Sabado"].dt.strftime("%d/%m/%Y") + " al Vie " + df["Fecha_Cierre_Viernes"].dt.strftime("%d/%m/%Y") + ")"
    df["AnioMes"] = df["Fecha"].dt.to_period("M")
    df = df.sort_values(by=["Num_Tambo", "Fecha", "N_Remito"])

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

        st.subheader("Ranking y Calidad por Tambo")
        
        ranking_data = []
        for (num_t, tambo_n), group in df_macro.groupby(["Num_Tambo", "Tambo"]):
            tot_l = group["Litros_Ticket"].sum()
            g_pond = calcular_promedio_ponderado(group, "Grasa")
            p_pond = calcular_promedio_ponderado(group, "Proteina")
            ranking_data.append({
                "Num_Tambo": str(num_t),
                "Tambo": str(tambo_n),
                "Litros_Ticket": tot_l,
                "Grasa_Ponderada": g_pond,
                "Proteina_Ponderada": p_pond
            })
        
        df_ranking = pd.DataFrame(ranking_data).sort_values("Litros_Ticket", ascending=False)
        
        def generar_pdf_panel_general_con_calidad(df_macro, periodo_titulo, total_litros, temp_prom, grasa_prom, prot_prom, ratio_gp, tambos_activos, df_ranking):
          ratio_str = f"{ratio_gp:.2f}".replace(".", ",") if pd.notna(ratio_gp) else "S/D"
          grasa_str = f"{grasa_prom:.2f}%".replace(".", ",") if pd.notna(grasa_prom) else "S/D"
          prot_str = f"{prot_prom:.2f}%".replace(".", ",") if pd.notna(prot_prom) else "S/D"

          metricas = [
              f"Tambos Activos: {tambos_activos} | Litros Totales: {formato_miles(total_litros)} L",
              f"Temp. Promedio: {formato_temp(temp_prom)} | Grasa Ponderada: {grasa_str} | Prot. Ponderada: {prot_str}",
              f"Ratio Grasa / Proteína: {ratio_str}",
              "Ranking y Calidad de Tambos por Volumen",
          ]
          headers = [("Código", 25), ("Nombre del Tambo", 85), ("Litros Totales", 35), ("Grasa Pond.", 22), ("Prot. Pond.", 23)]
          mapeo = [
              lambda r: str(r.Num_Tambo),
              lambda r: str(r.Tambo)[:28],
              lambda r: formato_miles(r.Litros_Ticket),
              lambda r: f"{r.Grasa_Ponderada:.2f}%".replace(".", ",") if pd.notna(r.Grasa_Ponderada) else "S/D",
              lambda r: f"{r.Proteina_Ponderada:.2f}%".replace(".", ",") if pd.notna(r.Proteina_Ponderada) else "S/D"
          ]
          return generar_pdf_base("Informe de Recolección y Calidad - Cooperativa", f"Período Evaluado: {periodo_titulo}", metricas, headers, df_ranking, mapeo)

        pdf_bytes = generar_pdf_panel_general_con_calidad(df_macro, periodo_texto, tot_litros, df_macro["Temperatura"].mean(), grasa_p, prot_p, ratio_gp, df_macro["Num_Tambo"].nunique(), df_ranking)
        st.download_button("📥 Descargar Informe PDF con Calidad", data=pdf_bytes, file_name=f"Informe_Calidad_{periodo_texto.replace(' ', '_')}.pdf", mime="application/pdf")
        
        df_ranking_show = df_ranking.copy()
        df_ranking_show["Litros_Ticket"] = df_ranking_show["Litros_Ticket"].apply(formato_miles)
        df_ranking_show["Grasa_Ponderada"] = df_ranking_show["Grasa_Ponderada"].apply(lambda x: f"{x:.2f}%".replace(".", ",") if pd.notna(x) else "S/D")
        df_ranking_show["Proteina_Ponderada"] = df_ranking_show["Proteina_Ponderada"].apply(lambda x: f"{x:.2f}%".replace(".", ",") if pd.notna(x) else "S/D")
        
        st.dataframe(df_ranking_show.rename(columns={
            "Tambo": "Nombre del Tambo", 
            "Num_Tambo": "Código", 
            "Litros_Ticket": "Litros Totales",
            "Grasa_Ponderada": "Grasa Ponderada",
            "Proteina_Ponderada": "Proteína Ponderada"
        }), hide_index=True, use_container_width=True)

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

          if v_temp and "Temperatura" in df_per: df_tabla_visual["Temperatura"] = df_per["Temperatura"].apply(lambda x: f"{x:.1f}°" if pd.notna(x) else "-")
          if v_grasa and "Grasa" in df_per: df_tabla_visual["Grasa"] = df_per["Grasa"].apply(lambda x: f"{x:.2f}%" if pd.notna(x) else "-")
          if v_prot and "Proteina" in df_per: df_tabla_visual["Proteína"] = df_per["Proteina"].apply(lambda x: f"{x:.2f}%" if pd.notna(x) else "-")
          if v_crios and "Crioscopia" in df_per: df_tabla_visual["Crioscopia"] = df_per["Crioscopia"].apply(lambda x: f"{x:.3f}" if pd.notna(x) else "-")
          if v_ufc and "UFC" in df_per: df_tabla_visual["UFC"] = df_per["UFC"].apply(lambda x: formato_miles(x) if pd.notna(x) else "-")
          if v_scc and "SCC" in df_per: df_tabla_visual["SCC"] = df_per["SCC"].apply(lambda x: formato_miles(x) if pd.notna(x) else "-")

          st.dataframe(df_tabla_visual, use_container_width=True, hide_index=True)

    elif vista_coop == "Envío Masivo Semanal":
      st.header("📤 Envío Masivo Semanal")
      st.info("Configuración de correos lista para enviar.")
  except Exception as e:
    st.error("Error en el Módulo Coopagro:")
    st.code(traceback.format_exc())

# =========================================================================
# MÓDULO 2: RECEPCIÓN Y CALIDAD MASTELLONE (FASÓN)
# =========================================================================
elif modulo_principal == "🚛 Recepción Mastellone (Fasón)":
  st.header("🚛 Recepción, Calidad y Producción Mastellone (Fasón)")
  import datetime
  now = datetime.datetime.now()

  def generar_pdf_mastellone_sin_logo(titulo, subtitulo, metricas, headers, df_datos, filas_mapeo, usable_width=190):
      from fpdf import FPDF
      pdf = FPDF(orientation="P", unit="mm", format="A4")
      pdf.set_auto_page_break(auto=True, margin=15)
      pdf.add_page()
      
      pdf.set_y(20)
      pdf.set_font("Arial", "B", 12)
      pdf.cell(0, 6, titulo, ln=True, align="C")
      pdf.ln(2)
      pdf.line(10, pdf.get_y(), 200, pdf.get_y())
      pdf.ln(4)

      pdf.set_font("Arial", "B", 11)
      pdf.cell(0, 7, subtitulo, ln=True)
      pdf.set_font("Arial", "", 10)
      for metrica in metricas: pdf.cell(0, 6, metrica, ln=True)

      pdf.ln(6)
      pdf.set_font("Arial", "B", 8)
      pdf.set_fill_color(200, 220, 255)

      suma_anchos = sum([w for _, w in headers])
      factor = usable_width / suma_anchos if suma_anchos > 0 else 1.0
      headers_ajustados = [(name, w * factor) for name, w in headers]

      for i, (col_name, col_w) in enumerate(headers_ajustados):
          pdf.cell(col_w, 8, col_name, 1, 1 if i == len(headers_ajustados) - 1 else 0, "C", fill=True)

      pdf.set_font("Arial", "", 8)
      for row in df_datos.itertuples(index=False):
          for i, (fn_mapeo, (_, col_w)) in enumerate(zip(filas_mapeo, headers_ajustados)):
              val = fn_mapeo(row)
              align = "L" if "Nombre" in headers_ajustados[i][0] or "Producto" in headers_ajustados[i][0] else "C"
              pdf.cell(col_w, 7, str(val), 1, 1 if i == len(headers_ajustados) - 1 else 0, align)

      output = pdf.output(dest="S")
      if isinstance(output, bytearray): return bytes(output)
      elif isinstance(output, str): return output.encode('latin1')
      return output
  
  try:
    with st.spinner("Sincronizando datos de Mastellone..."):
      xls_prod_mast = pd.ExcelFile(URL_PRODUCCION)
      hoja_prod_mast = "2026" if "2026" in xls_prod_mast.sheet_names else xls_prod_mast.sheet_names[-1]
      raw_prod = pd.read_excel(xls_prod_mast, sheet_name=hoja_prod_mast, skiprows=6)
      df_prod = pd.DataFrame()
      df_prod["Fecha"] = raw_prod.iloc[:, 0]
      df_prod["Lote"] = raw_prod.iloc[:, 1]
      df_prod["Litros Procesados"] = raw_prod.iloc[:, 3]
      df_prod["Producto Terminado"] = raw_prod.iloc[:, 5]
      df_prod["PNC"] = raw_prod.iloc[:, 6]
      df_prod = df_prod.dropna(subset=["Fecha", "Lote"])
      
      df_prod["Fecha"] = pd.to_datetime(df_prod["Fecha"], dayfirst=True, errors="coerce")
      df_prod = df_prod.dropna(subset=["Fecha"])
      for col in ["Litros Procesados", "Producto Terminado", "PNC"]: 
          df_prod[col] = pd.to_numeric(df_prod[col], errors="coerce").fillna(0)
      
      if len(df_prod) > 0: 
          df_prod["Producto"], df_prod["Grupo"] = zip(*df_prod["Lote"].astype(str).apply(clasificar_lote_general))
      else: 
          df_prod["Producto"], df_prod["Grupo"] = [], []
      
      df_prod["Año"] = df_prod["Fecha"].dt.year
      df_prod["Mes"] = df_prod["Fecha"].dt.month
      df_mastellone_prod = df_prod[df_prod["Grupo"] == "Mastellone"].copy()

      xls_mastellone = pd.ExcelFile(URL_MASTELLONE)
      sheet_mhsa = next((s for s in xls_mastellone.sheet_names if "mhsa" in s.lower()), xls_mastellone.sheet_names[1] if len(xls_mastellone.sheet_names) > 1 else xls_mastellone.sheet_names[0])
      
      df_mhsa_raw = pd.read_excel(URL_MASTELLONE, sheet_name=sheet_mhsa)
      
      df_mhsa = pd.DataFrame()
      
      fechas_col = df_mhsa_raw.iloc[:, 0]
      df_mhsa["Fecha"] = pd.to_datetime(fechas_col, format="%d/%m/%Y", errors="coerce").fillna(pd.to_datetime(fechas_col, dayfirst=True, errors="coerce")).dt.normalize()
      
      df_mhsa["N_Remito"] = df_mhsa_raw.iloc[:, 1]
      df_mhsa["Num_Tambo"] = df_mhsa_raw.iloc[:, 2].apply(limpiar_tambo)
      df_mhsa["Tambo"] = df_mhsa_raw.iloc[:, 3]
      df_mhsa["Litros_Ticket"] = pd.to_numeric(df_mhsa_raw.iloc[:, 4].astype(str).str.replace(',', '.'), errors="coerce").fillna(0)
      df_mhsa["Temperatura"] = pd.to_numeric(df_mhsa_raw.iloc[:, 7].astype(str).str.replace(',', '.'), errors="coerce") if len(df_mhsa_raw.columns) > 7 else pd.NaT
      
      df_mhsa = df_mhsa.dropna(subset=["Fecha", "Num_Tambo"])
      
      df_mhsa = df_mhsa[(df_mhsa["Fecha"].dt.year >= 2025) & (df_mhsa["Fecha"].dt.year <= 2028)]
      df_mhsa["Año"] = df_mhsa["Fecha"].dt.year
      df_mhsa["Mes"] = df_mhsa["Fecha"].dt.month

      _, _, df_lab_raw, df_bac_raw = cargar_datos_coopagro(URL_REMITOS, URL_MILKO, URL_BACSOMATIC)
      
      if not df_lab_raw.empty:
          df_lab_m = df_lab_raw.copy()
          col_sample = df_lab_m.columns[0]
          
          df_lab_m["Num_Tambo"] = df_lab_m[col_sample].astype(str).str.split().str[0].apply(limpiar_tambo)
          df_lab_m["Fecha_Extraida"] = pd.to_datetime(df_lab_m[col_sample].astype(str).str.split().str[-1].apply(extraer_fecha_texto), errors="coerce")
          col_date = next((c for c in df_lab_m.columns if any(x in c.lower() for x in ["fecha", "date", "analyzed"])), None)
          df_lab_m["Fecha"] = df_lab_m["Fecha_Extraida"].fillna(pd.to_datetime(df_lab_m[col_date], dayfirst=True, errors="coerce").dt.normalize() if col_date else pd.NaT)
          df_lab_m = df_lab_m.dropna(subset=["Fecha", "Num_Tambo"])
          
          df_lab_m["_sample_str"] = df_lab_m[col_sample].astype(str)
          df_lab_m = df_lab_m.sort_values(by=["_sample_str"]).drop_duplicates(subset=["Num_Tambo", "Fecha"], keep="first")
          
          map_cols = {}
          col_fat = next((c for c in df_lab_m.columns if "fat" in c.lower() or "grasa" in c.lower()), None)
          col_prot = next((c for c in df_lab_m.columns if "protein" in c.lower() or "proteina" in c.lower()), None)
          col_fp = next((c for c in df_lab_m.columns if "fp" in c.lower() or "crios" in c.lower()), None)
          if col_fat: map_cols[col_fat] = "Grasa_Lab"
          if col_prot: map_cols[col_prot] = "Proteina_Lab"
          if col_fp: map_cols[col_fp] = "Crioscopia_Lab"
          
          if map_cols:
              df_milko_clean = df_lab_m[["Num_Tambo", "Fecha"] + list(map_cols.keys())].rename(columns=map_cols)
              for c in map_cols.values(): 
                  df_milko_clean[c] = pd.to_numeric(df_milko_clean[c].astype(str).str.replace(",", "."), errors="coerce")
              df_mhsa = pd.merge(df_mhsa, df_milko_clean, on=["Num_Tambo", "Fecha"], how="left")
              if "Grasa_Lab" in df_mhsa: df_mhsa["Grasa"] = df_mhsa["Grasa_Lab"]
              if "Proteina_Lab" in df_mhsa: df_mhsa["Proteina"] = df_mhsa["Proteina_Lab"]
              if "Crioscopia_Lab" in df_mhsa: df_mhsa["Crioscopia"] = df_mhsa["Crioscopia_Lab"]

      # --- LECTURA EXACTA DEL BACSOMATIC (COLUMNAS F, Q y R) ---
      if not df_bac_raw.empty:
          df_bac_m = df_bac_raw.copy()
          
          # Buscamos la columna de identificación (Columna F -> índice 5)
          col_sample_bac = df_bac_m.columns[5] if len(df_bac_m.columns) > 5 else df_bac_m.columns[0]
          
          df_bac_m["Num_Tambo"] = df_bac_m[col_sample_bac].astype(str).str.split().str[0].apply(limpiar_tambo)
          df_bac_m["Fecha_Extraida"] = pd.to_datetime(df_bac_m[col_sample_bac].astype(str).str.split().str[-1].apply(extraer_fecha_texto), errors="coerce")
          
          col_date_bac = df_bac_m.columns[0] if len(df_bac_m.columns) > 0 else None
          df_bac_m["Fecha"] = df_bac_m["Fecha_Extraida"].fillna(pd.to_datetime(df_bac_m[col_date_bac], dayfirst=True, errors="coerce").dt.normalize() if col_date_bac else pd.NaT)
          df_bac_m = df_bac_m.dropna(subset=["Fecha", "Num_Tambo"])
          
          df_bac_m["_id_str"] = df_bac_m[col_sample_bac].astype(str)
          df_bac_m = df_bac_m.sort_values(by=["_id_str"]).drop_duplicates(subset=["Num_Tambo", "Fecha"], keep="first")
          
          df_bac_clean = pd.DataFrame()
          df_bac_clean["Num_Tambo"] = df_bac_m["Num_Tambo"]
          df_bac_clean["Fecha"] = df_bac_m["Fecha"]
          
          # Extracción directa por índices: Columna Q (16) = UFC | Columna R (17) = SCC
          if len(df_bac_m.columns) > 16:
              df_bac_clean["UFC_Val"] = pd.to_numeric(df_bac_m.iloc[:, 16].astype(str).str.replace(",", "."), errors="coerce")
          if len(df_bac_m.columns) > 17:
              df_bac_clean["SCC_Val"] = pd.to_numeric(df_bac_m.iloc[:, 17].astype(str).str.replace(",", "."), errors="coerce")
          
          df_bac_clean = df_bac_clean.dropna(subset=["Num_Tambo", "Fecha"])
          
          df_mhsa = pd.merge(df_mhsa, df_bac_clean, on=["Num_Tambo", "Fecha"], how="left")
          if "UFC_Val" in df_mhsa.columns: 
              df_mhsa["UFC"] = df_mhsa["UFC_Val"]
          if "SCC_Val" in df_mhsa.columns: 
              df_mhsa["SCC"] = df_mhsa["SCC_Val"]

    st.sidebar.subheader("Filtros Mastellone")
    anios_mhsa = df_mhsa["Año"].dropna().unique().tolist() if not df_mhsa.empty else []
    anios_prod = df_mastellone_prod["Año"].dropna().unique().tolist() if not df_mastellone_prod.empty else []
    todos_anios = sorted(list(set(anios_mhsa + anios_prod)))
    
    opciones_anio = ["Todos"] + (todos_anios if todos_anios else [now.year])
    filtro_anio = st.sidebar.selectbox("Año", opciones_anio, index=opciones_anio.index(now.year) if now.year in opciones_anio else 0, key="m_anio_mastellone")
    filtro_mes = st.sidebar.selectbox("Mes", ["Todos"] + list(range(1, 13)), index=now.month, key="m_mes_mastellone")

    df_filtrado = df_mastellone_prod.copy()
    if len(df_filtrado) > 0:
      if filtro_anio != "Todos": df_filtrado = df_filtrado[df_filtrado["Año"] == filtro_anio]
      if filtro_mes != "Todos": df_filtrado = df_filtrado[df_filtrado["Mes"] == filtro_mes]

    df_mhsa_f = df_mhsa.copy()
    if not df_mhsa_f.empty:
      if filtro_anio != "Todos": df_mhsa_f = df_mhsa_f[df_mhsa_f["Año"] == filtro_anio]
      if filtro_mes != "Todos": df_mhsa_f = df_mhsa_f[df_mhsa_f["Mes"] == filtro_mes]

    df_consolidado = pd.DataFrame()
    if len(df_filtrado) > 0:
      df_c_raw = df_filtrado.copy()
      df_c_raw["PT_Total"] = df_c_raw["Producto Terminado"] + df_c_raw["PNC"]
      df_c_raw["Ratio"] = df_c_raw.apply(lambda x: f"{(x['PT_Total'] / x['Litros Procesados'] * 100):.2f}%" if x["Litros Procesados"] > 0 else "0.00%", axis=1)
      df_consolidado = df_c_raw[["Fecha", "Lote", "Producto", "Litros Procesados", "PT_Total", "Ratio"]].rename(columns={"PT_Total": "Producto Terminado", "Ratio": "Ratio de Conversión (%)"})

    total_litros_ingresados = df_mhsa_f["Litros_Ticket"].sum() if not df_mhsa_f.empty else 0.0
    total_litros_proc = df_filtrado["Litros Procesados"].sum() if len(df_filtrado) > 0 else 0
    total_prod_consolidado = df_consolidado["Producto Terminado"].sum() if not df_consolidado.empty else 0
    ratio_ponderado = (total_prod_consolidado / total_litros_proc * 100) if total_litros_proc > 0 else 0

    c1, c2, c3 = st.columns(3)
    c1.metric("Litros Ingresados (MHSA)", formato_miles(total_litros_ingresados))
    c2.metric("Litros Procesados", formato_miles(total_litros_proc))
    c3.metric("Total Producto Terminado", formato_miles(total_prod_consolidado))
    
    c4, _, _ = st.columns(3)
    c4.metric("Ratio PT / Procesados", f"{ratio_ponderado:.2f}%")

    tab1, tab2 = st.tabs(["📑 Recepción y Calidad (MHSA)", "🏭 Producción Fasón"])
    
    with tab1:
        st.subheader("Recepción y Calidad de Tambos Mastellone")
        if not df_mhsa_f.empty:
            df_m_disp = df_mhsa_f.sort_values(by=["Fecha", "Num_Tambo"]).copy()
            df_m_disp["Fecha"] = df_m_disp["Fecha"].dt.strftime("%d/%m/%Y")
            df_m_disp["Litros"] = df_m_disp["Litros_Ticket"].apply(formato_miles)
            df_m_disp["Temperatura"] = df_m_disp["Temperatura"].apply(lambda x: f"{x:.1f}°" if pd.notna(x) else "-")
            
            if "Grasa" in df_m_disp: df_m_disp["Grasa"] = df_m_disp["Grasa"].apply(lambda x: f"{x:.2f}%" if pd.notna(x) else "-")
            if "Proteina" in df_m_disp: df_m_disp["Proteína"] = df_m_disp["Proteina"].apply(lambda x: f"{x:.2f}%" if pd.notna(x) else "-")
            if "Crioscopia" in df_m_disp: df_m_disp["Crioscopía"] = df_m_disp["Crioscopia"].apply(lambda x: f"{x:.3f}" if pd.notna(x) else "-")
            
            # Formato de valores Bacsomatic como strings
            if "UFC" in df_m_disp: df_m_disp["UFC"] = df_m_disp["UFC"].apply(lambda x: formato_miles(x) if pd.notna(x) else "-")
            if "SCC" in df_m_disp: df_m_disp["SCC"] = df_m_disp["SCC"].apply(lambda x: formato_miles(x) if pd.notna(x) else "-")
            
            df_m_disp = df_m_disp.rename(columns={"Num_Tambo": "Num Tambo"})
            
            cols = ["Fecha", "Num Tambo", "Tambo", "Litros", "Temperatura"]
            for col_extra in ["Grasa", "Proteína", "Crioscopía", "UFC", "SCC"]:
                if col_extra in df_m_disp: cols.append(col_extra)
            
            # --- LÓGICA DE ESTILOS PARA COLORES ROJOS (UFC > 200, SCC > 400) ---
            def highlight_bacsomatic(val, threshold):
                try:
                    if pd.notna(val) and str(val) != "-":
                        num_val = float(str(val).replace(".", "").replace(",", "."))
                        if num_val > threshold:
                            return 'color: red; font-weight: bold'
                except:
                    pass
                return ''
                
            df_style = df_m_disp[cols].style
            if "UFC" in cols:
                df_style = df_style.applymap(lambda x: highlight_bacsomatic(x, 200), subset=["UFC"])
            if "SCC" in cols:
                df_style = df_style.applymap(lambda x: highlight_bacsomatic(x, 400), subset=["SCC"])
            
            # Mostramos la tabla con estilos en la interfaz
            st.dataframe(df_style, use_container_width=True, hide_index=True)

            st.markdown("---")
            df_pdf_rec = df_m_disp.rename(columns={"Proteína": "Proteina", "Crioscopía": "Crioscopia"})
            
            headers_pdf_rec = [("Fecha", 25), ("Tambo", 70), ("Litros", 25), ("Temp", 20), ("Grasa", 25), ("Proteina", 25)]
            mapeo_pdf_rec = [
                lambda r: r.Fecha if pd.notna(r.Fecha) else "",
                lambda r: str(r.Tambo)[:25],
                lambda r: str(r.Litros),
                lambda r: str(getattr(r, "Temperatura", "-")),
                lambda r: str(getattr(r, "Grasa", "-")),
                lambda r: str(getattr(r, "Proteina", "-"))
            ]
            
            mes_str = MESES_ES.get(filtro_mes, str(filtro_mes)) if filtro_mes != "Todos" else "Todos"
            subt_rec = f"Período: {mes_str} {filtro_anio}"
            metricas_rec = [f"Total Litros Ingresados: {formato_miles(total_litros_ingresados)} L"]
            
            pdf_rec_bytes = generar_pdf_mastellone_sin_logo("Reporte de Recepción y Calidad MHSA", subt_rec, metricas_rec, headers_pdf_rec, df_pdf_rec, mapeo_pdf_rec)
            st.download_button("📥 Descargar Reporte Recepción PDF", data=pdf_rec_bytes, file_name="Reporte_Recepcion_MHSA.pdf", mime="application/pdf")
        else:
            st.info("No hay registros de recepción MHSA para el período seleccionado.")

    with tab2:
        st.subheader("Registro de Lotes Fasón (Mastellone - Código 840)")
        if not df_consolidado.empty:
            df_c_disp = df_consolidado.copy()
            df_c_disp["Fecha"] = df_c_disp["Fecha"].dt.strftime("%d/%m/%Y")
            df_c_disp["Litros Procesados"] = df_c_disp["Litros Procesados"].apply(formato_miles)
            df_c_disp["Producto Terminado"] = df_c_disp["Producto Terminado"].apply(formato_miles)
            st.dataframe(df_c_disp, use_container_width=True, hide_index=True)
            
            headers_pdf = [("Fecha", 25), ("Lote", 35), ("Producto", 65), ("Litros Proc.", 25), ("Prod. Term.", 25), ("Ratio", 15)]
            mapeo_pdf = [
                lambda r: r.Fecha if pd.notna(r.Fecha) else "",
                lambda r: str(r.Lote),
                lambda r: str(r.Producto),
                lambda r: str(r.Litros_Procesados),
                lambda r: str(r.Producto_Terminado),
                lambda r: str(r.Ratio_Conversion)
            ]

            df_consolidado_pdf = df_c_disp.rename(columns={
                "Litros Procesados": "Litros_Procesados",
                "Producto Terminado": "Producto_Terminado",
                "Ratio de Conversión (%)": "Ratio_Conversion",
            })

            ratio_ingresados = (total_prod_consolidado / total_litros_ingresados * 100) if total_litros_ingresados > 0 else 0
            mes_nombre_m = MESES_ES.get(filtro_mes, str(filtro_mes)) if filtro_mes != "Todos" else "Todos"
            subt_mast = f"Período: {mes_nombre_m} {filtro_anio}"
            
            metricas_prod = [
                f"Total Litros Ingresados: {formato_miles(total_litros_ingresados)} L",
                f"Total Litros Procesados: {formato_miles(total_litros_proc)} L",
                f"Total Producto Terminado: {formato_miles(total_prod_consolidado)} kg",
                f"Rendimiento (PT / Procesados): {ratio_ponderado:.2f}%",
                f"Rendimiento (PT / Ingresados): {ratio_ingresados:.2f}%"
            ]
            
            pdf_mast_bytes = generar_pdf_mastellone_sin_logo("Reporte de Producción Fasón - Mastellone", subt_mast, metricas_prod, headers_pdf, df_consolidado_pdf, mapeo_pdf)
            st.download_button("📥 Descargar Reporte Producción PDF", data=pdf_mast_bytes, file_name="Reporte_Produccion_Mastellone.pdf", mime="application/pdf")
        else:
            st.info("No hay producción de lotes Mastellone para el período seleccionado.")

  except Exception as e:
    st.error("Se produjo un error procesando los datos de Mastellone:")
    st.code(traceback.format_exc())
      
# =========================================================================
# MÓDULO 3: PRODUCCIÓN Y RENDIMIENTO COOPAGRO
# =========================================================================
elif modulo_principal == "🧀 Producción y Rendimiento":
  st.header("🧀 Producción y Rendimiento Coopagro")
  try:
    with st.spinner("Sincronizando datos de producción y recepción Coopagro..."):
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
      df_prod = df_prod.dropna(subset=['Fecha', 'Lote'])
      
      df_prod['Fecha'] = pd.to_datetime(df_prod['Fecha'], format="mixed", dayfirst=True, errors='coerce')
      df_prod = df_prod.dropna(subset=['Fecha'])
      
      for col in ["Litros Procesados", "Producto Terminado", "PNC"]: 
          df_prod[col] = pd.to_numeric(df_prod[col], errors='coerce').fillna(0)

      df_prod["Producto"], df_prod["Grupo"] = zip(*df_prod['Lote'].apply(clasificar_lote_general))
      df_prod['Año'] = df_prod['Fecha'].dt.year
      df_prod['Mes'] = df_prod['Fecha'].dt.month
      df_prod_coop = df_prod[df_prod['Grupo'] == 'Coopagro'].copy()

      url_remitos_limpia = f"{URL_REMITOS}&t={int(time.time())}"
      xls_remitos_p = pd.ExcelFile(url_remitos_limpia)
      hoja_remitos_p = next((s for s in xls_remitos_p.sheet_names if "od-pro-03" in s.lower()), xls_remitos_p.sheet_names[0])
      raw_recibo = pd.read_excel(xls_remitos_p, sheet_name=hoja_remitos_p, skiprows=4, usecols="B:K")
      df_recibo = pd.DataFrame()
      df_recibo['Fecha'] = pd.to_datetime(raw_recibo.iloc[:, 0], format="mixed", dayfirst=True, errors='coerce').dt.normalize()
      df_recibo['Litros Ingresados'] = pd.to_numeric(raw_recibo.iloc[:, 4], errors='coerce').fillna(0)
      df_recibo = df_recibo.dropna(subset=['Fecha'])
      df_recibo['Año'] = df_recibo['Fecha'].dt.year
      df_recibo['Mes'] = df_recibo['Fecha'].dt.month

    st.sidebar.subheader("Filtros Producción Coopagro")
    anios_p = sorted(df_prod_coop['Año'].unique().tolist()) if not df_prod_coop.empty else [2026]
    f_anio_p = st.sidebar.selectbox("Año Producción", ["Todos"] + anios_p, key="p_anio_coop")
    f_mes_p = st.sidebar.selectbox("Mes Producción", ["Todos"] + list(range(1, 13)), key="p_mes_coop")

    df_p_filtered = df_prod_coop.copy()
    if f_anio_p != "Todos": df_p_filtered = df_p_filtered[df_p_filtered['Año'] == f_anio_p]
    if f_mes_p != "Todos": df_p_filtered = df_p_filtered[df_p_filtered['Mes'] == f_mes_p]

    df_r_filtered = df_recibo.copy()
    if f_anio_p != "Todos": df_r_filtered = df_r_filtered[df_r_filtered['Año'] == f_anio_p]
    if f_mes_p != "Todos": df_r_filtered = df_r_filtered[df_r_filtered['Mes'] == f_mes_p]

    tot_ingresados_c = df_r_filtered['Litros Ingresados'].sum() if not df_r_filtered.empty else 0
    tot_proc_c = df_p_filtered['Litros Procesados'].sum() if not df_p_filtered.empty else 0
    
    tot_pt_neto = df_p_filtered['Producto Terminado'].sum() if not df_p_filtered.empty else 0
    tot_pnc_c = df_p_filtered['PNC'].sum() if not df_p_filtered.empty else 0
    tot_pt_c = tot_pt_neto + tot_pnc_c
    
    ratio_proc_c = (tot_pt_c / tot_proc_c * 100) if tot_proc_c > 0 else 0
    ratio_ing_c = (tot_pt_c / tot_ingresados_c * 100) if tot_ingresados_c > 0 else 0

    col1, col2, col3 = st.columns(3)
    col1.metric("Litros Ingresados", formato_miles(tot_ingresados_c))
    col2.metric("Litros Procesados", formato_miles(tot_proc_c))
    col3.metric("Producto Terminado", formato_miles(tot_pt_c))

    col4, col5, col6 = st.columns(3)
    col4.metric("PNC (No Conforme)", formato_miles(tot_pnc_c))
    col5.metric("Rendimiento (PT / Procesados)", f"{ratio_proc_c:.2f}%")
    col6.metric("Rendimiento (PT / Ingresados)", f"{ratio_ing_c:.2f}%")

    st.subheader("Detalle de Lotes de Producción Coopagro")
    if not df_p_filtered.empty:
        df_p_show = df_p_filtered.copy()
        df_p_show['Fecha_Dt'] = df_p_show['Fecha']
        df_p_show['Fecha'] = df_p_show['Fecha_Dt'].dt.strftime('%d/%m/%Y')
        
        df_p_show['Litros Procesados Num'] = df_p_show['Litros Procesados']
        df_p_show['PT_Total_Lote'] = df_p_show['Producto Terminado'] + df_p_show['PNC']
        
        df_p_show['Litros Procesados'] = df_p_show['Litros Procesados'].apply(formato_miles)
        df_p_show['Producto Terminado'] = df_p_show['PT_Total_Lote'].apply(formato_miles)
        df_p_show['PNC'] = df_p_show['PNC'].apply(formato_miles)
        df_p_show['Rendimiento Lote'] = df_p_show.apply(lambda x: f"{(x['PT_Total_Lote'] / x['Litros Procesados Num'] * 100):.2f}%" if x['Litros Procesados Num'] > 0 else "0.00%", axis=1)
        
        st.dataframe(df_p_show[['Fecha', 'Lote', 'Producto', 'Litros Procesados', 'Producto Terminado', 'PNC', 'Rendimiento Lote']], use_container_width=True, hide_index=True)
        
        headers_pdf_c = [("Fecha", 25), ("Lote", 35), ("Producto", 65), ("Litros Proc.", 25), ("Producto Term.", 25), ("Rend.", 15)]
        mapeo_pdf_c = [
            lambda r: r.Fecha if pd.notna(r.Fecha) else "",
            lambda r: str(r.Lote),
            lambda r: str(r.Producto),
            lambda r: str(r.Litros_Procesados),
            lambda r: str(r.Producto_Terminado),
            lambda r: str(r.Rendimiento_Lote)
        ]
        
        mes_nombre_pdf = MESES_ES.get(f_mes_p, str(f_mes_p)) if f_mes_p != "Todos" else "Todos los meses"
        anio_pdf = str(f_anio_p) if f_anio_p != "Todos" else "General"
        subtitulo_periodo = f"Período: {mes_nombre_pdf} {anio_pdf}" if f_mes_p != "Todos" else f"Período: Año {anio_pdf}"

        metricas_pdf_coop = [
            f"Total Litros Ingresados: {formato_miles(tot_ingresados_c)} L",
            f"Total Litros Procesados: {formato_miles(tot_proc_c)} L",
            f"Total Producto Terminado: {formato_miles(tot_pt_c)} kg",
            f"Ratio Litros Procesados / Producto Terminado: {ratio_proc_c:.2f}%",
            f"Ratio Litros Ingresados / Producto Terminado: {ratio_ing_c:.2f}%"
        ]
        
        df_pdf_prep = df_p_show.rename(columns={"Litros Procesados": "Litros_Procesados", "Producto Terminado": "Producto_Terminado", "Rendimiento Lote": "Rendimiento_Lote"})
        pdf_coop_bytes = generar_pdf_base("Reporte de Producción y Rendimiento - Coopagro", subtitulo_periodo, metricas_pdf_coop, headers_pdf_c, df_pdf_prep, mapeo_pdf_c)
        
        prefijo_nombre = f"{f_mes_p}- " if f_mes_p != "Todos" else "General - "
        nombre_archivo_pdf = f"{prefijo_nombre}Reporte de Producción Coopagro {mes_nombre_pdf} {anio_pdf}.pdf"
        
        st.download_button("📥 Descargar Reporte PDF Coopagro", data=pdf_coop_bytes, file_name=nombre_archivo_pdf, mime="application/pdf")
    else:
        st.info("No hay registros de producción para el período seleccionado.")

  except Exception as e:
    st.error(f"Error en el Módulo de Producción y Rendimiento: {e}")
    st.code(traceback.format_exc())

# =========================================================================
# MÓDULO 4: INSUMOS, INVENTARIO Y COSTOS
# =========================================================================
elif modulo_principal == "📦 Insumos, Inventario y Costos":
    st.header("📦 Gestión de Insumos, Inventario y Costos Variables")
  
    def generar_pdf_costos_mes(periodo_texto, fecha_stock_txt, valor_total, tinas_mes, kilos_mes, costo_insumos_total_mes, costo_por_kilo, costo_aditivos_kilo, costo_envasado_kilo, costo_cip_kilo, df_datos):
        titulo = "Reporte Mensual de Insumos, Inventario y Costos"
        subtitulo = f"Período Evaluado: {periodo_texto}  |  Stock al: {fecha_stock_txt}"
        metricas = [
            f"Capital Inmovilizado (Stock Relevado): $ {formato_miles(valor_total)}",
            f"Tinas Producidas: {tinas_mes} | Kilos de Queso: {formato_miles(kilos_mes)} kg",
            f"Costo Total Insumos del Mes: $ {formato_miles(costo_insumos_total_mes)}",
            f"Costo Variable Total / Kilo Producido: $ {costo_por_kilo:,.2f}".replace(",", "."),
            f"   - Desglose / kg: Aditivos: $ {costo_aditivos_kilo:,.2f}".replace(",", ".") + f" | Envasado: $ {costo_envasado_kilo:,.2f}".replace(",", ".") + f" | CIP: $ {costo_cip_kilo:,.2f}".replace(",", ".")
        ]
        
        # Preparamos un DataFrame con nombres limpios para evitar el conflicto de itertuples()
        df_pdf = df_datos.copy()
        df_pdf['Stock_Fisico_Fmt'] = df_pdf['Stock Base Físico']
        df_pdf['Precio_Unit_Fmt'] = df_pdf['Precio Unitario']
        df_pdf['Valorizacion_Fisica_Fmt'] = df_pdf['Valorización Física ($)']

        headers = [("Insumo", 60), ("Cat.", 20), ("Stock Físico", 25), ("Unidad", 15), ("Precio Unit.", 30), ("Valorización ($)", 50)]
        mapeo = [
            lambda r: str(getattr(r, "Insumo", ""))[:28],
            lambda r: str(getattr(r, "Categoría", ""))[:10],
            lambda r: formato_miles(getattr(r, "Stock_Fisico_Fmt", 0)),
            lambda r: str(getattr(r, "Unidad", ""))[:8],
            lambda r: f"$ {getattr(r, 'Precio_Unit_Fmt', 0):,.2f}".replace(",", "."),
            lambda r: f"$ {getattr(r, 'Valorizacion_Fisica_Fmt', 0):,.2f}".replace(",", ".")
        ]
        return generar_pdf_base(titulo, subtitulo, metricas, headers, df_pdf, mapeo)

    try:
        with st.spinner("Descargando base de datos de insumos..."):
            import time
            import re
            
            # 1. Leer archivo de Movimientos (Stock e Ingresos)
            URL_MOVIMIENTOS = f"https://docs.google.com/spreadsheets/d/{SHEET_INSUMOS_ID}/export?format=xlsx"
            xls_movimientos = pd.ExcelFile(URL_MOVIMIENTOS)
            
            # 2. Leer archivo del Maestro
            URL_MAESTRO = f"https://docs.google.com/spreadsheets/d/{SHEET_MAESTRO_ID}/export?format=xlsx"
            xls_maestro = pd.ExcelFile(URL_MAESTRO)
            
            # Búsqueda inteligente de pestañas
            sheet_maestro = next((s for s in xls_maestro.sheet_names if "maestro" in s.lower()), None)
            sheet_stock = next((s for s in xls_movimientos.sheet_names if "stock" in s.lower()), None)
            sheet_ingresos = next((s for s in xls_movimientos.sheet_names if "ingresos" in s.lower()), None)
            
            df_maestro = pd.read_excel(xls_maestro, sheet_name=sheet_maestro).dropna(how='all') if sheet_maestro else pd.DataFrame()
            df_stock_form = pd.read_excel(xls_movimientos, sheet_name=sheet_stock).dropna(how='all') if sheet_stock else pd.DataFrame()
            df_ingresos_form = pd.read_excel(xls_movimientos, sheet_name=sheet_ingresos).dropna(how='all') if sheet_ingresos else pd.DataFrame()

            # Limpiar nombres de columnas
            df_maestro.columns = df_maestro.columns.astype(str).str.strip()
            df_stock_form.columns = df_stock_form.columns.astype(str).str.strip()
            df_ingresos_form.columns = df_ingresos_form.columns.astype(str).str.strip()

            # Llave de cruce ultra-robusta
            if 'Insumo' in df_maestro.columns:
                df_maestro['Insumo'] = df_maestro['Insumo'].astype(str).str.strip()
                df_maestro['Insumo_Key'] = df_maestro['Insumo'].str.lower().str.replace(r'[^a-z0-9]', '', regex=True)
            else:
                df_maestro['Insumo_Key'] = 'desconocido'
                st.error("⚠️ La columna 'Insumo' no se encontró en el Maestro.")

            # Asegurar columnas requeridas
            cols_requeridas_maestro = {
                'Insumo': 'Desconocido', 'Categoría': 'General', 'Unidad': 'un',
                'Precio Unitario': 0.0, 'Consumo por tina': 0.0, 'Stock de seguridad': 0.0,
                'Demora proveedor (dias)': 0.0, 'Consumo Diario Promedio': 0.0
            }
            for col, val_def in cols_requeridas_maestro.items():
                if col not in df_maestro.columns:
                    df_maestro[col] = val_def

            cols_num_m = ['Precio Unitario', 'Consumo por tina', 'Stock de seguridad', 'Demora proveedor (dias)', 'Consumo Diario Promedio']
            for col in cols_num_m:
                df_maestro[col] = pd.to_numeric(df_maestro[col].astype(str).str.replace(',', '.'), errors='coerce').fillna(0.0)

            # Filtros Sidebar
            st.sidebar.markdown("---")
            st.sidebar.subheader("📅 Filtro de Costos y Stock")
            anios_disponibles = [2026, 2027]
            meses_disponibles = list(MESES_ES.keys())
            
            filtro_anio_costo = st.sidebar.selectbox("Año de Análisis", anios_disponibles, index=0, key="costo_anio")
            filtro_mes_costo = st.sidebar.selectbox("Mes de Análisis", meses_disponibles, format_func=lambda m: MESES_ES[m], index=8, key="costo_mes")

            # =====================================================================
            # INGRESOS Y PRECIOS
            # =====================================================================
            if not df_ingresos_form.empty and 'Cantidad recibida' in df_ingresos_form.columns:
                df_ingresos_form['Insumo'] = df_ingresos_form['Insumo'].astype(str).str.strip()
                df_ingresos_form['Insumo_Key'] = df_ingresos_form['Insumo'].str.lower().str.replace(r'[^a-z0-9]', '', regex=True)
                
                df_ingresos_form['Cantidad recibida'] = pd.to_numeric(df_ingresos_form['Cantidad recibida'].astype(str).str.replace(',', '.'), errors='coerce').fillna(0.0)
                df_ingresos_form['Costo total (pesos)'] = pd.to_numeric(df_ingresos_form.get('Costo total (pesos)', 0).astype(str).str.replace(',', '.'), errors='coerce').fillna(0.0)
                
                col_fecha_ing = next((c for c in df_ingresos_form.columns if 'fecha' in c.lower()), next((c for c in df_ingresos_form.columns if 'marca' in c.lower()), None))
                
                if col_fecha_ing:
                    df_ingresos_form['Fecha_Dt'] = pd.to_datetime(df_ingresos_form[col_fecha_ing], format="mixed", dayfirst=True, errors='coerce')
                    df_ingresos_mes = df_ingresos_form[(df_ingresos_form['Fecha_Dt'].dt.year == filtro_anio_costo) & (df_ingresos_form['Fecha_Dt'].dt.month == filtro_mes_costo)]
                else:
                    df_ingresos_mes = df_ingresos_form

                compras_totales = df_ingresos_mes.groupby('Insumo_Key')['Cantidad recibida'].sum().reset_index()
                compras_totales.rename(columns={'Cantidad recibida': 'Total Ingresado'}, inplace=True)
                
                df_ingresos_form['Precio Calculado'] = df_ingresos_form.apply(lambda x: x['Costo total (pesos)'] / x['Cantidad recibida'] if x['Cantidad recibida'] > 0 else 0, axis=1)
                precios_nuevos = df_ingresos_form[df_ingresos_form['Precio Calculado'] > 0].groupby('Insumo_Key')['Precio Calculado'].last().reset_index()
            else:
                compras_totales = pd.DataFrame(columns=['Insumo_Key', 'Total Ingresado'])
                precios_nuevos = pd.DataFrame(columns=['Insumo_Key', 'Precio Calculado'])

            # =====================================================================
            # STOCK FÍSICO
            # =====================================================================
            fecha_maxima_stock = pd.NaT

            if not df_stock_form.empty:
                cols_base = [c for c in df_stock_form.columns if 'marca' in c.lower() or 'fecha' in c.lower()]
                nombres_viejos = ['insumo', 'stock fisico real', 'columna 5']
                cols_viejas = [c for c in df_stock_form.columns if c.lower() in nombres_viejos]
                df_stock_limpio = df_stock_form.drop(columns=cols_viejas, errors='ignore')
                cols_insumos = [c for c in df_stock_limpio.columns if c not in cols_base and not c.lower().startswith('unnamed')]

                df_stock_long = df_stock_limpio.melt(
                    id_vars=cols_base,
                    value_vars=cols_insumos,
                    var_name='Insumo_Form',
                    value_name='Stock fisico real'
                )

                df_stock_long = df_stock_long.dropna(subset=['Stock fisico real'])
                df_stock_long['Stock fisico real'] = pd.to_numeric(df_stock_long['Stock fisico real'].astype(str).str.replace(',', '.'), errors='coerce')
                df_stock_long = df_stock_long.dropna(subset=['Stock fisico real'])
                df_stock_long['Insumo_Key'] = df_stock_long['Insumo_Form'].astype(str).str.lower().str.replace(r'[^a-z0-9]', '', regex=True)

                col_fecha_stock = next((c for c in cols_base if 'fecha' in c.lower()), cols_base[0] if cols_base else None)
                
                if col_fecha_stock:
                    df_stock_long['Fecha_Dt'] = pd.to_datetime(df_stock_long[col_fecha_stock], format="mixed", dayfirst=True, errors='coerce')
                    mask = (df_stock_long['Fecha_Dt'].dt.year == filtro_anio_costo) & (df_stock_long['Fecha_Dt'].dt.month == filtro_mes_costo)
                    df_stock_mes = df_stock_long[mask]
                    
                    if not df_stock_mes.empty:
                        ultimo_stock = df_stock_mes.sort_values('Fecha_Dt').groupby('Insumo_Key').last().reset_index()
                        fecha_maxima_stock = df_stock_mes['Fecha_Dt'].max()
                    else:
                        ultimo_stock = pd.DataFrame(columns=['Insumo_Key', 'Stock Base Físico'])
                else:
                    ultimo_stock = df_stock_long.groupby('Insumo_Key').last().reset_index()
                
                ultimo_stock.rename(columns={'Stock fisico real': 'Stock Base Físico'}, inplace=True)
            else:
                ultimo_stock = pd.DataFrame(columns=['Insumo_Key', 'Stock Base Físico'])

            # =====================================================================
            # PRODUCCIÓN (Tinas y Kilos)
            # =====================================================================
            tinas_mes, kilos_mes, litros_procesados_mes = 0, 0.0, 0.0
            try:
                xls_prod_ins = pd.ExcelFile(URL_PRODUCCION)
                hoja_p_ins = "2026" if "2026" in xls_prod_ins.sheet_names else xls_prod_ins.sheet_names[-1]
                raw_p_ins = pd.read_excel(xls_prod_ins, sheet_name=hoja_p_ins, skiprows=6)
                df_p_ins = pd.DataFrame()
                df_p_ins['Fecha'] = pd.to_datetime(raw_p_ins.iloc[:, 0], format="mixed", dayfirst=True, errors='coerce')
                df_p_ins['Lote'] = raw_p_ins.iloc[:, 1]
                df_p_ins['Litros Procesados'] = pd.to_numeric(raw_p_ins.iloc[:, 3], errors='coerce').fillna(0)
                df_p_ins['Prod Terminado'] = pd.to_numeric(raw_p_ins.iloc[:, 5], errors='coerce').fillna(0)
                df_p_ins['PNC'] = pd.to_numeric(raw_p_ins.iloc[:, 6], errors='coerce').fillna(0)
                df_p_ins = df_p_ins.dropna(subset=['Fecha', 'Lote'])
                
                _, df_p_ins['Grupo'] = zip(*df_p_ins['Lote'].apply(clasificar_lote_general))
                df_p_mes = df_p_ins[(df_p_ins['Grupo'] == 'Coopagro') & (df_p_ins['Fecha'].dt.year == filtro_anio_costo) & (df_p_ins['Fecha'].dt.month == filtro_mes_costo)]
                
                litros_procesados_mes = df_p_mes['Litros Procesados'].sum()
                kilos_mes = (df_p_mes['Prod Terminado'] + df_p_mes['PNC']).sum()
                tinas_mes = round(litros_procesados_mes / 8000) if litros_procesados_mes > 0 else 0
            except Exception:
                pass

            # =====================================================================
            # CÁLCULOS MAESTROS Y MERGE
            # =====================================================================
            df_master_calc = pd.merge(df_maestro, ultimo_stock[['Insumo_Key', 'Stock Base Físico']], on='Insumo_Key', how='left').fillna(0)
            df_master_calc = pd.merge(df_master_calc, compras_totales, on='Insumo_Key', how='left').fillna(0)
            
            if not precios_nuevos.empty and 'Insumo_Key' in precios_nuevos.columns:
                df_master_calc = pd.merge(df_master_calc, precios_nuevos, on='Insumo_Key', how='left')
                if 'Precio Calculado' in df_master_calc.columns:
                    df_master_calc['Precio Unitario'] = df_master_calc['Precio Calculado'].combine_first(df_master_calc['Precio Unitario'])

            df_master_calc = df_master_calc[~df_master_calc['Insumo'].str.lower().isin(['nan', '0', 'desconocido', ''])]

            df_master_calc['Consumo Teórico Mes'] = tinas_mes * df_master_calc['Consumo por tina']
            
            # Stock Físico
            df_master_calc['Valorización Física ($)'] = df_master_calc['Stock Base Físico'] * df_master_calc['Precio Unitario']

            # Stock Proyectado
            df_master_calc['Stock Actual'] = (df_master_calc['Stock Base Físico'] + df_master_calc['Total Ingresado']) - df_master_calc['Consumo Teórico Mes']
            df_master_calc['Stock Actual'] = df_master_calc['Stock Actual'].apply(lambda x: max(0.0, x))
            
            df_master_calc['Punto de Pedido'] = (df_master_calc['Consumo Diario Promedio'] * df_master_calc['Demora proveedor (dias)']) + df_master_calc['Stock de seguridad']
            df_master_calc['Valorización Proyectada ($)'] = df_master_calc['Stock Actual'] * df_master_calc['Precio Unitario']
            
            df_master_calc['Estado'] = df_master_calc.apply(
                lambda x: '🔴 Crítico' if x['Stock Actual'] <= x['Stock de seguridad'] else ('🟡 Reponer' if x['Stock Actual'] <= x['Punto de Pedido'] else '🟢 Normal'), axis=1
            )

            # --- CÁLCULOS POR CATEGORÍA ---
            df_master_calc['Costo Total Insumos Mes'] = df_master_calc['Consumo Teórico Mes'] * df_master_calc['Precio Unitario']
            
            costo_insumos_total_mes = df_master_calc['Costo Total Insumos Mes'].sum()
            costo_por_kilo = (costo_insumos_total_mes / kilos_mes) if kilos_mes > 0 else 0.0
            
            df_aditivos = df_master_calc[df_master_calc['Categoría'].str.strip().str.lower() == 'aditivos']
            df_envasado = df_master_calc[df_master_calc['Categoría'].str.strip().str.lower() == 'envasado']
            df_cip = df_master_calc[df_master_calc['Categoría'].str.strip().str.lower() == 'cip']
            
            costo_aditivos_total = df_aditivos['Costo Total Insumos Mes'].sum()
            costo_envasado_total = df_envasado['Costo Total Insumos Mes'].sum()
            costo_cip_total = df_cip['Costo Total Insumos Mes'].sum()
            
            costo_aditivos_kilo = (costo_aditivos_total / kilos_mes) if kilos_mes > 0 else 0.0
            costo_envasado_kilo = (costo_envasado_total / kilos_mes) if kilos_mes > 0 else 0.0
            costo_cip_kilo = (costo_cip_total / kilos_mes) if kilos_mes > 0 else 0.0

        # =====================================================================
        # VISUALIZACIÓN EN TRES PESTAÑAS
        # =====================================================================
        tab_inv1, tab_inv2, tab_inv3 = st.tabs([
            "📊 Alertas y Reposición", 
            "💰 Stock Valorizado", 
            f"🧀 Costos y Rendimiento ({MESES_ES[filtro_mes_costo]} {filtro_anio_costo})"
        ])

        with tab_inv1:
            st.subheader(f"Control de Alertas — {MESES_ES[filtro_mes_costo]} {filtro_anio_costo}")
            
            st.info("**Guía rápida de lectura:**\n"
                    "• **Stock Proyectado:** Último recuento físico + Ingresos del mes - Consumo teórico de las tinas producidas en el mes.\n"
                    "• **Punto de Pedido:** Consumo promedio durante los días de demora del proveedor + tu stock de seguridad.\n"
                    "• **Estado:** Te avisa 🟡 **Reponer** si el stock proyectado perfora el punto de pedido, o 🔴 **Crítico** si toca tu stock de seguridad.")
            
            c2, c3, c4 = st.columns(3)
            c2.metric(f"Tinas ({MESES_ES[filtro_mes_costo][:3]})", tinas_mes)
            c3.metric("Insumos Críticos", len(df_master_calc[df_master_calc['Estado'] == '🔴 Crítico']))
            c4.metric("Insumos a Reponer", len(df_master_calc[df_master_calc['Estado'] == '🟡 Reponer']))

            df_mostrar = df_master_calc[['Insumo', 'Categoría', 'Stock Actual', 'Unidad', 'Punto de Pedido', 'Estado']].copy()
            df_mostrar = df_mostrar.rename(columns={'Stock Actual': 'Stock Proyectado'})
            st.dataframe(df_mostrar, use_container_width=True, hide_index=True)

        with tab_inv2:
            if pd.notna(fecha_maxima_stock):
                titulo_valorizado = f"Stock al {fecha_maxima_stock.day} de {MESES_ES[filtro_mes_costo]} de {filtro_anio_costo}"
            else:
                titulo_valorizado = f"Inventario y Valorización ({MESES_ES[filtro_mes_costo]} {filtro_anio_costo})"
                
            st.subheader(titulo_valorizado)
            
            valor_total_fisico = df_master_calc['Valorización Física ($)'].sum()
            st.metric("Capital Inmovilizado (Stock Relevado)", f"$ {formato_miles(valor_total_fisico)}")

            df_valorizado = df_master_calc[['Insumo', 'Categoría', 'Stock Base Físico', 'Unidad', 'Precio Unitario', 'Valorización Física ($)']].copy()
            df_valorizado = df_valorizado.rename(columns={'Stock Base Físico': 'Stock Físico (Relevado)'})
            df_valorizado = df_valorizado.sort_values('Valorización Física ($)', ascending=False)
            
            df_valorizado['Precio Unitario'] = df_valorizado['Precio Unitario'].apply(lambda x: f"$ {x:,.2f}".replace(",", "."))
            df_valorizado['Valorización Física ($)'] = df_valorizado['Valorización Física ($)'].apply(lambda x: f"$ {x:,.2f}".replace(",", "."))
            
            st.dataframe(df_valorizado, use_container_width=True, hide_index=True)

        with tab_inv3:
            st.subheader(f"Costos Variables ({MESES_ES[filtro_mes_costo]} {filtro_anio_costo})")
            
            k1, k2, k3 = st.columns(3)
            k1.metric("Kilos de Queso", f"{formato_miles(kilos_mes)} kg")
            k2.metric("Costo Total Insumos", f"$ {formato_miles(costo_insumos_total_mes)}")
            k3.metric("Costo Total / Kilo Producido", f"$ {costo_por_kilo:,.2f}".replace(",", "."))
            
            st.markdown("---")
            st.markdown("##### Desglose de Costo por Kilo según Categoría")
            c_ad, c_env, c_cip = st.columns(3)
            c_ad.metric("🧀 Aditivos / kg", f"$ {costo_aditivos_kilo:,.2f}".replace(",", "."))
            c_env.metric("📦 Envasado / kg", f"$ {costo_envasado_kilo:,.2f}".replace(",", "."))
            c_cip.metric("🧼 CIP / kg", f"$ {costo_cip_kilo:,.2f}".replace(",", "."))

            st.markdown("---")
            periodo_pdf_str = f"{MESES_ES[filtro_mes_costo]} {filtro_anio_costo}"
            fecha_stock_txt = f"{fecha_maxima_stock.strftime('%d/%m/%Y')}" if pd.notna(fecha_maxima_stock) else "Sin recuento registrado"
            
            pdf_costos_bytes = generar_pdf_costos_mes(
                periodo_pdf_str, fecha_stock_txt, valor_total_fisico, tinas_mes, kilos_mes, 
                costo_insumos_total_mes, costo_por_kilo, costo_aditivos_kilo, costo_envasado_kilo, costo_cip_kilo, df_master_calc
            )
            st.download_button(
                "📥 Descargar Reporte Completo PDF",
                data=pdf_costos_bytes,
                file_name=f"Reporte_Mensual_{filtro_mes_costo}_{filtro_anio_costo}.pdf",
                mime="application/pdf",
                use_container_width=True
            )

            df_receta = df_master_calc[df_master_calc['Consumo por tina'] > 0].copy()
            df_receta['Costo en Tina ($)'] = df_receta['Consumo por tina'] * df_receta['Precio Unitario']
            
            df_receta_show = df_receta[['Insumo', 'Categoría', 'Consumo por tina', 'Unidad', 'Precio Unitario', 'Costo en Tina ($)', 'Costo Total Insumos Mes']].copy()
            df_receta_show['Precio Unitario'] = df_receta_show['Precio Unitario'].apply(lambda x: f"$ {x:,.2f}".replace(",", "."))
            df_receta_show['Costo en Tina ($)'] = df_receta_show['Costo en Tina ($)'].apply(lambda x: f"$ {x:,.2f}".replace(",", "."))
            df_receta_show['Costo Total Insumos Mes'] = df_receta_show['Costo Total Insumos Mes'].apply(lambda x: f"$ {x:,.2f}".replace(",", "."))
            
            st.dataframe(df_receta_show, use_container_width=True, hide_index=True)

    except Exception as e:
        st.error(f"Error procesando el módulo de insumos: {e}")
        st.code(traceback.format_exc())
