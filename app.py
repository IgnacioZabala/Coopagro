import os
import re
import datetime
import traceback

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
FILE_ID_BACSOMATIC = "1SKBiDh4-EyELoYwlvqxB6QXErzYAdqPI"
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
  match = re.search(r'(\d+)\s*([\-\/\s])\s*(\d+)', s)
  if match:
      num1 = match.group(1)
      sep = match.group(2)
      num2 = match.group(3)
      if sep in ['-', '/'] or (len(num1) <= 2 and len(num2) == 1):
          return f"T{num1}-{num2}"
  if s and s[0].isdigit(): return f"T{s}"
  return s

def extraer_id_tambo(texto) -> str:
  if pd.isna(texto): return ""
  
  s_temp = str(texto).strip()
  if re.match(r'^\d{4}-\d{2}-\d{2}(?:\s\d{2}:\d{2}:\d{2})?$', s_temp):
      try:
          dt = pd.to_datetime(s_temp)
          s_temp = f"{dt.day}-{dt.month}"
      except:
          pass
          
  s = s_temp.upper()
  match = re.search(r'(\d+)\s*([\-\/\s])\s*(\d+)', s)
  if match:
      num1 = match.group(1)
      sep = match.group(2)
      num2 = match.group(3)
      if sep in ['-', '/'] or (len(num1) <= 2 and len(num2) == 1):
          return f"T{num1}-{num2}"
          
  match_single = re.match(r'^[^\d]*(\d+)', s)
  if match_single: return f"T{match_single.group(1)}"
  return s.strip()

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
# FUNCIONES DE PDF BLINDADAS
# =========================================================================
def limpiar_valor_str(val):
    if pd.isna(val): return "-"
    s = str(val).strip()
    if s.lower() == "nan" or s == "<na>" or s == "": return "-"
    return s

def parse_valor_celda(val):
  if pd.isna(val) or val == "-": return 0.0
  s = str(val).replace("%", "").strip()
  if s.lower() == "nan": return 0.0
  try:
    s = s.replace(".", "").replace(",", ".")
    return float(s)
  except ValueError:
    return 0.0

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

  for row in df_datos.itertuples(index=False):
    for i, (fn_mapeo, (_, col_w)) in enumerate(zip(filas_mapeo, headers_ajustados)):
      val_crudo = fn_mapeo(row)
      val_limpio = limpiar_valor_str(val_crudo)
      col_name = headers_ajustados[i][0]
      align = "L" if "Nombre" in col_name or "Producto" in col_name or "Insumo" in col_name or "Tambo" in col_name else "C"
      
      is_red = False
      if "UFC" in col_name or "SCC" in col_name:
          num_val = parse_valor_celda(val_limpio)
          if "UFC" in col_name and num_val > 200: is_red = True
          if "SCC" in col_name and num_val > 400: is_red = True
      
      if is_red:
          pdf.set_text_color(255, 0, 0)
          pdf.set_font("Arial", "B", 8)
      else:
          pdf.set_text_color(0, 0, 0)
          pdf.set_font("Arial", "", 8)

      pdf.cell(col_w, 7, val_limpio, 1, 1 if i == len(headers_ajustados) - 1 else 0, align)
  
  pdf.set_text_color(0, 0, 0)

  output = pdf.output(dest="S")
  if isinstance(output, bytearray): return bytes(output)
  elif isinstance(output, str): return output.encode('latin1')
  return output

def generar_pdf_bytes(df_productor, tambo_nombre, tambo_id, periodo_texto, args_visibles, es_mensual=False, es_interno=False):
  # Títulos condicionales
  if es_interno:
      titulo = "Resumen mensual de recolección (USO INTERNO)" if es_mensual else "Resumen semanal de recolección (USO INTERNO)"
  else:
      titulo = "Resumen mensual de recolección" if es_mensual else "Resumen semanal de recolección"
      
  subtitulo = f"Productor: {tambo_nombre} (Código #{tambo_id})"
  temp_prom = df_productor["Temperatura"].mean() if "Temperatura" in df_productor else float("nan")

  metricas = [f"Período: {periodo_texto}", f"Total Litros: {formato_miles(df_productor['Litros_Ticket'].sum())} L"]
  if args_visibles["temp"]: metricas.append(f"Temperatura Promedio: {formato_temp(temp_prom)}")

  partes_solidos = []
  if args_visibles["grasa"] and "Grasa" in df_productor and pd.notna(df_productor["Grasa"].mean()): partes_solidos.append(f"Grasa: {df_productor['Grasa'].mean():.2f}%".replace(".", ","))
  if args_visibles["prot"] and "Proteina" in df_productor and pd.notna(df_productor["Proteina"].mean()): partes_solidos.append(f"Proteína: {df_productor['Proteina'].mean():.2f}%".replace(".", ","))
  
  if es_interno:
      su_mean = (df_productor["Grasa"].mean() if "Grasa" in df_productor else 0) + (df_productor["Proteina"].mean() if "Proteina" in df_productor else 0)
      if su_mean > 0: partes_solidos.append(f"Sólidos Útiles: {su_mean:.2f}%".replace(".", ","))

  if args_visibles["crios"] and "Crioscopia" in df_productor and pd.notna(df_productor["Crioscopia"].mean()): partes_solidos.append(f"Crioscopia: {df_productor['Crioscopia'].mean():.3f}".replace(".", ","))
  if args_visibles["ufc"] and "UFC" in df_productor and pd.notna(df_productor["UFC"].mean()): partes_solidos.append(f"UFC <200: {formato_miles(df_productor['UFC'].mean())}")
  if args_visibles["scc"] and "SCC" in df_productor and pd.notna(df_productor["SCC"].mean()): partes_solidos.append(f"SCC <400: {formato_miles(df_productor['SCC'].mean())}")
  if partes_solidos: metricas.append(f"Promedios Lab -> {' | '.join(partes_solidos)}")

  # Ajustamos el ancho si es interno para que entren las columnas nuevas
  w_fecha, w_remito, w_litros = (18, 22, 18) if es_interno else (26, 34, 30)
  
  headers = [("Fecha", w_fecha), ("N° Remito", w_remito), ("Litros", w_litros)]
  mapeo = [
      lambda r: getattr(r, "Fecha").strftime("%d/%m/%Y") if pd.notna(getattr(r, "Fecha")) else "",
      lambda r: str(getattr(r, "N_Remito")) if pd.notna(getattr(r, "N_Remito")) else "-",
      lambda r: formato_miles(getattr(r, "Litros_Ticket")) if pd.notna(getattr(r, "Litros_Ticket")) else "0"
  ]
  
  if args_visibles["temp"]: 
      headers.append(("Temp", 12 if es_interno else 18))
      mapeo.append(lambda r: formato_temp(getattr(r, "Temperatura", pd.NA)))
      
  if args_visibles["grasa"]: 
      headers.append(("Grasa", 14 if es_interno else 20))
      mapeo.append(lambda r: f"{getattr(r, 'Grasa'):.2f}%".replace(".", ",") if pd.notna(getattr(r, 'Grasa', pd.NA)) else "-")
      
  if args_visibles["prot"]: 
      headers.append(("Prot", 14 if es_interno else 20))
      mapeo.append(lambda r: f"{getattr(r, 'Proteina'):.2f}%".replace(".", ",") if pd.notna(getattr(r, 'Proteina', pd.NA)) else "-")

  if es_interno:
      headers.append(("% SU", 15))
      mapeo.append(lambda r: f"{(getattr(r, 'Grasa', 0) or 0) + (getattr(r, 'Proteina', 0) or 0):.2f}%".replace(".", ",") if pd.notna(getattr(r, 'Grasa', pd.NA)) and pd.notna(getattr(r, 'Proteina', pd.NA)) else "-")
      
      headers.append(("Kg SU", 18))
      def calc_kg(r):
          g = getattr(r, 'Grasa', 0) or 0
          p = getattr(r, 'Proteina', 0) or 0
          l = getattr(r, 'Litros_Ticket', 0) or 0
          d = getattr(r, 'Densidad', 1.030)
          if pd.isna(d): d = 1.030
          su = g + p
          if su > 0: return f"{l * d * (su/100):,.1f}".replace(",", ".")
          return "-"
      mapeo.append(calc_kg)

  if args_visibles["crios"]: 
      headers.append(("Crios", 16 if es_interno else 22))
      mapeo.append(lambda r: f"{getattr(r, 'Crioscopia'):.3f}".replace(".", ",") if pd.notna(getattr(r, 'Crioscopia', pd.NA)) else "-")
      
  if args_visibles["ufc"]: 
      headers.append(("UFC", 20 if es_interno else 24))
      mapeo.append(lambda r: formato_miles(getattr(r, 'UFC', pd.NA)) if pd.notna(getattr(r, 'UFC', pd.NA)) else "-")
      
  if args_visibles["scc"]: 
      headers.append(("SCC", 20 if es_interno else 25))
      mapeo.append(lambda r: formato_miles(getattr(r, 'SCC', pd.NA)) if pd.notna(getattr(r, 'SCC', pd.NA)) else "-")

  return generar_pdf_base(titulo, subtitulo, metricas, headers, df_productor, mapeo)

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
    for col in ["Grasa", "Proteina", "Crioscopia", "Densidad", "UFC", "SCC"]:
      if col not in df.columns: df[col] = pd.NA

    if not df_contactos.empty and "Tambo_Maestro" in df_contactos.columns:
        mapeo_nombres = dict(zip(df_contactos["Num_Tambo"], df_contactos["Tambo_Maestro"]))
        df["Tambo"] = df["Num_Tambo"].map(mapeo_nombres).combine_first(df["Tambo"])
        
    df["Tambo"] = df["Tambo"].replace("#REF!", "Desconocido")

    df = df.sort_values(by=["Num_Tambo", "Fecha", "N_Remito"])
    df["lab_index"] = df.groupby(["Num_Tambo", "Fecha"]).cumcount()

    if not df_lab_raw.empty:
      df_lab = df_lab_raw.copy()
      col_sample = next((c for c in df_lab.columns if any(x in c.lower() for x in ["sample", "number", "tambo", "muestra"])), df_lab.columns[0])
      df_lab["Num_Tambo"] = df_lab[col_sample].astype(str).apply(extraer_id_tambo)
      
      df_lab["Fecha_Extraida"] = pd.to_datetime(df_lab[col_sample].astype(str).apply(extraer_fecha_texto), errors="coerce")
      col_date = next((c for c in df_lab.columns if any(x in c.lower() for x in ["fecha", "date", "time", "analyzed"])), None)
      df_lab["Fecha"] = df_lab["Fecha_Extraida"].fillna(pd.to_datetime(df_lab[col_date], errors="coerce").dt.normalize() if col_date else pd.NaT)
      df_lab = df_lab.dropna(subset=["Fecha", "Num_Tambo"])
      
      map_cols = {}
      col_fat = next((c for c in df_lab.columns if "fat" in c.lower() or "grasa" in c.lower()), None)
      col_prot = next((c for c in df_lab.columns if "protein" in c.lower() or "proteina" in c.lower()), None)
      col_fp = next((c for c in df_lab.columns if c.lower() == "fp" or "crios" in c.lower()), None)
      col_dens = next((c for c in df_lab.columns if "dens" in c.lower()), None) # NUEVO: Densidad

      if col_fat: map_cols[col_fat] = "Grasa_Lab"
      if col_prot: map_cols[col_prot] = "Proteina_Lab"
      if col_fp: map_cols[col_fp] = "Crioscopia_Lab"
      if col_dens: map_cols[col_dens] = "Densidad_Lab" # NUEVO: Densidad

      for c_orig in map_cols.keys():
          df_lab[c_orig] = pd.to_numeric(df_lab[c_orig].astype(str).str.replace(",", "."), errors="coerce")

      # Filtro de seguridad: descarta lecturas erróneas con grasa menor a 1%
      if col_fat:
          df_lab = df_lab[df_lab[col_fat] > 1.0]

      if col_date:
          df_lab["_has_time"] = df_lab[col_sample].astype(str).str.contains(r"\d{2}:\d{2}", regex=True)
          df_lab["_sort_time"] = pd.to_datetime(df_lab[col_date], errors="coerce")
          df_lab = df_lab.sort_values(by=["Num_Tambo", "Fecha", "_has_time", "_sort_time"], ascending=[True, True, False, False])
      else:
          df_lab["_sample_str"] = df_lab[col_sample].astype(str)
          df_lab = df_lab.sort_values(by=["Num_Tambo", "Fecha", "_sample_str"], ascending=[True, True, False])
      
      df_lab["lab_index"] = df_lab.groupby(["Num_Tambo", "Fecha"]).cumcount()
      
      if map_cols:
        df_milko_clean = df_lab[["Num_Tambo", "Fecha", "lab_index"] + list(map_cols.keys())].rename(columns=map_cols)
        for c in map_cols.values(): df_milko_clean[c] = pd.to_numeric(df_milko_clean[c].astype(str).str.replace(",", "."), errors="coerce")
        
        df = pd.merge(df, df_milko_clean, on=["Num_Tambo", "Fecha", "lab_index"], how="left")
                    
        if "Grasa_Lab" in df: df["Grasa"] = df["Grasa_Lab"].combine_first(df["Grasa"])
        if "Proteina_Lab" in df: df["Proteina"] = df["Proteina_Lab"].combine_first(df["Proteina"])
        if "Crioscopia_Lab" in df: df["Crioscopia"] = df["Crioscopia_Lab"].combine_first(df["Crioscopia"])
        if "Densidad_Lab" in df: 
            if "Densidad" in df: df["Densidad"] = df["Densidad_Lab"].combine_first(df["Densidad"])
            else: df["Densidad"] = df["Densidad_Lab"]

    if not df_bac_raw.empty:
      df_bac = df_bac_raw.copy()
      col_id = next((c for c in df_bac.columns if any(x in c.lower() for x in ["id usuario", "sample", "tambo"])), df_bac.columns[0])
      df_bac["Num_Tambo"] = df_bac[col_id].astype(str).apply(extraer_id_tambo)
      
      df_bac["Fecha_Extraida"] = pd.to_datetime(df_bac[col_id].astype(str).apply(extraer_fecha_texto), errors="coerce")
      col_date_bac = next((c for c in df_bac.columns if any(x in c.lower() for x in ["fecha", "date", "analyzed"])), None)
      df_bac["Fecha"] = df_bac["Fecha_Extraida"].fillna(pd.to_datetime(df_bac[col_date_bac], errors="coerce").dt.normalize() if col_date_bac else pd.NaT)
      df_bac = df_bac.dropna(subset=["Fecha", "Num_Tambo"])
      
      if col_date_bac:
          df_bac["_sort_time"] = pd.to_datetime(df_bac[col_date_bac], errors="coerce")
          df_bac = df_bac.sort_values(by=["Num_Tambo", "Fecha", "_sort_time"], ascending=[True, True, False])
      else:
          df_bac["_sample_str"] = df_bac[col_id].astype(str)
          df_bac = df_bac.sort_values(by=["Num_Tambo", "Fecha", "_sample_str"], ascending=[True, True, False])
          
      df_bac["lab_index"] = df_bac.groupby(["Num_Tambo", "Fecha"]).cumcount()
      
      map_cols_bac = {}
      col_ufc = next((c for c in df_bac.columns if "ufc" in c.lower()), None)
      col_scc = next((c for c in df_bac.columns if any(x in c.lower() for x in ["scc", "celulas", "somáticas"])), None)
      if col_ufc: map_cols_bac[col_ufc] = "UFC_Val"
      if col_scc: map_cols_bac[col_scc] = "SCC_Val"
      
      if map_cols_bac:
        df_bac_clean = df_bac[["Num_Tambo", "Fecha", "lab_index"] + list(map_cols_bac.keys())].rename(columns=map_cols_bac)
        for c in map_cols_bac.values(): df_bac_clean[c] = pd.to_numeric(df_bac_clean[c].astype(str).str.replace(",", "."), errors="coerce")
        
        df = pd.merge(df, df_bac_clean, on=["Num_Tambo", "Fecha", "lab_index"], how="left")
                    
        if "UFC_Val" in df: df["UFC"] = df["UFC_Val"].combine_first(df["UFC"])
        if "SCC_Val" in df: df["SCC"] = df["SCC_Val"].combine_first(df["SCC"])

    df = df.drop(columns=["lab_index"], errors="ignore")
    # Llenamos densidades faltantes con un estándar para no fallar el cálculo de Kg (1.030 kg/L)
    df["Densidad"] = df.get("Densidad", pd.Series(dtype=float)).fillna(1.030) 

    df["Fecha_Cierre_Viernes"] = df["Fecha"] + pd.to_timedelta((4 - df["Fecha"].dt.weekday) % 7, unit="D")
    df["Fecha_Inicio_Sabado"] = df["Fecha_Cierre_Viernes"] - pd.Timedelta(days=6)
    df["Ciclo_Semana"] = "Viernes " + df["Fecha_Cierre_Viernes"].dt.strftime("%d/%m/%Y") + " (Sáb " + df["Fecha_Inicio_Sabado"].dt.strftime("%d/%m/%Y") + " al Vie " + df["Fecha_Cierre_Viernes"].dt.strftime("%d/%m/%Y") + ")"
    df["AnioMes"] = df["Fecha"].dt.to_period("M")
    df = df.sort_values(by=["Num_Tambo", "Fecha", "N_Remito"])

    st.sidebar.markdown("---")
    # SE ELIMINÓ LA VISTA "Envío Masivo Semanal"
    vista_coop = st.sidebar.radio("Sección Coopagro:", ["Panel de Control General", "Reporte Diario de Recibos y Laboratorio", "Gestión y Reportes por Tambo"])

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
            
            # CÁLCULO DE SÓLIDOS
            g_val = g_pond if pd.notna(g_pond) else 0
            p_val = p_pond if pd.notna(p_pond) else 0
            porcentaje_su = g_val + p_val
            
            d_pond = calcular_promedio_ponderado(group, "Densidad") if "Densidad" in group.columns else float("nan")
            densidad_calc = d_pond if pd.notna(d_pond) and d_pond > 0 else 1.030
            
            kg_solidos_utiles = tot_l * densidad_calc * (porcentaje_su / 100) if porcentaje_su > 0 else 0
            
            ranking_data.append({
                "Num_Tambo": str(num_t),
                "Tambo": str(tambo_n),
                "Litros_Ticket": tot_l,
                "Grasa_Ponderada": g_pond,
                "Proteina_Ponderada": p_pond,
                "Porcentaje_SU": porcentaje_su,
                "Kg_Solidos_Utiles": kg_solidos_utiles
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
              "Ranking y Calidad de Tambos por Volumen (Interno)",
          ]
          headers = [("Código", 15), ("Nombre del Tambo", 55), ("Litros", 22), ("Grasa Pond.", 16), ("Prot. Pond.", 16), ("% SU", 15), ("Kg SU", 20)]
          mapeo = [
              lambda r: str(r.Num_Tambo),
              lambda r: str(r.Tambo)[:28],
              lambda r: formato_miles(r.Litros_Ticket),
              lambda r: f"{r.Grasa_Ponderada:.2f}%".replace(".", ",") if pd.notna(r.Grasa_Ponderada) else "S/D",
              lambda r: f"{r.Proteina_Ponderada:.2f}%".replace(".", ",") if pd.notna(r.Proteina_Ponderada) else "S/D",
              lambda r: f"{r.Porcentaje_SU:.2f}%".replace(".", ",") if r.Porcentaje_SU > 0 else "S/D",
              lambda r: f"{r.Kg_Solidos_Utiles:,.1f}".replace(",", ".") if r.Kg_Solidos_Utiles > 0 else "-"
          ]
          return generar_pdf_base("Informe de Recolección y Calidad - Cooperativa", f"Período Evaluado: {periodo_titulo}", metricas, headers, df_ranking, mapeo)

        pdf_bytes = generar_pdf_panel_general_con_calidad(df_macro, periodo_texto, tot_litros, df_macro["Temperatura"].mean(), grasa_p, prot_p, ratio_gp, df_macro["Num_Tambo"].nunique(), df_ranking)
        st.download_button("🔒 Descargar Informe PDF Interno (Con Sólidos)", data=pdf_bytes, file_name=f"Informe_Calidad_Interno_{periodo_texto.replace(' ', '_')}.pdf", mime="application/pdf")
        
        df_ranking_show = df_ranking.copy()
        df_ranking_show["Litros_Ticket"] = df_ranking_show["Litros_Ticket"].apply(formato_miles)
        df_ranking_show["Grasa_Ponderada"] = df_ranking_show["Grasa_Ponderada"].apply(lambda x: f"{x:.2f}%".replace(".", ",") if pd.notna(x) else "S/D")
        df_ranking_show["Proteina_Ponderada"] = df_ranking_show["Proteina_Ponderada"].apply(lambda x: f"{x:.2f}%".replace(".", ",") if pd.notna(x) else "S/D")
        df_ranking_show["Porcentaje_SU"] = df_ranking_show["Porcentaje_SU"].apply(lambda x: f"{x:.2f}%".replace(".", ",") if x > 0 else "S/D")
        df_ranking_show["Kg_Solidos_Utiles"] = df_ranking_show["Kg_Solidos_Utiles"].apply(lambda x: f"{x:,.1f} kg".replace(",", ".") if x > 0 else "-")
        
        st.dataframe(df_ranking_show.rename(columns={
            "Tambo": "Nombre del Tambo", 
            "Num_Tambo": "Código", 
            "Litros_Ticket": "Litros Totales",
            "Grasa_Ponderada": "Grasa Pond.",
            "Proteina_Ponderada": "Proteína Pond.",
            "Porcentaje_SU": "% Sólidos Útiles",
            "Kg_Solidos_Utiles": "Total Sólidos (Kg)"
        }), hide_index=True, use_container_width=True)

    elif vista_coop == "Reporte Diario de Recibos y Laboratorio":
      st.header("📅 Reporte Diario de Recibos y Laboratorio")
      
      fechas_disponibles = sorted(df["Fecha"].dropna().unique(), reverse=True)
      if len(fechas_disponibles) > 0:
          fecha_sel = st.selectbox("Seleccione la Fecha de Recolección:", fechas_disponibles, format_func=lambda x: pd.to_datetime(x).strftime("%d/%m/%Y"))
          
          df_dia = df[df["Fecha"] == fecha_sel].sort_values(by=["Num_Tambo", "N_Remito"])
          
          if not df_dia.empty:
              tot_l_dia = df_dia["Litros_Ticket"].sum()
              temp_p_dia = df_dia["Temperatura"].mean()
              grasa_p_dia = calcular_promedio_ponderado(df_dia, "Grasa")
              prot_p_dia = calcular_promedio_ponderado(df_dia, "Proteina")
              
              c1, c2, c3, c4 = st.columns(4)
              c1.metric("Litros del Día", formato_miles(tot_l_dia))
              c2.metric("Remitos / Tambos", f"{len(df_dia)} ({df_dia['Num_Tambo'].nunique()})")
              c3.metric("Temp. Promedio", formato_temp(temp_p_dia))
              c4.metric("Grasa / Prot. Pond.", f"{grasa_p_dia:.2f}% / {prot_p_dia:.2f}%" if pd.notna(grasa_p_dia) and pd.notna(prot_p_dia) else "S/D")
              
              st.subheader(f"Detalle de Recepción y Calidad — {pd.to_datetime(fecha_sel).strftime('%d/%m/%Y')}")
              
              df_dia_show = pd.DataFrame()
              df_dia_show["Fecha"] = df_dia["Fecha"].dt.strftime("%d/%m/%Y")
              df_dia_show["N° Remito"] = df_dia["N_Remito"]
              df_dia_show["Código"] = df_dia["Num_Tambo"]
              df_dia_show["Tambo"] = df_dia["Tambo"]
              df_dia_show["Litros"] = df_dia["Litros_Ticket"].apply(formato_miles)
              df_dia_show["Temp"] = df_dia["Temperatura"].apply(formato_temp)
              df_dia_show["Grasa"] = df_dia["Grasa"].apply(lambda x: f"{x:.2f}%" if pd.notna(x) else "-")
              df_dia_show["Proteína"] = df_dia["Proteina"].apply(lambda x: f"{x:.2f}%" if pd.notna(x) else "-")
              df_dia_show["Crioscopía"] = df_dia["Crioscopia"].apply(lambda x: f"{x:.3f}" if pd.notna(x) else "-")
              df_dia_show["UFC <200"] = df_dia["UFC"].apply(lambda x: formato_miles(x) if pd.notna(x) else "-")
              df_dia_show["SCC <400"] = df_dia["SCC"].apply(lambda x: formato_miles(x) if pd.notna(x) else "-")
              
              def highlight_bacsomatic(val, threshold):
                  try:
                      if pd.notna(val) and str(val) != "-":
                          num_val = float(str(val).replace(".", "").replace(",", "."))
                          if num_val > threshold: return 'color: red; font-weight: bold'
                  except: pass
                  return ''

              df_style_dia = df_dia_show.style
              if "UFC <200" in df_dia_show.columns:
                  df_style_dia = df_style_dia.map(lambda x: highlight_bacsomatic(x, 200), subset=["UFC <200"])
              if "SCC <400" in df_dia_show.columns:
                  df_style_dia = df_style_dia.map(lambda x: highlight_bacsomatic(x, 400), subset=["SCC <400"])

              st.dataframe(df_style_dia, use_container_width=True, hide_index=True)
              
              headers_pdf_dia = [("Remito", 22), ("Código", 18), ("Tambo", 50), ("Litros", 22), ("Temp", 15), ("Grasa", 16), ("Prot", 16), ("UFC", 20), ("SCC", 20)]
              mapeo_pdf_dia = [
                  lambda r: str(getattr(r, "N_Remito", "-")),
                  lambda r: str(getattr(r, "Num_Tambo", "-")),
                  lambda r: str(getattr(r, "Tambo", "-"))[:24],
                  lambda r: formato_miles(getattr(r, "Litros_Ticket", 0)),
                  lambda r: formato_temp(getattr(r, "Temperatura", pd.NaT)),
                  lambda r: f"{getattr(r, 'Grasa'):.2f}%".replace(".", ",") if pd.notna(getattr(r, 'Grasa', pd.NaT)) else "-",
                  lambda r: f"{getattr(r, 'Proteina'):.2f}%".replace(".", ",") if pd.notna(getattr(r, 'Proteina', pd.NaT)) else "-",
                  lambda r: formato_miles(getattr(r, 'UFC', pd.NaT)) if pd.notna(getattr(r, 'UFC', pd.NaT)) else "-",
                  lambda r: formato_miles(getattr(r, 'SCC', pd.NaT)) if pd.notna(getattr(r, 'SCC', pd.NaT)) else "-"
              ]
              
              metricas_dia = [
                  f"Fecha del Reporte: {pd.to_datetime(fecha_sel).strftime('%d/%m/%Y')} | Total Litros: {formato_miles(tot_l_dia)} L",
                  f"Tambos Recolectados: {df_dia['Num_Tambo'].nunique()} | Temperatura Promedio: {formato_temp(temp_p_dia)}"
              ]
              
              pdf_dia_bytes = generar_pdf_base("Reporte Diario de Recepción y Calidad", "Cooperativa Agropecuaria (Coopagro)", metricas_dia, headers_pdf_dia, df_dia, mapeo_pdf_dia)
              st.download_button("📥 Descargar Reporte Diario PDF", data=pdf_dia_bytes, file_name=f"Reporte_Diario_{pd.to_datetime(fecha_sel).strftime('%Y%m%d')}.pdf", mime="application/pdf", use_container_width=True)
          else:
              st.info("No hay registros para la fecha seleccionada.")
      else:
          st.info("No hay fechas disponibles en los remitos.")

    elif vista_coop == "Gestión y Reportes por Tambo":
      st.header("📄 Reportes por Tambo")
      tipo_reporte_opcion = st.sidebar.radio("Período de Reporte:", ["Semanal", "Mensual"])
      
      if tipo_reporte_opcion == "Semanal":
        ciclos = df[["Fecha_Cierre_Viernes", "Ciclo_Semana"]].drop_duplicates().sort_values("Fecha_Cierre_Viernes", ascending=False)["Ciclo_Semana"].tolist()
        periodo_sel = st.sidebar.selectbox("1. Seleccione Semana:", ciclos) if ciclos else ""
        df_periodo = df[df["Ciclo_Semana"] == periodo_sel].sort_values(["Fecha", "N_Remito"]) if ciclos else pd.DataFrame()
        es_mensual = False
      else:
        meses = sorted(df["AnioMes"].unique(), reverse=True)
        periodo_sel = st.sidebar.selectbox("1. Seleccione Mes:", meses, format_func=lambda p: f"{MESES_ES.get(p.month)} {p.year}") if meses else None
        df_periodo = df[df["AnioMes"] == periodo_sel].sort_values(["Fecha", "N_Remito"]) if meses else pd.DataFrame()
        es_mensual = True
      
      st.sidebar.subheader("⚙️ Elementos del Reporte")
      v_temp = st.sidebar.checkbox("Temperatura", True)
      v_grasa = st.sidebar.checkbox("Grasa", True)
      v_prot = st.sidebar.checkbox("Proteína", True)
      v_crios = st.sidebar.checkbox("Crioscopia", True)
      v_ufc = st.sidebar.checkbox("UFC <200", True)
      v_scc = st.sidebar.checkbox("SCC <400", True)
      args_vis = {"temp": v_temp, "grasa": v_grasa, "prot": v_prot, "crios": v_crios, "ufc": v_ufc, "scc": v_scc}

      if not df_periodo.empty:
        st.subheader("📦 Descarga Masiva")
        st.info(f"Se encontraron movimientos para **{df_periodo['Num_Tambo'].nunique()} tambos** en el período seleccionado.")
        
        tipo_zip = st.radio("Formato del reporte masivo:", ["Para Productores (Sin sólidos)", "Uso Interno (Con sólidos)"])
        es_interno_zip = (tipo_zip == "Uso Interno (Con sólidos)")
        
        if st.button("Generar ZIP con Todos los Reportes"):
          with st.spinner("Generando PDFs y comprimiendo en ZIP..."):
            import io
            import zipfile
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
              for t_id in df_periodo["Num_Tambo"].unique():
                df_t_loop = df_periodo[df_periodo["Num_Tambo"] == t_id].copy()
                if df_t_loop.empty: continue
                t_nom_loop = df_t_loop["Tambo"].iloc[0]
                
                if es_mensual:
                  periodo_str_pdf = f"{MESES_ES.get(periodo_sel.month)} {periodo_sel.year}"
                  tipo_str = "INTERNO" if es_interno_zip else "mensual"
                  nom_arch = f"Reporte {tipo_str} {t_nom_loop} - {periodo_str_pdf}.pdf"
                else:
                  f_ini = df_t_loop['Fecha_Inicio_Sabado'].iloc[0]
                  f_fin = df_t_loop['Fecha_Cierre_Viernes'].iloc[0]
                  periodo_str_pdf = f"{f_ini:%d/%m/%Y} al {f_fin:%d/%m/%Y}"
                  tipo_str = "INTERNO" if es_interno_zip else "semanal"
                  nom_arch = f"Reporte {tipo_str} {t_nom_loop} - Semana {f_ini:%d-%m} al {f_fin:%d-%m}.pdf" 
                  
                pdf_bytes_loop = generar_pdf_bytes(df_t_loop, t_nom_loop, t_id, periodo_str_pdf, args_vis, es_mensual, es_interno=es_interno_zip)
                zip_file.writestr(nom_arch, pdf_bytes_loop)
                
            st.download_button(
                label="📥 Descargar Archivo ZIP",
                data=zip_buffer.getvalue(),
                file_name=f"Reportes_{'Internos' if es_interno_zip else 'Productores'}_{'Mensuales' if es_mensual else 'Semanales'}.zip",
                mime="application/zip",
                type="primary"
            )
            
        st.markdown("---")
        
        st.subheader("📄 Vista y Descarga Individual")
        mapeo_tambos_activos = df_periodo[["Tambo", "Num_Tambo"]].drop_duplicates().sort_values("Tambo")
        t_nombre = st.sidebar.selectbox("2. Seleccione Tambo (Vista Individual):", mapeo_tambos_activos["Tambo"].tolist())
        t_id = mapeo_tambos_activos.loc[mapeo_tambos_activos["Tambo"] == t_nombre, "Num_Tambo"].values[0]
        
        df_per = df_periodo[df_periodo["Num_Tambo"] == str(t_id)].copy()

        if not df_per.empty:
          # CÁLCULO DE SÓLIDOS PARA LA UI
          df_per["Porcentaje_SU"] = df_per["Grasa"].fillna(0) + df_per["Proteina"].fillna(0)
          df_per["Kg_SU"] = df_per["Litros_Ticket"] * df_per["Densidad"] * (df_per["Porcentaje_SU"] / 100)
          
          if es_mensual:
            periodo_pdf = f"{MESES_ES.get(periodo_sel.month)} {periodo_sel.year}"
            nom_arch_prod = f"Reporte mensual {t_nombre} - {periodo_pdf}.pdf"
            nom_arch_int = f"Reporte INTERNO {t_nombre} - {periodo_pdf}.pdf"
          else:
            f_ini = df_per['Fecha_Inicio_Sabado'].iloc[0]
            f_fin = df_per['Fecha_Cierre_Viernes'].iloc[0]
            periodo_pdf = f"{f_ini:%d/%m/%Y} al {f_fin:%d/%m/%Y}"
            nom_arch_prod = f"Reporte semanal {t_nombre} - Semana {f_ini:%d-%m} al {f_fin:%d-%m}.pdf"
            nom_arch_int = f"Reporte INTERNO {t_nombre} - Semana {f_ini:%d-%m} al {f_fin:%d-%m}.pdf"

          st.markdown(f"**Resumen {'Mensual' if es_mensual else 'Semanal'} - {t_nombre} (#{t_id})**")
          
          pdf_b_productor = generar_pdf_bytes(df_per, t_nombre, t_id, periodo_pdf, args_vis, es_mensual, es_interno=False)
          pdf_b_interno = generar_pdf_bytes(df_per, t_nombre, t_id, periodo_pdf, args_vis, es_mensual, es_interno=True)

          b1, b2 = st.columns(2)
          b1.download_button("📥 Descargar PDF (Productor)", data=pdf_b_productor, file_name=nom_arch_prod, mime="application/pdf", use_container_width=True)
          b2.download_button("🔒 Descargar PDF Interno (Con Sólidos)", data=pdf_b_interno, file_name=nom_arch_int, mime="application/pdf", use_container_width=True)

          df_tabla_visual = pd.DataFrame()
          df_tabla_visual["Fecha"] = df_per["Fecha"].dt.strftime("%d/%m/%Y")
          df_tabla_visual["N° Remito"] = df_per["N_Remito"]
          df_tabla_visual["Litros"] = df_per["Litros_Ticket"].apply(formato_miles)

          if v_temp and "Temperatura" in df_per: df_tabla_visual["Temperatura"] = df_per["Temperatura"].apply(lambda x: f"{x:.1f}°" if pd.notna(x) else "-")
          if v_grasa and "Grasa" in df_per: df_tabla_visual["Grasa"] = df_per["Grasa"].apply(lambda x: f"{x:.2f}%" if pd.notna(x) else "-")
          if v_prot and "Proteina" in df_per: df_tabla_visual["Proteína"] = df_per["Proteina"].apply(lambda x: f"{x:.2f}%" if pd.notna(x) else "-")
          
          # Agregamos Sólidos a la tabla UI
          df_tabla_visual["% Sólidos Útiles"] = df_per["Porcentaje_SU"].apply(lambda x: f"{x:.2f}%" if pd.notna(x) and x > 0 else "-")
          df_tabla_visual["Kg Sólidos Útiles"] = df_per["Kg_SU"].apply(lambda x: f"{x:,.1f}".replace(",", ".") if pd.notna(x) and x > 0 else "-")
          
          if v_crios and "Crioscopia" in df_per: df_tabla_visual["Crioscopia"] = df_per["Crioscopia"].apply(lambda x: f"{x:.3f}" if pd.notna(x) else "-")
          if v_ufc and "UFC" in df_per: df_tabla_visual["UFC <200"] = df_per["UFC"].apply(lambda x: formato_miles(x) if pd.notna(x) else "-")
          if v_scc and "SCC" in df_per: df_tabla_visual["SCC <400"] = df_per["SCC"].apply(lambda x: formato_miles(x) if pd.notna(x) else "-")

          def highlight_bacsomatic(val, threshold):
              try:
                  if pd.notna(val) and str(val) != "-":
                      num_val = float(str(val).replace(".", "").replace(",", "."))
                      if num_val > threshold: return 'color: red; font-weight: bold'
              except: pass
              return ''

          df_style_coop = df_tabla_visual.style
          if "UFC <200" in df_tabla_visual.columns:
              df_style_coop = df_style_coop.map(lambda x: highlight_bacsomatic(x, 200), subset=["UFC <200"])
          if "SCC <400" in df_tabla_visual.columns:
              df_style_coop = df_style_coop.map(lambda x: highlight_bacsomatic(x, 400), subset=["SCC <400"])

          st.dataframe(df_style_coop, use_container_width=True, hide_index=True)

  except Exception as e:
    st.error("Error en el Módulo Coopagro:")
    st.code(traceback.format_exc())

