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
      # FIX: skiprows=5 asegura que la fila 6 sea el encabezado sin comerse el primer dato
      raw_prod = pd.read_excel(URL_PRODUCCION, skiprows=5)
      df_prod = pd.DataFrame()
      df_prod["Fecha"] = raw_prod.iloc[:, 0]
      df_prod["Lote"] = raw_prod.iloc[:, 1]
      df_prod["Litros Procesados"] = raw_prod.iloc[:, 3]
      df_prod["Producto Terminado"] = raw_prod.iloc[:, 5]
      df_prod["PNC"] = raw_prod.iloc[:, 6]
      df_prod = df_prod.dropna(subset=["Fecha"])
      
      df_prod["Fecha"] = pd.to_datetime(df_prod["Fecha"], dayfirst=True, errors="coerce")
      df_prod = df_prod.dropna(subset=["Fecha"])
      
      # FIX: Reemplazar comas por puntos en los kgs/litros para evitar NaN
      for col in ["Litros Procesados", "Producto Terminado", "PNC"]: 
          df_prod[col] = pd.to_numeric(df_prod[col].astype(str).str.replace(",", "."), errors="coerce").fillna(0)
      
      if len(df_prod) > 0: 
          df_prod["Producto"] = df_prod["Lote"].astype(str).apply(lambda x: "Muzzarella Export. Mastellone" if "840" in str(x) else "Otro")
          df_prod["Grupo"] = df_prod["Lote"].astype(str).apply(lambda x: "Mastellone" if "840" in str(x) else "Coopagro")
      else: 
          df_prod["Producto"], df_prod["Grupo"] = [], []
      
      df_prod["Año"] = df_prod["Fecha"].dt.year
      df_prod["Mes"] = df_prod["Fecha"].dt.month
      df_mastellone_prod = df_prod[df_prod["Grupo"] == "Mastellone"].copy()

      # 2. Cargar Recepción Diaria (Búsqueda inteligente y segura de la hoja MHSA)
      df_mhsa = pd.DataFrame()
      try:
          xls_remitos = pd.ExcelFile(URL_REMITOS)
          # FIX: Coincidencia exacta de la solapa MHSA
          sheet_mhsa = next((s for s in xls_remitos.sheet_names if s.strip().upper() == "MHSA"), None)
          
          if not sheet_mhsa:
              st.sidebar.warning("No se encontró la solapa exacta 'MHSA'.")
          else:
              df_mhsa_raw = pd.read_excel(URL_REMITOS, sheet_name=sheet_mhsa, dtype=str)
              
              df_mhsa = pd.DataFrame()
              
              cols_lower = [str(c).lower() for c in df_mhsa_raw.columns]
              idx_fecha = next((i for i, c in enumerate(cols_lower) if "fecha" in c), 0)
              idx_num_tambo = next((i for i, c in enumerate(cols_lower) if "tambo" in c and ("n" in c or "num" in c)), 2 if len(cols_lower) > 2 else 0)
              idx_tambo = next((i for i, c in enumerate(cols_lower) if "tambo" in c and "n" not in c and "num" not in c), 3 if len(cols_lower) > 3 else 0)
              idx_litros = next((i for i, c in enumerate(cols_lower) if "litro" in c), 4 if len(cols_lower) > 4 else 0)
              idx_temp = next((i for i, c in enumerate(cols_lower) if "temperatura" in c or "temp" in c), min(7, len(cols_lower)-1))

              df_mhsa["Fecha"] = df_mhsa_raw.iloc[:, idx_fecha]
              df_mhsa["Num_Tambo"] = df_mhsa_raw.iloc[:, idx_num_tambo]
              df_mhsa["Tambo"] = df_mhsa_raw.iloc[:, idx_tambo]
              df_mhsa["Litros_Ticket"] = df_mhsa_raw.iloc[:, idx_litros]
              df_mhsa["Temperatura"] = df_mhsa_raw.iloc[:, idx_temp]
              
              df_mhsa["Num_Tambo"] = df_mhsa["Num_Tambo"].apply(limpiar_tambo)
              
              # FIX: Eliminación de format="mixed"
              df_mhsa["Fecha"] = pd.to_datetime(df_mhsa["Fecha"], dayfirst=True, errors="coerce").dt.normalize()
              df_mhsa = df_mhsa.dropna(subset=["Fecha", "Num_Tambo"])
              
              # FIX: Reemplazar comas por puntos en MHSA
              df_mhsa["Litros_Ticket"] = pd.to_numeric(df_mhsa["Litros_Ticket"].astype(str).str.replace(",", "."), errors="coerce").fillna(0)
              df_mhsa["Temperatura"] = pd.to_numeric(df_mhsa["Temperatura"].astype(str).str.replace(",", "."), errors="coerce")
              
              df_mhsa["Año"] = df_mhsa["Fecha"].dt.year
              df_mhsa["Mes"] = df_mhsa["Fecha"].dt.month
              df_mhsa["orden_remito"] = df_mhsa.groupby(["Num_Tambo", "Fecha"]).cumcount() + 1
      except Exception as e:
          st.sidebar.warning(f"Aviso de carga MHSA: {e}")
          df_mhsa = pd.DataFrame()
      
      # 3. Cargar Laboratorio y cruzar con Recepción
      _, _, df_lab_raw, df_bac_raw = cargar_datos_coopagro(URL_REMITOS, URL_LAB, URL_BACSOMATIC)
      
      if not df_mhsa.empty:
          # --- Procesar MilkoScan ---
          if not df_lab_raw.empty:
              df_lab_m = df_lab_raw.copy()
              col_sample = df_lab_m.columns[0]
              
              df_lab_m["Num_Tambo"] = df_lab_m[col_sample].astype(str).str.split().str[0].apply(limpiar_tambo)
              df_lab_m["Fecha_Extraida"] = df_lab_m[col_sample].astype(str).str.split().str[-1].apply(extraer_fecha_texto)
              
              col_date = next((c for c in df_lab_m.columns if any(x in c.lower() for x in ["fecha", "date", "analyzed"])), None)
              if col_date:
                  df_lab_m["Fecha_Analisis"] = pd.to_datetime(df_lab_m[col_date], errors="coerce").dt.normalize()
                  df_lab_m["Fecha"] = df_lab_m["Fecha_Extraida"].combine_first(df_lab_m["Fecha_Analisis"])
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
              df_bac_m["Fecha_Extraida"] = df_bac_m[col_sample_bac].astype(str).str.split().str[-1].apply(extraer_fecha_texto)
              
              col_date_bac = next((c for c in df_bac_m.columns if any(x in c.lower() for x in ["fecha", "date", "analyzed"])), None)
              if col_date_bac:
                  df_bac_m["Fecha_Analisis"] = pd.to_datetime(df_bac_m[col_date_bac], errors="coerce").dt.normalize()
                  df_bac_m["Fecha"] = df_bac_m["Fecha_Extraida"].combine_first(df_bac_m["Fecha_Analisis"])
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
    # FILTROS POR DEFECTO (MES Y AÑO ACTUAL)
    # ==========================================
    st.sidebar.subheader("Filtros Mastellone")
    anios_mhsa = df_mhsa["Año"].dropna().unique().tolist() if not df_mhsa.empty else []
    anios_prod = df_mastellone_prod["Año"].dropna().unique().tolist() if not df_mastellone_prod.empty else []
    todos_anios = sorted(list(set(anios_mhsa + anios_prod)))
    
    opciones_anio = ["Todos"] + (todos_anios if todos_anios else [now.year])
    default_anio_idx = opciones_anio.index(now.year) if now.year in opciones_anio else 0
    filtro_anio = st.sidebar.selectbox("Año", opciones_anio, index=default_anio_idx, key="m_anio_mastellone")

    opciones_mes = ["Todos"] + list(range(1, 13))
    default_mes_idx = opciones_mes.index(now.month) if now.month in opciones_mes else 0
    filtro_mes = st.sidebar.selectbox("Mes", opciones_mes, index=default_mes_idx, key="m_mes_mastellone")

    # ==========================================
    # CÁLCULOS Y MÉTRICAS
    # ==========================================
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

    df_mhsa_filtrado = df_mhsa.copy()
    if not df_mhsa_filtrado.empty:
        if filtro_anio != "Todos": df_mhsa_filtrado = df_mhsa_filtrado[df_mhsa_filtrado["Año"] == filtro_anio]
        if filtro_mes != "Todos": df_mhsa_filtrado = df_mhsa_filtrado[df_mhsa_filtrado["Mes"] == filtro_mes]
        total_litros_ingresados = df_mhsa_filtrado["Litros_Ticket"].sum()
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
