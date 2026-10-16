import io
import numpy as np
import pandas as pd
import streamlit as st


def render():
  st.subheader("Reporte: Cuentas por Cobrar Pendientes (CxC)")

  uploaded_files = st.file_uploader(
      "Sube los archivos necesarios (Anterior/Ctas por cobrar, Pagos Odoo,"
      " Facturas Odoo, Muestras, y Archivos 12121/12122/12211/12212):",
      type=["xlsx"],
      accept_multiple_files=True,
      key="uploader_cxc",
  )

  if st.button("Generar Reporte CxC", type="primary"):
    if not uploaded_files:
      st.error("Debes subir los archivos requeridos.")
      return

    file_dict = {f.name: f for f in uploaded_files}

    with st.spinner("Procesando consolidado CxC..."):
      ant_file = next(
          (
              f
              for n, f in file_dict.items()
              if "anterior" in n.lower() or "ctas por cobrar" in n.lower()
          ),
          None,
      )
      ant_by_comp = {}
      ant_by_nrocomp = {}

      if ant_file:
        xl_ant = pd.ExcelFile(ant_file)
        sheet_ant = (
            "Documentos Pendientes"
            if "Documentos Pendientes" in xl_ant.sheet_names
            else (
                "Anterior"
                if "Anterior" in xl_ant.sheet_names
                else xl_ant.sheet_names[0]
            )
        )
        ant_file.seek(0)
        df_ant = pd.read_excel(ant_file, sheet_name=sheet_ant, skiprows=6)

        for _, row in df_ant.iterrows():
          comp = str(row.iloc[6]).strip() if pd.notna(row.iloc[6]) else ""
          nro_asiento = (
              str(row.iloc[10]).strip()
              if len(row) > 10 and pd.notna(row.iloc[10])
              else ""
          )
          datos_gestion = {
              "Comprobante": comp,
              "Status": row.iloc[18] if len(row) > 18 else np.nan,
              "Vendedor": row.iloc[19] if len(row) > 19 else np.nan,
              "Cotización": row.iloc[20] if len(row) > 20 else np.nan,
              "Observación": row.iloc[21] if len(row) > 21 else np.nan,
              "Revisión": row.iloc[22] if len(row) > 22 else np.nan,
              "Comentarios de Ventas": row.iloc[23] if len(row) > 23 else np.nan,
          }
          if comp and comp not in ("nan", "", "None") and comp not in ant_by_comp:
            ant_by_comp[comp] = datos_gestion
          if (
              nro_asiento
              and nro_asiento not in ("nan", "", "None")
              and nro_asiento not in ant_by_nrocomp
          ):
            ant_by_nrocomp[nro_asiento] = datos_gestion

      pagos_map = {}
      pagos_file = next(
          (f for n, f in file_dict.items() if "pagos" in n.lower()), None
      )
      if pagos_file:
        pagos_file.seek(0)
        df_pagos = pd.read_excel(pagos_file)
        for _, row in df_pagos.iterrows():
          nombre = str(row.iloc[1]).strip()
          fact_conc = str(row.iloc[9]).strip() if pd.notna(row.iloc[9]) else ""
          if nombre and nombre not in ("nan", "", "None"):
            pagos_map[nombre] = fact_conc

      fact_map = {}
      fact_file = next(
          (f for n, f in file_dict.items() if "facturas" in n.lower()), None
      )
      if fact_file:
        fact_file.seek(0)
        df_fact = pd.read_excel(fact_file)
        for _, row in df_fact.iterrows():
          comp = str(row.iloc[1]).strip()
          if comp and comp not in ("nan", "", "None"):
            fact_map[comp] = {
                "Vendedor": row.iloc[6],
                "Cotización": row.iloc[7],
            }

      muestras_set = set()
      muestras_file = next(
          (f for n, f in file_dict.items() if "muestras" in n.lower()), None
      )
      if muestras_file:
        muestras_file.seek(0)
        df_muestras = pd.read_excel(muestras_file)
        col_orden = [c for c in df_muestras.columns if "Orden" in str(c)]
        if col_orden:
          muestras_set = set(
              df_muestras[col_orden[0]]
              .dropna()
              .astype(str)
              .str.strip()
              .unique()
          )
        else:
          muestras_set = set(
              df_muestras.iloc[:, 7].dropna().astype(str).str.strip().unique()
          )

      cuentas_config = [
          ("12121", "12121 - FACTURAS POR COBRAR M.N."),
          ("12122", "12122 - FACTURAS POR COBRAR M.E."),
          ("12211", "12211 - ANTICIPOS DE CLIENTES M.N."),
          ("12212", "12212 - ANTICIPOS DE CLIENTES M.E."),
      ]

      records = []
      for cta_codigo, cta_nombre_defecto in cuentas_config:
        f_match = next(
            (
                f
                for n, f in file_dict.items()
                if n.startswith(cta_codigo) and n.endswith(".xlsx")
            ),
            None,
        )
        if not f_match:
          continue

        f_match.seek(0)
        raw = pd.read_excel(f_match, header=None)

        cta_final = cta_nombre_defecto
        for r_i in range(min(12, len(raw))):
          v0 = str(raw.iloc[r_i, 0]).strip() if pd.notna(raw.iloc[r_i, 0]) else ""
          if cta_codigo in v0:
            cta_final = v0
            break

        ruc_actual = ""
        cliente_actual = ""

        for idx in range(7, len(raw)):
          row = raw.iloc[idx]
          c1 = str(row[1]).strip() if pd.notna(row[1]) else ""
          c2 = str(row[2]).strip() if pd.notna(row[2]) else ""
          if c1 and c1 not in ("nan", "RUC / DNI"):
            ruc_actual = c1
          if c2 and c2 not in ("nan", "Cliente"):
            cliente_actual = c2

          doc_raw = str(row[5]).strip() if pd.notna(row[5]) else ""
          tipo_raw = str(row[8]).strip() if pd.notna(row[8]) else ""
          abono_soles = str(row[12]).strip() if pd.notna(row[12]) else ""
          glosa_raw = str(row[10]).strip() if pd.notna(row[10]) else ""

          es_acumulado = (
              ("ACUMULADO" in abono_soles)
              or ("ACUMULADO" in glosa_raw)
              or ("ACUMULADO" in doc_raw)
          )
          es_cabecera = tipo_raw == "Tipo de Libro" or doc_raw == "Documento"

          fecha_op = row[6]
          tiene_datos = (
              (tipo_raw != "" and tipo_raw != "nan")
              or pd.notna(fecha_op)
              or (doc_raw != "" and doc_raw != "nan")
          )

          if tiene_datos and not es_acumulado and not es_cabecera:
            td = row[3] if pd.notna(row[3]) else ""
            serie = (
                str(row[4]).strip()
                if pd.notna(row[4]) and str(row[4]).strip() != "nan"
                else ""
            )
            doc = doc_raw if doc_raw != "nan" else ""
            fecha_vcto = row[7]
            tipo_libro = tipo_raw if tipo_raw != "nan" else ""
            nro_comp = (
                str(row[9]).strip()
                if pd.notna(row[9]) and str(row[9]).strip() != "nan"
                else ""
            )
            glosa = glosa_raw if glosa_raw != "nan" else ""

            debe_s = row[11]
            haber_s = row[12]
            soles = row[13]
            debe_d = row[14]
            haber_d = row[15]
            dolares = row[16]

            g_comp = ""
            tipo_up = tipo_libro.upper()
            if tipo_up == "VENTAS":
              g_comp = f"{serie}-{doc}" if serie else doc
            elif tipo_up == "APERTURA":
              g_comp = doc
            elif tipo_up in ("BANCO", "VARIOS"):
              if (
                  nro_comp in ant_by_nrocomp
                  and ant_by_nrocomp[nro_comp]["Comprobante"]
              ):
                g_comp = ant_by_nrocomp[nro_comp]["Comprobante"]
              elif doc in ant_by_comp:
                g_comp = doc

              if not g_comp:
                if tipo_up == "BANCO":
                  if nro_comp in pagos_map and pagos_map[nro_comp]:
                    g_comp = pagos_map[nro_comp]
                  elif doc in pagos_map and pagos_map[doc]:
                    g_comp = pagos_map[doc]
                  else:
                    g_comp = doc
                elif tipo_up == "VARIOS":
                  if (
                      "Diferencia en tasa de cambio" in doc
                      or "Diferencia en tasa de cambio" in glosa
                  ):
                    g_comp = nro_comp
                  else:
                    g_comp = doc
            else:
              g_comp = doc

            status = np.nan
            vendedor = np.nan
            cotizacion = np.nan
            obs = np.nan
            rev = np.nan
            coment = np.nan

            info_ant = ant_by_comp.get(g_comp) or ant_by_nrocomp.get(nro_comp)
            if info_ant:
              status = info_ant["Status"]
              vendedor = info_ant["Vendedor"]
              cotizacion = info_ant["Cotización"]
              obs = info_ant["Observación"]
              rev = info_ant["Revisión"]
              coment = info_ant["Comentarios de Ventas"]

            if pd.isna(vendedor) or str(vendedor).strip() in (
                "",
                "nan",
                "None",
            ):
              if g_comp in fact_map:
                vendedor = fact_map[g_comp]["Vendedor"]
                cotizacion = fact_map[g_comp]["Cotización"]

            if pd.notna(cotizacion) and str(cotizacion).strip() in muestras_set:
              status = "Muestra"

            records.append({
                "CTA CTBLE": cta_final,
                "RUC / DNI": str(ruc_actual),
                "Cliente": cliente_actual,
                "TD": td,
                "Serie": serie,
                "Documento": doc,
                "Comprobante": g_comp,
                "Fecha Operación": fecha_op,
                "Fecha Vcto": fecha_vcto,
                "Tipo de Libro": tipo_libro,
                "Nro de Comprobante": nro_comp,
                "Glosa": glosa,
                "DEBE": debe_s,
                "HABER": haber_s,
                "SOLES": soles,
                "DEBE_ME": debe_d,
                "HABER_ME": haber_d,
                "DOLARES": dolares,
                "Status": status,
                "Vendedor": vendedor,
                "Cotización": cotizacion,
                "Observación": obs,
                "Revisión": rev,
                "Comentarios de Ventas": coment,
            })

      df_final = pd.DataFrame(records)

      output = io.BytesIO()
      with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_final.to_excel(
            writer, sheet_name="Documentos Pendientes", index=False
        )
      output.seek(0)

      filename = "Documentos_Pendientes_Consolidado.xlsx"
      st.success(f"Reporte generado con {len(df_final):,} registros.")
      st.download_button(
          label="📥 Descargar Archivo Excel",
          data=output,
          file_name=filename,
          mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      )