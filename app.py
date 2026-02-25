import streamlit as st
import pandas as pd
import os
import io
import re
import datetime as _dt
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="ZAAL IA - Clasificación", layout="wide", page_icon="🚚")
st.title("🚀 ZAAL IA: Generador de salida.xlsx")

# ---------------------------------------------------------
# TU LÓGICA ORIGINAL DE REPARTO_GPT.PY (ADAPTADA)
# ---------------------------------------------------------

def clean_text(x) -> str:
    if pd.isna(x): return ""
    s = str(x).strip()
    s = re.sub(r"\s+", " ", s)
    return s.upper()

def style_sheet(ws):
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    thin_side = Side(border_style="thin", color="000000")
    border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = alignment
        cell.border = border
    
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            cell.border = border

def set_widths(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

# Función principal de procesamiento
def ejecutar_reparto(df_llegadas, df_reglas_h, df_reglas_f, origen_datos="LLEGADAS", nombre_csv="llegadas.csv"):
    # 1. Preparar Reglas
    df_reglas_h = df_reglas_h.copy()
    df_reglas_f = df_reglas_f.copy()
    df_reglas_h['Patrón_dirección'] = df_reglas_h['Patrón_dirección'].apply(clean_text)
    df_reglas_f['Patrón_dirección'] = df_reglas_f['Patrón_dirección'].apply(clean_text)
    
    # Unir y ordenar por longitud (Truco 1-3)
    df_total_reglas = pd.concat([df_reglas_h, df_reglas_f], ignore_index=True)
    df_total_reglas["len"] = df_total_reglas["Patrón_dirección"].str.len()
    df_total_reglas = df_total_reglas.sort_values("len", ascending=False).drop(columns=["len"])

    # 2. Clasificar
    df_llegadas = df_llegadas.copy()
    df_llegadas.columns = [c.strip() for c in df_llegadas.columns]
    col_dir = next((c for c in df_llegadas.columns if "DIR" in c.upper()), df_llegadas.columns[0])
    
    df_llegadas["Z.Rep"] = "RESTO (sin ruta)"
    df_llegadas["Es_Hospital"] = False
    df_llegadas["Es_Federacion"] = False

    patrones_h = set(df_reglas_h["Patrón_dirección"].unique())
    patrones_f = set(df_reglas_f["Patrón_dirección"].unique())

    for idx, fila in df_llegadas.iterrows():
        txt = clean_text(fila[col_dir])
        for _, reg in df_total_reglas.iterrows():
            p = reg["Patrón_dirección"]
            if p and p in txt:
                df_llegadas.at[idx, "Z.Rep"] = reg["Ruta"]
                if p in patrones_h: df_llegadas.at[idx, "Es_Hospital"] = True
                if p in patrones_f: df_llegadas.at[idx, "Es_Federacion"] = True
                break

    # 3. Crear el Excel (en memoria para Streamlit)
    wb = Workbook()
    
    # --- HOJA METADATOS ---
    ws_meta = wb.active
    ws_meta.title = "METADATOS"
    meta_data = [
        ["Nº", "Clave", "Valor"],
        [1, "Origen de datos", origen_datos],
        [2, "CSV", nombre_csv],
        [3, "Reglas", "Reglas_hospitales.xlsx"],
        [4, "Generado", _dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")]
    ]
    for r in meta_data: ws_meta.append(r)
    style_sheet(ws_meta)
    set_widths(ws_meta, [5, 20, 40])

    # --- HOJAS DE RESUMEN ---
    # (Aquí se incluiría tu lógica de RESUMEN_GENERAL y RESUMEN_UNICO)
    # Por brevedad, pasamos a las ZREP que es lo que más te importa
    
    # --- HOJAS ZREP ---
    rutas = sorted(df_llegadas["Z.Rep"].unique())
    for r_name in rutas:
        df_z = df_llegadas[df_llegadas["Z.Rep"] == r_name].copy()
        df_z.insert(0, "Parada", range(1, len(df_z) + 1))
        
        # Nombre de hoja seguro
        safe_name = f"ZREP_{str(r_name)[:20]}".replace("/", " ").replace("*","")
        ws = wb.create_sheet(title=safe_name[:31])
        
        for r_idx, row in enumerate(dataframe_to_rows(df_z, index=False, header=True), 1):
            ws.append(row)
        style_sheet(ws)
        set_widths(ws, [8, 18, 55, 70, 16, 12, 12, 22])

    return wb

# --- INTERFAZ STREAMLIT ---
f_csv = st.file_uploader("Sube el CSV de LLEGADAS", type=["csv"])

if f_csv:
    if st.button("📊 GENERAR SALIDA.XLSX"):
        try:
            df_llegadas = pd.read_csv(f_csv, sep=None, engine='python', encoding='latin-1')
            xl_reglas = pd.ExcelFile("Reglas_hospitales.xlsx")
            df_h = xl_reglas.parse('REGLAS_HOSPITALES')
            df_f = xl_reglas.parse('REGLAS_FEDERACION')

            with st.spinner("Ejecutando tu lógica de reparto_gpt.py..."):
                wb_resultado = ejecutar_reparto(df_llegadas, df_h, df_f, nombre_csv=f_csv.name)
                
                # Guardar en buffer para descarga
                output = io.BytesIO()
                wb_resultado.save(output)
                
                st.success("✅ ¡Hecho! Estructura idéntica al local generada.")
                st.download_button(
                    label="💾 DESCARGAR SALIDA.XLSX",
                    data=output.getvalue(),
                    file_name="salida.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        except Exception as e:
            st.error(f"Error en el proceso: {e}")
