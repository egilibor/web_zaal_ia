import streamlit as st
import pandas as pd
import io
import datetime as _dt
import re
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows

# --- LAS FUNCIONES DE TU REPARTO_GPT.PY ---

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

# --- INTERFAZ DE STREAMLIT ---

st.set_page_config(page_title="ZAAL IA", layout="wide")
st.title("🚚 ZAAL IA: Generador de salida.xlsx")

f_csv = st.file_uploader("Sube el CSV de LLEGADAS", type=["csv"])

if f_csv and st.button("EJECUTAR CLASIFICACIÓN"):
    try:
        # Cargar datos
        df_llegadas = pd.read_csv(f_csv, sep=None, engine='python', encoding='latin-1')
        xl_reglas = pd.ExcelFile("Reglas_hospitales.xlsx")
        df_h_reg = xl_reglas.parse('REGLAS_HOSPITALES')
        df_f_reg = xl_reglas.parse('REGLAS_FEDERACION')

        # --- LÓGICA DE PROCESAMIENTO (TU REPARTO_GPT.PY) ---
        df_h_reg['Patrón_dirección'] = df_h_reg['Patrón_dirección'].apply(clean_text)
        df_f_reg['Patrón_dirección'] = df_f_reg['Patrón_dirección'].apply(clean_text)
        
        df_total_reglas = pd.concat([df_h_reg, df_f_reg], ignore_index=True)
        df_total_reglas["len"] = df_total_reglas["Patrón_dirección"].str.len()
        df_total_reglas = df_total_reglas.sort_values("len", ascending=False).drop(columns=["len"])

        col_dir = next((c for c in df_llegadas.columns if "DIR" in c.upper()), df_llegadas.columns[0])
        df_llegadas["Z.Rep"] = "RESTO (sin ruta)"
        
        # Marcar Hospitales y Federación
        patrones_h = set(df_h_reg["Patrón_dirección"].unique())
        patrones_f = set(df_f_reg["Patrón_dirección"].unique())
        df_llegadas["Es_Hospital"] = False
        df_llegadas["Es_Federacion"] = False

        for idx, fila in df_llegadas.iterrows():
            txt = clean_text(fila[col_dir])
            for _, reg in df_total_reglas.iterrows():
                p = reg["Patrón_dirección"]
                if p and p in txt:
                    df_llegadas.at[idx, "Z.Rep"] = reg["Ruta"]
                    if p in patrones_h: df_llegadas.at[idx, "Es_Hospital"] = True
                    if p in patrones_f: df_llegadas.at[idx, "Es_Federacion"] = True
                    break

        # --- CONSTRUCCIÓN DEL EXCEL (TU ESTRUCTURA) ---
        wb = Workbook()
        
        # 1. METADATOS
        ws_meta = wb.active
        ws_meta.title = "METADATOS"
        ws_meta.append(["Nº", "Clave", "Valor"])
        ws_meta.append([1, "Origen de datos", "LLEGADAS"])
        ws_meta.append([2, "CSV", f_csv.name])
        ws_meta.append([3, "Reglas", "Reglas_hospitales.xlsx"])
        ws_meta.append([4, "Generado", _dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")])
        style_sheet(ws_meta)
        set_widths(ws_meta, [5, 20, 40])

        # 2. RESUMEN_GENERAL
        ws_res_gen = wb.create_sheet("RESUMEN_GENERAL")
        # Aquí va tu lógica de conteo por bloques
        h_mask = df_llegadas["Es_Hospital"]
        f_mask = df_llegadas["Es_Federacion"]
        r_mask = ~(h_mask | f_mask)
        
        res_data = [
            ["Nº", "Bloque", "Paradas", "Expediciones", "Kilos"],
            [1, "HOSPITALES", df_llegadas[h_mask][col_dir].nunique(), df_llegadas[h_mask]["Expediciones"].sum(), df_llegadas[h_mask]["Kilos"].sum()],
            [2, "FEDERACION", df_llegadas[f_mask][col_dir].nunique(), df_llegadas[f_mask]["Expediciones"].sum(), df_llegadas[f_mask]["Kilos"].sum()],
            [3, "RESTO (todas rutas)", df_llegadas[r_mask][col_dir].nunique(), df_llegadas[r_mask]["Expediciones"].sum(), df_llegadas[r_mask]["Kilos"].sum()]
        ]
        for row in res_data: ws_res_gen.append(row)
        style_sheet(ws_res_gen)

        # 3. HOSPITALES y FEDERACION (Hojas detalle)
        for name, mask in [("HOSPITALES", h_mask), ("FEDERACION", f_mask)]:
            ws = wb.create_sheet(name)
            df_sub = df_llegadas[mask]
            for r in dataframe_to_rows(df_sub, index=False, header=True): ws.append(r)
            style_sheet(ws)

        # 4. ZREPS
        for r_name in sorted(df_llegadas["Z.Rep"].unique()):
            safe_name = f"ZREP_{str(r_name)[:20]}".replace("/", " ")[:31]
            ws = wb.create_sheet(title=safe_name)
            df_z = df_llegadas[df_llegadas["Z.Rep"] == r_name].copy()
            df_z.insert(0, "Parada", range(1, len(df_z) + 1))
            for row in dataframe_to_rows(df_z, index=False, header=True): ws.append(row)
            style_sheet(ws)
            set_widths(ws, [8, 18, 55, 70, 16, 12, 12, 22])

        # Salida
        buf = io.BytesIO()
        wb.save(buf)
        st.success("✅ Clasificación finalizada.")
        st.download_button("💾 DESCARGAR SALIDA.XLSX", buf.getvalue(), "salida.xlsx")

    except Exception as e:
        st.error(f"Error: {e}")
