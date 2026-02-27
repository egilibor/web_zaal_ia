import sys
import shutil
import tempfile
import subprocess
import urllib.parse
from pathlib import Path

import streamlit as st
import pandas as pd

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="ZAAL IA - Gestión Profesional", layout="wide")
st.title("🚚 ZAAL IA: Control de Rutas Castellón")

# --- RUTAS DE ARCHIVOS DEL REPOSITORIO ---
REPO_DIR = Path(__file__).resolve().parent
SCRIPT_GPT = REPO_DIR / "reparto_gpt.py"
SCRIPT_GEMINI = REPO_DIR / "reparto_gemini.py"
REGLAS_HOSP = REPO_DIR / "Reglas_hospitales.xlsx"

# Carpeta de trabajo temporal por sesión
if "workdir" not in st.session_state:
    st.session_state.workdir = Path(tempfile.mkdtemp(prefix="zaal_final_"))
workdir = st.session_state.workdir

# --- PASO 1: CARGA DEL ARCHIVO ---
st.sidebar.header("📂 1. Carga de Datos")
csv_file = st.sidebar.file_uploader("Subir CSV de llegadas", type=["csv"])

if csv_file:
    # Guardar el CSV en la carpeta temporal
    input_path = workdir / "llegadas.csv"
    input_path.write_bytes(csv_file.getbuffer())
    
    # --- PASO 2: CLASIFICACIÓN (reparto_gpt.py) ---
    st.sidebar.header("⚙️ 2. Procesamiento Inicial")
    if st.sidebar.button("Generar salida.xlsx (Fase 1)"):
        with st.spinner("Ejecutando reparto_gpt.py..."):
            # Aseguramos que las reglas estén en el directorio de trabajo
            if REGLAS_HOSP.exists():
                shutil.copy(REGLAS_HOSP, workdir / "Reglas_hospitales.xlsx")
            
            # Ejecución del primer script
            res = subprocess.run(
                [sys.executable, str(SCRIPT_GPT), "--csv", "llegadas.csv", "--out", "salida.xlsx"],
                cwd=workdir, capture_output=True, text=True
            )
            
            if res.returncode == 0:
                st.session_state.fase1_ok = True
                st.sidebar.success("✅ Archivo salida.xlsx generado.")
            else:
                st.sidebar.error(f"Error en Fase 1: {res.stderr}")

# --- PASO 3: SELECCIÓN MANUAL DE RUTA (Sheets) ---
# Solo mostramos esto si el archivo salida.xlsx existe físicamente
if (workdir / "salida.xlsx").exists():
    st.divider()
    st.subheader("🎯 Selección de Ruta para Reparto")
    
    try:
        # Abrimos el Excel generado para leer los nombres de las pestañas
        xl = pd.ExcelFile(workdir / "salida.xlsx")
        hojas_disponibles = [h for h in xl.sheet_names if not any(x in h.upper() for x in ["METADATOS", "RESUMEN", "LOG"])]
        
        col_sel, col_opt = st.columns([3, 1])
        with col_sel:
            ruta_seleccionada = st.selectbox("Elige la ruta que quieres optimizar ahora:", hojas_disponibles)
        
        # --- PASO 4: OPTIMIZACIÓN A PETICIÓN (reparto_gemini.py) ---
        with col_opt:
            st.write("") # Alineación visual
            if st.button("🚀 OPTIMIZAR AHORA", type="primary"):
                # Obtenemos el índice de la hoja elegida
                idx_hoja = xl.sheet_names.index(ruta_seleccionada)
                
                with st.spinner(f"Optimizando {ruta_seleccionada}..."):
                    # Ejecución del segundo script SOLO para la hoja seleccionada
                    cmd = [
                        sys.executable, str(SCRIPT_GEMINI), 
                        "--seleccion", str(idx_hoja), 
                        "--in", "salida.xlsx", 
                        "--out", "PLAN_FINAL.xlsx"
                    ]
                    res_opt = subprocess.run(cmd, cwd=workdir, capture_output=True, text=True)
                    
                    if res_opt.returncode == 0:
                        st.session_state.fase2_ok = True
                        st.session_state.ruta_activa = ruta_seleccionada
                    else:
                        st.error(f"Error en Optimización: {res_opt.stderr}")

    except Exception as e:
        st.error(f"Error al leer las rutas de salida.xlsx: {e}")

# --- PASO 5: RESULTADO Y NAVEGACIÓN ---
if st.session_state.get("fase2_ok") and (workdir / "PLAN_FINAL.xlsx").exists():
    st.divider()
    st.subheader(f"📍 Ruta Optimizada: {st.session_state.ruta_activa}")
    
    # Cargamos los datos ya optimizados por Gemini
    df_plan = pd.read_excel(workdir / "PLAN_FINAL.xlsx", sheet_name=st.session_state.ruta_activa)
    
    # Buscamos columnas de dirección y población
    c_dir = next((c for c in df_plan.columns if "DIR" in str(c).upper()), None)
    c_pob = next((c for c in df_plan.columns if "POB" in str(c).upper()), "")

    if c_dir:
        # Preparamos las direcciones (sin reordenar nada, respetando a Gemini)
        direcciones = [urllib.parse.quote(f"{f[c_dir]}, {f[c_pob]}".strip(", ")) for _, f in df_plan.iterrows()]
        
        st.write(f"Se han generado **{len(direcciones)} paradas**.")
        
        # Bloques de 9 paradas para Google Maps
        for i in range(0, len(direcciones), 9):
            t = direcciones[i:i+9]
            # Salida desde Vall d'Uxo solo en el primer tramo
            origen = urllib.parse.quote("Vall d'Uxo, Castellon") if i == 0 else ""
            prefix = "0" if origen else "3"
            
            url = f"http://googleusercontent.com/maps.google.com/{prefix}{origen}&destination={t[-1]}&waypoints={'|'.join(t[:-1])}"
            
            # Etiqueta clara para el chófer
            st.link_button(f"🚗 TRAMO {i//9 + 1}: Empezar en {df_plan.iloc[i][c_dir]}", url, use_container_width=True)
        
        st.divider()
        st.download_button("💾 DESCARGAR PLAN EXCEL", (workdir / "PLAN_FINAL.xlsx").read_bytes(), "Plan_Zaal.xlsx")
    else:
        st.error("No se detecta la columna de direcciones en el archivo optimizado.")
