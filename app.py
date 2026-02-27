import sys
import shutil
import tempfile
import subprocess
import urllib.parse
from pathlib import Path

import streamlit as st
import pandas as pd

# --- CONFIGURACIÓN ---
st.set_page_config(page_title="ZAAL IA - Control de Rutas", layout="wide")
st.title("🚚 ZAAL IA: Gestión por Selección")

# --- RUTAS DE SCRIPTS ---
REPO_DIR = Path(__file__).resolve().parent
SCRIPT_GPT = REPO_DIR / "reparto_gpt.py"
SCRIPT_GEMINI = REPO_DIR / "reparto_gemini.py"

# Carpeta de trabajo por sesión
if "workdir" not in st.session_state:
    st.session_state.workdir = Path(tempfile.mkdtemp(prefix="zaal_final_"))
workdir = st.session_state.workdir

# --- 1. CARGA DE ARCHIVO ---
csv_file = st.sidebar.file_uploader("1. Sube el CSV de llegadas", type=["csv"])

if csv_file:
    input_path = workdir / "llegadas.csv"
    input_path.write_bytes(csv_file.getbuffer())
    
    # --- 2. EJECUCIÓN FASE 1 (reparto_gpt.py) ---
    if st.sidebar.button("2. Ejecutar Clasificación"):
        with st.spinner("Generando salida.xlsx..."):
            # Copiamos reglas si existen
            reglas = REPO_DIR / "Reglas_hospitales.xlsx"
            if reglas.exists():
                shutil.copy(reglas, workdir / "Reglas_hospitales.xlsx")
            
            # Lanzamos el script que crea las pestañas
            res = subprocess.run(
                [sys.executable, str(SCRIPT_GPT), "--csv", "llegadas.csv", "--out", "salida.xlsx"],
                cwd=workdir, capture_output=True, text=True
            )
            
            if res.returncode == 0:
                st.session_state.fase1_ok = True
                st.rerun() # Forzamos recarga para que aparezca el selector
            else:
                st.error(f"Error en Fase 1: {res.stderr}")

# --- 3. ANÁLISIS DE SHEETS Y SELECCIÓN ---
# Si el archivo ya ha sido generado, lo leemos y mostramos las rutas
f_salida = workdir / "salida.xlsx"
if f_salida.exists():
    st.divider()
    st.subheader("📋 Rutas encontradas en salida.xlsx")
    
    try:
        xl = pd.ExcelFile(f_salida)
        # Filtramos pestañas técnicas
        hojas = [h for h in xl.sheet_names if not any(x in h.upper() for x in ["METADATOS", "RESUMEN", "LOG"])]
        
        col1, col2 = st.columns([3, 1])
        with col1:
            ruta_sel = st.selectbox("Selecciona la ruta que quieres procesar:", hojas)
        
        # --- 4. OPTIMIZACIÓN A PETICIÓN (reparto_gemini.py) ---
        with col2:
            st.write("") # Alineación
            if st.button("🚀 Optimizar Ruta", type="primary"):
                idx = xl.sheet_names.index(ruta_sel)
                with st.spinner(f"Optimizando {ruta_sel}..."):
                    cmd = [sys.executable, str(SCRIPT_GEMINI), "--seleccion", str(idx), "--in", "salida.xlsx", "--out", "PLAN.xlsx"]
                    res_opt = subprocess.run(cmd, cwd=workdir, capture_output=True, text=True)
                    
                    if res_opt.returncode == 0:
                        st.session_state.fase2_ok = True
                        st.session_state.ruta_nombre = ruta_sel
                    else:
                        st.error(f"Error en Optimización: {res_opt.stderr}")
                        
    except Exception as e:
        st.error(f"No se pudo leer salida.xlsx: {e}")

# --- 5. RESULTADO FINAL (Google Maps) ---
f_plan = workdir / "PLAN.xlsx"
if st.session_state.get("fase2_ok") and f_plan.exists():
    st.divider()
    st.subheader(f"📍 Navegación: {st.session_state.ruta_nombre}")
    
    df = pd.read_excel(f_plan, sheet_name=st.session_state.ruta_nombre)
    c_dir = next((c for c in df.columns if "DIR" in str(c).upper()), None)
    c_pob = next((c for c in df.columns if "POB" in str(c).upper()), "")

    if c_dir:
        direcciones = [urllib.parse.quote(f"{f[c_dir]}, {f[c_pob]}".strip(", ")) for _, f in df.iterrows()]
        
        
        
        for i in range(0, len(direcciones), 9):
            t = direcciones[i:i+9]
            # Primer tramo sale de Vall d'Uxo
            origen = urllib.parse.quote("Vall d'Uxo, Castellon") if i == 0 else ""
            prefix = "0" if origen else "3"
            url = f"http://googleusercontent.com/maps.google.com/{prefix}{origen}&destination={t[-1]}&waypoints={'|'.join(t[:-1])}"
            
            st.link_button(f"🚗 TRAMO {i//9 + 1}: {df.iloc[i][c_dir]}", url, use_container_width=True)
