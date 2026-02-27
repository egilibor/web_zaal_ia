import sys
import uuid
import shutil
import tempfile
import subprocess
import urllib.parse
from pathlib import Path

import streamlit as st
import pandas as pd

# --- CONFIGURACIÓN ---
st.set_page_config(page_title="ZAAL IA - Logística Universal", layout="wide", page_icon="🚚")
st.title("🚀 ZAAL IA: Portal de Reparto Inteligente")

# --- RUTAS DE PROYECTO ---
REPO_DIR = Path(__file__).resolve().parent
SCRIPT_GPT = REPO_DIR / "reparto_gpt.py"
SCRIPT_GEMINI = REPO_DIR / "reparto_gemini.py"

def get_workdir():
    if "workdir" not in st.session_state:
        st.session_state.workdir = Path(tempfile.mkdtemp(prefix="zaal_app_"))
    return Path(st.session_state.workdir)

workdir = get_workdir()

# --- MONITOR DE EJECUCIÓN ---
def run_ia_process(cmd, cwd):
    log_area = st.empty()
    full_log = ""
    try:
        process = subprocess.Popen(
            cmd, cwd=str(cwd), stdout=subprocess.PIPE, stderr=subprocess.STDOUT, 
            text=True, bufsize=1, universal_newlines=True
        )
        for line in process.stdout:
            full_log += line
            log_area.code(full_log)
        process.wait()
        return process.returncode
    except Exception as e:
        st.error(f"Fallo crítico: {e}")
        return 1

# --- INTERFAZ ---
tab1, tab2 = st.tabs(["🏗️ Generar Plan", "📍 Navegación"])

with tab1:
    st.subheader("Procesamiento de Rutas (Cualquier Zona)")
    csv_file = st.file_uploader("Sube el CSV de llegadas", type=["csv"])
    
    if csv_file and st.button("🚀 INICIAR OPTIMIZACIÓN", type="primary"):
        input_path = workdir / "llegadas.csv"
        input_path.write_bytes(csv_file.getbuffer())
        
        # Copia de reglas necesarias
        if (REPO_DIR / "Reglas_hospitales.xlsx").exists():
            shutil.copy(REPO_DIR / "Reglas_hospitales.xlsx", workdir / "Reglas_hospitales.xlsx")

        # FASE 1
        st.info("Fase 1: Clasificando por zonas...")
        rc1 = run_ia_process([sys.executable, str(SCRIPT_GPT), "--csv", "llegadas.csv", "--out", "salida.xlsx"], workdir)
        
        if rc1 == 0:
            # FASE 2: Aquí es donde Gemini aplica su mapa interno
            st.info("Fase 2: Aplicando optimización geográfica universal...")
            xl = pd.ExcelFile(workdir / "salida.xlsx")
            hojas = [h for h in xl.sheet_names if not any(x in h.upper() for x in ["METADATOS", "RESUMEN"])]
            rango = f"0-{len(hojas)-1}"
            
            # El script de Gemini ahora manda en el orden
            rc2 = run_ia_process([sys.executable, str(SCRIPT_GEMINI), "--seleccion", rango, "--in", "salida.xlsx", "--out", "PLAN.xlsx"], workdir)
            
            if rc2 == 0:
                st.success("✅ Plan optimizado y listo para descargar.")

    # Descarga
    if (workdir / "PLAN.xlsx").exists():
        st.download_button("💾 DESCARGAR PLAN.XLSX", (workdir / "PLAN.xlsx").read_bytes(), "PLAN.xlsx", use_container_width=True)

with tab2:
    st.subheader("📍 Navegación por Tramos (Orden del Plan)")
    f_user = st.file_uploader("Subir PLAN.xlsx", type=["xlsx"], key="nav_uploader")
    
    path_plan = None
    if f_user:
        path_plan = workdir / "nav_temp.xlsx"
        path_plan.write_bytes(f_user.getbuffer())
    elif (workdir / "PLAN.xlsx").exists():
        path_plan = workdir / "PLAN.xlsx"

    if path_plan:
        # Cargamos el Excel respetando el orden original de las filas
        xl = pd.ExcelFile(path_plan)
        hojas = [h for h in xl.sheet_names if not any(x in h.upper() for x in ["METADATOS", "RESUMEN", "LOG"])]
        sel = st.selectbox("Selecciona la ruta a navegar:", hojas)
        
        df = pd.read_excel(path_plan, sheet_name=sel)
        
        # BUSCADOR DE COLUMNAS
        c_dir = next((c for c in df.columns if "DIR" in str(c).upper()), None)
        c_pob = next((c for c in df.columns if "POB" in str(c).upper()), "")

        if c_dir:
            # IMPORTANTE: No usamos sort_values(). Usamos el orden tal cual viene del Excel.
            direcciones = [urllib.parse.quote(f"{f[c_dir]}, {f[c_pob]}") for _, f in df.iterrows()]
            
            st.info(f"🚩 Ruta: {sel} | Paradas: {len(direcciones)}")
            
            
            
            for i in range(0, len(direcciones), 9):
                t = direcciones[i:i+9]
                origen = urllib.parse.quote("Vall d'Uxo, Castellon") if i == 0 else ""
                
                # Generamos la URL de Google Maps
                # origin=0 (con coordenadas/dirección) o origin=5 (ubicación actual)
                if origen:
                    url = f"https://www.google.com/maps/dir/?api=1&origin={origen}&destination={t[-1]}&waypoints={'|'.join(t[:-1])}"
                else:
                    url = f"https://www.google.com/maps/dir/?api=1&destination={t[-1]}&waypoints={'|'.join(t[:-1])}"
                
                # Etiqueta dinámica: muestra la primera y última parada del tramo
                st.link_button(f"🚗 TRAMO {i//9 + 1}: {df.iloc[i][c_dir]} ➡️ {df.iloc[i+len(t)-1][c_dir]}", url, use_container_width=True)
