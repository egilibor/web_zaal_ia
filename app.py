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
st.set_page_config(page_title="ZAAL IA - Optimizador Individual", layout="wide", page_icon="🚚")
st.title("🚀 ZAAL IA: Optimización de Ruta por Selección")

# --- RUTAS ---
REPO_DIR = Path(__file__).resolve().parent
SCRIPT_GPT = REPO_DIR / "reparto_gpt.py"
SCRIPT_GEMINI = REPO_DIR / "reparto_gemini.py"

def get_workdir():
    if "workdir" not in st.session_state:
        st.session_state.workdir = Path(tempfile.mkdtemp(prefix="zaal_puesto_"))
    return Path(st.session_state.workdir)

workdir = get_workdir()

# --- INTERFAZ ---
st.sidebar.header("1. Carga de Datos")
csv_file = st.sidebar.file_uploader("Subir CSV de llegadas", type=["csv"])

if csv_file:
    # Guardamos el CSV inicial
    (workdir / "llegadas.csv").write_bytes(csv_file.getbuffer())
    
    if st.sidebar.button("📦 Clasificar envíos"):
        with st.spinner("Clasificando por zonas..."):
            res = subprocess.run([sys.executable, str(SCRIPT_GPT), "--csv", "llegadas.csv", "--out", "salida.xlsx"], 
                                 cwd=workdir, capture_output=True, text=True)
            if res.returncode == 0:
                st.session_state.clasificado = True
                st.success("✅ Clasificación lista.")
            else:
                st.error(f"Error: {res.stderr}")

# --- SELECCIÓN Y OPTIMIZACIÓN ---
if st.session_state.get("clasificado"):
    st.divider()
    st.subheader("2. Selección de Ruta para Optimizar")
    
    xl = pd.ExcelFile(workdir / "salida.xlsx")
    hojas = [h for h in xl.sheet_names if not any(x in h.upper() for x in ["METADATOS", "RESUMEN"])]
    
    col1, col2 = st.columns([2, 1])
    with col1:
        ruta_sel = st.selectbox("Elige la ruta que vas a repartir ahora:", hojas)
    with col2:
        idx_sel = hojas.index(ruta_sel)
        btn_opt = st.button("🧠 Optimizar esta Ruta ahora", type="primary")

    if btn_opt:
        with st.spinner(f"Optimizando geográficamente {ruta_sel}..."):
            # Llamamos a Gemini SOLO para la ruta elegida (usando su índice)
            # El parámetro --seleccion acepta el índice de la hoja
            cmd = [sys.executable, str(SCRIPT_GEMINI), "--seleccion", str(idx_sel), "--in", "salida.xlsx", "--out", "PLAN_INDIVIDUAL.xlsx"]
            res = subprocess.run(cmd, cwd=workdir, capture_output=True, text=True)
            
            if res.returncode == 0:
                st.session_state.optimizado = True
                st.session_state.ruta_actual = ruta_sel
                st.success(f"✅ {ruta_sel} optimizada.")
            else:
                st.error("Error en la optimización. Probablemente Gemini tardó demasiado.")

# --- RESULTADOS Y MAPA ---
if st.session_state.get("optimizado"):
    st.divider()
    st.subheader(f"📍 Hoja de Ruta: {st.session_state.ruta_actual}")
    
    df = pd.read_excel(workdir / "PLAN_INDIVIDUAL.xlsx", sheet_name=st.session_state.ruta_actual)
    
    # Identificar columnas para Maps
    c_dir = next((c for c in df.columns if "DIR" in str(c).upper()), None)
    c_pob = next((c for c in df.columns if "POB" in str(c).upper()), "")

    if c_dir:
        # Mostramos la tabla para verificar orden
        st.dataframe(df[[c_dir, c_pob]].head(10), use_container_width=True)
        
        # Generar Tramos
        direcciones = [urllib.parse.quote(f"{f[c_dir]}, {f[c_pob]}") for _, f in df.iterrows()]
        
        st.write("### 📲 Enlaces para el móvil")
        for i in range(0, len(direcciones), 9):
            t = direcciones[i:i+9]
            # Primer tramo sale de Vall d'Uxo
            origen = urllib.parse.quote("Vall d'Uxo, Castellon") if i == 0 else ""
            
            url = f"http://googleusercontent.com/maps.google.com/{'0' if origen else '3'}{origen}&destination={t[-1]}&waypoints={'|'.join(t[:-1])}"
            
            # Etiqueta visual
            desc_tramo = f"Tramo {i//9 + 1}: {df.iloc[i][c_dir]} ({df.iloc[i][c_pob]})"
            st.link_button(f"🗺️ {desc_tramo}", url, use_container_width=True)
