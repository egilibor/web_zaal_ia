import sys
import uuid
import shutil
import tempfile
import subprocess
import urllib.parse
import os
from pathlib import Path
import streamlit as st
import pandas as pd

# --- CONFIGURACIÓN ---
st.set_page_config(page_title="ZAAL IA - Panel de Control", layout="wide")

# --- RUTAS DEL REPOSITORIO ---
REPO_DIR = Path(__file__).resolve().parent
SCRIPT_GPT = REPO_DIR / "reparto_gpt.py"
SCRIPT_GEMINI = REPO_DIR / "reparto_gemini.py"
REGLAS_HOSP = REPO_DIR / "Reglas_hospitales.xlsx"

# Carpeta de trabajo en /tmp (Evita el error de inotify)
if "workdir" not in st.session_state:
    st.session_state.workdir = Path(tempfile.gettempdir()) / f"zaal_{uuid.uuid4().hex[:8]}"
    st.session_state.workdir.mkdir(parents=True, exist_ok=True)

workdir = st.session_state.workdir

# --- SIDEBAR: ESTADO DEL SISTEMA (Estilo GPT) ---
with st.sidebar:
    st.header("Estado")
    st.code(f"Workdir: {workdir}")
    
    st.write("### Archivos del Sistema")
    st.write(f"GPT: `reparto_gpt.py` {'✅' if SCRIPT_GPT.exists() else '❌'}")
    st.write(f"Gemini: `reparto_gemini.py` {'✅' if SCRIPT_GEMINI.exists() else '❌'}")
    st.write(f"Reglas: `Reglas_hospitales.xlsx` {'✅' if REGLAS_HOSP.exists() else '❌'}")
    
    st.divider()
    if st.button("Reset sesión"):
        for key in list(st.session_state.keys()):
            del st.session_state[key]
        st.rerun()

# --- CUERPO PRINCIPAL ---
st.title("🚚 ZAAL IA: Sistema de Reparto")

# 1) Subida de archivo
st.subheader("1) Subir CSV de llegadas")
csv_file = st.file_uploader("Arrastra el archivo aquí", type=["csv"], label_visibility="collapsed")

if csv_file:
    (workdir / "llegadas.csv").write_bytes(csv_file.getbuffer())
    
    # 2) Generación de salida.xlsx
    st.subheader("2) Generar salida.xlsx")
    if st.button("Generar salida.xlsx", type="primary"):
        # Misma etiqueta en el spinner para evitar confusiones
        with st.spinner("Generando salida.xlsx..."):
            if REGLAS_HOSP.exists():
                shutil.copy(REGLAS_HOSP, workdir / "Reglas_hospitales.xlsx")
            
            # Ejecución
            subprocess.run([sys.executable, str(SCRIPT_GPT), "--csv", "llegadas.csv", "--out", "salida.xlsx"], 
                           cwd=str(workdir))
            st.rerun()

# 3) Selección de ruta (Solo si existe el archivo)
f_salida = workdir / "salida.xlsx"
if f_salida.exists():
    st.divider()
    st.subheader("3) Seleccionar Ruta de salida.xlsx")
    
    try:
        xl = pd.ExcelFile(f_salida)
        hojas = [h for h in xl.sheet_names if not any(x in h.upper() for x in ["METADATOS", "RESUMEN"])]
        
        col_sel, col_opt = st.columns([3, 1])
        with col_sel:
            ruta_sel = st.selectbox("Rutas detectadas:", hojas)
        
        with col_opt:
            st.write("") # Alineación
            if st.button(f"Optimizar {ruta_sel}"):
                idx = xl.sheet_names.index(ruta_sel)
                with st.spinner(f"Optimizar {ruta_sel}..."):
                    subprocess.run([
                        sys.executable, str(SCRIPT_GEMINI), 
                        "--seleccion", str(idx), 
                        "--in", "salida.xlsx", 
                        "--out", "PLAN.xlsx"
                    ], cwd=str(workdir))
                    st.session_state.ready = True
                    st.session_state.nombre = ruta_sel
                    st.rerun()
    except Exception as e:
        st.error(f"Error al leer salida.xlsx: {e}")

# 4) Resultados y Mapas
f_plan = workdir / "PLAN.xlsx"
if st.session_state.get("ready") and f_plan.exists():
    st.divider()
    st.subheader(f"📍 Mapa: {st.session_state.nombre}")
    
    df = pd.read_excel(f_plan, sheet_name=st.session_state.nombre)
    c_dir = next((c for c in df.columns if "DIR" in str(c).upper()), None)
    c_pob = next((c for c in df.columns if "POB" in str(c).upper()), "")

    if c_dir:
        direcciones = [urllib.parse.quote(f"{f[c_dir]}, {f[c_pob]}") for _, f in df.iterrows()]
        
        # Bloques de navegación
        for i in range(0, len(direcciones), 9):
            t = direcciones[i:i+9]
            origen = urllib.parse.quote("Vall d'Uxo, Castellon") if i == 0 else ""
            prefix = "0" if origen else "3"
            url = f"http://googleusercontent.com/maps.google.com/{prefix}{origen}&destination={t[-1]}&waypoints={'|'.join(t[:-1])}"
            st.link_button(f"🚗 TRAMO {i//9 + 1}: Empezar en {df.iloc[i][c_dir]}", url, use_container_width=True)
    
    st.download_button("💾 Descargar PLAN.xlsx", f_plan.read_bytes(), "PLAN_ZAAL.xlsx")
