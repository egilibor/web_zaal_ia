import streamlit as st
import pandas as pd
import subprocess
import sys
import os
import uuid
from pathlib import Path

# --- CONFIGURACIÓN ---
st.set_page_config(page_title="ZAAL IA - Panel Estado", layout="wide")

# Rutas fijas del proyecto
REPO_DIR = Path(__file__).resolve().parent
PYTHON_EXE = sys.executable

# Carpeta de trabajo limpia en /tmp (para evitar el error de inotify)
if "workdir" not in st.session_state:
    st.session_state.workdir = Path("/tmp") / f"reparto_{uuid.uuid4().hex[:6]}"
    st.session_state.workdir.mkdir(parents=True, exist_ok=True)

workdir = st.session_state.workdir

# --- SIDEBAR: ESTADO (Tal cual la captura que enviaste) ---
with st.sidebar:
    st.header("Estado")
    st.write(f"**Run:** `{uuid.uuid4().hex[:7]}`")
    st.write(f"**Workdir:** `{workdir}`")
    st.write(f"**Repo dir:** `{REPO_DIR}`")
    st.write(f"**Python:** `{PYTHON_EXE}`")
    
    st.divider()
    # Verificación de archivos en tiempo real
    gpt_exists = (REPO_DIR / "reparto_gpt.py").exists()
    gemini_exists = (REPO_DIR / "reparto_gemini.py").exists()
    reglas_exists = (REPO_DIR / "Reglas_hospitales.xlsx").exists()
    
    st.write(f"GPT: `reparto_gpt.py` exists = **{gpt_exists}**")
    st.write(f"Gemini: `reparto_gemini.py` exists = **{gemini_exists}**")
    st.write(f"Reglas: `Reglas_hospitales.xlsx` exists = **{reglas_exists}**")
    
    if st.button("Reset sesión"):
        st.session_state.clear()
        st.rerun()

# --- CUERPO PRINCIPAL ---
st.title("🚚 ZAAL IA: Sistema de Reparto")

# 1) SUBIR CSV
st.subheader("1) Subir CSV de llegadas")
csv_file = st.file_uploader("CSV de llegadas", type=["csv"], label_visibility="collapsed")

if csv_file:
    # Guardar archivo
    input_path = workdir / "llegadas.csv"
    with open(input_path, "wb") as f:
        f.write(csv_file.getbuffer())

    # 2) EJECUTAR FASE 1
    st.subheader("2) Ejecutar (genera salida.xlsx)")
    if st.button("Ejecutar", type="primary"):
        with st.status("Ejecutando clasificación...", expanded=True) as status:
            # Limpieza previa
            if (workdir / "salida.xlsx").exists(): os.remove(workdir / "salida.xlsx")
            
            # Ejecución directa
            subprocess.run([PYTHON_EXE, str(REPO_DIR / "reparto_gpt.py"), "--csv", "llegadas.csv", "--out", "salida.xlsx"], cwd=str(workdir))
            
            if (workdir / "salida.xlsx").exists():
                status.update(label="✅ Archivo salida.xlsx generado", state="complete")
                st.rerun()
            else:
                st.error("Error: El script no generó salida.xlsx")

# 3) SELECCIÓN Y OPTIMIZACIÓN (Tu flujo solicitado)
if (workdir / "salida.xlsx").exists():
    st.divider()
    st.subheader("3) Seleccionar Ruta y Optimizar")
    
    xl = pd.ExcelFile(workdir / "salida.xlsx")
    hojas = [h for h in xl.sheet_names if not any(x in h.upper() for x in ["METADATOS", "RESUMEN"])]
    
    col1, col2 = st.columns([3, 1])
    with col1:
        ruta_sel = st.selectbox("Elige la hoja:", hojas)
    with col2:
        st.write("")
        if st.button(f"Optimizar {ruta_sel}"):
            idx = xl.sheet_names.index(ruta_sel)
            with st.spinner(f"Gemini optimizando {ruta_sel}..."):
                subprocess.run([PYTHON_EXE, str(REPO_DIR / "reparto_gemini.py"), "--seleccion", str(idx), "--in", "salida.xlsx", "--out", "PLAN.xlsx"], cwd=str(workdir))
                st.session_state.finalizado = True
                st.session_state.ruta_nombre = ruta_sel
                st.rerun()

# 4) RESULTADOS
if st.session_state.get("finalizado") and (workdir / "PLAN.xlsx").exists():
    st.success(f"✅ Ruta {st.session_state.ruta_nombre} optimizada correctamente.")
    
    col_dl1, col_dl2 = st.columns(2)
    with col_dl1:
        st.download_button("Descargar salida.xlsx", (workdir / "salida.xlsx").read_bytes(), "salida.xlsx")
    with col_dl2:
        st.download_button("Descargar PLAN.xlsx", (workdir / "PLAN.xlsx").read_bytes(), "PLAN_ZAAL.xlsx")

    # Botones de Google Maps
    df = pd.read_excel(workdir / "PLAN.xlsx", sheet_name=st.session_state.ruta_nombre)
    c_dir = next((c for c in df.columns if "DIR" in str(c).upper()), None)
    if c_dir:
        st.write("### 📍 Enlaces de Navegación")
        direcciones = [urllib.parse.quote(str(d)) for d in df[c_dir].tolist()]
        for i in range(0, len(direcciones), 9):
            t = direcciones[i:i+9]
            url = f"https://www.google.com/maps/dir/?api=1&origin=Vall+dUxo&destination={t[-1]}&waypoints={'|'.join(t[:-1])}"
            st.link_button(f"🚗 TRAMO {i//9 + 1}", url, use_container_width=True)
