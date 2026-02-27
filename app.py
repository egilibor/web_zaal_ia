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
st.set_page_config(page_title="ZAAL IA - Control Directo", layout="wide")
st.title("🚚 ZAAL IA: Gestión de Rutas (Paso a Paso)")

# --- RUTAS DE PROYECTO ---
REPO_DIR = Path(__file__).resolve().parent
SCRIPT_GPT = REPO_DIR / "reparto_gpt.py"
SCRIPT_GEMINI = REPO_DIR / "reparto_gemini.py"
REGLAS_HOSP = REPO_DIR / "Reglas_hospitales.xlsx"

# Gestión de carpeta de trabajo por sesión
if "workdir" not in st.session_state:
    st.session_state.workdir = Path(tempfile.mkdtemp(prefix="zaal_final_"))
workdir = st.session_state.workdir

# --- PASO 1: CARGA DEL ARCHIVO ---
st.sidebar.header("1. Carga de Datos")
csv_file = st.sidebar.file_uploader("Subir CSV de llegadas", type=["csv"])

if csv_file:
    # Guardar CSV original
    input_path = workdir / "llegadas.csv"
    input_path.write_bytes(csv_file.getbuffer())
    
    # --- PASO 2: CLASIFICACIÓN (FASE 1) ---
    st.sidebar.header("2. Clasificación")
    if st.sidebar.button("📦 EJECUTAR FASE 1 (Generar salida.xlsx)"):
        with st.spinner("Clasificando envíos por rutas..."):
            # Copiamos reglas si existen para que el script las encuentre
            if REGLAS_HOSP.exists():
                shutil.copy(REGLAS_HOSP, workdir / "Reglas_hospitales.xlsx")
            
            # Ejecución
            res = subprocess.run([sys.executable, str(SCRIPT_GPT), "--csv", "llegadas.csv", "--out", "salida.xlsx"], 
                                 cwd=workdir, capture_output=True, text=True)
            
            if res.returncode == 0:
                st.session_state.fase1_ok = True
                st.sidebar.success("✅ salida.xlsx generado.")
            else:
                st.sidebar.error(f"Error en Fase 1: {res.stderr}")

# --- PASO 3: SELECCIÓN MANUAL (Leer Sheets) ---
if (workdir / "salida.xlsx").exists():
    st.divider()
    st.subheader("🎯 Paso 3: Selección de Ruta")
    
    try:
        xl = pd.ExcelFile(workdir / "salida.xlsx")
        # Filtramos hojas que no son de reparto
        hojas_disponibles = [h for h in xl.sheet_names if not any(x in h.upper() for x in ["METADATOS", "RESUMEN", "LOG"])]
        
        col_sel, col_opt = st.columns([3, 1])
        with col_sel:
            ruta_elegida = st.selectbox("Selecciona la pestaña de salida.xlsx que quieres procesar:", hojas_disponibles)
        
        # --- PASO 4: OPTIMIZACIÓN A PETICIÓN ---
        with col_opt:
            st.write("") # Alineación
            if st.button("🚀 OPTIMIZAR RUTA ELEGIDA", type="primary"):
                # Obtenemos el índice real de la hoja en el Excel original
                idx_real = xl.sheet_names.index(ruta_elegida)
                
                with st.spinner(f"Optimizando geográficamente {ruta_elegida}..."):
                    # Llamamos a Gemini SOLO para esa ruta específica
                    cmd = [sys.executable, str(SCRIPT_GEMINI), "--seleccion", str(idx_real), "--in", "salida.xlsx", "--out", "PLAN_FINAL.xlsx"]
                    res_opt = subprocess.run(cmd, cwd=workdir, capture_output=True, text=True)
                    
                    if res_opt.returncode == 0:
                        st.session_state.fase2_ok = True
                        st.session_state.ruta_nombre = ruta_elegida
                        st.success(f"✅ Optimización de {ruta_elegida} completada.")
                    else:
                        st.error(f"Error en Fase 2: {res_opt.stderr}")

    except Exception as e:
        st.error(f"No se pudo leer el archivo de salida: {e}")

# --- PASO 5: RESULTADO (MAPS) ---
if st.session_state.get("fase2_ok") and (workdir / "PLAN_FINAL.xlsx").exists():
    st.divider()
    st.subheader(f"📍 Navegación Optimizada: {st.session_state.ruta_nombre}")
    
    # Leemos la ruta específica del plan optimizado
    df = pd.read_excel(workdir / "PLAN_FINAL.xlsx", sheet_name=st.session_state.ruta_nombre)
    
    # Buscamos columnas críticas (Dirección y Población)
    c_dir = next((c for c in df.columns if "DIR" in str(c).upper()), None)
    c_pob = next((c for c in df.columns if "POB" in str(c).upper() or "LOC" in str(c).upper()), "")

    if c_dir:
        # Preparamos las paradas (respetando el orden exacto de la IA)
        direcciones = [urllib.parse.quote(f"{f[c_dir]}, {f[c_pob]}".strip(", ")) for _, f in df.iterrows()]
        
        st.info(f"Total: {len(direcciones)} paradas. Se han dividido en tramos de 9 para Google Maps.")
        
        # Generación de botones por tramos
        for i in range(0, len(direcciones), 9):
            t = direcciones[i:i+9]
            # Primer tramo sale de Vall d'Uxo, los siguientes usan ubicación actual
            origen = urllib.parse.quote("Vall d'Uxo, Castellon") if i == 0 else ""
            prefix = "0" if origen else "3"
            
            url = f"http://googleusercontent.com/maps.google.com/{prefix}{origen}&destination={t[-1]}&waypoints={'|'.join(t[:-1])}"
            
            # Etiqueta del botón (Primera calle del tramo)
            st.link_button(f"🚗 TRAMO {i//9 + 1}: Empezar en {df.iloc[i][c_dir]}", url, use_container_width=True)
        
        # Opción de descarga del archivo final
        st.download_button("💾 Descargar PLAN_FINAL.xlsx", (workdir / "PLAN_FINAL.xlsx").read_bytes(), "PLAN_ZAAL.xlsx")
    else:
        st.error("No se ha encontrado la columna de dirección en la hoja optimizada.")
