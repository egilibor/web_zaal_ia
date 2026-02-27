import sys
import uuid
import shutil
import tempfile
import subprocess
import urllib.parse
from pathlib import Path
import streamlit as st
import pandas as pd

# --- CONFIGURACIÓN DE LA PÁGINA ---
st.set_page_config(page_title="ZAAL IA - Control Profesional", layout="wide")
st.title("🚀 ZAAL IA: Gestión de Rutas Castellón")

# --- RUTAS DE ARCHIVOS EN EL REPOSITORIO ---
REPO_DIR = Path(__file__).resolve().parent
SCRIPT_GPT = REPO_DIR / "reparto_gpt.py"
SCRIPT_GEMINI = REPO_DIR / "reparto_gemini.py"
REGLAS_HOSP = REPO_DIR / "Reglas_hospitales.xlsx"

# --- GESTIÓN DE CARPETA DE TRABAJO (En /tmp para evitar errores de inotify) ---
if "workdir" not in st.session_state:
    # Creamos una subcarpeta única en el temporal del sistema
    temp_path = Path(tempfile.gettempdir()) / f"zaal_session_{uuid.uuid4().hex[:6]}"
    temp_path.mkdir(parents=True, exist_ok=True)
    st.session_state.workdir = temp_path

workdir = st.session_state.workdir

# --- PASO 1: CARGA DEL ARCHIVO CSV ---
st.sidebar.header("1. Entrada de Datos")
csv_file = st.sidebar.file_uploader("Sube el CSV de llegadas", type=["csv"])

if csv_file:
    # Guardamos el CSV en la carpeta temporal de trabajo
    input_path = workdir / "llegadas.csv"
    input_path.write_bytes(csv_file.getbuffer())
    
    # --- PASO 2: CLASIFICACIÓN (Ejecución de reparto_gpt.py) ---
    st.sidebar.header("2. Procesamiento")
    if st.sidebar.button("📦 EJECUTAR FASE 1"):
        with st.spinner("Clasificando envíos..."):
            # Copiamos las reglas necesarias a la carpeta de trabajo para el script
            if REGLAS_HOSP.exists():
                shutil.copy(REGLAS_HOSP, workdir / "Reglas_hospitales.xlsx")
            
            # Ejecutamos el script de clasificación
            res = subprocess.run(
                [sys.executable, str(SCRIPT_GPT), "--csv", "llegadas.csv", "--out", "salida.xlsx"],
                cwd=str(workdir), capture_output=True, text=True
            )
            
            if res.returncode == 0:
                st.session_state.fase1_completada = True
                st.sidebar.success("✅ Archivo salida.xlsx generado.")
            else:
                st.sidebar.error(f"Error en Fase 1: {res.stderr}")

# --- PASO 3: SELECCIÓN DE RUTA (Lectura de Sheets de salida.xlsx) ---
f_salida = workdir / "salida.xlsx"
if f_salida.exists():
    st.divider()
    st.subheader("🎯 Paso 3: Selección de Ruta")
    
    try:
        # Abrimos el Excel para leer las pestañas generadas
        xl = pd.ExcelFile(f_salida)
        # Filtramos hojas que no son rutas de reparto
        hojas_rutas = [h for h in xl.sheet_names if not any(x in h.upper() for x in ["METADATOS", "RESUMEN", "LOG"])]
        
        col_sel, col_opt = st.columns([3, 1])
        with col_sel:
            ruta_seleccionada = st.selectbox("Elige la ruta que quieres optimizar ahora:", hojas_rutas)
        
        # --- PASO 4: OPTIMIZACIÓN A PETICIÓN (Ejecución de reparto_gemini.py) ---
        with col_opt:
            st.write("") # Alineación visual
            if st.button("🚀 OPTIMIZAR ESTA HOJA", type="primary"):
                # Buscamos el índice real de la hoja en el Excel original
                idx_hoja = xl.sheet_names.index(ruta_seleccionada)
                
                with st.spinner(f"Optimizando {ruta_seleccionada}..."):
                    # Ejecutamos Gemini solo para la hoja seleccionada
                    cmd_gemini = [
                        sys.executable, str(SCRIPT_GEMINI), 
                        "--seleccion", str(idx_hoja), 
                        "--in", "salida.xlsx", 
                        "--out", "PLAN_FINAL.xlsx"
                    ]
                    res_opt = subprocess.run(cmd_gemini, cwd=str(workdir), capture_output=True, text=True)
                    
                    if res_opt.returncode == 0:
                        st.session_state.fase2_completada = True
                        st.session_state.ruta_activa = ruta_seleccionada
                    else:
                        st.error(f"Error en Optimización: {res_opt.stderr}")

    except Exception as e:
        st.error(f"Error al leer salida.xlsx: {e}")

# --- PASO 5: RESULTADO FINAL (Botones de Google Maps) ---
f_plan = workdir / "PLAN_FINAL.xlsx"
if st.session_state.get("fase2_completada") and f_plan.exists():
    st.divider()
    st.subheader(f"📍 Hoja de Ruta Optimizada: {st.session_state.ruta_activa}")
    
    # Leemos la hoja optimizada
    df_opt = pd.read_excel(f_plan, sheet_name=st.session_state.ruta_activa)
    
    # Detectamos columnas de Dirección y Población
    c_dir = next((c for c in df_opt.columns if "DIR" in str(c).upper()), None)
    c_pob = next((c for c in df_opt.columns if "POB" in str(c).upper()), "")

    if c_dir:
        # Preparamos las direcciones para los enlaces
        direcciones = [urllib.parse.quote(f"{f[c_dir]}, {f[c_pob]}".strip(", ")) for _, f in df_opt.iterrows()]
        
        st.write(f"Se han generado **{len(direcciones)} paradas** ordenadas geográficamente.")
        
        # Generación de botones por tramos de 9 paradas
        for i in range(0, len(direcciones), 9):
            t = direcciones[i:i+9]
            # Salida desde Vall d'Uxo solo para el primer tramo
            origen = urllib.parse.quote("Vall d'Uxo, Castellon") if i == 0 else ""
            prefix = "0" if origen else "3"
            
            url = f"http://googleusercontent.com/maps.google.com/{prefix}{origen}&destination={t[-1]}&waypoints={'|'.join(t[:-1])}"
            
            st.link_button(f"🚗 TRAMO {i//9 + 1}: Empezar en {df_opt.iloc[i][c_dir]}", url, use_container_width=True)
        
        st.divider()
        st.download_button("💾 DESCARGAR EXCEL FINAL", f_plan.read_bytes(), "Plan_Zaal_IA.xlsx")
    else:
        st.error("No se han encontrado columnas de dirección en el archivo optimizado.")
