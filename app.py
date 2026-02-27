import sys
import uuid
import shutil
import tempfile
import subprocess
import urllib.parse
from pathlib import Path
import streamlit as st
import pandas as pd

# --- CONFIGURACIÓN DE INTERFAZ ---
st.set_page_config(page_title="ZAAL IA - Control de Reparto", layout="wide", page_icon="🚚")
st.title("🚀 ZAAL IA: Gestión de Rutas por Pasos")

# --- RUTAS DE SISTEMA ---
REPO_DIR = Path(__file__).resolve().parent
SCRIPT_GPT = REPO_DIR / "reparto_gpt.py"
SCRIPT_GEMINI = REPO_DIR / "reparto_gemini.py"

# --- GESTIÓN DE SESIÓN ---
def get_workdir():
    if "workdir" not in st.session_state:
        st.session_state.workdir = Path(tempfile.mkdtemp(prefix="zaal_strict_"))
    return Path(st.session_state.workdir)

workdir = get_workdir()

# --- 1. CARGA DEL ARCHIVO ---
st.sidebar.header("📁 Paso 1: Carga")
csv_file = st.sidebar.file_uploader("Sube el CSV de llegadas", type=["csv"])

if csv_file:
    input_path = workdir / "llegadas.csv"
    input_path.write_bytes(csv_file.getbuffer())
    
    # --- 2. CLASIFICACIÓN (FASE 1) ---
    st.sidebar.header("⚡ Paso 2: Clasificación")
    if st.sidebar.button("Ejecutar Clasificación General"):
        st.info("Iniciando Fase 1: Clasificando paquetes por zonas...")
        log_area = st.empty()
        
        # Ejecutamos con monitor de salida para evitar cuelgues visuales
        process = subprocess.Popen(
            [sys.executable, str(SCRIPT_GPT), "--csv", "llegadas.csv", "--out", "salida.xlsx"],
            cwd=str(workdir), stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True
        )
        
        output = ""
        for line in process.stdout:
            output += line
            log_area.code(output) # Ves el progreso real aquí
        
        rc = process.wait()
        if rc == 0:
            st.session_state.fase1_completada = True
            st.success("✅ Clasificación finalizada. Archivo 'salida.xlsx' generado.")
        else:
            st.error("❌ Fallo en la Clasificación. Revisa los logs arriba.")

# --- 3. SELECCIÓN MANUAL ---
if st.session_state.get("fase1_completada"):
    st.divider()
    st.subheader("🎯 Paso 3: Selección de Ruta")
    
    try:
        xl = pd.ExcelFile(workdir / "salida.xlsx")
        # Filtramos hojas que no son rutas
        hojas_rutas = [h for h in xl.sheet_names if not any(x in h.upper() for x in ["METADATOS", "RESUMEN"])]
        
        col_sel, col_btn = st.columns([3, 1])
        with col_sel:
            ruta_seleccionada = st.selectbox("¿Qué ruta quieres optimizar ahora?", hojas_rutas)
        
        # --- 4. OPTIMIZACIÓN A PETICIÓN ---
        with col_btn:
            st.write("") # Espaciado
            if st.button("🧠 Optimizar Ruta", type="primary"):
                st.session_state.ruta_en_proceso = ruta_seleccionada
                idx_hoja = hojas_rutas.index(ruta_seleccionada)
                
                with st.status(f"Optimizando {ruta_seleccionada}...", expanded=True) as status:
                    cmd_gemini = [
                        sys.executable, str(SCRIPT_GEMINI), 
                        "--seleccion", str(idx_hoja), 
                        "--in", "salida.xlsx", 
                        "--out", "PLAN_UNICO.xlsx"
                    ]
                    res = subprocess.run(cmd_gemini, cwd=str(workdir), capture_output=True, text=True)
                    
                    if res.returncode == 0:
                        st.session_state.fase2_completada = True
                        status.update(label="✅ Optimización Geográfica lista", state="complete")
                    else:
                        st.error(f"Error en Fase 2: {res.stderr}")

    except Exception as e:
        st.error(f"Error al leer las rutas: {e}")

# --- 5. RESULTADO (BOTONES MAPS) ---
if st.session_state.get("fase2_completada"):
    st.divider()
    st.subheader(f"📍 Hoja de Ruta Optimizada: {st.session_state.ruta_en_proceso}")
    
    # Leemos la ruta ya optimizada (respetando el orden de la IA)
    df_opt = pd.read_excel(workdir / "PLAN_UNICO.xlsx", sheet_name=st.session_state.ruta_en_proceso)
    
    # Buscamos columnas de dirección y población
    c_dir = next((c for c in df_opt.columns if "DIR" in str(c).upper()), None)
    c_pob = next((c for c in df_opt.columns if "POB" in str(c).upper()), "")

    if c_dir:
        # Preparamos las direcciones para los enlaces
        direcciones = [urllib.parse.quote(f"{f[c_dir]}, {f[c_pob]}") for _, f in df_opt.iterrows()]
        
        st.info(f"Se han generado {len(direcciones)} paradas siguiendo el hilo lógico de la carretera.")
        
        

        # Generación de botones por tramos de 9
        cols_maps = st.columns(3)
        for i in range(0, len(direcciones), 9):
            t = direcciones[i:i+9]
            origen = urllib.parse.quote("Vall d'Uxo, Castellon") if i == 0 else ""
            
            # URL de Google Maps: 0 para origen fijo, 3 para ubicación actual
            prefix = "0" if origen else "3"
            url = f"http://googleusercontent.com/maps.google.com/{prefix}{origen}&destination={t[-1]}&waypoints={'|'.join(t[:-1])}"
            
            with cols_maps[(i//9) % 3]:
                st.link_button(f"🚗 TRAMO {i//9 + 1}", url, use_container_width=True)
                st.caption(f"De: {df_opt.iloc[i][c_dir]}")
    else:
        st.error("No se ha encontrado la columna de dirección en el archivo optimizado.")
