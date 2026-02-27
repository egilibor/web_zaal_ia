import sys
import uuid
import shutil
import tempfile
import subprocess
import urllib.parse
import os
import time
from pathlib import Path
import streamlit as st
import pandas as pd

# --- CONFIGURACIÓN DE ESTADO ---
st.set_page_config(page_title="ZAAL IA - Logística Profesional", layout="wide", page_icon="🚚")
st.title("🚀 ZAAL IA: Sistema de Reparto Universal")

# --- RUTAS DE SCRIPTS ---
REPO_DIR = Path(__file__).resolve().parent
SCRIPT_GPT = REPO_DIR / "reparto_gpt.py"
SCRIPT_GEMINI = REPO_DIR / "reparto_gemini.py"

# --- GESTIÓN DE DIRECTORIO ---
def get_workdir():
    if "workdir" not in st.session_state:
        st.session_state.workdir = Path(tempfile.mkdtemp(prefix="zaal_pro_"))
    return Path(st.session_state.workdir)

workdir = get_workdir()

# --- MOTOR DE EJECUCIÓN (Anticuelgues) ---
def run_ia_task(cmd, description):
    """Ejecuta la IA con una barra de progreso real para que la web no se duerma"""
    with st.spinner(f"Ejecutando {description}..."):
        try:
            # Usamos un timeout largo pero definido para evitar cuelgues infinitos
            process = subprocess.run(
                cmd, 
                cwd=str(workdir), 
                capture_output=True, 
                text=True, 
                timeout=300 # 5 minutos máximo por fase
            )
            if process.returncode != 0:
                st.error(f"Error en {description}: {process.stderr}")
                return False
            return True
        except subprocess.TimeoutExpired:
            st.error(f"La IA ha tardado demasiado en {description}. Reintenta con menos datos.")
            return False
        except Exception as e:
            st.error(f"Error inesperado: {e}")
            return False

# --- INTERFAZ ---
tab1, tab2 = st.tabs(["🏗️ Generador de Plan", "📍 Navegación Móvil"])

with tab1:
    st.subheader("Optimización Geográfica Universal")
    st.info("Este sistema usa inteligencia geográfica para ordenar cualquier zona (interior o costa) sin usar el abecedario.")
    
    csv_file = st.file_uploader("Sube el CSV de llegadas", type=["csv"])
    
    if csv_file:
        input_path = workdir / "llegadas.csv"
        input_path.write_bytes(csv_file.getbuffer())
        
        # Copiamos reglas si existen
        if (REPO_DIR / "Reglas_hospitales.xlsx").exists():
            shutil.copy(REPO_DIR / "Reglas_hospitales.xlsx", workdir / "Reglas_hospitales.xlsx")

        if st.button("🚀 GENERAR PLAN COMPLETO", type="primary"):
            # Fase 1
            if run_ia_task([sys.executable, str(SCRIPT_GPT), "--csv", "llegadas.csv", "--out", "salida.xlsx"], "Clasificación (Fase 1)"):
                
                # Fase 2
                try:
                    xl = pd.ExcelFile(workdir / "salida.xlsx")
                    hojas = [h for h in xl.sheet_names if not any(x in h.upper() for x in ["METADATOS", "RESUMEN"])]
                    rango = f"0-{len(hojas)-1}"
                    
                    if run_ia_task([sys.executable, str(SCRIPT_GEMINI), "--seleccion", rango, "--in", "salida.xlsx", "--out", "PLAN.xlsx"], "Optimización Geográfica (Fase 2)"):
                        st.success("✅ ¡Éxito! Plan generado correctamente.")
                        st.balloons()
                except Exception as e:
                    st.error(f"Error al procesar el archivo de salida: {e}")

    # Descarga
    if (workdir / "PLAN.xlsx").exists():
        st.download_button("💾 DESCARGAR PLAN.XLSX", (workdir / "PLAN.xlsx").read_bytes(), "PLAN.xlsx", use_container_width=True)

with tab2:
    st.subheader("📍 Navegación por Tramos")
    f_user = st.file_uploader("Subir PLAN.xlsx para ruta", type=["xlsx"], key="nav_uploader")
    
    path_nav = None
    if f_user:
        path_nav = workdir / "nav_user.xlsx"
        path_nav.write_bytes(f_user.getbuffer())
    elif (workdir / "PLAN.xlsx").exists():
        path_nav = workdir / "PLAN.xlsx"

    if path_nav:
        xl_nav = pd.ExcelFile(path_nav)
        hojas_nav = [h for h in xl_nav.sheet_names if not any(x in h.upper() for x in ["METADATOS", "RESUMEN", "LOG"])]
        sel_ruta = st.selectbox("Selecciona la ruta para el chófer:", hojas_nav)
        
        df = pd.read_excel(path_nav, sheet_name=sel_ruta)
        
        # Identificar columnas
        c_dir = next((c for c in df.columns if "DIR" in str(c).upper()), None)
        c_pob = next((c for c in df.columns if "POB" in str(c).upper()), "")

        if c_dir:
            # RESPETO TOTAL AL ORDEN DE LA IA: No reordenamos nada aquí.
            direcciones = [urllib.parse.quote(f"{f[c_dir]}, {f[c_pob]}") for _, f in df.iterrows() if len(str(f[c_dir])) > 3]
            
            st.write(f"📦 **{len(direcciones)} paradas** detectadas en el orden óptimo.")
            
            
            
            for i in range(0, len(direcciones), 9):
                t = direcciones[i:i+9]
                origen = urllib.parse.quote("Vall d'Uxo, Castellon") if i == 0 else ""
                
                # URL de Google Maps (Modo Navegación)
                if origen:
                    url = f"https://www.google.com/maps/dir/?api=1&origin={origen}&destination={t[-1]}&waypoints={'|'.join(t[:-1])}"
                else:
                    url = f"https://www.google.com/maps/dir/?api=1&destination={t[-1]}&waypoints={'|'.join(t[:-1])}"
                
                # Botón de tramo con información visual de la calle
                calle_inicio = df.iloc[i][c_dir]
                st.link_button(f"🚗 TRAMO {i//9 + 1}: Empezar en {calle_inicio}", url, use_container_width=True)
