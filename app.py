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
st.set_page_config(page_title="ZAAL IA - Fase 1 Debug", layout="wide", page_icon="🚚")
st.title("🚀 ZAAL IA: Monitor de Clasificación")

# --- RUTAS ---
REPO_DIR = Path(__file__).resolve().parent
SCRIPT_GPT = REPO_DIR / "reparto_gpt.py"
SCRIPT_GEMINI = REPO_DIR / "reparto_gemini.py"

def get_workdir():
    if "workdir" not in st.session_state:
        st.session_state.workdir = Path(tempfile.mkdtemp(prefix="zaal_diag_"))
    return Path(st.session_state.workdir)

workdir = get_workdir()

# --- MONITOR DE LOGS EN VIVO ---
def ejecutar_con_consola(cmd, titulo):
    st.write(f"### {titulo}")
    consola = st.empty()
    output_acumulado = ""
    
    process = subprocess.Popen(
        cmd, cwd=str(workdir), stdout=subprocess.PIPE, stderr=subprocess.STDOUT, 
        text=True, bufsize=1, universal_newlines=True
    )
    
    # Leemos la salida línea a línea mientras ocurre
    for line in process.stdout:
        output_acumulado += line
        consola.code(output_acumulado) # Esto muestra el log en tiempo real
        
    rc = process.wait()
    return rc

# --- INTERFAZ ---
with st.sidebar:
    st.header("1. Entrada de Datos")
    csv_file = st.file_uploader("Subir CSV de llegadas", type=["csv"])

if csv_file:
    # Guardar el archivo para el proceso
    (workdir / "llegadas.csv").write_bytes(csv_file.getbuffer())
    
    if st.button("📦 CLASIFICAR (FASE 1)", type="primary"):
        # Verificación de archivos antes de empezar
        if not SCRIPT_GPT.exists():
            st.error(f"Error: No encuentro {SCRIPT_GPT}")
        else:
            # Ejecución con monitor
            rc = ejecutar_con_consola(
                [sys.executable, str(SCRIPT_GPT), "--csv", "llegadas.csv", "--out", "salida.xlsx"],
                "Monitor de Clasificación GPT"
            )
            
            if rc == 0:
                st.session_state.paso1_ok = True
                st.success("✅ Clasificación terminada.")
            else:
                st.error("❌ La clasificación se ha detenido con errores.")

# --- SELECCIÓN DE RUTA (Solo si el Paso 1 terminó) ---
if st.session_state.get("paso1_ok"):
    st.divider()
    try:
        xl = pd.ExcelFile(workdir / "salida.xlsx")
        hojas = [h for h in xl.sheet_names if not any(x in h.upper() for x in ["METADATOS", "RESUMEN"])]
        
        col1, col2 = st.columns([2, 1])
        with col1:
            ruta_sel = st.selectbox("Ruta a optimizar:", hojas)
        with col2:
            idx = hojas.index(ruta_sel)
            if st.button("🧠 OPTIMIZAR ESTA RUTA"):
                rc_opt = ejecutar_con_consola(
                    [sys.executable, str(SCRIPT_GEMINI), "--seleccion", str(idx), "--in", "salida.xlsx", "--out", "PLAN_FINAL.xlsx"],
                    f"Optimizando {ruta_sel}..."
                )
                if rc_opt == 0:
                    st.session_state.paso2_ok = True
                    st.session_state.ruta_nombre = ruta_sel
    except Exception as e:
        st.error(f"Error leyendo el resultado: {e}")

# --- RESULTADO FINAL ---
if st.session_state.get("paso2_ok"):
    st.divider()
    df = pd.read_excel(workdir / "PLAN_FINAL.xlsx", sheet_name=st.session_state.ruta_nombre)
    
    c_dir = next((c for c in df.columns if "DIR" in str(c).upper()), None)
    c_pob = next((c for c in df.columns if "POB" in str(c).upper()), "")

    if c_dir:
        st.write(f"### 📍 Mapa de {st.session_state.ruta_nombre}")
        direcciones = [urllib.parse.quote(f"{f[c_dir]}, {f[c_pob]}") for _, f in df.iterrows()]
        
        for i in range(0, len(direcciones), 9):
            t = direcciones[i:i+9]
            origen = urllib.parse.quote("Vall d'Uxo, Castellon") if i == 0 else ""
            url = f"http://googleusercontent.com/maps.google.com/{'0' if origen else '3'}{origen}&destination={t[-1]}&waypoints={'|'.join(t[:-1])}"
            st.link_button(f"🚗 TRAMO {i//9 + 1}", url, use_container_width=True)
