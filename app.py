import sys
import uuid
import shutil
import tempfile
import subprocess
import urllib.parse
from pathlib import Path

import streamlit as st
import pandas as pd

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="ZAAL IA - Gestión de Reparto", layout="wide", page_icon="🚚")
st.title("🚀 ZAAL IA: Portal de Reparto Automatizado")

# --- PATHS EN REPOSITORIO ---
REPO_DIR = Path(__file__).resolve().parent
SCRIPT_REPARTO = REPO_DIR / "reparto_gpt.py"
SCRIPT_GEMINI = REPO_DIR / "reparto_gemini.py"
REGLAS_REPO = REPO_DIR / "Reglas_hospitales.xlsx"

# -------------------------
# UTILIDADES
# -------------------------
def ensure_workdir() -> Path:
    if "workdir" not in st.session_state:
        st.session_state.workdir = Path(tempfile.mkdtemp(prefix="reparto_"))
        st.session_state.run_id = str(uuid.uuid4())[:8]
    return st.session_state.workdir

def save_upload(uploaded_file, dst: Path) -> Path:
    dst.write_bytes(uploaded_file.getbuffer())
    return dst

def run_process(cmd: list[str], cwd: Path) -> tuple[int, str, str]:
    try:
        p = subprocess.run(cmd, cwd=str(cwd), capture_output=True, text=True, timeout=300)
        return p.returncode, p.stdout, p.stderr
    except Exception as e:
        return 1, "", str(e)

# -------------------------
# ESTADO
# -------------------------
workdir = ensure_workdir()

with st.sidebar:
    st.header("Control de Sesión")
    if st.button("Limpiar todo y reiniciar"):
        shutil.rmtree(workdir, ignore_errors=True)
        for key in list(st.session_state.keys()): del st.session_state[key]
        st.rerun()
    st.divider()
    st.write(f"ID Sesión: `{st.session_state.run_id}`")

# -------------------------
# MENÚ PRINCIPAL
# -------------------------
opcion = st.selectbox("Menú de Operaciones:", ["Asignación de Reparto", "Google Maps (Rutas Móvil)"])
st.divider()

# -------------------------
# 1) ASIGNACIÓN DE REPARTO
# -------------------------
if opcion == "Asignación de Reparto":
    st.subheader("Generar Clasificación y Plan de Carga")
    csv_file = st.file_uploader("Sube el CSV de llegadas", type=["csv"])

    if csv_file:
        save_upload(csv_file, workdir / "llegadas.csv")
        (workdir / "Reglas_hospitales.xlsx").write_bytes(REGLAS_REPO.read_bytes())

        if st.button("🚀 EJECUTAR PROCESO COMPLETO", type="primary"):
            # Fase 1: Clasificación
            cmd_gpt = [sys.executable, str(SCRIPT_REPARTO), "--csv", "llegadas.csv", "--reglas", "Reglas_hospitales.xlsx", "--out", "salida.xlsx"]
            with st.spinner("Procesando clasificación..."):
                rc1, out1, err1 = run_process(cmd_gpt, cwd=workdir)
            
            if rc1 == 0:
                # Fase 2: Optimización
                cmd_gemini = [sys.executable, str(SCRIPT_GEMINI), "--seleccion", "1-9", "--in", "salida.xlsx", "--out", "PLAN.xlsx"]
                with st.spinner("Procesando optimización..."):
                    rc2, out2, err2 = run_process(cmd_gemini, cwd=workdir)
                
                if rc2 == 0:
                    st.success("✅ ¡Archivos generados!")
                else:
                    st.error("Falló la optimización."); st.code(err2)
            else:
                st.error("Falló la clasificación."); st.code(err1)

    # --- DESCARGAS SIEMPRE VISIBLES ---
    salida_p = workdir / "salida.xlsx"
    plan_p = workdir / "PLAN.xlsx"

    if salida_p.exists() or plan_p.exists():
        st.markdown("### 📥 Descargar Resultados")
        c1, c2 = st.columns(2)
        with c1:
            if salida_p.exists():
                st.download_button("💾 DESCARGAR SALIDA.XLSX", salida_p.read_bytes(), "salida.xlsx", use_container_width=True)
        with c2:
            if plan_p.exists():
                st.download_button("💾 DESCARGAR PLAN.XLSX", plan_p.read_bytes(), "PLAN.xlsx", use_container_width=True)

# -------------------------
# 2) GOOGLE MAPS
# -------------------------
elif opcion == "Google Maps (Rutas Móvil)":
    st.subheader("📍 Generar Enlaces para el Chófer")
    
    # Subida manual opcional
    f_user = st.file_uploader("Sube un archivo PLAN.xlsx para extraer rutas", type=["xlsx"])
    
    path_plan = None
    if f_user:
        path_plan = save_upload(f_user, workdir / "temp_plan.xlsx")
    elif (workdir / "PLAN.xlsx").exists():
        path_plan = workdir / "PLAN.xlsx"
        st.success("Utilizando PLAN.xlsx de la sesión actual.")

    if path_plan:
        try:
            xl = pd.ExcelFile(path_plan)
            # CORRECCIÓN: Filtrar solo las hojas que NO son técnicas (Metadatos, Resumen, etc.)
            hojas_excluir = ["METADATOS", "RESUMEN", "LOG", "INSTRUCCIONES", "HOJA1"]
            hojas_validas = [h for h in xl.sheet_names if not any(ex in h.upper() for ex in hojas_excluir)]
            
            if hojas_validas:
                sel = st.selectbox("Selecciona la Ruta / Zona:", hojas_validas)
                df = pd.read_excel(path_plan, sheet_name=sel)
                
                # Identificar columnas de dirección
                c_dir = next((c for c in df.columns if "DIREC" in c.upper()), None)
                c_pob = next((c for c in df.columns if "POB" in c.upper()), "")

                if c_dir:
                    # Limpiar y preparar direcciones
                    direcciones = []
                    for _, fila in df.iterrows():
                        addr = f"{fila[c_dir]}, {fila[c_pob]}".strip(", ")
                        if addr:
                            direcciones.append(urllib.parse.quote(addr))
                    
                    st.write(f"**Ruta seleccionada:** {sel} ({len(direcciones)} paradas)")
                    
                    # Tramos de 9 paradas
                    for i in range(0, len(direcciones), 9):
                        t = direcciones[i:i+9]
                        destino = t[-1]
                        waypoints = t[:-1]
                        
                        # URL oficial de navegación (Navigation Mode)
                        link = f"https://www.google.com/maps/dir/?api=1&destination={destino}"
                        if waypoints:
                            link += f"&waypoints={'|'.join(waypoints)}"
                        
                        st.link_button(f"🗺️ Iniciar Tramo {i+1} al {i+len(t)}", link, use_container_width=True)
                else:
                    st.error(f"No encuentro la columna 'DIRECCION' en la hoja '{sel}'.")
            else:
                st.warning("No he encontrado ninguna hoja de ruta válida en el archivo.")
        except Exception as e:
            st.error(f"Error al leer el archivo: {e}")
