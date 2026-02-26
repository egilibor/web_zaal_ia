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

def reset_session_dir():
    wd = st.session_state.get("workdir")
    if wd and isinstance(wd, Path):
        shutil.rmtree(wd, ignore_errors=True)
    st.session_state.workdir = Path(tempfile.mkdtemp(prefix="reparto_"))
    st.session_state.run_id = str(uuid.uuid4())[:8]

def save_upload(uploaded_file, dst: Path) -> Path:
    dst.write_bytes(uploaded_file.getbuffer())
    return dst

def run_process(cmd: list[str], cwd: Path, timeout_s: int = 300) -> tuple[int, str, str]:
    try:
        p = subprocess.run(cmd, cwd=str(cwd), capture_output=True, text=True, timeout=timeout_s)
        return p.returncode, p.stdout, p.stderr
    except subprocess.TimeoutExpired as e:
        return 124, e.stdout or "", f"TIMEOUT tras {timeout_s}s"

def show_logs(stdout: str, stderr: str):
    if stdout.strip():
        st.subheader("STDOUT")
        st.code(stdout)
    if stderr.strip():
        st.subheader("STDERR")
        st.code(stderr)

# -------------------------
# ESTADO Y VERIFICACIONES
# -------------------------
workdir = ensure_workdir()

with st.sidebar:
    st.header("Estado del Sistema")
    st.write(f"ID Sesión: `{st.session_state.run_id}`")
    if st.button("Limpiar y Reiniciar"):
        reset_session_dir()
        st.rerun()
    st.divider()
    st.write(f"Motor GPT: {'✅' if SCRIPT_REPARTO.exists() else '❌'}")
    st.write(f"Motor Gemini: {'✅' if SCRIPT_GEMINI.exists() else '❌'}")
    st.write(f"Reglas: {'✅' if REGLAS_REPO.exists() else '❌'}")

# -------------------------
# MENÚ PRINCIPAL
# -------------------------
opcion = st.selectbox("Seleccione operación:", ["Asignación de Reparto", "Google Maps (Rutas Móvil)"])
st.divider()

# -------------------------
# 1) ASIGNACIÓN DE REPARTO
# -------------------------
if opcion == "Asignación de Reparto":
    st.subheader("Generar Nuevo Plan")
    csv_file = st.file_uploader("Subir CSV de llegadas", type=["csv"])

    if csv_file:
        save_upload(csv_file, workdir / "llegadas.csv")
        (workdir / "Reglas_hospitales.xlsx").write_bytes(REGLAS_REPO.read_bytes())

        if st.button("Ejecutar Procesamiento Completo", type="primary"):
            # Fase 1: Clasificación
            cmd_gpt = [sys.executable, str(SCRIPT_REPARTO), "--csv", "llegadas.csv", "--reglas", "Reglas_hospitales.xlsx", "--out", "salida.xlsx"]
            with st.spinner("Clasificando..."):
                rc, out, err = run_process(cmd_gpt, cwd=workdir)
            
            if rc == 0:
                # Fase 2: Optimización
                cmd_gemini = [sys.executable, str(SCRIPT_GEMINI), "--seleccion", "1-9", "--in", "salida.xlsx", "--out", "PLAN.xlsx"]
                with st.spinner("Optimizando..."):
                    rc2, out2, err2 = run_process(cmd_gemini, cwd=workdir)
                
                if rc2 == 0:
                    st.success("✅ ¡Plan generado!")
                    st.session_state.plan_listo = True
                else:
                    st.error("Error en optimización"); show_logs(out2, err2)
            else:
                st.error("Error en clasificación"); show_logs(out, err)

    # Descargas
    plan_path = workdir / "PLAN.xlsx"
    if plan_path.exists():
        st.download_button("💾 DESCARGAR PLAN.XLSX", plan_path.read_bytes(), "PLAN.xlsx")

# -------------------------
# 2) GOOGLE MAPS (RUTAS MÓVIL)
# -------------------------
elif opcion == "Google Maps (Rutas Móvil)":
    st.subheader("📍 Navegación por Tramos")
    
    # NUEVO: Pedir el fichero directamente
    f_plan_user = st.file_uploader("Subir archivo PLAN.xlsx optimizado", type=["xlsx"], help="Sube el archivo generado previamente para crear los enlaces de Maps")

    # Decidir qué archivo usar: el subido o el que esté en el workdir
    plan_source = None
    if f_plan_user:
        plan_source = f_plan_user
    elif (workdir / "PLAN.xlsx").exists():
        plan_source = workdir / "PLAN.xlsx"
        st.info("Utilizando el Plan generado en la sesión actual.")

    if plan_source:
        try:
            xl = pd.ExcelFile(plan_source)
            hojas_zrep = [h for h in xl.sheet_names if "ZREP" in h.upper()]

            if hojas_zrep:
                ruta_sel = st.selectbox("Selecciona la ruta (ZREP):", hojas_zrep)
                df = pd.read_excel(plan_source, sheet_name=ruta_sel)
                
                col_dir = next((c for c in df.columns if "DIREC" in c.upper()), None)
                col_pob = next((c for c in df.columns if "POB" in c.upper()), "")

                if col_dir:
                    direcciones = []
                    for _, fila in df.iterrows():
                        addr = f"{fila[col_dir]}, {fila[col_pob]}".strip(", ")
                        direcciones.append(urllib.parse.quote(addr))

                    st.success(f"Ruta: {ruta_sel} | {len(direcciones)} paradas detectadas.")
                    
                    # Generar botones por tramos
                    for i in range(0, len(direcciones), 9):
                        tramos = direcciones[i : i + 9]
                        destino = tramos[-1]
                        waypoints = tramos[:-1]
                        
                        url = f"https://www.google.com/maps/dir/?api=1&destination={destino}"
                        if waypoints:
                            url += f"&waypoints={'|'.join(waypoints)}"
                        
                        st.link_button(f"🗺️ Tramo {i+1} - {i+len(tramos)}", url, use_container_width=True)
                else:
                    st.error("No se encontró la columna de dirección.")
            else:
                st.warning("El archivo no contiene pestañas de ruta (ZREP).")
        except Exception as e:
            st.error(f"Error al leer el archivo: {e}")
    else:
        st.info("Sube un archivo `PLAN.xlsx` para generar los tramos de navegación.")
