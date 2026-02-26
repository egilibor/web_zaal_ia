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
# Se asume que los scripts y archivos de reglas están en la misma carpeta que este app.py
REPO_DIR = Path(__file__).resolve().parent
SCRIPT_REPARTO = REPO_DIR / "reparto_gpt.py"
SCRIPT_GEMINI = REPO_DIR / "reparto_gemini.py"
REGLAS_REPO = REPO_DIR / "Reglas_hospitales.xlsx"

# -------------------------
# UTILIDADES
# -------------------------
def ensure_workdir() -> Path:
    """Crea y asegura un directorio de trabajo temporal para la sesión."""
    if "workdir" not in st.session_state:
        st.session_state.workdir = Path(tempfile.mkdtemp(prefix="reparto_"))
        st.session_state.run_id = str(uuid.uuid4())[:8]
    return st.session_state.workdir

def reset_session_dir():
    """Limpia el directorio temporal y reinicia la sesión."""
    wd = st.session_state.get("workdir")
    if wd and isinstance(wd, Path):
        shutil.rmtree(wd, ignore_errors=True)
    st.session_state.workdir = Path(tempfile.mkdtemp(prefix="reparto_"))
    st.session_state.run_id = str(uuid.uuid4())[:8]

def save_upload(uploaded_file, dst: Path) -> Path:
    """Guarda el archivo subido en el destino especificado."""
    dst.write_bytes(uploaded_file.getbuffer())
    return dst

def run_process(cmd: list[str], cwd: Path, timeout_s: int = 300) -> tuple[int, str, str]:
    """Ejecuta un proceso externo de forma segura."""
    try:
        p = subprocess.run(
            cmd,
            cwd=str(cwd),
            capture_output=True,
            text=True,
            timeout=timeout_s,
        )
        return p.returncode, p.stdout, p.stderr
    except subprocess.TimeoutExpired as e:
        stdout = e.stdout or ""
        stderr = e.stderr or ""
        return 124, stdout, f"TIMEOUT tras {timeout_s}s\n{stderr}"

def show_logs(stdout: str, stderr: str):
    """Muestra los registros de salida en caso de error."""
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
    st.write(f"ID Ejecución: `{st.session_state.run_id}`")
    if st.button("Limpiar y Reiniciar Sesión"):
        reset_session_dir()
        st.rerun()
    
    st.divider()
    # Verificaciones de archivos críticos en el repositorio
    st.write(f"Motor GPT: {'✅' if SCRIPT_REPARTO.exists() else '❌'}")
    st.write(f"Motor Gemini: {'✅' if SCRIPT_GEMINI.exists() else '❌'}")
    st.write(f"Reglas: {'✅' if REGLAS_REPO.exists() else '❌'}")

# Detener si faltan archivos base
if not SCRIPT_REPARTO.exists() or not SCRIPT_GEMINI.exists() or not REGLAS_REPO.exists():
    st.error("Faltan archivos críticos en el servidor. Revisa el repositorio.")
    st.stop()

# -------------------------
# MENÚ PRINCIPAL
# -------------------------
opcion = st.selectbox("Seleccione una operación:", ["Asignación de Reparto", "Google Maps (Rutas Móvil)"])

st.divider()

# -------------------------
# 1) ASIGNACIÓN DE REPARTO
# -------------------------
if opcion == "Asignación de Reparto":
    st.subheader("1. Carga de Datos")
    csv_file = st.file_uploader("Subir CSV de llegadas", type=["csv"])

    if not csv_file:
        st.info("Por favor, sube el archivo CSV para comenzar.")
        st.stop()

    # Preparar entorno de trabajo
    csv_path = save_upload(csv_file, workdir / "llegadas.csv")
    (workdir / "Reglas_hospitales.xlsx").write_bytes(REGLAS_REPO.read_bytes())

    st.subheader("2. Procesamiento")
    if st.button("Ejecutar Procesos", type="primary"):
        # ---- FASE 1: GPT (Clasificación) ----
        cmd_gpt = [
            sys.executable, str(SCRIPT_REPARTO),
            "--csv", "llegadas.csv",
            "--reglas", "Reglas_hospitales.xlsx",
            "--out", "salida.xlsx",
        ]
        
        with st.spinner("Clasificando envíos..."):
            rc, out, err = run_process(cmd_gpt, cwd=workdir)
        
        if rc != 0:
            st.error("Error en la clasificación (reparto_gpt.py)")
            show_logs(out, err)
            st.stop()

        # ---- FASE 2: GEMINI (Optimización) ----
        # Se ejecuta con una selección por defecto o según lógica previa
        cmd_gemini = [
            sys.executable, str(SCRIPT_GEMINI),
            "--seleccion", "1-9",
            "--in", "salida.xlsx",
            "--out", "PLAN.xlsx",
        ]

        with st.spinner("Optimizando rutas de carga..."):
            rc2, out2, err2 = run_process(cmd_gemini, cwd=workdir)

        if rc2 != 0:
            st.error("Error en la optimización (reparto_gemini.py)")
            show_logs(out2, err2)
            st.stop()

        st.success("✅ Procesamiento completado con éxito.")

    # Descarga de resultados si existen
    salida_path = workdir / "salida.xlsx"
    plan_path = workdir / "PLAN.xlsx"

    if salida_path.exists() and plan_path.exists():
        col1, col2 = st.columns(2)
        with col1:
            st.download_button("Descargar Clasificación (salida.xlsx)", data=salida_path.read_bytes(), file_name="salida.xlsx")
        with col2:
            st.download_button("Descargar Plan de Carga (PLAN.xlsx)", data=plan_path.read_bytes(), file_name="PLAN.xlsx")

# -------------------------
# 2) GOOGLE MAPS (RUTAS MÓVIL)
# -------------------------
elif opcion == "Google Maps (Rutas Móvil)":
    st.subheader("📍 Preparación de Navegación por Tramos")
    plan_path = workdir / "PLAN.xlsx"

    if not plan_path.exists():
        st.warning("No se ha generado ningún Plan de Carga aún. Ve a 'Asignación de Reparto' primero.")
        st.stop()

    try:
        # Cargar el plan optimizado para extraer las rutas
        xl = pd.ExcelFile(plan_path)
        hojas_zrep = [h for h in xl.sheet_names if "ZREP" in h.upper()]

        if not hojas_zrep:
            st.error("No se encontraron rutas optimizadas en el archivo PLAN.xlsx.")
            st.stop()

        ruta_seleccionada = st.selectbox("Selecciona la ruta para el conductor:", hojas_zrep)

        if ruta_seleccionada:
            df = pd.read_excel(plan_path, sheet_name=ruta_seleccionada)
            
            # Identificación de columnas de dirección y población
            col_dir = next((c for c in df.columns if "DIREC" in c.upper()), None)
            col_pob = next((c for c in df.columns if "POB" in c.upper()), "")

            if not col_dir:
                st.error("No se pudo localizar la columna de dirección en la hoja.")
                st.stop()

            # Formatear direcciones para la URL de Google Maps
            direcciones_urls = []
            for _, fila in df.iterrows():
                # Combinamos dirección y población para mayor precisión
                direccion_completa = f"{fila[col_dir]}, {fila[col_pob]}".strip(", ")
                direcciones_urls.append(urllib.parse.quote(direccion_completa))

            st.info(f"Ruta: {ruta_seleccionada} | Total Paradas: {len(direcciones_urls)}")
            st.write("Selecciona un tramo para iniciar la navegación (Máx. 9 paradas por tramo):")

            # Generación de botones por tramos de 9 paradas
            for i in range(0, len(direcciones_urls), 9):
                tramos = direcciones_urls[i : i + 9]
                destino = tramos[-1]
                puntos_paso = tramos[:-1]
                
                inicio = i + 1
                fin = i + len(tramos)
                
                # Construcción de la URL de navegación de Google Maps
                url_final = f"https://www.google.com/maps/dir/?api=1&destination={destino}"
                if puntos_paso:
                    url_final += f"&waypoints={'|'.join(puntos_paso)}"
                
                st.link_button(f"🗺️ Iniciar Tramo: Paradas {inicio} - {fin}", url_final, use_container_width=True)

    except Exception as e:
        st.error(f"Error al procesar las rutas: {e}")
