import sys
import uuid
import shutil
import tempfile
import subprocess
import urllib.parse
import time
from pathlib import Path

import streamlit as st
import pandas as pd

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="ZAAL IA - Logística", layout="wide", page_icon="🚚")
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

def run_process(cmd: list[str], cwd: Path):
    try:
        # Aumentamos el timeout a 10 minutos por seguridad
        p = subprocess.run(cmd, cwd=str(cwd), capture_output=True, text=True, timeout=600)
        return p.returncode, p.stdout, p.stderr
    except Exception as e:
        return 1, "", f"Error de ejecución: {str(e)}"

# -------------------------
# INICIALIZACIÓN
# -------------------------
workdir = ensure_workdir()

with st.sidebar:
    st.header("⚙️ Control")
    if st.button("🔄 Reiniciar Aplicación"):
        shutil.rmtree(workdir, ignore_errors=True)
        for key in list(st.session_state.keys()): del st.session_state[key]
        st.rerun()
    st.divider()
    st.info(f"ID Sesión: {st.session_state.run_id}")

# -------------------------
# MENÚ
# -------------------------
opcion = st.selectbox("Operación:", ["1. Asignación de Reparto", "2. Google Maps (Rutas Móvil)"])
st.divider()

# -------------------------
# 1) ASIGNACIÓN DE REPARTO
# -------------------------
if opcion == "1. Asignación de Reparto":
    st.subheader("Clasificación y Optimización de Rutas")
    csv_file = st.file_uploader("Sube el CSV de llegadas", type=["csv"])

    if csv_file:
        save_upload(csv_file, workdir / "llegadas.csv")
        if REGLAS_REPO.exists():
            (workdir / "Reglas_hospitales.xlsx").write_bytes(REGLAS_REPO.read_bytes())

        if st.button("🚀 INICIAR PROCESO COMPLETO", type="primary"):
            with st.status("Ejecutando motores de IA...", expanded=True) as status:
                
                # FASE 1: CLASIFICACIÓN
                st.write("⏳ Fase 1: Clasificando envíos...")
                cmd_gpt = [sys.executable, str(SCRIPT_REPARTO), "--csv", "llegadas.csv", "--reglas", "Reglas_hospitales.xlsx", "--out", "salida.xlsx"]
                rc1, out1, err1 = run_process(cmd_gpt, cwd=workdir)
                
                if rc1 != 0:
                    status.update(label="❌ Error en Clasificación", state="error")
                    st.error(err1)
                else:
                    # FASE 2: OPTIMIZACIÓN DINÁMICA
                    st.write("⏳ Fase 2: Calculando rango y optimizando rutas...")
                    time.sleep(1) # Pausa técnica para asegurar escritura del archivo
                    
                    try:
                        # Leemos el archivo generado para saber cuántas hojas tiene exactamente
                        temp_xl = pd.ExcelFile(workdir / "salida.xlsx")
                        hojas_reales = temp_xl.sheet_names
                        total_hojas = len(hojas_reales)
                        # El rango DEBE ser exacto para evitar el ValueError
                        rango_seguro = f"0-{total_hojas-1}"
                        st.write(f"📦 Detectadas {total_hojas} hojas. Procesando rango {rango_seguro}...")
                        
                        cmd_gemini = [
                            sys.executable, str(SCRIPT_GEMINI), 
                            "--seleccion", rango_seguro, 
                            "--in", "salida.xlsx", 
                            "--out", "PLAN.xlsx"
                        ]
                        rc2, out2, err2 = run_process(cmd_gemini, cwd=workdir)
                        
                        if rc2 != 0:
                            status.update(label="❌ Error en Optimización", state="error")
                            st.error(err2)
                        else:
                            status.update(label="✅ Todo completado con éxito", state="complete")
                            st.success(f"Plan generado con {total_hojas} rutas optimizadas.")
                    except Exception as e:
                        status.update(label="❌ Error al leer salida.xlsx", state="error")
                        st.error(f"No se pudo calcular el rango: {e}")

    # Descargas
    s_path, p_path = workdir / "salida.xlsx", workdir / "PLAN.xlsx"
    if s_path.exists() or p_path.exists():
        st.markdown("### 📥 Descargas")
        c1, c2 = st.columns(2)
        with c1:
            if s_path.exists(): st.download_button("💾 DESCARGAR SALIDA.XLSX", s_path.read_bytes(), "salida.xlsx", use_container_width=True)
        with c2:
            if p_path.exists(): st.download_button("💾 DESCARGAR PLAN.XLSX", p_path.read_bytes(), "PLAN.xlsx", use_container_width=True)

# -------------------------
# 2) GOOGLE MAPS
# -------------------------
elif opcion == "2. Google Maps (Rutas Móvil)":
    st.subheader("📍 Navegación (Origen: Vall d'Uxo)")
    
    f_user = st.file_uploader("Subir PLAN.xlsx (Opcional)", type=["xlsx"])
    path_plan = save_upload(f_user, workdir / "temp_plan.xlsx") if f_user else (workdir / "PLAN.xlsx" if (workdir / "PLAN.xlsx").exists() else None)

    if path_plan:
        try:
            xl = pd.ExcelFile(path_plan)
            # Filtramos hojas que no son de reparto
            ignorar = ["METADATOS", "LOG", "INSTRUCCIONES", "RESUMEN_GENERAL", "RESUMEN"]
            hojas = [h for h in xl.sheet_names if h.upper() not in ignorar]
            
            if hojas:
                sel = st.selectbox(f"Selecciona Ruta ({len(hojas)} totales):", hojas)
                df = pd.read_excel(path_plan, sheet_name=sel)
                
                # Buscador flexible de columnas
                c_dir = next((c for c in df.columns if "DIR" in str(c).upper()), None)
                c_pob = next((c for c in df.columns if "POB" in str(c).upper() or "LOC" in str(c).upper()), "")

                if c_dir:
                    # ORIGEN FIJO: Vall d'Uxo
                    origen_url = urllib.parse.quote("Vall d'Uxo, Castellon")
                    
                    direcciones = []
                    for _, fila in df.iterrows():
                        addr = f"{fila[c_dir]}, {fila[c_pob]}".strip(", ")
                        if len(addr) > 5: direcciones.append(urllib.parse.quote(addr))
                    
                    st.info(f"🚩 Ruta: {sel} | Paradas: {len(direcciones)}")
                    
                    # Tramos de 9 paradas
                    for i in range(0, len(direcciones), 9):
                        t = direcciones[i:i+9]
                        destino = t[-1]
                        waypoints = t[:-1]
                        
                        # URL Profesional: Origin -> Waypoints -> Destination
                        link = f"https://www.google.com/maps/dir/?api=1&origin={origen_url}&destination={destino}"
                        if waypoints:
                            link += f"&waypoints={'|'.join(waypoints)}"
                        
                        st.link_button(f"🚗 Iniciar Tramo {i+1} a {i+len(t)}", link, use_container_width=True)
                else:
                    st.error("No se encuentra la columna de dirección.")
            else:
                st.warning("No hay rutas detectadas en este archivo.")
        except Exception as e:
            st.error(f"Error: {e}")
