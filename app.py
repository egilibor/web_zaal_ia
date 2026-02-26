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

def run_process(cmd: list[str], cwd: Path):
    try:
        p = subprocess.run(cmd, cwd=str(cwd), capture_output=True, text=True, timeout=600)
        return p.returncode, p.stdout, p.stderr
    except Exception as e:
        return 1, "", str(e)

# -------------------------
# ESTADO
# -------------------------
workdir = ensure_workdir()

with st.sidebar:
    st.header("⚙️ Panel de Control")
    if st.button("🗑️ Borrar Todo y Reiniciar"):
        shutil.rmtree(workdir, ignore_errors=True)
        for key in list(st.session_state.keys()): del st.session_state[key]
        st.rerun()
    st.divider()
    st.info(f"ID Sesión: {st.session_state.run_id}")

# -------------------------
# MENÚ
# -------------------------
opcion = st.selectbox("Seleccione operación:", ["1. Asignación de Reparto", "2. Google Maps (Rutas Móvil)"])
st.divider()

# -------------------------
# 1) ASIGNACIÓN DE REPARTO
# -------------------------
if opcion == "1. Asignación de Reparto":
    st.subheader("Generar Clasificación y Plan de Carga")
    csv_file = st.file_uploader("Sube el CSV de llegadas", type=["csv"])

    if csv_file:
        save_upload(csv_file, workdir / "llegadas.csv")
        if REGLAS_REPO.exists():
            (workdir / "Reglas_hospitales.xlsx").write_bytes(REGLAS_REPO.read_bytes())

        if st.button("🚀 INICIAR PROCESAMIENTO", type="primary"):
            with st.status("Procesando datos...", expanded=True) as status:
                # FASE 1
                st.write("⏳ Clasificando envíos...")
                cmd_gpt = [sys.executable, str(SCRIPT_REPARTO), "--csv", "llegadas.csv", "--reglas", "Reglas_hospitales.xlsx", "--out", "salida.xlsx"]
                rc1, out1, err1 = run_process(cmd_gpt, cwd=workdir)
                
                if rc1 != 0:
                    st.error("Error en Clasificación"); st.code(err1)
                    status.update(label="❌ Error en Fase 1", state="error")
                else:
                    # FASE 2
                    st.write("⏳ Optimizando rutas con Gemini...")
                    cmd_gemini = [sys.executable, str(SCRIPT_GEMINI), "--seleccion", "1-9", "--in", "salida.xlsx", "--out", "PLAN.xlsx"]
                    rc2, out2, err2 = run_process(cmd_gemini, cwd=workdir)
                    
                    if rc2 != 0:
                        st.error("Error en Optimización"); st.code(err2)
                        status.update(label="❌ Error en Fase 2", state="error")
                    else:
                        status.update(label="✅ Proceso completado", state="complete")
                        st.success("Archivos generados correctamente.")

    # Descargas
    salida_p = workdir / "salida.xlsx"
    plan_p = workdir / "PLAN.xlsx"
    if salida_p.exists() or plan_p.exists():
        st.markdown("### 📥 Descargas")
        c1, c2 = st.columns(2)
        with c1:
            if salida_p.exists(): st.download_button("💾 SALIDA.XLSX", salida_p.read_bytes(), "salida.xlsx", use_container_width=True)
        with c2:
            if plan_p.exists(): st.download_button("💾 PLAN.XLSX", plan_p.read_bytes(), "PLAN.xlsx", use_container_width=True)

# -------------------------
# 2) GOOGLE MAPS
# -------------------------
elif opcion == "2. Google Maps (Rutas Móvil)":
    st.subheader("📍 Navegación por Tramos")
    
    f_user = st.file_uploader("Subir PLAN.xlsx optimizado (opcional)", type=["xlsx"])
    
    path_plan = None
    if f_user:
        path_plan = save_upload(f_user, workdir / "temp_plan.xlsx")
    elif (workdir / "PLAN.xlsx").exists():
        path_plan = workdir / "PLAN.xlsx"
        st.success("Usando archivo de la sesión actual.")

    if path_plan:
        try:
            # Leemos el archivo SIN filtros agresivos para ver qué hay dentro
            xl = pd.ExcelFile(path_plan)
            todas_las_hojas = xl.sheet_names
            
            # Filtro mínimo: solo quitamos lo que SEGURO no es una ruta
            excluir_basico = ["METADATOS", "LOG", "INSTRUCCIONES"]
            hojas_finales = [h for h in todas_las_hojas if h.upper() not in excluir_basico]
            
            st.write(f"📂 **Hojas detectadas:** {len(hojas_finales)} de {len(todas_las_hojas)}")
            
            if hojas_finales:
                sel = st.selectbox("Selecciona la Ruta a visualizar:", hojas_finales)
                
                # Cargar la hoja seleccionada
                df = pd.read_excel(path_plan, sheet_name=sel)
                
                # Identificar columnas de dirección (buscamos cualquier cosa que contenga "DIR")
                col_dir = next((c for c in df.columns if "DIR" in str(c).upper()), None)
                col_pob = next((c for c in df.columns if "POB" in str(c).upper() or "LOC" in str(c).upper()), "")

                if col_dir:
                    direcciones = []
                    for _, fila in df.iterrows():
                        d = str(fila[col_dir]).strip()
                        p = str(fila[col_pob]).strip() if col_pob in df.columns else ""
                        full_addr = f"{d}, {p}".strip(", ")
                        if len(full_addr) > 5:
                            direcciones.append(urllib.parse.quote(full_addr))
                    
                    st.info(f"📍 **{sel}**: {len(direcciones)} paradas encontradas.")
                    
                    # Generar botones por tramos de 9
                    for i in range(0, len(direcciones), 9):
                        t = direcciones[i:i+9]
                        destino = t[-1]
                        waypoints = t[:-1]
                        
                        link = f"https://www.google.com/maps/dir/?api=1&destination={destino}"
                        if waypoints:
                            link += f"&waypoints={'|'.join(waypoints)}"
                        
                        st.link_button(f"🚗 Iniciar Tramo {i+1} - {i+len(t)}", link, use_container_width=True)
                else:
                    st.error(f"No encuentro columna de dirección en '{sel}'. Columnas disponibles: {list(df.columns)}")
            else:
                st.warning("No se encontraron hojas de ruta válidas en el archivo.")
                with st.expander("Ver todas las hojas encontradas (Debug)"):
                    st.write(todas_las_hojas)
                    
        except Exception as e:
            st.error(f"Error técnico: {e}")
