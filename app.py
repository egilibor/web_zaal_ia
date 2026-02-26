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
# MENÚ PRINCIPAL
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
            with st.status("Ejecutando motores de IA...", expanded=True) as status:
                
                # FASE 1: CLASIFICACIÓN
                st.write("⏳ Fase 1: Clasificando envíos...")
                cmd_gpt = [sys.executable, str(SCRIPT_REPARTO), "--csv", "llegadas.csv", "--reglas", "Reglas_hospitales.xlsx", "--out", "salida.xlsx"]
                rc1, out1, err1 = run_process(cmd_gpt, cwd=workdir)
                
                if rc1 != 0:
                    st.error("Error en Fase 1"); st.code(err1)
                    status.update(label="❌ Fallo en Fase 1", state="error")
                else:
                    # --- CÁLCULO DINÁMICO DEL RANGO PARA GEMINI ---
                    st.write("⏳ Fase 2: Calculando rutas y optimizando...")
                    try:
                        temp_xl = pd.ExcelFile(workdir / "salida.xlsx")
                        num_total_hojas = len(temp_xl.sheet_names)
                        # Detectamos todas las hojas para optimizarlas todas
                        rango_dinamico = f"0-{num_total_hojas-1}"
                    except:
                        rango_dinamico = "0-50" # Fallback amplio
                    
                    # FASE 2: OPTIMIZACIÓN
                    cmd_gemini = [
                        sys.executable, str(SCRIPT_GEMINI), 
                        "--seleccion", rango_dinamico, 
                        "--in", "salida.xlsx", 
                        "--out", "PLAN.xlsx"
                    ]
                    rc2, out2, err2 = run_process(cmd_gemini, cwd=workdir)
                    
                    if rc2 != 0:
                        st.error("Error en Fase 2"); st.code(err2)
                        status.update(label="❌ Fallo en Fase 2", state="error")
                    else:
                        status.update(label="✅ Plan generado correctamente", state="complete")
                        st.success(f"Optimización finalizada para {num_total_hojas} hojas.")

    # Descargas
    salida_p = workdir / "salida.xlsx"
    plan_p = workdir / "PLAN.xlsx"
    if salida_p.exists() or plan_p.exists():
        st.markdown("### 📥 Descargar Resultados")
        c1, c2 = st.columns(2)
        with c1:
            if salida_p.exists(): st.download_button("💾 DESCARGAR SALIDA.XLSX", salida_p.read_bytes(), "salida.xlsx", use_container_width=True)
        with c2:
            if plan_p.exists(): st.download_button("💾 DESCARGAR PLAN.XLSX", plan_p.read_bytes(), "PLAN.xlsx", use_container_width=True)

# -------------------------
# 2) GOOGLE MAPS
# -------------------------
elif opcion == "2. Google Maps (Rutas Móvil)":
    st.subheader("📍 Enlaces de Navegación (Origen: Vall d'Uxo)")
    
    f_user = st.file_uploader("Subir archivo PLAN.xlsx manualmente", type=["xlsx"])
    
    path_plan = None
    if f_user:
        path_plan = save_upload(f_user, workdir / "temp_plan.xlsx")
    elif (workdir / "PLAN.xlsx").exists():
        path_plan = workdir / "PLAN.xlsx"
        st.info("Utilizando el plan generado en la sesión actual.")

    if path_plan:
        try:
            xl = pd.ExcelFile(path_plan)
            # Filtro para ver todas las hojas de ruta reales
            excluir = ["METADATOS", "LOG", "INSTRUCCIONES", "RESUMEN_GENERAL", "RESUMEN"]
            hojas_validas = [h for h in xl.sheet_names if h.upper() not in excluir]
            
            if hojas_validas:
                sel = st.selectbox(f"Selecciona la Ruta ({len(hojas_validas)} disponibles):", hojas_validas)
                df = pd.read_excel(path_plan, sheet_name=sel)
                
                # Identificar columnas
                c_dir = next((c for c in df.columns if "DIR" in str(c).upper()), None)
                c_pob = next((c for c in df.columns if "POB" in str(c).upper() or "LOC" in str(c).upper()), "")

                if c_dir:
                    # Punto de origen fijo
                    origen_fijo = urllib.parse.quote("Vall d'Uxo, Castellon")
                    
                    direcciones = []
                    for _, fila in df.iterrows():
                        d = str(fila[c_dir]).strip()
                        p = str(fila[c_pob]).strip() if c_pob in df.columns else ""
                        full_addr = f"{d}, {p}".strip(", ")
                        if len(full_addr) > 5:
                            direcciones.append(urllib.parse.quote(full_addr))
                    
                    st.write(f"🗺️ **Ruta:** {sel} | **Paradas:** {len(direcciones)}")
                    
                    # Generar tramos de 9 paradas para Maps
                    # Cada tramo empieza en Vall d'Uxo
                    for i in range(0, len(direcciones), 9):
                        t = direcciones[i:i+9]
                        destino = t[-1]
                        waypoints = t[:-1]
                        
                        # URL de Google Maps con origen en Vall d'Uxo
                        # Estructura: origin=OR&destination=DEST&waypoints=W1|W2...
                        link = f"https://www.google.com/maps/dir/?api=1&origin={origen_fijo}&destination={destino}"
                        if waypoints:
                            link += f"&waypoints={'|'.join(waypoints)}"
                        
                        st.link_button(f"🚗 Iniciar Tramo {i+1} a {i+len(t)}", link, use_container_width=True)
                else:
                    st.error(f"No se detecta columna de dirección en la hoja '{sel}'.")
            else:
                st.warning("No se han detectado hojas de ruta en el archivo.")
        except Exception as e:
            st.error(f"Error al procesar el archivo: {e}")
