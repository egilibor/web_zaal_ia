import sys
import uuid
import shutil
import tempfile
import subprocess
import urllib.parse
import time
import re
from pathlib import Path

import streamlit as st
import pandas as pd

# --- CONFIGURACIÓN ---
st.set_page_config(page_title="ZAAL IA - Gestión de Reparto", layout="wide", page_icon="🚚")
st.title("🚀 ZAAL IA: Portal de Reparto Automatizado")

# --- PATHS ---
REPO_DIR = Path(__file__).resolve().parent
SCRIPT_REPARTO = REPO_DIR / "reparto_gpt.py"
SCRIPT_GEMINI = REPO_DIR / "reparto_gemini.py"
REGLAS_REPO = REPO_DIR / "Reglas_hospitales.xlsx"

# -------------------------
# Utilidades
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

workdir = ensure_workdir()

with st.sidebar:
    st.header("⚙️ Control")
    if st.button("🔄 Reiniciar Aplicación"):
        shutil.rmtree(workdir, ignore_errors=True)
        for key in list(st.session_state.keys()): del st.session_state[key]
        st.rerun()

opcion = st.selectbox("Operación:", ["1. Asignación de Reparto", "2. Google Maps (Rutas Móvil)"])
st.divider()

# -------------------------
# 1) ASIGNACIÓN DE REPARTO
# -------------------------
if opcion == "1. Asignación de Reparto":
    st.subheader("Clasificación y Optimización Total")
    csv_file = st.file_uploader("Sube el CSV de llegadas", type=["csv"])

    if csv_file:
        save_upload(csv_file, workdir / "llegadas.csv")
        if REGLAS_REPO.exists():
            (workdir / "Reglas_hospitales.xlsx").write_bytes(REGLAS_REPO.read_bytes())

        if st.button("🚀 GENERAR PLAN COMPLETO (TODAS LAS RUTAS)", type="primary"):
            with st.status("Procesando...", expanded=True) as status:
                
                # FASE 1: CLASIFICACIÓN
                st.write("⏳ Fase 1: Clasificando envíos...")
                cmd_gpt = [sys.executable, str(SCRIPT_REPARTO), "--csv", "llegadas.csv", "--reglas", "Reglas_hospitales.xlsx", "--out", "salida.xlsx"]
                rc1, out1, err1 = run_process(cmd_gpt, cwd=workdir)
                
                if rc1 == 0:
                    st.write("⏳ Fase 2: Sincronizando hojas para optimización...")
                    time.sleep(1) 
                    
                    try:
                        xl = pd.ExcelFile(workdir / "salida.xlsx")
                        # Gemini ignora estas tres hojas internamente
                        hojas_excluidas = ["METADATOS", "RESUMEN_GENERAL", "RESUMEN"]
                        hojas_validas = [h for h in xl.sheet_names if h.upper() not in hojas_excluidas]
                        
                        # Definimos el rango exacto 0 a N-1
                        rango_final = f"0-{len(hojas_validas)-1}"
                        st.write(f"📦 Detectadas {len(hojas_validas)} rutas. Procesando rango {rango_final}...")
                        
                        cmd_gemini = [sys.executable, str(SCRIPT_GEMINI), "--seleccion", rango_final, "--in", "salida.xlsx", "--out", "PLAN.xlsx"]
                        rc2, out2, err2 = run_process(cmd_gemini, cwd=workdir)
                        
                        # SISTEMA DE RESCATE AUTOMÁTICO SI EL RANGO FALLA
                        if rc2 != 0 and "Rango válido" in err2:
                            match = re.search(r"Rango válido: 0\.\.(\d+)", err2)
                            if match:
                                actual_max = match.group(1)
                                st.warning(f"Ajustando rango a 0-{actual_max}...")
                                cmd_gemini[2] = f"0-{actual_max}"
                                rc2, out2, err2 = run_process(cmd_gemini, cwd=workdir)

                        if rc2 == 0:
                            status.update(label="✅ Proceso completado", state="complete")
                            st.success(f"Plan generado. Revisa Onda-Alcora en el archivo.")
                        else:
                            status.update(label="❌ Error en Optimización", state="error")
                            st.error(err2)
                    except Exception as e:
                        st.error(f"Error técnico: {e}")
                else:
                    status.update(label="❌ Error en Clasificación", state="error")
                    st.error(err1)

    # Descargas
    s_path, p_path = workdir / "salida.xlsx", workdir / "PLAN.xlsx"
    if s_path.exists() or p_path.exists():
        st.markdown("### 📥 Descargas")
        c1, c2 = st.columns(2)
        if s_path.exists(): c1.download_button("💾 DESCARGAR SALIDA.XLSX", s_path.read_bytes(), "salida.xlsx", use_container_width=True)
        if p_path.exists(): c2.download_button("💾 DESCARGAR PLAN.XLSX", p_path.read_bytes(), "PLAN.xlsx", use_container_width=True)

# -------------------------
# 2) GOOGLE MAPS
# -------------------------
elif opcion == "2. Google Maps (Rutas Móvil)":
    st.subheader("📍 Navegación (Origen: Vall d'Uxo)")
    
    f_user = st.file_uploader("Subir PLAN.xlsx para procesar", type=["xlsx"])
    path_plan = save_upload(f_user, workdir / "temp.xlsx") if f_user else (workdir / "PLAN.xlsx" if (workdir / "PLAN.xlsx").exists() else None)

    if path_plan:
        try:
            xl = pd.ExcelFile(path_plan)
            # Mostramos TODAS las hojas excepto las de sistema
            ignorar = ["METADATOS", "LOG", "INSTRUCCIONES", "RESUMEN_GENERAL", "RESUMEN"]
            hojas = [h for h in xl.sheet_names if h.upper() not in ignorar]
            
            if hojas:
                sel = st.selectbox(f"Selecciona Ruta ({len(hojas)} encontradas):", hojas)
                df = pd.read_excel(path_plan, sheet_name=sel)
                
                # Buscador flexible de columnas
                c_dir = next((c for c in df.columns if "DIR" in str(c).upper()), None)
                c_pob = next((c for c in df.columns if "POB" in str(c).upper() or "LOC" in str(c).upper()), "")

                if c_dir:
                    # ORIGEN FIJO: Vall d'Uxo
                    origen_fijo = "Vall d'Uxo, Castellon"
                    origen_encoded = urllib.parse.quote(origen_fijo)
                    
                    direcciones = [urllib.parse.quote(f"{fila[c_dir]}, {fila[c_pob]}".strip(", ")) for _, fila in df.iterrows() if len(str(fila[c_dir])) > 5]
                    
                    st.info(f"🚩 Ruta: {sel} | Paradas: {len(direcciones)}")
                    
                    # Tramos de 9 paradas
                    for i in range(0, len(direcciones), 9):
                        t = direcciones[i:i+9]
                        destino = t[-1]
                        waypoints = t[:-1]
                        
                        # URL oficial con origen Vall d'Uxo
                        url = f"https://www.google.com/maps/dir/?api=1&origin={origen_encoded}&destination={destino}"
                        if waypoints:
                            url += f"&waypoints={'|'.join(waypoints)}"
                        
                        st.link_button(f"🚗 Abrir Tramo {i+1} a {i+len(t)}", url, use_container_width=True)
                else:
                    st.error("No se encontró la columna de dirección.")
            else:
                st.warning("No hay rutas válidas.")
        except Exception as e:
            st.error(f"Error: {e}")
