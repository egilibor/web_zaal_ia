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
st.set_page_config(page_title="ZAAL IA - Logística", layout="wide", page_icon="🚚")
st.title("🚀 ZAAL IA: Portal de Reparto Automatizado")

# --- PATHS ---
REPO_DIR = Path(__file__).resolve().parent
SCRIPT_REPARTO = REPO_DIR / "reparto_gpt.py"
SCRIPT_GEMINI = REPO_DIR / "reparto_gemini.py"
REGLAS_REPO = REPO_DIR / "Reglas_hospitales.xlsx"

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

# -------------------------
# 1) ASIGNACIÓN DE REPARTO
# -------------------------
opcion = st.sidebar.selectbox("Operación:", ["1. Asignación de Reparto", "2. Google Maps (Rutas Móvil)"])

if opcion == "1. Asignación de Reparto":
    st.subheader("Optimización de Macro-Ruta (CP) y Micro-Ruta (Callejero)")
    csv_file = st.file_uploader("Sube el CSV de llegadas", type=["csv"])

    if csv_file:
        save_upload(csv_file, workdir / "llegadas.csv")
        if REGLAS_REPO.exists():
            (workdir / "Reglas_hospitales.xlsx").write_bytes(REGLAS_REPO.read_bytes())

        if st.button("🚀 GENERAR PLAN OPTIMIZADO", type="primary"):
            with st.status("Ejecutando motores de IA...", expanded=True) as status:
                
                # FASE 1: CLASIFICACIÓN
                st.write("⏳ Fase 1: Clasificando envíos...")
                cmd_gpt = [sys.executable, str(SCRIPT_REPARTO), "--csv", "llegadas.csv", "--reglas", "Reglas_hospitales.xlsx", "--out", "salida.xlsx"]
                rc1, out1, err1 = run_process(cmd_gpt, cwd=workdir)
                
                if rc1 == 0:
                    # FASE 2: OPTIMIZACIÓN GEOGRÁFICA
                    st.write("⏳ Fase 2: Aplicando inteligencia de ruta (Traveling Salesman)...")
                    try:
                        xl = pd.ExcelFile(workdir / "salida.xlsx")
                        hojas_validas = [h for h in xl.sheet_names if not any(x in h.upper() for x in ["METADATOS", "RESUMEN"])]
                        rango = f"0-{len(hojas_validas)-1}"
                        
                        # IMPORTANTE: Aquí el script de Gemini debe recibir la instrucción de NO usar orden alfabético.
                        # Asumimos que el script de Gemini ya tiene el prompt de "repartidor local".
                        cmd_gemini = [sys.executable, str(SCRIPT_GEMINI), "--seleccion", rango, "--in", "salida.xlsx", "--out", "PLAN.xlsx"]
                        rc2, out2, err2 = run_process(cmd_gemini, cwd=workdir)
                        
                        # Auto-corrección de rango si falla
                        if rc2 != 0 and "Rango válido" in err2:
                            match = re.search(r"Rango válido: 0\.\.(\d+)", err2)
                            if match:
                                cmd_gemini[3] = f"0-{match.group(1)}"
                                rc2, out2, err2 = run_process(cmd_gemini, cwd=workdir)

                        if rc2 == 0:
                            status.update(label="✅ Plan generado con éxito", state="complete")
                            st.success("Rutas optimizadas. Se ha priorizado la cercanía geográfica por CP.")
                        else:
                            st.error(f"Fallo en optimización: {err2}")
                    except Exception as e:
                        st.error(f"Error de proceso: {e}")

    if (workdir / "PLAN.xlsx").exists():
        st.download_button("💾 DESCARGAR PLAN OPTIMIZADO", (workdir / "PLAN.xlsx").read_bytes(), "PLAN.xlsx")

# -------------------------
# 2) GOOGLE MAPS (ORDEN GEOGRÁFICO)
# -------------------------
elif opcion == "2. Google Maps (Rutas Móvil)":
    st.subheader("📍 Navegación Geográfica (Sin Abecedario)")
    
    path_plan = workdir / "PLAN.xlsx" if (workdir / "PLAN.xlsx").exists() else None
    
    if path_plan:
        xl = pd.ExcelFile(path_plan)
        hojas = [h for h in xl.sheet_names if not any(x in h.upper() for x in ["METADATOS", "RESUMEN", "LOG"])]
        
        sel = st.selectbox("Selecciona Ruta:", hojas)
        df = pd.read_excel(path_plan, sheet_name=sel)
        
        # BUSCADOR DE COLUMNAS
        c_dir = next((c for c in df.columns if "DIR" in str(c).upper()), None)
        c_pob = next((c for c in df.columns if "POB" in str(c).upper() or "LOC" in str(c).upper()), "")
        c_cp = next((c for c in df.columns if "CP" in str(c).upper() or "POSTAL" in str(c).upper()), None)

        if c_dir:
            # NO ORDENAMOS AQUÍ. Respetamos el orden que nos ha dado el PLAN.xlsx (Gemini)
            st.write(f"📂 Mostrando paradas en el orden optimizado por la IA...")

            direcciones = []
            for _, fila in df.iterrows():
                addr = f"{fila[c_dir]}, {fila[c_pob]}".strip(", ")
                direcciones.append(urllib.parse.quote(addr))

            st.info(f"🚩 Ruta: {sel} | {len(direcciones)} paradas.")

            # ORIGEN FIJO VALL D'UXO
            origen_fijo = urllib.parse.quote("Vall d'Uxo, Castellon")

            # Tramos de 9 paradas
            for i in range(0, len(direcciones), 9):
                t = direcciones[i:i+9]
                
                # Solo el primer tramo sale de Vall d'Uxo
                if i == 0:
                    url = f"https://www.google.com/maps/dir/?api=1&origin={origen_fijo}&destination={t[-1]}"
                else:
                    url = f"https://www.google.com/maps/dir/?api=1&destination={t[-1]}"
                
                if len(t) > 1:
                    url += f"&waypoints={'|'.join(t[:-1])}"
                
                st.link_button(f"🚗 Abrir Tramo {i+1}-{i+len(t)} (Siguiente parada más cercana)", url, use_container_width=True)
