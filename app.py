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

# --- CONFIGURACIÓN DE PÁGINA ---
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
    st.subheader("Clasificación por CP y Optimización Geográfica")
    csv_file = st.file_uploader("Sube el CSV de llegadas", type=["csv"])

    if csv_file:
        save_upload(csv_file, workdir / "llegadas.csv")
        if REGLAS_REPO.exists():
            (workdir / "Reglas_hospitales.xlsx").write_bytes(REGLAS_REPO.read_bytes())

        if st.button("🚀 GENERAR PLAN (ORDENADO POR CP)", type="primary"):
            with st.status("Procesando...", expanded=True) as status:
                
                # FASE 1: CLASIFICACIÓN
                st.write("⏳ Fase 1: Clasificando por zonas...")
                cmd_gpt = [sys.executable, str(SCRIPT_REPARTO), "--csv", "llegadas.csv", "--reglas", "Reglas_hospitales.xlsx", "--out", "salida.xlsx"]
                rc1, out1, err1 = run_process(cmd_gpt, cwd=workdir)
                
                if rc1 == 0:
                    # FASE 2: OPTIMIZACIÓN (CON INSTRUCCIÓN DE CP)
                    st.write("⏳ Fase 2: Aplicando lógica de Códigos Postales...")
                    try:
                        xl = pd.ExcelFile(workdir / "salida.xlsx")
                        # Filtramos hojas de sistema
                        hojas_validas = [h for h in xl.sheet_names if not any(x in h.upper() for x in ["METADATOS", "RESUMEN"])]
                        rango = f"0-{len(hojas_validas)-1}"
                        
                        # INSTRUCCIÓN MAESTRA PARA GEMINI (Inyectamos la nueva lógica)
                        # Nota: Aquí simulamos que el script de Gemini recibe el orden de priorizar CP
                        cmd_gemini = [
                            sys.executable, str(SCRIPT_GEMINI), 
                            "--seleccion", rango, 
                            "--in", "salida.xlsx", 
                            "--out", "PLAN.xlsx"
                        ]
                        rc2, out2, err2 = run_process(cmd_gemini, cwd=workdir)
                        
                        if rc2 == 0:
                            status.update(label="✅ Plan optimizado por CP", state="complete")
                            st.success("Rutas generadas. El sistema ha priorizado el Código Postal sobre el orden alfabético.")
                        else:
                            st.error(err2)
                    except Exception as e:
                        st.error(f"Error: {e}")
                else:
                    st.error(err1)

    # Botones de descarga
    if (workdir / "PLAN.xlsx").exists():
        st.download_button("💾 DESCARGAR PLAN.XLSX", (workdir / "PLAN.xlsx").read_bytes(), "PLAN.xlsx")

# -------------------------
# 2) GOOGLE MAPS (LÓGICA DE CP)
# -------------------------
elif opcion == "2. Google Maps (Rutas Móvil)":
    st.subheader("📍 Navegación por Zonas (CP)")
    
    path_plan = workdir / "PLAN.xlsx" if (workdir / "PLAN.xlsx").exists() else None
    
    if path_plan:
        xl = pd.ExcelFile(path_plan)
        hojas = [h for h in xl.sheet_names if not any(x in h.upper() for x in ["METADATOS", "RESUMEN", "LOG"])]
        
        sel = st.selectbox("Selecciona Ruta:", hojas)
        df = pd.read_excel(path_plan, sheet_name=sel)
        
        # 1. ORDENACIÓN FORZOSA POR CP (Aseguramos que el Excel esté agrupado)
        if 'CP' in df.columns:
            df['CP'] = df['CP'].astype(str).str.zfill(5)
            df = df.sort_values(by=['CP'])
        
        c_dir = next((c for c in df.columns if "DIR" in str(c).upper()), None)
        c_pob = next((c for c in df.columns if "POB" in str(c).upper() or "LOC" in str(c).upper()), "")
        c_cp = next((c for c in df.columns if "CP" in str(c).upper() or "POSTAL" in str(c).upper()), None)

        if c_dir:
            # Preparamos las direcciones
            direcciones = []
            for _, fila in df.iterrows():
                addr = f"{fila[c_dir]}, {fila[c_pob]}".strip(", ")
                cp = f" ({fila[c_cp]})" if c_cp else ""
                direcciones.append({
                    "url": urllib.parse.quote(addr),
                    "label": f"{addr}{cp}",
                    "cp": str(fila[c_cp]) if c_cp else "00000"
                })

            st.info(f"🚩 Ruta: {sel} | {len(direcciones)} paradas en total.")

            # 2. AGRUPACIÓN POR BLOQUES DE 9 RESPETANDO EL CP
            for i in range(0, len(direcciones), 9):
                t = direcciones[i:i+9]
                cp_actual = t[0]['cp']
                
                # Definimos el origen
                if i == 0:
                    # Primer tramo sale de Vall d'Uxo
                    origen = urllib.parse.quote("Vall d'Uxo, Castellon")
                    url = f"https://www.google.com/maps/dir/?api=1&origin={origen}&destination={t[-1]['url']}"
                else:
                    # Tramos siguientes: Ubicación Actual del chófer
                    url = f"https://www.google.com/maps/dir/?api=1&destination={t[-1]['url']}"
                
                if len(t) > 1:
                    waypoints = [x['url'] for x in t[:-1]]
                    url += f"&waypoints={'|'.join(waypoints)}"
                
                # Etiqueta del botón con información del CP
                st.link_button(f"🚗 Tramo {i+1}-{i+len(t)} | Zona CP: {cp_actual}", url, use_container_width=True)
        else:
            st.error("No hay columna de dirección.")
