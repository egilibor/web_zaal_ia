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

# --- PATHS EN REPO ---
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

workdir = ensure_workdir()

# -------------------------
# MENÚ PRINCIPAL
# -------------------------
opcion = st.selectbox("Operación:", ["1. Asignación de Reparto", "2. Google Maps (Rutas Móvil)"])
st.divider()

# -------------------------
# 1) ASIGNACIÓN DE REPARTO
# -------------------------
if opcion == "1. Asignación de Reparto":
    st.subheader("Clasificación y Optimización (Todas las Rutas)")
    csv_file = st.file_uploader("Sube el CSV de llegadas", type=["csv"])

    if csv_file:
        save_upload(csv_file, workdir / "llegadas.csv")
        if REGLAS_REPO.exists():
            (workdir / "Reglas_hospitales.xlsx").write_bytes(REGLAS_REPO.read_bytes())

        if st.button("🚀 INICIAR PROCESO COMPLETO", type="primary"):
            with st.status("Ejecutando motores de IA...", expanded=True) as status:
                
                # FASE 1: CLASIFICACIÓN
                st.write("⏳ Fase 1: Clasificando envíos (salida.xlsx)...")
                cmd_gpt = [sys.executable, str(SCRIPT_REPARTO), "--csv", "llegadas.csv", "--reglas", "Reglas_hospitales.xlsx", "--out", "salida.xlsx"]
                rc1, out1, err1 = run_process(cmd_gpt, cwd=workdir)
                
                if rc1 != 0:
                    status.update(label="❌ Error en Fase 1", state="error")
                    st.error(err1)
                else:
                    # --- SOLUCIÓN: CÁLCULO DINÁMICO DE HOJAS ---
                    st.write("⏳ Fase 2: Detectando todas las rutas para optimizar...")
                    try:
                        temp_xl = pd.ExcelFile(workdir / "salida.xlsx")
                        # Gemini ignora hojas técnicas. Contamos solo las de reparto.
                        ignorar = ["METADATOS", "RESUMEN", "LOG"]
                        hojas_reparto = [h for h in temp_xl.sheet_names if not any(x in h.upper() for x in ignorar)]
                        
                        num_validas = len(hojas_reparto)
                        # Rango dinámico: desde la 0 hasta la última (N-1)
                        rango_dinamico = f"0-{num_validas-1}"
                        
                        st.write(f"📦 Detectadas {num_validas} rutas (incluyendo Onda-Alcora).")
                        
                        cmd_gemini = [
                            sys.executable, str(SCRIPT_GEMINI), 
                            "--seleccion", rango_dinamico, 
                            "--in", "salida.xlsx", 
                            "--out", "PLAN.xlsx"
                        ]
                        rc2, out2, err2 = run_process(cmd_gemini, cwd=workdir)
                        
                        # Si Gemini protesta por el índice, capturamos el error y ajustamos
                        if rc2 != 0 and "Rango válido" in err2:
                            match = re.search(r"Rango válido: 0\.\.(\d+)", err2)
                            if match:
                                actual_max = match.group(1)
                                cmd_gemini[2] = f"0-{actual_max}"
                                rc2, out2, err2 = run_process(cmd_gemini, cwd=workdir)

                        if rc2 == 0:
                            status.update(label="✅ Proceso completado", state="complete")
                            st.success(f"Plan generado con {num_validas} rutas optimizadas.")
                        else:
                            status.update(label="❌ Error en Fase 2", state="error")
                            st.error(err2)
                    except Exception as e:
                        st.error(f"Error técnico al sincronizar: {e}")

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
    f_user = st.file_uploader("Subir PLAN.xlsx para Maps", type=["xlsx"])
    p_path = save_upload(f_user, workdir / "temp.xlsx") if f_user else (workdir / "PLAN.xlsx" if (workdir / "PLAN.xlsx").exists() else None)

    if p_path:
        try:
            xl = pd.ExcelFile(p_path)
            hojas = [h for h in xl.sheet_names if not any(x in h.upper() for x in ["METADATOS", "RESUMEN", "LOG"])]
            
            if hojas:
                sel = st.selectbox("Selecciona Ruta:", hojas)
                df = pd.read_excel(p_path, sheet_name=sel)
                
                c_dir = next((c for c in df.columns if "DIR" in str(c).upper()), None)
                c_pob = next((c for c in df.columns if "POB" in str(c).upper() or "LOC" in str(c).upper()), "")

                if c_dir:
                    # ORIGEN FIJO
                    origen = urllib.parse.quote("Vall d'Uxo, Castellon")
                    direcciones = [urllib.parse.quote(f"{f[c_dir]}, {f[c_pob]}".strip(", ")) for _, f in df.iterrows() if len(str(f[c_dir])) > 5]
                    
                    st.info(f"🚩 Ruta: {sel} | Paradas: {len(direcciones)}")
                    for i in range(0, len(direcciones), 9):
                        t = direcciones[i:i+9]
                        # URL oficial con origen Vall d'Uxo
                        url = f"https://www.google.com/maps/dir/?api=1&origin={origen}&destination={t[-1]}"
                        if t[:-1]: url += f"&waypoints={'|'.join(t[:-1])}"
                        st.link_button(f"🚗 Abrir Tramo {i+1} a {i+len(t)}", url, use_container_width=True)
                else:
                    st.error("No se encontró la columna de dirección.")
            else:
                st.warning("No hay rutas en el archivo.")
        except Exception as e:
            st.error(f"Error: {e}")
