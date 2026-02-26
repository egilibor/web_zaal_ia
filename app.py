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
    st.header("⚙️ Control de Sesión")
    if st.button("🗑️ Reiniciar Todo"):
        shutil.rmtree(workdir, ignore_errors=True)
        for key in list(st.session_state.keys()): del st.session_state[key]
        st.rerun()
    st.info(f"ID: {st.session_state.run_id}")

opcion = st.selectbox("Operación:", ["1. Asignación de Reparto", "2. Google Maps (Rutas Móvil)"])
st.divider()

# -------------------------
# 1) ASIGNACIÓN DE REPARTO
# -------------------------
if opcion == "1. Asignación de Reparto":
    csv_file = st.file_uploader("Sube el CSV de llegadas", type=["csv"])

    if csv_file:
        save_upload(csv_file, workdir / "llegadas.csv")
        if REGLAS_REPO.exists():
            (workdir / "Reglas_hospitales.xlsx").write_bytes(REGLAS_REPO.read_bytes())

        if st.button("🚀 GENERAR PLAN COMPLETO", type="primary"):
            with st.status("Ejecutando motores de IA...", expanded=True) as status:
                # FASE 1
                st.write("⏳ Fase 1: Clasificando envíos...")
                cmd_gpt = [sys.executable, str(SCRIPT_REPARTO), "--csv", "llegadas.csv", "--reglas", "Reglas_hospitales.xlsx", "--out", "salida.xlsx"]
                rc1, out1, err1 = run_process(cmd_gpt, cwd=workdir)
                
                if rc1 == 0:
                    st.write("⏳ Fase 2: Optimizando rutas...")
                    try:
                        xl = pd.ExcelFile(workdir / "salida.xlsx")
                        # Filtro más amplio para no dejar ninguna fuera (incluida la Ruta 9)
                        hojas_a_procesar = [h for h in xl.sheet_names if not any(x in h.upper() for x in ["METADATOS", "RESUMEN_GENERAL", "RESUMEN", "LOG"])]
                        
                        # Probamos con el rango máximo detectado
                        rango = f"0-{len(hojas_a_procesar)-1}"
                        st.write(f"📦 Rutas detectadas: {len(hojas_a_procesar)}. Procesando todas...")
                        
                        cmd_gemini = [sys.executable, str(SCRIPT_GEMINI), "--seleccion", rango, "--in", "salida.xlsx", "--out", "PLAN.xlsx"]
                        rc2, out2, err2 = run_process(cmd_gemini, cwd=workdir)
                        
                        # AUTO-CORRECCIÓN: Si Gemini dice que el rango es menor, lo ajustamos al vuelo
                        if rc2 != 0 and "Rango válido" in err2:
                            match = re.search(r"Rango válido: 0\.\.(\d+)", err2)
                            if match:
                                nuevo_rango = f"0-{match.group(1)}"
                                st.write(f"🔄 Ajustando rango automáticamente a {nuevo_rango}...")
                                cmd_gemini[2] = nuevo_rango
                                rc2, out2, err2 = run_process(cmd_gemini, cwd=workdir)

                        if rc2 == 0:
                            status.update(label="✅ Plan generado correctamente", state="complete")
                            st.success(f"Se han optimizado todas las rutas disponibles.")
                        else:
                            status.update(label="❌ Error en Fase 2", state="error")
                            st.error(err2)
                    except Exception as e:
                        st.error(f"Error de lectura: {e}")
                else:
                    status.update(label="❌ Error en Fase 1", state="error")
                    st.error(err1)

    # Descargas
    s_p, p_p = workdir / "salida.xlsx", workdir / "PLAN.xlsx"
    if s_p.exists() or p_p.exists():
        st.markdown("### 📥 Descargas")
        c1, c2 = st.columns(2)
        if s_p.exists(): c1.download_button("💾 DESCARGAR SALIDA.XLSX", s_p.read_bytes(), "salida.xlsx", use_container_width=True)
        if p_p.exists(): c2.download_button("💾 DESCARGAR PLAN.XLSX", p_p.read_bytes(), "PLAN.xlsx", use_container_width=True)

# -------------------------
# 2) GOOGLE MAPS
# -------------------------
elif opcion == "2. Google Maps (Rutas Móvil)":
    st.subheader("📍 Navegación (Origen: Vall d'Uxo)")
    f_user = st.file_uploader("Subir PLAN.xlsx (opcional)", type=["xlsx"])
    p_path = save_upload(f_user, workdir / "temp_p.xlsx") if f_user else (workdir / "PLAN.xlsx" if (workdir / "PLAN.xlsx").exists() else None)

    if p_path:
        xl = pd.ExcelFile(p_path)
        # Mostrar absolutamente todas las hojas que no sean de sistema
        ignorar = ["METADATOS", "LOG", "INSTRUCCIONES", "RESUMEN"]
        hojas = [h for h in xl.sheet_names if not any(x in h.upper() for x in ignorar)]
        
        if hojas:
            sel = st.selectbox("Selecciona la ruta (Asegúrate de ver la Ruta 9):", hojas)
            df = pd.read_excel(p_path, sheet_name=sel)
            
            c_dir = next((c for c in df.columns if "DIR" in str(c).upper()), None)
            c_pob = next((c for c in df.columns if "POB" in str(c).upper() or "LOC" in str(c).upper()), "")

            if c_dir:
                # ORIGEN FIJO: Vall d'Uxo
                origen_enc = urllib.parse.quote("Vall d'Uxo, Castellon")
                # Limpiamos direcciones vacías
                direcciones = [urllib.parse.quote(f"{f[c_dir]}, {f[c_pob]}".strip(", ")) for _, f in df.iterrows() if len(str(f[c_dir])) > 4]
                
                st.info(f"🚩 Ruta: {sel} | Total paradas: {len(direcciones)}")
                
                # Tramos de 9 paradas
                for i in range(0, len(direcciones), 9):
                    tramo = direcciones[i:i+9]
                    # URL: origin=Vall d'Uxo & destination=Ultima parada del tramo & waypoints=Resto de paradas
                    url_maps = f"https://www.google.com/maps/dir/?api=1&origin={origen_enc}&destination={tramo[-1]}"
                    if len(tramo) > 1:
                        url_maps += f"&waypoints={'|'.join(tramo[:-1])}"
                    
                    st.link_button(f"🚗 Abrir Tramo: Paradas {i+1} a {i+len(tramo)}", url_maps, use_container_width=True)
            else:
                st.error("No se ha encontrado la columna de dirección en esta hoja.")
        else:
            st.warning("No se han detectado rutas en el archivo.")
