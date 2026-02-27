import sys
import uuid
import shutil
import tempfile
import subprocess
import urllib.parse
from pathlib import Path

import streamlit as st
import pandas as pd

# --- CONFIGURACIÓN ---
st.set_page_config(page_title="ZAAL IA - Logística", layout="wide", page_icon="🚚")
st.title("🚀 ZAAL IA: Generador de Rutas Optimizado")

# --- LÓGICA DE MACRO-RUTA (Interior) ---
ORDEN_PUEBLOS = {
    "VALL D'UXO": 0, "ALFONDEGUILLA": 1, "ARTANA": 2, "ESLIDA": 3,
    "BETXI": 4, "ONDA": 5, "RIBESALBES": 6, "FANZARA": 7,
    "ALCORA": 8, "L'ALCORA": 8, "FIGUEROLES": 9, "LUCENA": 10,
    "VISTABELLA": 11, "TOGA": 12, "CIRAT": 13, "MONTANEJOS": 14
}

def obtener_prioridad(poblacion):
    pob = str(poblacion).upper().strip()
    return ORDEN_PUEBLOS.get(pob, 99)

# --- GESTIÓN DE DIRECTORIO ---
def ensure_workdir():
    if "workdir" not in st.session_state:
        st.session_state.workdir = Path(tempfile.mkdtemp(prefix="zaal_"))
    return Path(st.session_state.workdir)

workdir = ensure_workdir()

def run_process(cmd: list[str], cwd: Path):
    try:
        p = subprocess.run(cmd, cwd=str(cwd), capture_output=True, text=True, timeout=600)
        return p.returncode, p.stdout, p.stderr
    except Exception as e:
        return 1, "", str(e)

# --- INTERFAZ ---
with st.sidebar:
    st.header("⚙️ Herramientas")
    if st.button("🗑️ Limpiar Sesión"):
        shutil.rmtree(workdir, ignore_errors=True)
        for key in list(st.session_state.keys()): del st.session_state[key]
        st.rerun()

opcion = st.selectbox("Elige Operación:", ["1. Generar Plan de Reparto", "2. Google Maps (Rutas Móvil)"])
st.divider()

# -------------------------
# 1) GENERAR PLAN
# -------------------------
if opcion == "1. Generar Plan de Reparto":
    st.subheader("Fase 1 y 2: Clasificación y Optimización Geográfica")
    csv_file = st.file_uploader("Sube el CSV de llegadas", type=["csv"])

    if csv_file:
        # Guardar archivo subido
        input_path = workdir / "llegadas.csv"
        input_path.write_bytes(csv_file.getbuffer())
        
        # Copiar reglas si existen en el repo
        reglas_repo = Path("Reglas_hospitales.xlsx")
        if reglas_repo.exists():
            shutil.copy(reglas_repo, workdir / "Reglas_hospitales.xlsx")

        if st.button("🚀 INICIAR PROCESO COMPLETO", type="primary"):
            with st.status("Ejecutando motores de IA...", expanded=True) as status:
                
                # FASE 1: REPARTO_GPT
                st.write("⏳ Fase 1: Clasificando envíos...")
                cmd1 = [sys.executable, "reparto_gpt.py", "--csv", "llegadas.csv", "--reglas", "Reglas_hospitales.xlsx", "--out", "salida.xlsx"]
                rc1, out1, err1 = run_process(cmd1, cwd=workdir)
                
                if rc1 == 0:
                    # FASE 2: REPARTO_GEMINI
                    st.write("⏳ Fase 2: Optimizando macro-ruta (Pueblos)...")
                    # Detectar rango de hojas automáticamente
                    xl_temp = pd.ExcelFile(workdir / "salida.xlsx")
                    hojas = [h for h in xl_temp.sheet_names if not any(x in h.upper() for x in ["METADATOS", "RESUMEN"])]
                    rango = f"0-{len(hojas)-1}"
                    
                    cmd2 = [sys.executable, "reparto_gemini.py", "--seleccion", rango, "--in", "salida.xlsx", "--out", "PLAN.xlsx"]
                    rc2, out2, err2 = run_process(cmd2, cwd=workdir)
                    
                    if rc2 == 0:
                        status.update(label="✅ ¡Todo listo!", state="complete")
                        st.success("Plan generado correctamente siguiendo el eje Vall d'Uxo -> Onda -> Alcora.")
                    else:
                        st.error(f"Error en Fase 2: {err2}")
                else:
                    st.error(f"Error en Fase 1: {err1}")

    # --- AQUÍ ESTÁN LOS ENLACES (Botones de descarga) ---
    s_xlsx = workdir / "salida.xlsx"
    p_xlsx = workdir / "PLAN.xlsx"
    
    if s_xlsx.exists() or p_xlsx.exists():
        st.markdown("### 📥 Descargar Archivos Generados")
        col1, col2 = st.columns(2)
        with col1:
            if s_xlsx.exists():
                st.download_button("💾 DESCARGAR SALIDA.XLSX", s_xlsx.read_bytes(), "salida.xlsx", use_container_width=True)
        with col2:
            if p_xlsx.exists():
                st.download_button("💾 DESCARGAR PLAN.XLSX", p_xlsx.read_bytes(), "PLAN.xlsx", use_container_width=True)

# -------------------------
# 2) GOOGLE MAPS
# -------------------------
elif opcion == "2. Google Maps (Rutas Móvil)":
    st.subheader("📍 Navegación Inteligente")
    f_user = st.file_uploader("Subir PLAN.xlsx para navegar", type=["xlsx"])
    
    path_final = None
    if f_user:
        path_final = workdir / "temp_nav.xlsx"
        path_final.write_bytes(f_user.getbuffer())
    elif (workdir / "PLAN.xlsx").exists():
        path_final = workdir / "PLAN.xlsx"

    if path_final:
        xl = pd.ExcelFile(path_final)
        hojas = [h for h in xl.sheet_names if not any(x in h.upper() for x in ["METADATOS", "RESUMEN", "LOG"])]
        sel = st.selectbox("Selecciona la ruta para el chófer:", hojas)
        
        df = pd.read_excel(path_final, sheet_name=sel)
        
        # APLICAR ORDEN GEOGRÁFICO ANTES DE GENERAR MAPAS
        if 'POBLACION' in df.columns:
            df['peso_geo'] = df['POBLACION'].apply(obtener_prioridad)
            df = df.sort_values(by=['peso_geo', 'CP', 'DIRECCION'], ascending=[True, True, True])

        c_dir = next((c for c in df.columns if "DIR" in str(c).upper()), None)
        c_pob = next((c for c in df.columns if "POB" in str(c).upper()), "")
        
        if c_dir:
            direcciones = [urllib.parse.quote(f"{f[c_dir]}, {f[c_pob]}") for _, f in df.iterrows() if len(str(f[c_dir])) > 3]
            st.info(f"🚩 Ruta: {sel} | Paradas: {len(direcciones)}")
            
            for i in range(0, len(direcciones), 9):
                t = direcciones[i:i+9]
                origen = urllib.parse.quote("Vall d'Uxo, Castellon") if i == 0 else ""
                
                if origen:
                    url = f"https://www.google.com/maps/dir/?api=1&origin={origen}&destination={t[-1]}&waypoints={'|'.join(t[:-1])}"
                else:
                    url = f"https://www.google.com/maps/dir/?api=1&destination={t[-1]}&waypoints={'|'.join(t[:-1])}"
                
                st.link_button(f"🚗 TRAMO {i//9 + 1} ({len(t)} paradas)", url, use_container_width=True)
