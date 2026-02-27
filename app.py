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
st.title("🚀 ZAAL IA: Generador de Rutas (Orden de Carretera)")

# --- LA BIBLIA DEL REPARTIDOR (Ruta Real) ---
# Definimos el orden de paso saliendo de Vall d'Uxo. 
# Si entra un pueblo nuevo, solo hay que añadirlo aquí con su número de orden.
SECUENCIA_RUTA = {
    "VALL D'UXO": 1,
    "ALFONDEGUILLA": 2,
    "ARTANA": 3,
    "ESLIDA": 4,
    "AIN": 5,
    "ALCUDIA DE VEO": 6,
    "BETXI": 7,
    "ONDA": 8,
    "RIBESALBES": 9,
    "FANZARA": 10,
    "ALCORA": 11,
    "L'ALCORA": 11,
    "FIGUEROLES": 12,
    "LUCENA": 13,
    "VISTABELLA": 14
}

def asignar_orden(fila, col_pob):
    pueblo = str(fila[col_pob]).upper().strip()
    # Si el pueblo está en la lista, su peso es el número asignado (1, 2, 3...)
    # Si no está, le damos un 999 para que vaya al final.
    return SECUENCIA_RUTA.get(pueblo, 999)

# --- RUTAS ---
REPO_DIR = Path(__file__).resolve().parent
SCRIPT_GPT = REPO_DIR / "reparto_gpt.py"
SCRIPT_GEMINI = REPO_DIR / "reparto_gemini.py"

def get_workdir():
    if "workdir" not in st.session_state:
        st.session_state.workdir = Path(tempfile.mkdtemp(prefix="zaal_fix_"))
    return Path(st.session_state.workdir)

workdir = get_workdir()

# --- INTERFAZ ---
opcion = st.sidebar.selectbox("Menú:", ["1. Generar Plan", "2. Google Maps (Navegación)"])

if opcion == "1. Generar Plan":
    st.subheader("Fase 1 y 2: Clasificación y Optimización")
    csv_file = st.file_uploader("Sube el CSV de llegadas", type=["csv"])
    
    if csv_file and st.button("🚀 INICIAR PROCESO"):
        input_path = workdir / "llegadas.csv"
        input_path.write_bytes(csv_file.getbuffer())
        
        with st.status("Procesando...", expanded=True) as status:
            st.write("⏳ Ejecutando Clasificación...")
            subprocess.run([sys.executable, str(SCRIPT_GPT), "--csv", "llegadas.csv", "--out", "salida.xlsx"], cwd=workdir)
            
            st.write("⏳ Ejecutando Optimización Geográfica...")
            xl = pd.ExcelFile(workdir / "salida.xlsx")
            rango = f"0-{len(xl.sheet_names)-1}"
            subprocess.run([sys.executable, str(SCRIPT_GEMINI), "--seleccion", rango, "--in", "salida.xlsx", "--out", "PLAN.xlsx"], cwd=workdir)
            
            status.update(label="✅ Plan Generado", state="complete")

    if (workdir / "PLAN.xlsx").exists():
        st.download_button("💾 DESCARGAR PLAN.XLSX", (workdir / "PLAN.xlsx").read_bytes(), "PLAN.xlsx", use_container_width=True)

elif opcion == "2. Google Maps (Navegación)":
    st.subheader("📍 Tramos de Ruta (Orden Geográfico Real)")
    f_user = st.file_uploader("Subir PLAN.xlsx", type=["xlsx"])
    
    path_file = None
    if f_user:
        path_file = workdir / "manual_nav.xlsx"
        path_file.write_bytes(f_user.getbuffer())
    elif (workdir / "PLAN.xlsx").exists():
        path_file = workdir / "PLAN.xlsx"

    if path_file:
        xl = pd.ExcelFile(path_file)
        hojas = [h for h in xl.sheet_names if not any(x in h.upper() for x in ["METADATOS", "RESUMEN"])]
        sel = st.selectbox("Selecciona la ruta:", hojas)
        
        df = pd.read_excel(path_file, sheet_name=sel)
        
        # BUSCAR COLUMNAS
        c_pob = next((c for c in df.columns if "POB" in str(c).upper()), None)
        c_dir = next((c for c in df.columns if "DIR" in str(c).upper()), None)
        c_cp = next((c for c in df.columns if "CP" in str(c).upper()), None)

        if c_pob and c_dir:
            # APLICAMOS EL ORDEN DE CARRETERA
            df['orden_logico'] = df.apply(lambda x: asignar_orden(x, c_pob), axis=1)
            
            # ORDENAMOS: 1º Por el pueblo (carretera), 2º Por CP, 3º Por dirección
            df = df.sort_values(by=['orden_logico', c_cp if c_cp else c_pob, c_dir])
            
            st.success(f"Ruta '{sel}' optimizada por sentido de marcha.")
            
            direcciones = [urllib.parse.quote(f"{f[c_dir]}, {f[c_pob]}") for _, f in df.iterrows()]
            
            

            for i in range(0, len(direcciones), 9):
                t = direcciones[i:i+9]
                origen = urllib.parse.quote("Vall d'Uxo, Castellon") if i == 0 else ""
                
                if origen:
                    url = f"https://www.google.com/maps/dir/?api=1&origin={origen}&destination={t[-1]}&waypoints={'|'.join(t[:-1])}"
                else:
                    url = f"https://www.google.com/maps/dir/?api=1&destination={t[-1]}&waypoints={'|'.join(t[:-1])}"
                
                # Mostramos el pueblo dominante en el tramo para que el chófer sepa dónde va
                pueblo_tramo = df.iloc[i][c_pob]
                st.link_button(f"🚗 TRAMO {i//9 + 1}: {pueblo_tramo} ({len(t)} paradas)", url, use_container_width=True)
