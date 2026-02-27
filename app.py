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
st.set_page_config(page_title="ZAAL IA - Logística Interior", layout="wide", page_icon="🚚")
st.title("🚀 ZAAL IA: Optimización de Rutas Interior")

# --- LÓGICA DE MACRO-RUTA (El "Hilo" de los pueblos) ---
# Definimos el orden lógico saliendo de Vall d'Uxo hacia el interior.
# Cuanto menor el número, antes se visita.
ORDEN_PUEBLOS = {
    "VALL D'UXO": 0,
    "ALFONDEGUILLA": 1,
    "ARTANA": 2,
    "ESLIDA": 3,
    "BETXI": 4,
    "ONDA": 5,
    "RIBESALBES": 6,
    "FANZARA": 7,
    "ALCORA": 8,
    "L'ALCORA": 8,
    "FIGUEROLES": 9,
    "LUCENA": 10,
    "VISTABELLA": 11,
    "TOGA": 12,
    "CIRAT": 13,
    "MONTANEJOS": 14
}

def obtener_prioridad(poblacion):
    pob = str(poblacion).upper().strip()
    # Si el pueblo está en nuestra lista, devolvemos su orden, si no, lo mandamos al final (99)
    return ORDEN_PUEBLOS.get(pob, 99)

# --- UTILIDADES DE ARCHIVO ---
workdir = Path(tempfile.gettempdir()) / "reparto_zaal"
workdir.mkdir(exist_ok=True)

def run_process(cmd: list[str], cwd: Path):
    try:
        p = subprocess.run(cmd, cwd=str(cwd), capture_output=True, text=True, timeout=600)
        return p.returncode, p.stdout, p.stderr
    except Exception as e:
        return 1, "", str(e)

# -------------------------
# INTERFAZ
# -------------------------
opcion = st.sidebar.selectbox("Menú:", ["1. Generar Plan", "2. Google Maps (Rutas)"])

if opcion == "1. Generar Plan":
    st.subheader("Fase de Optimización Macro y Micro")
    csv_file = st.file_uploader("Sube el CSV de llegadas", type=["csv"])
    
    if csv_file and st.button("🚀 OPTIMIZAR TODO"):
        with st.status("Calculando ruta óptima...", expanded=True) as status:
            # (Aquí irían tus scripts reparto_gpt y reparto_gemini)
            # Simulamos que se genera salida.xlsx
            st.write("⏳ Aplicando lógica de proximidad a los pueblos del interior...")
            # ... (Llamadas a subprocess) ...
            st.success("Plan generado. Los pueblos ahora siguen la carretera, no el abecedario.")

elif opcion == "2. Google Maps (Rutas)":
    st.subheader("📍 Navegación por Sentido de Marcha")
    f_user = st.file_uploader("Subir PLAN.xlsx", type=["xlsx"])
    
    if f_user:
        xl = pd.ExcelFile(f_user)
        hojas = [h for h in xl.sheet_names if "RESUMEN" not in h.upper()]
        sel = st.selectbox("Selecciona la ruta:", hojas)
        
        df = pd.read_excel(f_user, sheet_name=sel)
        
        # --- EL MOTOR DE ORDENACIÓN REAL ---
        if 'POBLACION' in df.columns:
            # 1. Creamos una columna temporal para el peso geográfico
            df['peso_geo'] = df['POBLACION'].apply(obtener_prioridad)
            
            # 2. ORDENAMOS: Primero por el peso del pueblo, luego por CP, luego por dirección
            df = df.sort_values(by=['peso_geo', 'CP', 'DIRECCION'], ascending=[True, True, True])
            
            st.success(f"Ruta reordenada siguiendo el eje: Vall d'Uxo -> Onda -> Alcora...")
        
        # Generar botones de Google Maps
        c_dir = next((c for c in df.columns if "DIR" in str(c).upper()), None)
        c_pob = next((c for c in df.columns if "POB" in str(c).upper()), "")
        
        if c_dir:
            direcciones = [urllib.parse.quote(f"{f[c_dir]}, {f[c_pob]}") for _, f in df.iterrows()]
            
            for i in range(0, len(direcciones), 9):
                t = direcciones[i:i+9]
                # Primer tramo sale de Vall d'Uxo
                origen = urllib.parse.quote("Vall d'Uxo, Castellon") if i == 0 else ""
                
                if origen:
                    url = f"https://www.google.com/maps/dir/?api=1&origin={origen}&destination={t[-1]}&waypoints={'|'.join(t[:-1])}"
                else:
                    # Tramos siguientes: desde ubicación actual
                    url = f"https://www.google.com/maps/dir/?api=1&destination={t[-1]}&waypoints={'|'.join(t[:-1])}"
                
                st.link_button(f"🚗 TRAMO {i//9 + 1}: {df.iloc[i][c_pob]} ({len(t)} paradas)", url, use_container_width=True)
