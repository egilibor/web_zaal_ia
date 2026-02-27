import sys
import os
import subprocess
import urllib.parse
import streamlit as st
import pandas as pd

# --- CONFIGURACIÓN ---
st.set_page_config(page_title="ZAAL IA - Gestión Directa", layout="wide")
st.title("🚚 ZAAL IA: Sistema de Reparto")

# --- PASO 1: CARGA ---
csv_file = st.sidebar.file_uploader("1. Sube el CSV de llegadas", type=["csv"])

if csv_file:
    # Guardamos el archivo en la carpeta local (donde están tus scripts)
    with open("llegadas.csv", "wb") as f:
        f.write(csv_file.getbuffer())
    
    # --- PASO 2: EJECUCIÓN FASE 1 ---
    if st.sidebar.button("2. Generar salida.xlsx"):
        # Eliminamos salida.xlsx previo para no leer datos viejos
        if os.path.exists("salida.xlsx"):
            os.remove("salida.xlsx")
            
        with st.spinner("Ejecutando clasificación..."):
            # Ejecución directa en la misma carpeta
            res = subprocess.run([sys.executable, "reparto_gpt.py", "--csv", "llegadas.csv", "--out", "salida.xlsx"], 
                                 capture_output=True, text=True)
            
            if res.returncode == 0:
                st.success("✅ Clasificación terminada en 1 segundo.")
            else:
                st.error(f"Error en script: {res.stderr}")

# --- PASO 3: MOSTRAR RUTAS (LEER salida.xlsx) ---
if os.path.exists("salida.xlsx"):
    st.divider()
    st.subheader("📋 Selecciona la ruta de salida.xlsx")
    
    try:
        # Cargamos el Excel físicamente
        xl = pd.ExcelFile("salida.xlsx")
        hojas = [h for h in xl.sheet_names if not any(x in h.upper() for x in ["METADATOS", "RESUMEN"])]
        
        col1, col2 = st.columns([3, 1])
        with col1:
            ruta_sel = st.selectbox("Rutas disponibles:", hojas)
        
        # --- PASO 4: OPTIMIZAR SELECCIÓN ---
        with col2:
            st.write("") # Espaciado
            if st.button("🚀 Optimizar esta ruta", type="primary"):
                idx = xl.sheet_names.index(ruta_sel)
                with st.spinner(f"Optimizando {ruta_sel}..."):
                    # Ejecutamos Gemini solo para esa hoja
                    subprocess.run([sys.executable, "reparto_gemini.py", "--seleccion", str(idx), "--in", "salida.xlsx", "--out", "PLAN.xlsx"], 
                                   capture_output=True, text=True)
                    st.session_state.listo = True
                    st.session_state.nombre = ruta_sel
                    
    except Exception as e:
        st.error(f"Error leyendo el Excel: {e}")

# --- PASO 5: RESULTADOS (GOOGLE MAPS) ---
if st.session_state.get("listo") and os.path.exists("PLAN.xlsx"):
    st.divider()
    st.subheader(f"📍 Navegación: {st.session_state.nombre}")
    
    df = pd.read_excel("PLAN.xlsx", sheet_name=st.session_state.nombre)
    c_dir = next((c for c in df.columns if "DIR" in str(c).upper()), None)
    c_pob = next((c for c in df.columns if "POB" in str(c).upper()), "")

    if c_dir:
        direcciones = [urllib.parse.quote(f"{f[c_dir]}, {f[c_pob]}") for _, f in df.iterrows()]
        
        for i in range(0, len(direcciones), 9):
            t = direcciones[i:i+9]
            origen = urllib.parse.quote("Vall d'Uxo, Castellon") if i == 0 else ""
            prefix = "0" if origen else "3"
            url = f"http://googleusercontent.com/maps.google.com/{prefix}{origen}&destination={t[-1]}&waypoints={'|'.join(t[:-1])}"
            st.link_button(f"🚗 TRAMO {i//9 + 1}: {df.iloc[i][c_dir]}", url, use_container_width=True)
