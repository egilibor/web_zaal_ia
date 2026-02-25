import streamlit as st
import pandas as pd
import os
import subprocess
import time
import io
import glob

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="ZAAL IA - Logística", layout="wide", page_icon="🚚")

# Forzar directorio local
os.chdir(os.path.dirname(os.path.abspath(__file__)))

st.title("🚀 ZAAL IA: Portal de Reparto Automatizado")

# --- PASO 1: CLASIFICACIÓN ---
st.header("1️⃣ Fase de Clasificación")
f_csv = st.file_uploader("Sube el CSV de LLEGADAS", type=["csv"])

if st.button("EJECUTAR CLASIFICACIÓN"):
    if f_csv:
        if os.path.exists("salida.xlsx"):
            try: os.remove("salida.xlsx")
            except: st.error("⚠️ Cierra 'salida.xlsx' antes de continuar.")
        
        with open("llegadas.csv", "wb") as f: f.write(f_csv.getbuffer())
        
        with st.spinner("Clasificando rutas..."):
            cmd = ["python", "reparto_gpt.py", "--csv", "llegadas.csv", "--reglas", "Reglas_hospitales.xlsx", "--out", "salida.xlsx"]
            subprocess.run(cmd, input="\n", capture_output=True, text=True)
            
            if os.path.exists("salida.xlsx"):
                st.success("✅ Clasificación completada.")
                st.rerun()

# --- VISTA PREVIA Y DESCARGA INTERMEDIA ---
if os.path.exists("salida.xlsx"):
    st.markdown("---")
    with st.expander("👀 VER Y GUARDAR ARCHIVO INTERMEDIO", expanded=True):
        with open("salida.xlsx", "rb") as f:
            data_int = f.read()
            # AQUÍ ESTÁ EL CAMBIO: Usamos "stretch" para que el CMD no proteste
            st.dataframe(pd.read_excel(io.BytesIO(data_int)).head(10), width="stretch")
            st.download_button("💾 GUARDAR 'SALIDA.XLSX' EN MI PC", data_int, "salida.xlsx")

    st.markdown("---")

    # --- PASO 2: OPTIMIZACIÓN AUTOMÁTICA ---
    st.header("2️⃣ Fase de Optimización (Plan de Carga)")
    
    if st.button("🚀 GENERAR PLAN FINAL (TODAS LAS RUTAS)"):
        # Limpieza de planes antiguos
        for f in glob.glob("PLAN_*.xlsx"): 
            try: os.remove(f)
            except: pass
            
        try:
            xl = pd.ExcelFile("salida.xlsx")
            zonas = [h for h in xl.sheet_names if not any(x in h.upper() for x in ["RESUMEN", "LOG"])]
            rango_auto = f"0-{len(zonas)-1}"
            
            with st.spinner(f"Optimizando {len(zonas)} rutas..."):
                subprocess.run(["python", "reparto_gemini.py"], input=f"{rango_auto}\n", text=True)
                
                planes = glob.glob("PLAN_*.xlsx")
                if planes:
                    plan_nombre = planes[0]
                    st.success(f"🎯 Plan Final: {plan_nombre}")
                    
                    with open(plan_nombre, "rb") as f_plan:
                        data_final = f_plan.read()
                        # Volvemos a usar "stretch" aquí también
                        st.dataframe(pd.read_excel(io.BytesIO(data_final)).head(10), width="stretch")
                        st.download_button("💾 GUARDAR PLAN FINAL EN MI PC", data_final, plan_nombre)
        except Exception as e:
            st.error(f"Error al leer las zonas: {e}")