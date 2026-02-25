import streamlit as st
import pandas as pd
import os
import io
import glob
from datetime import datetime

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="ZAAL IA - Logística", layout="wide", page_icon="🚚")

st.title("🚀 ZAAL IA: Portal de Reparto Automatizado")

# Función que reemplaza a 'reparto_gpt.py' (Tu lógica original de clasificación)
def ejecutar_clasificacion_interna(df_llegadas, df_hosp_reg, df_fed_reg):
    # Unir reglas y aplicar TRUCO 1-3
    df_reglas = pd.concat([df_hosp_reg, df_fed_reg], ignore_index=True)
    df_reglas['len'] = df_reglas['Patrón_dirección'].astype(str).str.len()
    df_reglas = df_reglas.sort_values(by='len', ascending=False).drop(columns=['len'])
    
    # Identificar columnas del CSV
    cols = df_llegadas.columns
    col_dir = next((c for c in cols if 'DIR' in c.upper()), cols[0])
    
    df_llegadas['Ruta_Asignada'] = "RESTO"
    df_llegadas['Bloque'] = "RESTO"
    
    patrones_hosp = set(df_hosp_reg['Patrón_dirección'].astype(str).str.upper().strip())

    for idx, fila in df_llegadas.iterrows():
        direccion = str(fila[col_dir]).upper()
        for _, regla in df_reglas.iterrows():
            patron = str(regla['Patrón_dirección']).upper().strip()
            if patron and patron != "NAN" and patron in direccion:
                df_llegadas.at[idx, 'Ruta_Asignada'] = regla['Ruta']
                df_llegadas.at[idx, 'Bloque'] = "HOSPITALES" if patron in patrones_hosp else "FEDERACION"
                break
    return df_llegadas

# --- PASO 1: CLASIFICACIÓN ---
st.header("1️⃣ Fase de Clasificación")
f_csv = st.file_uploader("Sube el CSV de LLEGADAS", type=["csv"])

if st.button("EJECUTAR CLASIFICACIÓN"):
    if f_csv:
        try:
            # Leer datos
            df_llegadas = pd.read_csv(f_csv, sep=None, engine='python', encoding='latin-1')
            df_hosp = pd.read_excel("Reglas_hospitales.xlsx", sheet_name='REGLAS_HOSPITALES')
            df_fed = pd.read_excel("Reglas_hospitales.xlsx", sheet_name='REGLAS_FEDERACION')
            
            with st.spinner("Clasificando rutas..."):
                df_procesado = ejecutar_clasificacion_interna(df_llegadas, df_hosp, df_fed)
                
                # Crear el archivo salida.xlsx en memoria
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    # Hoja Metadatos
                    pd.DataFrame({'Clave': ['CSV', 'Fecha'], 'Valor': [f_csv.name, datetime.now()]}).to_excel(writer, sheet_name='METADATOS')
                    # Resumen
                    df_procesado.groupby('Ruta_Asignada').size().to_excel(writer, sheet_name='RESUMEN_UNICO')
                    # Hojas por ruta
                    for ruta in df_procesado['Ruta_Asignada'].unique():
                        df_temp = df_procesado[df_procesado['Ruta_Asignada'] == ruta]
                        df_temp.to_excel(writer, sheet_name=str(ruta)[:30].replace('/','-'), index=False)
                
                st.session_state['archivo_salida'] = output.getvalue()
                st.success("✅ Clasificación completada con éxito.")
                
                st.download_button(
                    label="💾 DESCARGAR 'SALIDA.XLSX'",
                    data=st.session_state['archivo_salida'],
                    file_name="salida.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        except Exception as e:
            st.error(f"Error: {e}")

# --- PASO 2: OPTIMIZACIÓN ---
st.header("2️⃣ Fase de Optimización (Plan de Carga)")
st.info("Nota: La optimización avanzada requiere que el script 'reparto_gemini.py' esté integrado de la misma forma.")

if st.button("🚀 GENERAR PLAN FINAL"):
    st.warning("Para ejecutar la Fase 2 en la nube, necesitamos integrar el código de 'reparto_gemini.py' aquí dentro, igual que hemos hecho con la clasificación.")
