import streamlit as st
import pandas as pd
import os
import io
from datetime import datetime

# --- CONFIGURACIÓN ---
st.set_page_config(page_title="ZAAL IA - Logística", layout="wide", page_icon="🚚")
st.title("🚀 ZAAL IA: Portal de Reparto Automatizado")

# --- FUNCIÓN DE CLASIFICACIÓN (Tu lógica de local) ---
def procesar_clasificacion(df_llegadas, df_hosp, df_fed):
    # Unir reglas
    df_reglas = pd.concat([df_hosp, df_fed], ignore_index=True)
    
    # Limpieza de patrones (Evita el error de 'Series' y 'strip')
    df_reglas['Patrón_dirección'] = df_reglas['Patrón_dirección'].astype(str).str.strip().str.upper()
    df_reglas = df_reglas[df_reglas['Patrón_dirección'] != 'NAN']
    
    # Truco 1-3: Ordenar por longitud para que el más específico gane
    df_reglas['len'] = df_reglas['Patrón_dirección'].str.len()
    df_reglas = df_reglas.sort_values(by='len', ascending=False).drop(columns=['len'])
    
    # Identificar columna de dirección en el CSV
    cols_csv = df_llegadas.columns
    col_dir = next((c for c in cols_csv if 'DIR' in c.upper() or 'ENTREGA' in c.upper()), cols_csv[0])
    
    # Columnas de Kilos y Expediciones para los resúmenes
    col_exp = next((c for c in cols_csv if 'EXP' in c.upper()), None)
    col_kil = next((c for c in cols_csv if 'KILO' in c.upper() or 'KG' in c.upper()), None)

    df_llegadas['Ruta_Asignada'] = "RESTO"
    df_llegadas['Bloque'] = "RESTO"
    
    # Patrones de Hospitales para el resumen
    list_hosp = set(df_hosp['Patrón_dirección'].astype(str).str.strip().str.upper())

    # Bucle de clasificación
    for idx, fila in df_llegadas.iterrows():
        dir_texto = str(fila[col_dir]).upper()
        for _, regla in df_reglas.iterrows():
            patron = regla['Patrón_dirección']
            if patron in dir_texto:
                df_llegadas.at[idx, 'Ruta_Asignada'] = regla['Ruta']
                df_llegadas.at[idx, 'Bloque'] = "HOSPITALES" if patron in list_hosp else "FEDERACION"
                break
    
    return df_llegadas, col_dir, col_exp, col_kil

# --- INTERFAZ STREAMLIT ---
st.header("1️⃣ Fase de Clasificación")
f_csv = st.file_uploader("Sube el CSV de LLEGADAS", type=["csv"])

if st.button("EJECUTAR CLASIFICACIÓN"):
    if f_csv:
        try:
            # Leer archivos
            df_llegadas = pd.read_csv(f_csv, sep=None, engine='python', encoding='latin-1')
            df_hosp = pd.read_excel("Reglas_hospitales.xlsx", sheet_name='REGLAS_HOSPITALES')
            df_fed = pd.read_excel("Reglas_hospitales.xlsx", sheet_name='REGLAS_FEDERACION')
            
            with st.spinner("Clasificando..."):
                df_res, col_dir, col_exp, col_kil = procesar_clasificacion(df_llegadas, df_hosp, df_fed)
                
                # Generar el Excel idéntico al local
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    # METADATOS
                    pd.DataFrame({'Clave': ['CSV', 'Fecha'], 'Valor': [f_csv.name, datetime.now().strftime("%Y-%m-%d %H:%M")]}).to_excel(writer, sheet_name='METADATOS', index=False)
                    
                    # RESUMEN_GENERAL
                    res_gen = df_res.groupby('Bloque').size().reset_index(name='Paradas')
                    res_gen.to_excel(writer, sheet_name='RESUMEN_GENERAL', index=False)
                    
                    # HOJAS POR RUTA (ZREP)
                    rutas = sorted(df_res['Ruta_Asignada'].unique())
                    for r in rutas:
                        df_ruta = df_res[df_res['Ruta_Asignada'] == r].copy()
                        # Nombre de pestaña como en local
                        nombre_hoja = f"ZREP_{str(r)[:25]}".replace('/', '-')
                        df_ruta.to_excel(writer, sheet_name=nombre_hoja[:31], index=False)
                
                st.success("✅ ¡Hecho! Ya puedes descargar el archivo procesado.")
                st.download_button(
                    label="💾 DESCARGAR SALIDA.XLSX",
                    data=output.getvalue(),
                    file_name="salida.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        except Exception as e:
            st.error(f"Error técnico: {e}")
