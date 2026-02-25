
import streamlit as st
import pandas as pd
from pathlib import Path
import io

# Configuración de la página
st.set_page_config(page_title="ZAAL Logística IA", layout="wide")

st.title("🚚 ZAAL - Sistema de Clasificación Logística")
st.markdown("Sube tu archivo `.csv` de llegadas para organizar las 13 rutas automáticamente.")

# --- LÓGICA DE PROCESAMIENTO ---
def clasificar_envios(df, df_reglas):
    # Aseguramos que las columnas necesarias existan
    df.columns = df.columns.str.strip()
    
    # Unimos con las reglas de hospitales (basado en la dirección o nombre)
    # Aquí simulamos la lógica que tenías en reparto_gpt
    df['Ruta'] = "RUTA NO ASIGNADA"
    
    # Ejemplo de lógica simplificada para que no falle:
    for index, row in df.iterrows():
        destino = str(row['Dir. entrega']).upper()
        for _, regla in df_reglas.iterrows():
            if str(regla['Patron']).upper() in destino:
                df.at[index, 'Ruta'] = regla['Ruta']
                break
    
    return df

# --- INTERFAZ DE USUARIO ---
archivo_subido = st.file_uploader("Elige el archivo llegadas.csv", type="csv")

if archivo_subido is not None:
    try:
        # Leer el CSV subido
        df_llegadas = pd.read_csv(archivo_subido, sep=None, engine='python', encoding='latin-1')
        st.success("Archivo subido correctamente")
        
        # Intentar leer las reglas desde GitHub
        try:
            df_reglas = pd.read_excel("Reglas_hospitales.xlsx")
        except Exception as e:
            st.error(f"No se encontró el archivo de reglas en GitHub: {e}")
            df_reglas = None

        if df_reglas is not None:
            if st.button("🚀 Procesar Clasificación"):
                # Procesar
                resultado = clasificar_envios(df_llegadas, df_reglas)
                
                # Mostrar resultados
                st.subheader("Vista Previa del Reparto")
                st.dataframe(resultado.head(20))
                
                # Botón de Descarga Excel
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    resultado.to_excel(writer, index=False, sheet_name='Plan de Reparto')
                
                st.download_button(
                    label="📥 Descargar Plan Final (Excel)",
                    data=output.getvalue(),
                    file_name="Plan_Logistica_ZAAL.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
    except Exception as e:
        st.error(f"Error al leer el CSV: {e}")

else:
    st.info("Esperando archivo... Por favor, sube el .csv para empezar.")
