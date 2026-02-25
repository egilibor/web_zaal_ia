import streamlit as st
import pandas as pd
import os
import io

st.set_page_config(page_title="ZAAL Logística", layout="wide")
st.title("🚚 ZAAL - Clasificador de Rutas")

# 1. Localizar reglas
ruta_reglas = "Reglas_hospitales.xlsx"
if not os.path.exists(ruta_reglas):
    st.error("❌ No se encuentra Reglas_hospitales.xlsx")
    st.stop()

archivo_subido = st.file_uploader("Sube el archivo llegadas.csv", type=["csv"])

if archivo_subido:
    try:
        # Leer archivos
        df_llegadas = pd.read_csv(archivo_subido, sep=None, engine='python', encoding='latin-1')
        df_reglas = pd.read_excel(ruta_reglas, engine='openpyxl')
        
        # LIMPIEZA TOTAL: Quitamos espacios y pasamos a mayúsculas
        df_llegadas.columns = [c.strip() for c in df_llegadas.columns]
        df_reglas.columns = [c.strip() for c in df_reglas.columns]
        
        # Identificar la columna de dirección (por si no se llama exactamente 'Dir. entrega')
        col_dir = 'Dir. entrega' if 'Dir. entrega' in df_llegadas.columns else df_llegadas.columns[0]
        
        if st.button("🚀 Procesar Clasificación"):
            df_llegadas['Ruta'] = "RUTA NO ENCONTRADA" # Valor por defecto más claro
            
            # Lógica de búsqueda mejorada
            for idx, fila in df_llegadas.iterrows():
                texto_entrega = str(fila[col_dir]).upper().strip()
                
                for _, regla in df_reglas.iterrows():
                    patron = str(regla['Patron']).upper().strip()
                    
                    if patron and patron in texto_entrega:
                        df_llegadas.at[idx, 'Ruta'] = regla['Ruta']
                        break # Si encuentra una, para de buscar para esa fila

            # Mostrar cuántas se han clasificado
            encontradas = len(df_llegadas[df_llegadas['Ruta'] != "RUTA NO ENCONTRADA"])
            st.success(f"✅ Proceso terminado. Se han clasificado {encontradas} de {len(df_llegadas)} entregas.")
            
            st.dataframe(df_llegadas)

            # Preparar descarga
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_llegadas.to_excel(writer, index=False)
            
            st.download_button("📥 Descargar Resultado Final", output.getvalue(), "Plan_Rutas_ZAAL.xlsx")

    except Exception as e:
        st.error(f"Error: {e}")
