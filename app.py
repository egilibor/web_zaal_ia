import streamlit as st
import pandas as pd
import os
import io

st.set_page_config(page_title="ZAAL Logística", layout="wide")
st.title("🚚 ZAAL - Clasificador de Rutas")

# 1. Localizar el archivo de reglas
ruta_reglas = "Reglas_hospitales.xlsx"
if not os.path.exists(ruta_reglas):
    st.error(f"❌ No se encuentra el archivo {ruta_reglas}")
    st.stop()

archivo_subido = st.file_uploader("Sube el archivo llegadas.csv", type=["csv"])

if archivo_subido:
    try:
        # Leer archivos
        df_llegadas = pd.read_csv(archivo_subido, sep=None, engine='python', encoding='latin-1')
        df_reglas = pd.read_excel(ruta_reglas, engine='openpyxl')
        
        # Limpieza de nombres de columnas (quitar espacios invisibles)
        df_llegadas.columns = [c.strip() for c in df_llegadas.columns]
        df_reglas.columns = [c.strip() for c in df_reglas.columns]
        
        # Identificar las columnas clave
        # En el CSV de llegadas buscamos algo que se parezca a 'Dir. entrega'
        col_dir_llegadas = next((c for c in df_llegadas.columns if 'DIR' in c.upper() or 'ENTREGA' in c.upper()), df_llegadas.columns[0])
        
        # En el Excel de reglas usamos el nombre que me has dado
        col_patron_reglas = 'Patrón_dirección'
        col_ruta_reglas = 'Ruta' # Asegúrate de que en el Excel la columna de ruta se llame así

        if st.button("🚀 Procesar Clasificación"):
            if col_patron_reglas not in df_reglas.columns:
                st.error(f"❌ No encuentro la columna '{col_patron_reglas}' en el Excel. Las columnas que veo son: {list(df_reglas.columns)}")
                st.stop()

            df_llegadas['Ruta_Asignada'] = "SIN RUTA"
            
            # Lógica de comparación
            for idx, fila in df_llegadas.iterrows():
                direccion_cliente = str(fila[col_dir_llegadas]).upper().strip()
                
                for _, regla in df_reglas.iterrows():
                    palabra_clave = str(regla[col_patron_reglas]).upper().strip()
                    
                    if palabra_clave and palabra_clave in direccion_cliente:
                        df_llegadas.at[idx, 'Ruta_Asignada'] = regla[col_ruta_reglas]
                        break

            # Resultados
            encontrados = len(df_llegadas[df_llegadas['Ruta_Asignada'] != "SIN RUTA"])
            st.success(f"✅ ¡Hecho! Se han clasificado {encontrados} de {len(df_llegadas)} envíos.")
            
            st.dataframe(df_llegadas)

            # Preparar descarga
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_llegadas.to_excel(writer, index=False)
            
            st.download_button(
                label="📥 Descargar Resultado Final (Excel)",
                data=output.getvalue(),
                file_name="Plan_Logistica_ZAAL.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Hubo un problema: {e}")
