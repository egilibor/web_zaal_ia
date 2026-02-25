import streamlit as st
import pandas as pd
import os
import io

# Configuración de página
st.set_page_config(page_title="ZAAL Logística IA", layout="wide")

st.title("🚚 ZAAL - Clasificador de Rutas Inteligente")

# --- COMPROBACIÓN DEL EXCEL DE REGLAS ---
ruta_reglas = "Reglas_hospitales.xlsx"

if not os.path.exists(ruta_reglas):
    st.error(f"❌ No se encuentra el archivo '{ruta_reglas}' en GitHub. Asegúrate de que el nombre sea exacto.")
    st.stop()

# --- INTERFAZ DE CARGA ---
archivo_subido = st.file_uploader("Sube tu archivo CSV de llegadas", type=["csv"])

if archivo_subido is not None:
    try:
        # 1. Leer el CSV
        df_llegadas = pd.read_csv(archivo_subido, sep=None, engine='python', encoding='latin-1')
        
        # 2. Leer las reglas (Aquí es donde se usa openpyxl)
        df_reglas = pd.read_excel(ruta_reglas, engine='openpyxl')
        
        st.success("✅ Archivos cargados correctamente.")

        if st.button("🚀 Procesar Clasificación"):
            # Limpieza de columnas
            df_llegadas.columns = [str(c).strip() for c in df_llegadas.columns]
            df_reglas.columns = [str(c).strip() for c in df_reglas.columns]
            
            # Lógica de rutas
            df_llegadas['Ruta'] = "SIN ASIGNAR"
            for idx, fila in df_llegadas.iterrows():
                direccion = str(fila.get('Dir. entrega', '')).upper()
                for _, regla in df_reglas.iterrows():
                    patron = str(regla.get('Patron', '')).upper()
                    if patron in direccion and patron != "":
                        df_llegadas.at[idx, 'Ruta'] = regla.get('Ruta', 'RUTA X')
                        break

            st.subheader("Resultado")
            st.dataframe(df_llegadas.head(10))

            # Excel para descarga
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                df_llegadas.to_excel(writer, index=False)
            
            st.download_button(
                label="📥 Descargar Excel Final",
                data=buffer.getvalue(),
                file_name="Plan_Rutas_ZAAL.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Hubo un error: {e}")
