import streamlit as st
import pandas as pd
import os
import io

# --- TRUCO MAESTRO PARA OPENPYXL ---
try:
    import openpyxl
except ImportError:
    import subprocess
    import sys
    subprocess.check_call([sys.executable, "-m", "pip", "install", "openpyxl"])
    import openpyxl

st.set_page_config(page_title="ZAAL Logística", layout="wide")
st.title("🚚 ZAAL - Clasificador de Rutas")

# Intentamos localizar el Excel
# Si no lo encuentra, nos avisará con la ruta exacta que está viendo
ruta_reglas = "Reglas_hospitales.xlsx"
if not os.path.exists(ruta_reglas):
    # Si está dentro de una carpeta, intentamos buscarlo ahí
    ruta_alternativa = os.path.join("web_zaal_ia", "Reglas_hospitales.xlsx")
    if os.path.exists(ruta_alternativa):
        ruta_reglas = ruta_alternativa

if not os.path.exists(ruta_reglas):
    st.error(f"❌ No encuentro el archivo de reglas. Asegúrate de que 'Reglas_hospitales.xlsx' esté en GitHub.")
    st.stop()

archivo_subido = st.file_uploader("Sube tu CSV de llegadas", type=["csv"])

if archivo_subido:
    try:
        df_llegadas = pd.read_csv(archivo_subido, sep=None, engine='python', encoding='latin-1')
        df_reglas = pd.read_excel(ruta_reglas, engine='openpyxl')
        
        st.success("✅ Datos cargados. Pulsa el botón para procesar.")

        if st.button("🚀 Procesar Clasificación"):
            df_llegadas['Ruta'] = "SIN ASIGNAR"
            # Lógica simple de búsqueda
            for idx, fila in df_llegadas.iterrows():
                direc = str(fila.get('Dir. entrega', '')).upper()
                for _, regla in df_reglas.iterrows():
                    patron = str(regla.get('Patron', '')).upper()
                    if patron in direc and patron != "":
                        df_llegadas.at[idx, 'Ruta'] = regla.get('Ruta', 'RUTA X')
                        break
            
            st.dataframe(df_llegadas.head(10))
            
            # Preparar descarga
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_llegadas.to_excel(writer, index=False)
            
            st.download_button("📥 Descargar Plan Final", output.getvalue(), "Plan_ZAAL.xlsx")
            
    except Exception as e:
        st.error(f"Error al procesar: {e}")
