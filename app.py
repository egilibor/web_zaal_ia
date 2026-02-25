import streamlit as st
import pandas as pd
import os
import subprocess
import sys
import io

# --- SOLUCIÓN DE LIBRERÍAS ---
# Forzamos la instalación de openpyxl internamente para evitar el error de dependencias
try:
    import openpyxl
except ImportError:
    subprocess.check_call([sys.executable, "-m", "pip", "install", "openpyxl"])

# Configuración de página
st.set_page_config(page_title="ZAAL Logística IA", layout="wide")

st.title("🚚 ZAAL - Clasificador de Rutas Inteligente")
st.info("Sube el archivo 'llegadas.csv' para procesar el reparto diario.")

# --- COMPROBACIÓN DEL EXCEL DE REGLAS ---
# Buscamos el archivo en el repositorio
ruta_reglas = "Reglas_hospitales.xlsx"

if not os.path.exists(ruta_reglas):
    st.error(f"❌ Error crítico: No se encuentra el archivo '{ruta_reglas}' en GitHub.")
    st.stop()

# --- INTERFAZ DE CARGA ---
archivo_subido = st.file_uploader("Arrastra aquí tu archivo CSV", type=["csv"])

if archivo_subido is not None:
    try:
        # 1. Leer el CSV con codificación flexible para evitar errores de caracteres
        df_llegadas = pd.read_csv(archivo_subido, sep=None, engine='python', encoding='latin-1')
        st.success("✅ Archivo de llegadas cargado.")

        # 2. Leer las reglas de Excel
        df_reglas = pd.read_excel(ruta_reglas, engine='openpyxl')
        
        # Limpieza rápida de nombres de columnas
        df_llegadas.columns = [str(c).strip() for c in df_llegadas.columns]
        df_reglas.columns = [str(c).strip() for c in df_reglas.columns]

        if st.button("🚀 Procesar Clasificación"):
            # 3. Lógica de asignación de rutas
            df_llegadas['Ruta'] = "SIN ASIGNAR"
            
            # Buscamos coincidencias en la dirección de entrega
            for idx, fila in df_llegadas.iterrows():
                direccion = str(fila.get('Dir. entrega', '')).upper()
                for _, regla in df_reglas.iterrows():
                    patron = str(regla.get('Patron', '')).upper()
                    if patron in direccion and patron != "":
                        df_llegadas.at[idx, 'Ruta'] = regla.get('Ruta', 'RUTA X')
                        break

            # 4. Mostrar Resultados
            st.subheader("Vista Previa del Resultado")
            st.dataframe(df_llegadas.head(15))

            # 5. Generar Excel para descarga
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                df_llegadas.to_excel(writer, index=False, sheet_name='Plan_ZAAL')
            
            st.download_button(
                label="📥 Descargar Plan de Reparto (Excel)",
                data=buffer.getvalue(),
                file_name="Resultado_Rutas_ZAAL.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Hubo un problema al procesar los datos: {e}")
