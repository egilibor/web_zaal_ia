import streamlit as st
import pandas as pd
import os
import io

# Configuración profesional de la página
st.set_page_config(page_title="ZAAL Logística IA", page_icon="🚚", layout="wide")

st.title("🚚 ZAAL - Clasificador de Rutas Inteligente")
st.markdown("---")

# --- LOCALIZADOR DE ARCHIVOS ---
# Buscamos el Excel de reglas en la raíz o en la carpeta web_zaal_ia
nombre_excel = "Reglas_hospitales.xlsx"
ruta_excel = None

posibles_rutas = [
    nombre_excel,
    os.path.join("web_zaal_ia", nombre_excel),
    os.path.join(os.path.dirname(__file__), nombre_excel)
]

for ruta in posibles_rutas:
    if os.path.exists(ruta):
        ruta_excel = ruta
        break

# --- INTERFAZ DE USUARIO ---
if not ruta_excel:
    st.error(f"❌ No se encuentra el archivo '{nombre_excel}' en el repositorio.")
    st.info("Asegúrate de que el nombre sea exacto y esté subido a GitHub.")
    st.stop()

st.sidebar.header("Configuración")
archivo_subido = st.file_uploader("📂 Sube el archivo 'llegadas.csv'", type=["csv"])

if archivo_subido is not None:
    try:
        # 1. Leer el CSV con codificación robusta
        df_llegadas = pd.read_csv(archivo_subido, sep=None, engine='python', encoding='latin-1')
        
        # 2. Leer las reglas de Excel (requiere openpyxl en requirements.txt)
        df_reglas = pd.read_excel(ruta_excel, engine='openpyxl')
        
        st.success("✅ Datos cargados correctamente.")

        if st.button("🚀 Procesar Clasificación"):
            # Limpiar nombres de columnas para evitar errores de espacios
            df_llegadas.columns = [str(c).strip() for c in df_llegadas.columns]
            df_reglas.columns = [str(c).strip() for c in df_reglas.columns]
            
            # 3. Lógica de asignación de rutas
            df_llegadas['Ruta'] = "SIN ASIGNAR"
            
            # Buscamos la columna de dirección (ajustar si el nombre varía en el CSV)
            col_direccion = 'Dir. entrega' if 'Dir. entrega' in df_llegadas.columns else df_llegadas.columns[0]

            for idx, fila in df_llegadas.iterrows():
                direccion = str(fila[col_direccion]).upper()
                for _, regla in df_reglas.iterrows():
                    patron = str(regla.get('Patron', '')).upper()
                    if patron and patron in direccion:
                        df_llegadas.at[idx, 'Ruta'] = regla.get('Ruta', 'RUTA DESCONOCIDA')
                        break

            # 4. Mostrar Resultados
            st.subheader("📋 Previsualización del Reparto")
            st.dataframe(df_llegadas.head(20), use_container_width=True)

            # 5. Botón de Descarga
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_llegadas.to_excel(writer, index=False, sheet_name='Reparto_ZAAL')
            
            st.download_button(
                label="📥 Descargar Plan Final (Excel)",
                data=output.getvalue(),
                file_name="Plan_Logistica_ZAAL.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"⚠️ Error al procesar: {e}")
        st.info("Asegúrate de que el CSV tenga el formato correcto y use codificación estándar.")

else:
    st.info("👋 Bienvenido. Por favor, sube el archivo CSV de llegadas para empezar la clasificación.")
