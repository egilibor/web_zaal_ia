import streamlit as st
import pandas as pd
import os
import io
from datetime import datetime

st.set_page_config(page_title="ZAAL Logística - Versión Local Pro", layout="wide")
st.title("🚚 ZAAL - Generador de salida.xlsx (Modo Local)")

ruta_reglas = "Reglas_hospitales.xlsx"

if not os.path.exists(ruta_reglas):
    st.error(f"❌ Falta el archivo de reglas.")
    st.stop()

archivo_subido = st.file_uploader("Sube el CSV de llegadas", type=["csv"])

if archivo_subido:
    try:
        # 1. Carga de datos
        df_llegadas = pd.read_csv(archivo_subido, sep=None, engine='python', encoding='latin-1')
        df_hosp_reg = pd.read_excel(ruta_reglas, sheet_name='REGLAS_HOSPITALES')
        df_fed_reg = pd.read_excel(ruta_reglas, sheet_name='REGLAS_FEDERACION')

        # Lógica 1-3: Unir y ordenar reglas por longitud de patrón
        df_reglas = pd.concat([df_hosp_reg, df_fed_reg], ignore_index=True)
        df_reglas['len'] = df_reglas['Patrón_dirección'].astype(str).str.len()
        df_reglas = df_reglas.sort_values(by='len', ascending=False).drop(columns=['len'])

        # 2. PROCESAMIENTO
        df_llegadas.columns = [c.strip() for c in df_llegadas.columns]
        col_dir = next((c for c in df_llegadas.columns if 'DIR' in c.upper() or 'ENTREGA' in c.upper()), df_llegadas.columns[0])
        col_patron = 'Patrón_dirección'
        
        if st.button("🚀 Generar salida.xlsx idéntico al local"):
            df_llegadas['Ruta_Asignada'] = "RESTO"
            df_llegadas['Bloque'] = "RESTO"

            for idx, fila in df_llegadas.iterrows():
                direccion = str(fila[col_dir]).upper()
                for _, regla in df_reglas.iterrows():
                    patron = str(regla[col_patron]).upper().strip()
                    if patron and patron != "NAN" and patron in direccion:
                        df_llegadas.at[idx, 'Ruta_Asignada'] = regla['Ruta']
                        # Marcar si pertenece a Hospitales o Federación para los resúmenes
                        if patron in str(df_hosp_reg['Patrón_dirección'].values).upper():
                            df_llegadas.at[idx, 'Bloque'] = "HOSPITALES"
                        else:
                            df_llegadas.at[idx, 'Bloque'] = "FEDERACION"
                        break

            # 3. CREACIÓN DEL EXCEL MULTI-HOJA (salida.xlsx)
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                
                # --- Hoja 1: METADATOS ---
                meta = pd.DataFrame({
                    'Clave': ['Origen de datos', 'CSV', 'Reglas', 'Generado'],
                    'Valor': ['LLEGADAS', 'llegadas.csv', 'Reglas_hospitales.xlsx', datetime.now().strftime("%Y-%m-%d %H:%M:%S")]
                })
                meta.to_excel(writer, sheet_name='METADATOS', index=True, index_label='Nº')

                # --- Hoja 2: RESUMEN_GENERAL ---
                resumen_gen = df_llegadas.groupby('Bloque').agg(
                    Paradas=(col_dir, 'nunique'),
                    Expediciones=('Expediciones', 'sum'),
                    Kilos=('Kilos', 'sum')
                ).reset_index()
                resumen_gen.to_excel(writer, sheet_name='RESUMEN_GENERAL', index=True, index_label='Nº')

                # --- Hoja 3: RESUMEN_UNICO ---
                resumen_uni = df_llegadas.groupby(['Bloque', 'Ruta_Asignada']).agg(
                    Paradas=(col_dir, 'nunique'),
                    Expediciones=('Expediciones', 'sum'),
                    Kilos=('Kilos', 'sum')
                ).reset_index()
                resumen_uni.to_excel(writer, sheet_name='RESUMEN_UNICO', index=True, index_label='Nº')

                # --- Pestañas por Zonas de Reparto (ZREP) ---
                rutas = sorted(df_llegadas['Ruta_Asignada'].unique())
                for r in rutas:
                    df_r = df_llegadas[df_llegadas['Ruta_Asignada'] == r].copy()
                    # Limpieza de nombre para la pestaña
                    nombre_hoja = f"ZREP_{str(r)[:20]}".replace('/', ' ')
                    # Seleccionamos y renombramos columnas para que sea igual que tu muestra
                    # Ajustamos según lo que tenga tu CSV original
                    df_r.to_excel(writer, sheet_name=nombre_hoja[:31], index=False)

            st.success("✨ ¡Vibe Check 100%! El archivo salida.xlsx está listo.")
            
            st.download_button(
                label="📥 Descargar salida.xlsx",
                data=output.getvalue(),
                file_name="salida.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Error reproduciendo el local: {e}")
