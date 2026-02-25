import streamlit as st
import pandas as pd
import os
import io
from datetime import datetime

st.set_page_config(page_title="ZAAL Logística - Versión Local Pro", layout="wide")
st.title("🚚 ZAAL - Generador de salida.xlsx")

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

        # Limpiar nombres de columnas del CSV
        df_llegadas.columns = [c.strip() for c in df_llegadas.columns]
        cols_csv = df_llegadas.columns.tolist()

        # --- BUSCADOR INTELIGENTE DE COLUMNAS ---
        def buscar_col(lista, palabras):
            for p in palabras:
                for c in lista:
                    if p.upper() in c.upper(): return c
            return None

        col_dir = buscar_col(cols_csv, ['DIR', 'ENTREGA', 'DESTINO'])
        col_exp = buscar_col(cols_csv, ['EXPED', 'ENVIO', 'Nº'])
        col_kil = buscar_col(cols_csv, ['KILO', 'PESO', 'KGS'])
        col_pob = buscar_col(cols_csv, ['POB', 'CIUDAD', 'MUNICIPIO'])
        col_con = buscar_col(cols_csv, ['CONS', 'NOMBRE', 'CLIENTE'])
        col_bul = buscar_col(cols_csv, ['BULT', 'PAQUET'])

        # Validar si faltan columnas críticas
        if not col_exp or not col_kil:
            st.error(f"❌ No detecto columnas de Expediciones o Kilos. Columnas encontradas: {cols_csv}")
            st.stop()

        # 2. PROCESAMIENTO 1-3
        df_reglas = pd.concat([df_hosp_reg, df_fed_reg], ignore_index=True)
        df_reglas['len'] = df_reglas['Patrón_dirección'].astype(str).str.len()
        df_reglas = df_reglas.sort_values(by='len', ascending=False).drop(columns=['len'])

        if st.button("🚀 Generar salida.xlsx"):
            df_llegadas['Ruta_Asignada'] = "RESTO"
            df_llegadas['Bloque'] = "RESTO"

            # Diccionario para saber qué patrones son de Hospitales
            patrones_hosp = set(df_hosp_reg['Patrón_dirección'].astype(str).str.upper().strip())

            for idx, fila in df_llegadas.iterrows():
                direccion = str(fila[col_dir]).upper()
                for _, regla in df_reglas.iterrows():
                    patron = str(regla['Patrón_dirección']).upper().strip()
                    if patron and patron != "NAN" and patron in direccion:
                        df_llegadas.at[idx, 'Ruta_Asignada'] = regla['Ruta']
                        if patron in patrones_hosp:
                            df_llegadas.at[idx, 'Bloque'] = "HOSPITALES"
                        else:
                            df_llegadas.at[idx, 'Bloque'] = "FEDERACION"
                        break

            # 3. CREACIÓN DEL EXCEL (salida.xlsx)
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                
                # --- METADATOS ---
                meta = pd.DataFrame({
                    'Clave': ['Origen', 'CSV', 'Reglas', 'Fecha'],
                    'Valor': ['LLEGADAS', archivo_subido.name, 'Reglas_hospitales.xlsx', datetime.now().strftime("%Y-%m-%d %H:%M")]
                })
                meta.to_excel(writer, sheet_name='METADATOS', index=True)

                # --- RESUMEN_GENERAL ---
                res_gen = df_llegadas.groupby('Bloque').agg({
                    col_dir: 'nunique',
                    col_exp: 'sum',
                    col_kil: 'sum'
                }).reset_index()
                res_gen.columns = ['Bloque', 'Paradas', 'Expediciones', 'Kilos']
                res_gen.to_excel(writer, sheet_name='RESUMEN_GENERAL', index=False)

                # --- RESUMEN_UNICO ---
                res_uni = df_llegadas.groupby(['Bloque', 'Ruta_Asignada']).agg({
                    col_dir: 'nunique',
                    col_exp: 'sum',
                    col_kil: 'sum'
                }).reset_index()
                res_uni.columns = ['Tipo', 'Clave', 'Paradas', 'Expediciones', 'Kilos']
                res_uni.to_excel(writer, sheet_name='RESUMEN_UNICO', index=False)

                # --- Pestañas ZREP ---
                rutas = sorted(df_llegadas['Ruta_Asignada'].unique())
                for r in rutas:
                    df_r = df_llegadas[df_llegadas['Ruta_Asignada'] == r].copy()
                    # Numerar paradas
                    df_r.insert(0, 'Parada', range(1, len(df_r) + 1))
                    nombre_hoja = f"ZREP_{str(r)[:20]}".replace('/', '-')
                    df_r.to_excel(writer, sheet_name=nombre_hoja[:31], index=False)

            st.success("✨ ¡Archivo generado con éxito!")
            st.download_button("📥 Descargar salida.xlsx", output.getvalue(), "salida.xlsx")

    except Exception as e:
        st.error(f"Error detallado: {e}")
