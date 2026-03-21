import sys
import uuid
import shutil
import tempfile
import subprocess
import datetime
import pandas as pd
from pathlib import Path

import streamlit as st
from reordenar_rutas import reordenar_excel
from add_resumen_unico import generar_resumen_unico
from modulo_valencia_gestores import generar_libros_gestores
from openpyxl import load_workbook

# ==========================================================
# WORKDIR
# ==========================================================
hora_salida = st.sidebar.time_input(
    "Hora de salida",
    value=datetime.time(8, 30)
)

if "workdir" not in st.session_state:
    st.session_state.workdir = Path(tempfile.mkdtemp(prefix="reparto_"))
    st.session_state.run_id = str(uuid.uuid4())[:8]

workdir = st.session_state.workdir

with st.sidebar:
    st.write(f"Run ID: {st.session_state.run_id}")
    st.write(f"Workdir: {workdir}")
    if st.button("Reset sesión"):
        shutil.rmtree(workdir, ignore_errors=True)
        st.session_state.workdir = Path(tempfile.mkdtemp(prefix="reparto_"))
        st.session_state.run_id = str(uuid.uuid4())[:8]
        st.rerun()
    if st.button("🗑️ Limpiar caché geocodificación"):
        from geocodificador import limpiar_cache
        limpiar_cache()
        st.success("Caché limpiada correctamente")
# ==========================================================
# CONFIG
# ==========================================================
st.set_page_config(page_title="Reparto determinista", layout="wide")

if "GOOGLE_MAPS_API_KEY" not in st.secrets:
    st.error("⚠️ Falta la clave GOOGLE_MAPS_API_KEY en los secrets. Contacta con el administrador.")
    st.stop()

REPO_DIR = Path(__file__).resolve().parent
SCRIPT_REPARTO = REPO_DIR / "reparto_gpt.py"
REGLAS_REPO = REPO_DIR / "Reglas_hospitales.xlsx"

# ==========================================================
# DELEGACIÓN - PANTALLA DE INICIO
# ==========================================================
if "delegacion_activa" not in st.session_state:
    st.session_state.delegacion_activa = None

if st.session_state.delegacion_activa is None:
    st.markdown("## Selecciona la delegación")
    st.markdown("---")
    col1, col2 = st.columns(2)
    with col1:
        if st.button("🏙️ CASTELLÓN", use_container_width=True, type="primary"):
            st.session_state.delegacion_activa = "castellon"
            st.rerun()
    with col2:
        if st.button("🌆 VALENCIA", use_container_width=True, type="primary"):
            st.session_state.delegacion_activa = "valencia"
            st.rerun()
    st.stop()

delegacion = st.session_state.delegacion_activa

COORDENADAS_FILES = {
    "castellon": "Libro_de_Servicio_Castellon_con_coordenadas.xlsx",
    "valencia": "valencia_municipios_coordenadas.xlsx",
}

COORDENADAS_REPO = REPO_DIR / COORDENADAS_FILES[delegacion]

st.title(f"Reparto determinista — {delegacion.upper()}")

# Botón para cambiar delegación en sidebar
if st.sidebar.button("🔄 Cambiar delegación"):
    st.session_state.delegacion_activa = None
    st.session_state.pop("workdir", None)
    st.rerun()

# ==========================================================
# MENÚ HORIZONTAL
# ==========================================================

tab1, tab2, tab3 = st.tabs([
    "FASE 1 · Asignación reparto",
    "FASE 2 . Ajuste Manual",
    "FASE 2 · Reordenación topográfica"
])


# ==========================================================
# FASE 1
# ==========================================================

with tab1:

    st.subheader("Subir CSV de llegadas")

    csv_file = st.file_uploader(
        "CSV de llegadas",
        type=["csv"],
        key="fase1_csv"
    )

    if csv_file:

        input_csv = workdir / "llegadas.csv"
        input_csv.write_bytes(csv_file.getbuffer())

        (workdir / "Reglas_hospitales.xlsx").write_bytes(
            REGLAS_REPO.read_bytes()
        )

        if st.button("Generar salida.xlsx", key="fase1_btn"):

            cmd = [
                sys.executable,
                str(SCRIPT_REPARTO),
                "--csv", "llegadas.csv",
                "--reglas", "Reglas_hospitales.xlsx",
                "--out", "salida.xlsx",
            ]

            with st.spinner("Ejecutando reparto_gpt.py…"):
                p = subprocess.run(
                    cmd,
                    cwd=str(workdir),
                    capture_output=True,
                    text=True,
                )

            if p.returncode != 0:
                st.error("Error en reparto_gpt.py")
                st.code(p.stderr)
            else:
                salida = workdir / "salida.xlsx"
                if salida.exists():
                    generar_resumen_unico(str(salida))
                    st.success("Archivo generado correctamente")

                    st.download_button(
                        "Descargar salida.xlsx",
                        data=salida.read_bytes(),
                        file_name="salida.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
                else:
                    st.error("No se generó salida.xlsx")
    else:
        st.info("Sube un CSV para habilitar la ejecución.")

# ==========================================================
# FASE 2
# ==========================================================
with tab2:

    st.subheader("Ajuste manual de expediciones")

    excel_ajuste = st.file_uploader(
        "Subir salida.xlsx (generado en Fase 1)",
        type=["xlsx"],
        key="fase2_ajuste_excel"
    )

    if excel_ajuste:

        file_id = excel_ajuste.name + str(excel_ajuste.size)
        if st.session_state.get("ajuste_file_id") != file_id:
            ajuste_path = workdir / "ajuste_entrada.xlsx"
            ajuste_path.write_bytes(excel_ajuste.getbuffer())
            wb = load_workbook(ajuste_path)
            st.session_state["ajuste_wb"] = wb
            st.session_state["ajuste_file_id"] = file_id

        wb = st.session_state["ajuste_wb"]
        hojas_disponibles = wb.sheetnames
        hojas_operativas = [
            h for h in hojas_disponibles
            if h.startswith("ZREP_") or h in ("HOSPITALES", "FEDERACION")
        ]

        if not hojas_operativas:
            st.warning("No se encontraron hojas operativas (ZREP_, HOSPITALES, FEDERACION).")
        else:
            col_iz, col_der = st.columns(2)

            with col_iz:
                hoja_origen = st.selectbox("Hoja origen", hojas_operativas, key="hoja_origen")
            with col_der:
                hojas_destino = [h for h in hojas_operativas if h != hoja_origen]
                hoja_destino = st.selectbox("Hoja destino", hojas_destino, key="hoja_destino")

            def ws_to_df(wb, nombre):
                ws = wb[nombre]
                datos = list(ws.values)
                if not datos:
                    return pd.DataFrame()
                return pd.DataFrame(datos[1:], columns=datos[0])

            df_origen = ws_to_df(wb, hoja_origen)
            df_destino = ws_to_df(wb, hoja_destino)

            columnas_mostrar = [
                c for c in ["Exp", "Consignatario", "Población", "Dirección", "Kgs"]
                if c in df_origen.columns
            ]

            with st.expander(f"Ver hoja destino: {hoja_destino} ({len(df_destino)} expediciones)"):
                st.dataframe(df_destino[columnas_mostrar] if columnas_mostrar else df_destino, use_container_width=True)

            st.markdown(f"**{hoja_origen}** — {len(df_origen)} expediciones")

            seleccionar_todas = st.checkbox("Seleccionar todas", key="chk_master")

            seleccion = {}
            for idx, row in df_origen.iterrows():
                etiqueta = " · ".join(
                    str(row[c]) for c in columnas_mostrar if pd.notna(row.get(c))
                )
                seleccion[idx] = st.checkbox(etiqueta, value=seleccionar_todas, key=f"chk_{idx}")

            indices_seleccionados = [idx for idx, marcado in seleccion.items() if marcado]
            st.markdown(f"*{len(indices_seleccionados)} expedición(es) seleccionada(s)*")

            if indices_seleccionados:
                if st.button(f"Mover {len(indices_seleccionados)} exp. → {hoja_destino}", key="btn_mover"):

                    from openpyxl.utils.dataframe import dataframe_to_rows

                    df_src = ws_to_df(wb, hoja_origen)
                    df_dst = ws_to_df(wb, hoja_destino)

                    filas_a_mover = df_src.loc[indices_seleccionados]
                    df_src_nuevo = df_src.drop(index=indices_seleccionados).reset_index(drop=True)
                    df_dst_nuevo = pd.concat([df_dst, filas_a_mover], ignore_index=True)

                    for nombre_hoja, df_nuevo in [(hoja_origen, df_src_nuevo), (hoja_destino, df_dst_nuevo)]:
                        ws = wb[nombre_hoja]
                        ws.delete_rows(1, ws.max_row)
                        for r in dataframe_to_rows(df_nuevo, index=False, header=True):
                            ws.append(r)

                    st.session_state["ajuste_wb"] = wb
                    st.success(f"{len(indices_seleccionados)} expedición(es) movidas de '{hoja_origen}' a '{hoja_destino}'")
                    st.rerun()

            ajuste_salida = workdir / "ajuste_salida.xlsx"
            wb.save(ajuste_salida)
            st.download_button(
                "⬇️ Descargar Excel modificado",
                data=ajuste_salida.read_bytes(),
                file_name="ajuste_salida.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="btn_descarga_ajuste"
            )

    else:
        st.info("Sube el salida.xlsx de la Fase 1 para hacer ajustes manuales.")
        
# ==========================================================
# FASE 3
# ==========================================================

with tab3:

    st.subheader("Reordenar rutas existentes")

    archivo_excel = st.file_uploader(
        "Subir salida.xlsx modificado",
        type=["xlsx"],
        key="fase2_excel"
    )

    if archivo_excel:

        input_path = workdir / "entrada_fase2.xlsx"
        output_path = workdir / "salida_reordenada.xlsx"

        input_path.write_bytes(archivo_excel.getbuffer())

        if st.button("Reordenar rutas", key="fase2_btn"):

            try:
                
                if delegacion == "valencia":
                    lat_origen = 39.44069
                    lon_origen = -0.42589
                else:
                    lat_origen = 39.804106
                    lon_origen = -0.217351
                    
                paradas = reordenar_excel(
                    input_path,
                    output_path,
                    COORDENADAS_REPO,
                    lat_origen,
                    lon_origen,
                    api_key=st.secrets["GOOGLE_MAPS_API_KEY"],
                    delegacion=delegacion,
                    hora_salida=hora_salida,
                )
                
 
                generar_resumen_unico(str(output_path), paradas_por_hoja=paradas)
                
                if output_path.exists():

                    st.success("Rutas reordenadas correctamente")

                    st.download_button(
                        "Descargar salida_reordenada.xlsx",
                        data=output_path.read_bytes(),
                        file_name="salida_reordenada.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )

                else:
                    st.error("No se generó el archivo reordenado.")

            except Exception as e:
                import traceback
                st.error(f"Error en reordenación: {e}")
                st.code(traceback.format_exc())

        # -------------------------------------------------
        # DIVIDIR POR GESTORES (SOLO VALENCIA)
        # -------------------------------------------------

        if delegacion == "valencia" and output_path.exists():

            st.markdown("---")
            st.subheader("Dividir rutas para gestores de tráfico")

            if st.button("Generar Excel por gestor", key="fase3_btn"):

                try:

                    ruta_asignacion = REPO_DIR / "gestor_zonas.xlsx"

                    resultado = generar_libros_gestores(
                        ruta_excel_final=str(output_path),
                        ruta_asignacion=str(ruta_asignacion),
                        carpeta_salida=str(workdir)
                    )

                    if not resultado["ok"]:

                        st.error("Error generando libros de gestores")

                        for e in resultado["errores"]:
                            st.write(e)

                    else:

                        st.success("Archivos generados correctamente")

                        for gestor, ruta_archivo in resultado["archivos_generados"].items():

                            ruta = Path(ruta_archivo)

                            st.download_button(
                                label=f"Descargar Excel {gestor}",
                                data=ruta.read_bytes(),
                                file_name=ruta.name,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            )

                except Exception as e:

                    st.error(f"Error en generación de gestores: {e}")

    else:

        st.info("Sube el archivo para activar la reordenación.")



