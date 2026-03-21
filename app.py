import sys
import uuid
import shutil
import tempfile
import subprocess
import datetime
from pathlib import Path

import streamlit as st
from reordenar_rutas import reordenar_excel
from add_resumen_unico import generar_resumen_unico
from modulo_valencia_gestores import generar_libros_gestores

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

# ==========================================================
# FASE 3
# ==========================================================

