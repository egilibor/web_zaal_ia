import sys
import uuid
import shutil
import tempfile
import subprocess
from pathlib import Path

import streamlit as st
from reordenar_rutas import reordenar_excel
import importlib
import add_resumen_unico

# ==========================================================
# CONFIG
# ==========================================================

st.set_page_config(page_title="Reparto determinista", layout="wide")
st.title("Reparto determinista")

# -----------------------------------
# SELECCIÓN DE DELEGACIÓN
# -----------------------------------

delegacion = st.selectbox(
    "Delegación",
    ["Castellón", "Valencia"]
)

REPO_DIR = Path(__file__).resolve().parent
SCRIPT_REPARTO = REPO_DIR / "reparto_gpt.py"

if delegacion == "Castellón":
    REGLAS_REPO = REPO_DIR / "Reglas_hospitales.xlsx"
    COORDENADAS_REPO = REPO_DIR / "Libro_Servicio_Castellon.xlsx"
    LAT0 = 39.804106
    LON0 = -0.217351
elif delegacion == "Valencia":
    REGLAS_REPO = REPO_DIR / "Reglas_hospitales.xlsx"
    COORDENADAS_REPO = REPO_DIR / "valencia_municipios_coordenadas.xlsx"
    LAT0 = 39.44068
    LON0 = -0.42592
    
# ==========================================================
# WORKDIR POR SESIÓN
# ==========================================================

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


# ==========================================================
# MENÚ
# ==========================================================

tab1, tab2 = st.tabs([
    "FASE 1 · Asignación reparto",
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

        # Copiamos reglas al workdir
        reglas_path = workdir / "Reglas_hospitales.xlsx"
        reglas_path.write_bytes(REGLAS_REPO.read_bytes())

        if st.button("Generar reparto", key="fase1_btn"):

            # 🔹 Nombre único por ejecución
            unique_id = uuid.uuid4().hex[:10]
            nombre_salida = f"rutas_{unique_id}.xlsx"
            salida_path = workdir / nombre_salida

            cmd = [
                sys.executable,
                str(SCRIPT_REPARTO),
                "--csv", "llegadas.csv",
                "--reglas", "Reglas_hospitales.xlsx",
                "--out", nombre_salida,
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
                if salida_path.exists():

                    st.info(f"Archivo generado: {nombre_salida}")

                    # Ejecutamos módulo de resumen
                    importlib.reload(add_resumen_unico)
                    add_resumen_unico.generar_resumen_unico(str(salida_path))
                    #st.warning("add_resumen_unico ejecutado")

                    #st.success("Archivo generado correctamente")

                    st.download_button(
                        "Descargar archivo generado",
                        data=salida_path.read_bytes(),
                        file_name=nombre_salida,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
                else:
                    st.error("No se generó el archivo de salida")
    else:
        st.info("Sube un CSV para habilitar la ejecución.")


# ==========================================================
# FASE 2
# ==========================================================

with tab2:

    st.subheader("Reordenar rutas existentes")

    archivo_excel = st.file_uploader(
        "Subir archivo Excel",
        type=["xlsx"],
        key="fase2_excel"
    )

    if archivo_excel:

        input_path = workdir / "entrada_fase2.xlsx"
        output_unique = f"salida_reordenada_{uuid.uuid4().hex[:8]}.xlsx"
        output_path = workdir / output_unique

        input_path.write_bytes(archivo_excel.getbuffer())

        if st.button("Reordenar rutas", key="fase2_btn"):

            try:
                reordenar_excel(
                    input_path=input_path,
                    output_path=output_path,
                    ruta_coordenadas=COORDENADAS_REPO,
                )
                
                # 🔁 Regenerar RESUMEN_UNICO tras reordenar
                importlib.reload(add_resumen_unico)
                add_resumen_unico.generar_resumen_unico(str(output_path))

                if output_path.exists():
                    st.success("Rutas reordenadas correctamente")

                    st.download_button(
                        "Descargar archivo reordenado",
                        data=output_path.read_bytes(),
                        file_name=output_unique,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
                else:
                    st.error("No se generó el archivo reordenado.")

            except Exception as e:
                st.error(f"Error en reordenación: {e}")

    else:
        st.info("Sube el archivo para activar la reordenación.")








