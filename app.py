import uuid
import shutil
import tempfile
from pathlib import Path

import streamlit as st
from reordenar_rutas import reordenar_excel


# ==========================================================
# CONFIG
# ==========================================================

st.set_page_config(page_title="Reparto determinista", layout="wide")
st.title("Reparto determinista")

# ==========================================================
# WORKDIR
# ==========================================================

if "workdir" not in st.session_state:
    st.session_state.workdir = Path(tempfile.mkdtemp(prefix="reparto_"))
    st.session_state.run_id = str(uuid.uuid4())[:8]

workdir = st.session_state.workdir

with st.sidebar:
    st.write(f"Run ID: {st.session_state.run_id}")
    if st.button("Reset sesión"):
        shutil.rmtree(workdir, ignore_errors=True)
        st.session_state.workdir = Path(tempfile.mkdtemp(prefix="reparto_"))
        st.session_state.run_id = str(uuid.uuid4())[:8]
        st.rerun()


# ==========================================================
# MENÚ HORIZONTAL
# ==========================================================

tab1, tab2 = st.tabs([
    "FASE 1 · Asignación reparto",
    "FASE 2 · Reordenación topográfica"
])


# ==========================================================
# FASE 1 (SOLO VISUAL PARA PRUEBA)
# ==========================================================

with tab1:
    st.write("FASE 1 ACTIVA")
    st.info("Prueba visual. Sin ejecución real.")


# ==========================================================
# FASE 2 (PRUEBA REAL)
# ==========================================================

with tab2:

    st.write("FASE 2 ACTIVA")

    archivo_excel = st.file_uploader(
        "1) Subir salida.xlsx modificado",
        type=["xlsx"],
        key="excel"
    )

    archivo_coords = st.file_uploader(
        "2) Subir archivo de coordenadas",
        type=["xlsx"],
        key="coords"
    )

    if archivo_excel and archivo_coords:

        input_path = workdir / "entrada.xlsx"
        coords_path = workdir / "coords.xlsx"
        output_path = workdir / "salida_reordenada.xlsx"

        input_path.write_bytes(archivo_excel.getbuffer())
        coords_path.write_bytes(archivo_coords.getbuffer())

        if st.button("Reordenar rutas"):

            try:
                reordenar_excel(
                    input_path=input_path,
                    output_path=output_path,
                    ruta_coordenadas=coords_path,
                )
                st.success("Reordenación completada")

                if output_path.exists():
                    st.download_button(
                        "Descargar salida_reordenada.xlsx",
                        data=output_path.read_bytes(),
                        file_name="salida_reordenada.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )

            except Exception as e:
                st.error(f"Error: {e}")
    else:
        st.info("Sube ambos archivos para activar la reordenación.")
