import sys
import uuid
import shutil
import tempfile
import subprocess
from pathlib import Path

import streamlit as st
from reordenar_rutas import reordenar_excel


# ==========================================================
# CONFIGURACIÓN GENERAL
# ==========================================================

st.set_page_config(page_title="Reparto determinista", layout="wide")
st.title("Reparto determinista")

REPO_DIR = Path(__file__).resolve().parent
SCRIPT_REPARTO = REPO_DIR / "reparto_gpt.py"
REGLAS_REPO = REPO_DIR / "Reglas_hospitales.xlsx"


# ==========================================================
# UTILIDADES
# ==========================================================

def ensure_workdir() -> Path:
    if "workdir" not in st.session_state:
        st.session_state.workdir = Path(tempfile.mkdtemp(prefix="reparto_"))
        st.session_state.run_id = str(uuid.uuid4())[:8]
    return st.session_state.workdir


def reset_session_dir():
    wd = st.session_state.get("workdir")
    if wd and isinstance(wd, Path):
        shutil.rmtree(wd, ignore_errors=True)
    st.session_state.workdir = Path(tempfile.mkdtemp(prefix="reparto_"))
    st.session_state.run_id = str(uuid.uuid4())[:8]


def save_upload(uploaded_file, dst: Path) -> Path:
    dst.write_bytes(uploaded_file.getbuffer())
    return dst


def run_process(cmd: list[str], cwd: Path, timeout_s: int = 300):
    try:
        p = subprocess.run(
            cmd,
            cwd=str(cwd),
            capture_output=True,
            text=True,
            timeout=timeout_s,
        )
        return p.returncode, p.stdout, p.stderr
    except subprocess.TimeoutExpired as e:
        return 124, e.stdout or "", f"TIMEOUT tras {timeout_s}s"


# ==========================================================
# ESTADO
# ==========================================================

workdir = ensure_workdir()

with st.sidebar:
    st.header("Estado sesión")
    st.write(f"Run ID: `{st.session_state.run_id}`")
    st.write(f"Workdir: `{workdir}`")
    if st.button("Reset sesión"):
        reset_session_dir()
        st.rerun()


# ==========================================================
# MENÚ SUPERIOR HORIZONTAL
# ==========================================================

tab_fase1, tab_fase2 = st.tabs(
    [
        "FASE 1 · Asignación reparto",
        "FASE 2 · Reordenación topográfica",
    ]
)


# ==========================================================
# FASE 1
# ==========================================================

with tab_fase1:

    st.subheader("1) Subir CSV de llegadas")

    csv_file = st.file_uploader(
        "CSV de llegadas",
        type=["csv"],
        key="fase1_csv"
    )

    if csv_file:

        csv_path = save_upload(csv_file, workdir / "llegadas.csv")

        if REGLAS_REPO.exists():
            (workdir / "Reglas_hospitales.xlsx").write_bytes(REGLAS_REPO.read_bytes())
        else:
            st.error("No se encuentra Reglas_hospitales.xlsx en el repo.")
        
        st.subheader("2) Ejecutar")

        if st.button("Generar salida.xlsx", type="primary", key="fase1_btn"):

            cmd = [
                sys.executable,
                str(SCRIPT_REPARTO),
                "--csv", "llegadas.csv",
                "--reglas", "Reglas_hospitales.xlsx",
                "--out", "salida.xlsx",
            ]

            with st.spinner("Ejecutando reparto_gpt.py…"):
                rc, out, err = run_process(cmd, cwd=workdir)

            if rc != 0:
                st.error("Error en reparto_gpt.py")
                st.code(err)
            else:
                salida_path = workdir / "salida.xlsx"

                if salida_path.exists():
                    st.success("Archivo generado correctamente")

                    st.download_button(
                        "Descargar salida.xlsx",
                        data=salida_path.read_bytes(),
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

with tab_fase2:

    st.subheader("Reordenar rutas existentes")

    archivo_excel = st.file_uploader(
        "1) Subir salida.xlsx modificado",
        type=["xlsx"],
        key="fase2_excel"
    )

    archivo_coords = st.file_uploader(
        "2) Subir archivo de coordenadas",
        type=["xlsx"],
        key="fase2_coords"
    )

    if archivo_excel and archivo_coords:

        input_path = save_upload(archivo_excel, workdir / "entrada_fase2.xlsx")
        coords_path = save_upload(archivo_coords, workdir / "coords.xlsx")
        output_path = workdir / "salida_reordenada.xlsx"

        if st.button("Reordenar rutas", type="primary", key="fase2_btn"):

            try:
                reordenar_excel(
                    input_path=input_path,
                    output_path=output_path,
                    ruta_coordenadas=coords_path,
                )
            except Exception as e:
                st.error(f"Error en reordenación: {e}")
            else:
                if output_path.exists():
                    st.success("Rutas reordenadas correctamente")

                    st.download_button(
                        "Descargar salida_reordenada.xlsx",
                        data=output_path.read_bytes(),
                        file_name="salida_reordenada.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
                else:
                    st.error("No se generó el archivo de salida.")
    else:
        st.info("Sube ambos archivos para habilitar la reordenación.")
