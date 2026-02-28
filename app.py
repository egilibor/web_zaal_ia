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
st.title("Reparto determinista (Streamlit)")

REPO_DIR = Path(__file__).resolve().parent
SCRIPT_REPARTO = REPO_DIR / "reparto_gpt.py"
REGLAS_REPO = REPO_DIR / "Reglas_hospitales.xlsx"
COORDENADAS_REPO = REPO_DIR / "Libro_de_Servicio_Castellon_con_coordenadas.xlsx"


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


def run_process(cmd: list[str], cwd: Path, timeout_s: int = 300) -> tuple[int, str, str]:
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
        stdout = e.stdout or ""
        stderr = e.stderr or ""
        return 124, stdout, f"TIMEOUT tras {timeout_s}s\n{stderr}"


def show_logs(stdout: str, stderr: str):
    if stdout.strip():
        st.subheader("STDOUT")
        st.code(stdout)
    if stderr.strip():
        st.subheader("STDERR")
        st.code(stderr)


# ==========================================================
# ESTADO (SIDEBAR)
# ==========================================================

workdir = ensure_workdir()

with st.sidebar:
    st.header("Estado")
    st.write(f"Run: `{st.session_state.run_id}`")
    st.write(f"Workdir: `{workdir}`")
    st.write(f"Repo dir: `{REPO_DIR}`")
    st.write(f"Python: `{sys.executable}`")

    st.divider()
    st.write(f"reparto_gpt.py = `{SCRIPT_REPARTO.exists()}`")
    st.write(f"Reglas_hospitales.xlsx = `{REGLAS_REPO.exists()}`")
    st.write(f"Libro coordenadas = `{COORDENADAS_REPO.exists()}`")

    st.divider()
    if st.button("Reset sesión"):
        reset_session_dir()
        st.rerun()


# ==========================================================
# VERIFICACIONES DURAS
# ==========================================================

missing = []

if not SCRIPT_REPARTO.exists():
    missing.append("reparto_gpt.py")

if not REGLAS_REPO.exists():
    missing.append("Reglas_hospitales.xlsx")

if not COORDENADAS_REPO.exists():
    missing.append("Libro_de_Servicio_Castellon_con_coordenadas.xlsx")

if missing:
    st.error("Faltan archivos en el repo desplegado: " + ", ".join(missing))
    st.stop()

st.divider()


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
# FASE 1 · ASIGNACIÓN REPARTO
# ==========================================================

with tab_fase1:

    st.subheader("1) Subir CSV de llegadas")

    csv_file = st.file_uploader(
        "CSV de llegadas",
        type=["csv"],
        key="fase1_csv"
    )

    st.divider()

    if not csv_file:
        st.info("Sube el CSV para habilitar la ejecución.")
        st.stop()

    csv_path = save_upload(csv_file, workdir / "llegadas.csv")
    (workdir / "Reglas_hospitales.xlsx").write_bytes(REGLAS_REPO.read_bytes())

    st.subheader("2) Ejecutar (genera salida.xlsx)")

    if st.button("Ejecutar", type="primary", key="fase1_btn"):

        cmd_gpt = [
            sys.executable,
            str(SCRIPT_REPARTO),
            "--csv", "llegadas.csv",
            "--reglas", "Reglas_hospitales.xlsx",
            "--out", "salida.xlsx",
        ]

        with st.spinner("Ejecutando reparto_gpt.py…"):
            rc, out, err = run_process(cmd_gpt, cwd=workdir, timeout_s=300)

        if rc != 0:
            st.error("❌ Falló reparto_gpt.py")
            show_logs(out, err)
            st.stop()

        salida_path = workdir / "salida.xlsx"

        if not salida_path.exists():
            st.error("Terminó sin error, pero no encuentro `salida.xlsx`.")
            show_logs(out, err)
            st.stop()

        st.success("✅ salida.xlsx generado")

        st.download_button(
            "Descargar salida.xlsx",
            data=salida_path.read_bytes(),
            file_name="salida.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


# ==========================================================
# FASE 2 · REORDENACIÓN TOPOGRÁFICA
# ==========================================================

with tab_fase2:

    st.subheader("Reordenar rutas existentes")

    archivo_modificado = st.file_uploader(
        "Subir salida.xlsx modificado",
        type=["xlsx"],
        key="fase2_uploader"
    )

    if archivo_modificado:
    
        input_path = save_upload(archivo_modificado, workdir / "entrada_fase2.xlsx")
        output_path = workdir / "salida_reordenada.xlsx"
    
        if st.button("Reordenar rutas", type="primary", key="fase2_btn"):
    
            try:
                reordenar_excel(
                    input_path=input_path,
                    output_path=output_path,
                    ruta_coordenadas=COORDENADAS_REPO,
                )
            except Exception as e:
                st.error(f"Error en reordenación: {e}")
            else:
                if output_path.exists():
                    st.success("✅ Rutas reordenadas correctamente")
    
                    st.download_button(
                        "Descargar salida_reordenada.xlsx",
                        data=output_path.read_bytes(),
                        file_name="salida_reordenada.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
                else:
                    st.error("No se generó el archivo reordenado.")
    else:
        st.info("Sube el archivo para habilitar la reordenación.")
    
    input_path = save_upload(archivo_modificado, workdir / "entrada_fase2.xlsx")
    output_path = workdir / "salida_reordenada.xlsx"

    if st.button("Reordenar rutas", type="primary", key="fase2_btn"):

        try:
            reordenar_excel(
                input_path=input_path,
                output_path=output_path,
                ruta_coordenadas=COORDENADAS_REPO,
            )
        except Exception as e:
            st.error(f"Error en reordenación: {e}")
            st.stop()

        if not output_path.exists():
            st.error("No se generó el archivo reordenado.")
            st.stop()

        st.success("✅ Rutas reordenadas correctamente")

        st.download_button(
            "Descargar salida_reordenada.xlsx",
            data=output_path.read_bytes(),
            file_name="salida_reordenada.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

