# app.py — Streamlit wrapper robusto (sin pedir Reglas_hospitales.xlsx al usuario)
# - El Excel de reglas vive en el repo junto a los .py
# - El usuario solo sube el CSV de llegadas
# - Ejecuta reparto_gpt.py con sys.executable y ruta absoluta del script
# - Trabaja en un directorio temporal por sesión
# - Copia Reglas_hospitales.xlsx del repo al workdir para que el script lo encuentre
# - Muestra stdout/stderr si falla y permite descargar salida.xlsx

import sys
import uuid
import shutil
import tempfile
import subprocess
from pathlib import Path

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Reparto determinista", layout="wide")
st.title("Reparto determinista (Streamlit)")

REPO_DIR = Path(__file__).resolve().parent
SCRIPT_REPARTO = REPO_DIR / "reparto_gpt.py"          # Ajusta si cambia el nombre o la carpeta
REGLAS_REPO = REPO_DIR / "Reglas_hospitales.xlsx"     # Debe estar en el repo junto a app.py

# -------------------------
# Utilidades
# -------------------------
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

def run_cmd(cmd: list[str], cwd: Path) -> tuple[int, str, str]:
    p = subprocess.run(
        cmd,
        cwd=str(cwd),
        capture_output=True,
        text=True,
    )
    return p.returncode, p.stdout, p.stderr

def show_logs(stdout: str, stderr: str):
    if stdout.strip():
        st.subheader("STDOUT")
        st.code(stdout)
    if stderr.strip():
        st.subheader("STDERR")
        st.code(stderr)

# -------------------------
# Estado
# -------------------------
workdir = ensure_workdir()

with st.sidebar:
    st.header("Estado")
    st.write(f"Run: `{st.session_state.run_id}`")
    st.write(f"Workdir: `{workdir}`")
    st.write(f"Repo dir: `{REPO_DIR}`")
    st.write(f"Script: `{SCRIPT_REPARTO}`")
    st.write(f"Reglas: `{REGLAS_REPO}`")

    if st.button("Reset sesión"):
        reset_session_dir()
        st.rerun()

# -------------------------
# Verificaciones
# -------------------------
if not SCRIPT_REPARTO.exists():
    st.error(
        "No encuentro `reparto_gpt.py` en el repositorio.\n\n"
        "Asegúrate de que está en la misma carpeta que `app.py` o ajusta `SCRIPT_REPARTO`."
    )
    st.stop()

if not REGLAS_REPO.exists():
    st.error(
        "No encuentro `Reglas_hospitales.xlsx` en el repositorio.\n\n"
        "Súbelo al repo (misma carpeta que `app.py`) o ajusta `REGLAS_REPO`."
    )
    st.stop()

st.divider()

# -------------------------
# Input (solo CSV)
# -------------------------
st.subheader("1) Subir CSV de llegadas")
csv_file = st.file_uploader("CSV de llegadas", type=["csv"])

st.divider()

# Opciones de preview (solo visual)
st.subheader("2) Previsualización (opcional)")
col1, col2 = st.columns(2, gap="large")
with col1:
    sep = st.selectbox("Separador CSV", options=[";", ",", "TAB"], index=0)
    sep_val = "\t" if sep == "TAB" else sep
with col2:
    encoding = st.selectbox("Encoding", options=["utf-8", "latin1", "cp1252"], index=0)

preview_rows = st.slider("Filas de previsualización", 5, 50, 10)

st.divider()

# -------------------------
# Guardado + preview + ejecución
# -------------------------
if csv_file:
    # Guardar CSV en workdir con nombre esperado por el script
    csv_path = save_upload(csv_file, workdir / "llegadas.csv")

    # Copiar reglas del repo al workdir para que el script pueda usar un path simple
    reglas_path = workdir / "Reglas_hospitales.xlsx"
    reglas_path.write_bytes(REGLAS_REPO.read_bytes())

    # Preview
    try:
        df_prev = pd.read_csv(csv_path, sep=sep_val, encoding=encoding)
        st.dataframe(df_prev.head(preview_rows), use_container_width=True)
        st.caption(f"Columnas detectadas: {list(df_prev.columns)}")
    except Exception as e:
        st.warning("No he podido previsualizar el CSV con ese separador/encoding (no afecta a la ejecución del script).")
        st.exception(e)

    st.divider()

    st.subheader("3) Ejecutar")
    st.caption("Usa `Reglas_hospitales.xlsx` del repo. El usuario solo sube el CSV.")

    if st.button("Generar salida.xlsx", type="primary"):
        cmd = [
            sys.executable,
            str(SCRIPT_REPARTO),
            "--csv", "llegadas.csv",
            "--reglas", "Reglas_hospitales.xlsx",
            "--out", "salida.xlsx",
        ]

        st.info("Ejecutando…")
        rc, out, err = run_cmd(cmd, cwd=workdir)

        if rc != 0:
            st.error("❌ Falló reparto_gpt.py")
            show_logs(out, err)
            st.stop()

        salida_path = workdir / "salida.xlsx"
        if not salida_path.exists():
            st.error("El script terminó sin error, pero no encuentro `salida.xlsx`.")
            show_logs(out, err)
            st.stop()

        st.success("✅ salida.xlsx generada")

        # Preview del Excel si se puede
        try:
            df_out = pd.read_excel(salida_path)
            st.subheader("Vista previa de salida.xlsx")
            st.dataframe(df_out.head(preview_rows), use_container_width=True)
        except Exception:
            st.warning("No he podido previsualizar el Excel, pero el archivo está generado.")

        st.download_button(
            "Descargar salida.xlsx",
            data=salida_path.read_bytes(),
            file_name="salida.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
else:
    st.info("Sube el CSV para habilitar la ejecución.")
