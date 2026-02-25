import io
import os
import sys
import uuid
import shutil
import tempfile
import subprocess
from pathlib import Path

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Reparto determinista", layout="wide")

st.title("Reparto determinista")
st.caption("Sube archivos → ejecuta scripts → descarga Excel. Sin rutas locales.")

# --- Utilidades ---
def run_cmd(cmd: list[str], cwd: Path) -> tuple[int, str, str]:
    """Ejecuta comando y devuelve (returncode, stdout, stderr)."""
    p = subprocess.run(
        cmd,
        cwd=str(cwd),
        capture_output=True,
        text=True,
    )
    return p.returncode, p.stdout, p.stderr

def save_upload(uploaded_file, dst: Path) -> Path:
    dst.write_bytes(uploaded_file.getbuffer())
    return dst

def show_logs(stdout: str, stderr: str):
    if stdout.strip():
        st.subheader("STDOUT")
        st.code(stdout)
    if stderr.strip():
        st.subheader("STDERR")
        st.code(stderr)

# --- Sesión / workspace ---
if "workdir" not in st.session_state:
    st.session_state.workdir = Path(tempfile.mkdtemp(prefix="reparto_"))
    st.session_state.run_id = str(uuid.uuid4())[:8]

workdir: Path = st.session_state.workdir

with st.sidebar:
    st.header("Estado")
    st.write(f"Run: `{st.session_state.run_id}`")
    st.write(f"Workdir: `{workdir}`")

    if st.button("Reset sesión"):
        try:
            shutil.rmtree(workdir, ignore_errors=True)
        finally:
            st.session_state.workdir = Path(tempfile.mkdtemp(prefix="reparto_"))
            st.session_state.run_id = str(uuid.uuid4())[:8]
        st.rerun()

st.divider()

# --- Inputs ---
col1, col2 = st.columns(2, gap="large")

with col1:
    st.subheader("1) Subida de archivos")
    csv_file = st.file_uploader("CSV de llegadas", type=["csv"], key="csv")
    reglas_file = st.file_uploader("Excel reglas (hospitales/zonas/etc.)", type=["xlsx"], key="reglas")

    st.caption("Consejo: usa nombres simples y sin espacios raros.")

with col2:
    st.subheader("2) Opciones")
    sep = st.selectbox("Separador CSV", options=[";", ",", "TAB"], index=0)
    sep_val = "\t" if sep == "TAB" else sep
    encoding = st.selectbox("Encoding", options=["utf-8", "latin1", "cp1252"], index=0)

    preview_rows = st.slider("Filas de previsualización", 5, 50, 10)

st.divider()

# --- Guardado + preview ---
if csv_file and reglas_file:
    csv_path = save_upload(csv_file, workdir / "llegadas.csv")
    reglas_path = save_upload(reglas_file, workdir / "Reglas_hospitales.xlsx")

    st.subheader("Previsualización")
    try:
        df_prev = pd.read_csv(csv_path, sep=sep_val, encoding=encoding)
        st.dataframe(df_prev.head(preview_rows), use_container_width=True)
        st.caption(f"Columnas detectadas: {list(df_prev.columns)}")
    except Exception as e:
        st.error("No he podido leer el CSV con ese separador/encoding.")
        st.exception(e)

    st.divider()

    # --- Ejecución ---
    st.subheader("3) Ejecutar")
    run_reparto = st.button("Generar salida.xlsx", type="primary")

    if run_reparto:
        st.info("Ejecutando scripts…")

        # Ajusta aquí tus scripts y argumentos.
        # Importante: sys.executable para que en cloud apunte al python correcto.
        cmd = [
            sys.executable, "reparto_gpt.py",
            "--csv", str(csv_path.name),
            "--reglas", str(reglas_path.name),
            "--out", "salida.xlsx",
        ]

        rc, out, err = run_cmd(cmd, cwd=workdir)

        if rc != 0:
            st.error("Falló reparto_gpt.py")
            show_logs(out, err)
        else:
            st.success("OK: salida.xlsx generada.")
            salida_path = workdir / "salida.xlsx"

            # Vista rápida del Excel
            try:
                df_out = pd.read_excel(salida_path)
                st.dataframe(df_out.head(preview_rows), use_container_width=True)
            except Exception:
                st.warning("No he podido previsualizar el Excel, pero el archivo existe.")

            # Botón descarga
            st.download_button(
                "Descargar salida.xlsx",
                data=salida_path.read_bytes(),
                file_name="salida.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

else:
    st.warning("Sube el CSV y el Excel de reglas para habilitar la ejecución.")

st.divider()
st.caption("Si algo falla, quiero ver el STDERR aquí mismo. Eso es lo que evita perder horas.")
