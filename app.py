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

# --- Paths en repo ---
REPO_DIR = Path(__file__).resolve().parent
SCRIPT_REPARTO = REPO_DIR / "reparto_gpt.py"
SCRIPT_GEMINI = REPO_DIR / "reparto_gemini.py"
REGLAS_REPO = REPO_DIR / "Reglas_hospitales.xlsx"

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


def run_cmd_input(cmd: list[str], cwd: Path, stdin_text: str) -> tuple[int, str, str]:
    p = subprocess.run(
        cmd,
        cwd=str(cwd),
        input=stdin_text,
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


def list_plan_files(workdir: Path) -> list[Path]:
    return sorted(workdir.glob("PLAN_*.xlsx"), key=lambda x: x.stat().st_mtime, reverse=True)


# -------------------------
# Estado
# -------------------------
workdir = ensure_workdir()

with st.sidebar:
    st.header("Estado")
    st.write(f"Run: `{st.session_state.run_id}`")
    st.write(f"Workdir: `{workdir}`")
    st.write(f"Repo dir: `{REPO_DIR}`")
    st.write(f"Python: `{sys.executable}`")

    st.divider()
    st.write(f"Script GPT: `{SCRIPT_REPARTO}`")
    st.write(f"GPT exists: `{SCRIPT_REPARTO.exists()}`")
    st.write(f"Script Gemini: `{SCRIPT_GEMINI}`")
    st.write(f"Gemini exists: `{SCRIPT_GEMINI.exists()}`")
    st.write(f"Reglas: `{REGLAS_REPO}`")
    st.write(f"Reglas exists: `{REGLAS_REPO.exists()}`")

    st.divider()
    try:
        st.write("Repo files:", sorted([p.name for p in REPO_DIR.iterdir()]))
    except Exception as e:
        st.write("Repo files: (error)")
        st.write(str(e))

    if st.button("Reset sesión"):
        reset_session_dir()
        st.rerun()


# -------------------------
# Verificaciones
# -------------------------
missing = []
if not SCRIPT_REPARTO.exists():
    missing.append("reparto_gpt.py")
if not SCRIPT_GEMINI.exists():
    missing.append("reparto_gemini.py")
if not REGLAS_REPO.exists():
    missing.append("Reglas_hospitales.xlsx")

if missing:
    st.error(
        "Faltan archivos en el repo desplegado: " + ", ".join(missing) + "\n\n"
        "Revisa que estén en el branch desplegado (main) y en la misma carpeta que app.py."
    )
    st.stop()

st.divider()

# -------------------------
# Inputs
# -------------------------
st.subheader("1) Subir CSV de llegadas")
csv_file = st.file_uploader("CSV de llegadas", type=["csv"])

st.subheader("2) Previsualización (opcional, no afecta a ejecución)")
col1, col2 = st.columns(2, gap="large")
with col1:
    sep = st.selectbox("Separador CSV (solo para vista previa)", options=[";", ",", "TAB"], index=0)
    sep_val = "\t" if sep == "TAB" else sep
with col2:
    encoding = st.selectbox("Encoding (solo para vista previa)", options=["utf-8", "latin1", "cp1252"], index=0)

preview_rows = st.slider("Filas de previsualización", 5, 50, 10)

st.divider()

if not csv_file:
    st.info("Sube el CSV para habilitar la ejecución.")
    st.stop()

# Guardar CSV en workdir
csv_path = save_upload(csv_file, workdir / "llegadas.csv")

# Copiar reglas del repo al workdir
(workdir / "Reglas_hospitales.xlsx").write_bytes(REGLAS_REPO.read_bytes())

# Preview CSV (solo visual)
try:
    df_prev = pd.read_csv(csv_path, sep=sep_val, encoding=encoding)
    st.dataframe(df_prev.head(preview_rows), use_container_width=True)
    st.caption(f"Columnas detectadas: {list(df_prev.columns)}")
except Exception as e:
    st.warning("No he podido previsualizar el CSV con ese separador/encoding (no afecta al script).")
    st.exception(e)

st.divider()

# -------------------------
# Ejecución paso 1: reparto_gpt.py
# -------------------------
st.subheader("3) Ejecutar reparto_gpt.py (genera salida.xlsx)")

if st.button("Generar salida.xlsx", type="primary"):
    cmd = [
        sys.executable,
        str(SCRIPT_REPARTO),
        "--csv", "llegadas.csv",
        "--reglas", "Reglas_hospitales.xlsx",
        "--out", "salida.xlsx",
    ]
    st.write("CMD:", cmd)
    st.info("Ejecutando reparto_gpt.py…")

    rc, out, err = run_cmd(cmd, cwd=workdir)

    if rc != 0:
        st.error("❌ Falló reparto_gpt.py")
        show_logs(out, err)
        st.stop()

    salida_path = workdir / "salida.xlsx"
    if not salida_path.exists():
        st.error("Terminó sin error, pero no encuentro `salida.xlsx` en el workdir.")
        show_logs(out, err)
        st.stop()

    st.success("✅ salida.xlsx generada")

    # Preview salida.xlsx (si se puede)
    try:
        df_out = pd.read_excel(salida_path)
        st.dataframe(df_out.head(preview_rows), use_container_width=True)
    except Exception:
        st.warning("No he podido previsualizar salida.xlsx, pero el archivo existe.")

    st.download_button(
        "Descargar salida.xlsx",
        data=salida_path.read_bytes(),
        file_name="salida.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

st.divider()

# -------------------------
# Ejecución paso 2: reparto_gemini.py
# -------------------------
st.subheader("4) Ejecutar reparto_gemini.py (genera PLAN_*.xlsx)")
st.caption("Este script lee `salida.xlsx` y pide la selección por stdin (ej: 0, 1, 3-5).")

salida_exists = (workdir / "salida.xlsx").exists()
if not salida_exists:
    st.warning("Primero genera `salida.xlsx` con el paso 3.")
else:
    # Mostrar hojas e índices (mismo criterio de exclusión que usa el script)
    try:
        xl = pd.ExcelFile(workdir / "salida.xlsx")
        hojas_disp = [
            h for h in xl.sheet_names
            if not any(x in h.upper() for x in ["VINAROZ", "MORELLA", "RESUMEN"])
        ]
        st.write("Rutas disponibles (índice → hoja):")
        for i, h in enumerate(hojas_disp):
            st.write(f"[{i}] {h}")
    except Exception:
        st.warning("No he podido listar hojas de salida.xlsx, pero puedes intentar ejecutar igualmente.")

    seleccion = st.text_input("Selección (ej: 0, 1, 3-5)", value="0")

    colA, colB = st.columns(2, gap="large")
    with colA:
        run_plan = st.button("Generar PLAN_*.xlsx", type="secondary")
    with colB:
        st.caption("Consejo: usa índices tal como aparecen arriba.")

    if run_plan:
        # Borramos planes anteriores para detectar el nuevo con seguridad
        for f in workdir.glob("PLAN_*.xlsx"):
            try:
                f.unlink()
            except Exception:
                pass

        cmd2 = [sys.executable, str(SCRIPT_GEMINI)]
        st.write("CMD:", cmd2)
        st.info("Ejecutando reparto_gemini.py…")

        rc2, out2, err2 = run_cmd_input(cmd2, cwd=workdir, stdin_text=(seleccion.strip() + "\n"))

        if rc2 != 0:
            st.error("❌ Falló reparto_gemini.py")
            show_logs(out2, err2)
            st.stop()

        planes = list_plan_files(workdir)
        if not planes:
            st.error("Terminó sin error, pero no encuentro PLAN_*.xlsx en el workdir.")
            show_logs(out2, err2)
            st.stop()

        plan_path = planes[0]
        st.success(f"✅ Generado: {plan_path.name}")

        st.download_button(
            f"Descargar {plan_path.name}",
            data=plan_path.read_bytes(),
            file_name=plan_path.name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        # Preview opcional
        try:
            df_plan = pd.read_excel(plan_path)
            st.dataframe(df_plan.head(preview_rows), use_container_width=True)
        except Exception:
            st.warning("No he podido previsualizar el PLAN, pero el archivo está generado.")
