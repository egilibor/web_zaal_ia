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
    st.write(f"GPT: `{SCRIPT_REPARTO.name}` exists = `{SCRIPT_REPARTO.exists()}`")
    st.write(f"Gemini: `{SCRIPT_GEMINI.name}` exists = `{SCRIPT_GEMINI.exists()}`")
    st.write(f"Reglas: `{REGLAS_REPO.name}` exists = `{REGLAS_REPO.exists()}`")

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
# Verificaciones duras
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

# Copiar reglas del repo al workdir (para que el script las encuentre fácil)
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
# Paso 3: reparto_gpt.py
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
    st.write("CMD GPT:", cmd)
    st.info("Ejecutando reparto_gpt.py…")

    rc, out, err = run_process(cmd, cwd=workdir, timeout_s=300)

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

    st.download_button(
        "Descargar salida.xlsx",
        data=salida_path.read_bytes(),
        file_name="salida.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

st.divider()

# -------------------------
# Paso 4: reparto_gemini.py (sin input, con --seleccion)
# -------------------------
st.subheader("4) Ejecutar reparto_gemini.py (genera PLAN.xlsx)")
st.caption('Requiere que tu reparto_gemini.py acepte: --seleccion, --in, --out (sin input()).')

salida_path = workdir / "salida.xlsx"
if not salida_path.exists():
    st.warning("Primero genera `salida.xlsx` en el paso 3.")
    st.stop()

# Listado de hojas para ayudar al usuario (sin leer todas)
try:
    xl = pd.ExcelFile(salida_path)
    hojas = xl.sheet_names
    st.write("Hojas detectadas en salida.xlsx:")
    st.write(hojas)
except Exception:
    st.warning("No he podido listar hojas de salida.xlsx, pero puedes ejecutar igualmente.")

seleccion = st.text_input('Selección (ej: "0,1,3-5" o "all")', value="0")

if st.button("Generar PLAN.xlsx"):
    cmd2 = [
        sys.executable,
        str(SCRIPT_GEMINI),
        "--seleccion", seleccion.strip(),
        "--in", "salida.xlsx",
        "--out", "PLAN.xlsx",
    ]
    st.write("CMD GEMINI:", cmd2)
    st.info("Ejecutando reparto_gemini.py…")

    rc2, out2, err2 = run_process(cmd2, cwd=workdir, timeout_s=300)

    if rc2 != 0:
        st.error("❌ Falló reparto_gemini.py")
        show_logs(out2, err2)
        st.stop()

    plan_path = workdir / "PLAN.xlsx"
    if not plan_path.exists():
        st.error("Terminó sin error, pero no encuentro `PLAN.xlsx` en el workdir.")
        show_logs(out2, err2)
        st.stop()

    st.success("✅ PLAN.xlsx generado")

    st.download_button(
        "Descargar PLAN.xlsx",
        data=plan_path.read_bytes(),
        file_name="PLAN.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
