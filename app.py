import sys
import uuid
import shutil
import tempfile
import subprocess
from pathlib import Path

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
# MENÚ
# -------------------------
opcion = st.selectbox("Menú", ["Asignación reparto"])

st.divider()

# -------------------------
# OPCIÓN: Asignación reparto
# -------------------------
if opcion == "Asignación reparto":
    st.subheader("1) Subir CSV de llegadas")
    csv_file = st.file_uploader("CSV de llegadas", type=["csv"])

    st.divider()

    if not csv_file:
        st.info("Sube el CSV para habilitar la ejecución.")
        st.stop()

    # Guardar CSV en workdir
    csv_path = save_upload(csv_file, workdir / "llegadas.csv")

    # Copiar reglas del repo al workdir (para que el script las encuentre fácil)
    (workdir / "Reglas_hospitales.xlsx").write_bytes(REGLAS_REPO.read_bytes())

    st.subheader("2) Ejecutar (genera salida.xlsx y PLAN.xlsx)")
    st.caption('Gemini se ejecuta automáticamente con --seleccion "1-9".')

    if st.button("Ejecutar", type="primary"):
        # ---- GPT ----
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
            st.error("Terminó sin error, pero no encuentro `salida.xlsx` en el workdir.")
            show_logs(out, err)
            st.stop()

        # ---- GEMINI (selección fija 1-9) ----
        cmd_gemini = [
            sys.executable,
            str(SCRIPT_GEMINI),
            "--seleccion", "1-9",
            "--in", "salida.xlsx",
            "--out", "PLAN.xlsx",
        ]

        with st.spinner('Ejecutando reparto_gemini.py (selección 1-9)…'):
            rc2, out2, err2 = run_process(cmd_gemini, cwd=workdir, timeout_s=300)

        if rc2 != 0:
            st.error("❌ Falló reparto_gemini.py")
            show_logs(out2, err2)
            st.stop()

        plan_path = workdir / "PLAN.xlsx"
        if not plan_path.exists():
            st.error("Terminó sin error, pero no encuentro `PLAN.xlsx` en el workdir.")
            show_logs(out2, err2)
            st.stop()

        st.success("✅ Archivos generados: salida.xlsx y PLAN.xlsx")

        col_a, col_b = st.columns(2, gap="large")
        with col_a:
            st.download_button(
                "Descargar salida.xlsx",
                data=salida_path.read_bytes(),
                file_name="salida.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        with col_b:
            st.download_button(
                "Descargar PLAN.xlsx",
                data=plan_path.read_bytes(),
                file_name="PLAN.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
