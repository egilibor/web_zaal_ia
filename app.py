# reparto_gemini.py (sin input) — v2
# Uso:
#   python reparto_gemini.py --seleccion "0,1,3-5"
#   python reparto_gemini.py --seleccion "all"
# Opcional:
#   --in salida.xlsx
#   --out PLAN.xlsx
#
# Nota: sigue usando el mismo criterio de exclusión de hojas que tu versión:
# excluye hojas que contengan VINAROZ, MORELLA o RESUMEN.

import argparse
import re
from pathlib import Path
from datetime import datetime

import pandas as pd


EXCLUDE_TOKENS = ("VINAROZ", "MORELLA", "RESUMEN")


def parse_selection(sel: str, n: int) -> list[int]:
    """
    Convierte "0,1,3-5" en [0,1,3,4,5].
    Acepta "all".
    Valida rango 0..n-1.
    """
    sel = (sel or "").strip().lower()
    if sel in ("all", "*"):
        return list(range(n))

    if not sel:
        raise ValueError("Selección vacía. Usa --seleccion \"0\" o --seleccion \"all\".")

    out: set[int] = set()
    parts = [p.strip() for p in sel.split(",") if p.strip()]
    for p in parts:
        if "-" in p:
            a, b = p.split("-", 1)
            a = a.strip()
            b = b.strip()
            if not a.isdigit() or not b.isdigit():
                raise ValueError(f"Rango inválido: '{p}'")
            ia = int(a)
            ib = int(b)
            if ia > ib:
                ia, ib = ib, ia
            for k in range(ia, ib + 1):
                out.add(k)
        else:
            if not p.isdigit():
                raise ValueError(f"Índice inválido: '{p}'")
            out.add(int(p))

    bad = [i for i in sorted(out) if i < 0 or i >= n]
    if bad:
        raise ValueError(f"Índices fuera de rango: {bad}. Rango válido: 0..{n-1}")

    return sorted(out)


def eligible_sheets(xlsx_path: Path) -> list[str]:
    xl = pd.ExcelFile(xlsx_path)
    sheets = []
    for name in xl.sheet_names:
        up = name.upper()
        if any(tok in up for tok in EXCLUDE_TOKENS):
            continue
        sheets.append(name)
    return sheets


def plan_from_sheets(xlsx_path: Path, sheet_names: list[str]) -> dict[str, pd.DataFrame]:
    """
    Lee las hojas seleccionadas y las devuelve como dict hoja->DataFrame.
    Aquí NO invento ninguna lógica adicional.
    Si tu script original hacía transformaciones, dímelo y las replico 1:1.
    """
    out: dict[str, pd.DataFrame] = {}
    for sh in sheet_names:
        out[sh] = pd.read_excel(xlsx_path, sheet_name=sh)
    return out


def default_out_name() -> str:
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return f"PLAN_{stamp}.xlsx"


def main():
    ap = argparse.ArgumentParser(description="Genera plan desde salida.xlsx sin interacción (sin input).")
    ap.add_argument("--seleccion", required=True, help='Ej: "0,1,3-5" o "all"')
    ap.add_argument("--in", dest="in_path", default="salida.xlsx", help="Excel de entrada. Por defecto: salida.xlsx")
    ap.add_argument("--out", dest="out_path", default="", help="Excel de salida. Por defecto: PLAN_YYYYMMDD_HHMMSS.xlsx")
    args = ap.parse_args()

    xlsx_path = Path(args.in_path).resolve()
    if not xlsx_path.exists():
        raise FileNotFoundError(f"No existe el Excel de entrada: {xlsx_path}")

    sheets = eligible_sheets(xlsx_path)
    if not sheets:
        raise RuntimeError("No hay hojas elegibles (todas excluidas por VINAROZ/MORELLA/RESUMEN).")

    idxs = parse_selection(args.seleccion, len(sheets))
    selected = [sheets[i] for i in idxs]

    # Construye el plan (lectura 1:1)
    data = plan_from_sheets(xlsx_path, selected)

    out_name = args.out_path.strip() or default_out_name()
    out_path = Path(out_name).resolve()

    # Escribe un Excel con una pestaña por hoja seleccionada, más una pestaña RESUMEN simple
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        resumen_rows = []
        for sh, df in data.items():
            df.to_excel(writer, sheet_name=sh[:31], index=False)
            resumen_rows.append({"Hoja": sh, "Filas": int(df.shape[0]), "Columnas": int(df.shape[1])})

        pd.DataFrame(resumen_rows).to_excel(writer, sheet_name="RESUMEN_PLAN", index=False)

    print(f"OK: {out_path.name}")
    print("SELECCION:", args.seleccion)
    print("HOJAS:", selected)


if __name__ == "__main__":
    main()
