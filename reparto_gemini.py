# reparto_gemini.py — versión SIN input() para Streamlit
# Uso:
#   python reparto_gemini.py --seleccion "0,1,3-5" --in salida.xlsx --out PLAN.xlsx
#   python reparto_gemini.py --seleccion "all" --in salida.xlsx --out PLAN.xlsx
#
# Mantiene el criterio de exclusión de hojas: VINAROZ, MORELLA, RESUMEN.
# No inventa datos ni hace optimización: solo compone un PLAN con las hojas seleccionadas.

import argparse
from pathlib import Path
from datetime import datetime

import pandas as pd

EXCLUDE_TOKENS = ("VINAROZ", "MORELLA", "RESUMEN")


def eligible_sheets(xlsx_path: Path) -> list[str]:
    xl = pd.ExcelFile(xlsx_path)
    out = []
    for name in xl.sheet_names:
        up = name.upper()
        if any(tok in up for tok in EXCLUDE_TOKENS):
            continue
        out.append(name)
    return out


def parse_selection(sel: str, n: int) -> list[int]:
    sel = (sel or "").strip().lower()
    if sel in ("all", "*"):
        return list(range(n))
    if not sel:
        raise ValueError('Selección vacía. Usa --seleccion "0" o --seleccion "all".')

    out: set[int] = set()
    for part in [p.strip() for p in sel.split(",") if p.strip()]:
        if "-" in part:
            a, b = part.split("-", 1)
            a, b = a.strip(), b.strip()
            if not a.isdigit() or not b.isdigit():
                raise ValueError(f"Rango inválido: {part}")
            ia, ib = int(a), int(b)
            if ia > ib:
                ia, ib = ib, ia
            for k in range(ia, ib + 1):
                out.add(k)
        else:
            if not part.isdigit():
                raise ValueError(f"Índice inválido: {part}")
            out.add(int(part))

    bad = [i for i in sorted(out) if i < 0 or i >= n]
    if bad:
        raise ValueError(f"Índices fuera de rango: {bad}. Rango válido: 0..{n-1}")

    return sorted(out)


def default_out_name() -> str:
    return f"PLAN_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"


def main():
    ap = argparse.ArgumentParser(description="Genera PLAN desde salida.xlsx sin interacción (sin input).")
    ap.add_argument("--seleccion", required=True, help='Ej: "0,1,3-5" o "all"')
    ap.add_argument("--in", dest="in_path", default="salida.xlsx", help="Excel de entrada (por defecto: salida.xlsx)")
    ap.add_argument("--out", dest="out_path", default="", help="Excel de salida (por defecto: PLAN_YYYYMMDD_HHMMSS.xlsx)")
    args = ap.parse_args()

    in_path = Path(args.in_path)
    if not in_path.exists():
        raise FileNotFoundError(f"No existe el Excel de entrada: {in_path.resolve()}")

    sheets = eligible_sheets(in_path)
    if not sheets:
        raise RuntimeError("No hay hojas elegibles (todas excluidas por VINAROZ/MORELLA/RESUMEN).")

    idxs = parse_selection(args.seleccion, len(sheets))
    selected = [sheets[i] for i in idxs]

    out_name = args.out_path.strip() or default_out_name()
    out_path = Path(out_name)

    # Construcción 1:1: copiar hojas seleccionadas al PLAN
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        resumen = []
        for sh in selected:
            df = pd.read_excel(in_path, sheet_name=sh)
            df.to_excel(writer, sheet_name=sh[:31], index=False)
            resumen.append({"Hoja": sh, "Filas": int(df.shape[0]), "Columnas": int(df.shape[1])})

        pd.DataFrame(resumen).to_excel(writer, sheet_name="RESUMEN_PLAN", index=False)

    print(f"OK: {out_path.name}")
    print("SELECCION:", args.seleccion)
    print("HOJAS:", selected)


if __name__ == "__main__":
    main()
