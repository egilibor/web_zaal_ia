#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from __future__ import annotations

import argparse
import os
import datetime as _dt
import re
import math
from pathlib import Path
from typing import List, Tuple

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows


WORKDIR = Path(r"C:\REPARTO")
DEFAULT_CSV = str(WORKDIR / "llegadas.csv")
DEFAULT_REGLAS = str(WORKDIR / "Reglas_hospitales.xlsx")
DEFAULT_OUT = str(WORKDIR / "salida.xlsx")

ORIGEN_LAT = 39.804106
ORIGEN_LON = -0.217351
COORD_FILE = Path(__file__).parent / "Libro_de_Servicio_Castellon_con_coordenadas.xlsx"


# -------------------------------------------------
# UTILIDADES
# -------------------------------------------------

def clean_text(x) -> str:
    if pd.isna(x):
        return ""
    s = str(x).strip()
    s = re.sub(r"\s+", " ", s)
    return s


def parse_kg(x) -> float:
    if pd.isna(x):
        return 0.0
    s = str(x).strip()
    if s == "":
        return 0.0
    if re.search(r"\d+,\d+", s):
        s = s.replace(".", "").replace(",", ".")
    else:
        s = s.replace(",", ".")
    s = re.sub(r"[^0-9\.\-]", "", s)
    try:
        return float(s) if s != "" else 0.0
    except Exception:
        return 0.0


def parse_int(x) -> int:
    if pd.isna(x):
        return 0
    s = re.sub(r"[^0-9\-]", "", str(x))
    try:
        return int(s) if s != "" else 0
    except Exception:
        return 0


def norm(s: str) -> str:
    s = clean_text(s).upper()
    trans = str.maketrans({
        "Á":"A","É":"E","Í":"I","Ó":"O","Ú":"U","Ü":"U","Ñ":"N","Ç":"C",
        "À":"A","È":"E","Ì":"I","Ò":"O","Ù":"U","Ä":"A","Ë":"E","Ï":"I","Ö":"O",
        "Â":"A","Ê":"E","Î":"I","Ô":"O","Û":"U"
    })
    s = s.translate(trans)
    s = re.sub(r"[^\w\s]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def unique_join(values: List[str], sep: str = " / ") -> str:
    seen = set()
    out = []
    for v in values:
        v = clean_text(v)
        if not v or v in seen:
            continue
        seen.add(v)
        out.append(v)
    return sep.join(out)


# -------------------------------------------------
# EXCEL STYLE
# -------------------------------------------------

def style_sheet(ws):
    thin = Side(style="thin", color="D0D0D0")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    header_font = Font(bold=True)
    header_fill = PatternFill("solid", fgColor="F2F2F2")
    wrap = Alignment(wrap_text=False, vertical="center")
    top = Alignment(vertical="top")

    if ws.max_row >= 1:
        for cell in ws[1]:
            cell.font = header_font
            cell.fill = header_fill
            cell.border = border
            cell.alignment = top

    for r in ws.iter_rows(min_row=2, max_row=ws.max_row,
                          min_col=1, max_col=ws.max_column):
        for c in r:
            c.border = border
            c.alignment = wrap

    ws.freeze_panes = "A2"
    if ws.max_row >= 2:
        ws.auto_filter.ref = f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"


def set_widths(ws, widths: List[int]):
    for i, w in enumerate(widths, start=1):
        if i <= ws.max_column:
            ws.column_dimensions[get_column_letter(i)].width = w


# -------------------------------------------------
# CORE
# -------------------------------------------------

def run(csv_path: Path, reglas_path: Path, out_path: Path, origen: str) -> None:

    df = pd.read_csv(csv_path, sep=";", encoding="utf-8-sig", dtype=str, engine="python")
    df["Kgs"] = df["Kgs"].apply(parse_kg)
    df["Bultos"] = df["Btos."].apply(parse_int)
    df["Población"] = df["Población"].apply(clean_text)
    df["Dirección"] = df["Dir_OK"].apply(clean_text)
    df["Z.Rep"] = df["Z.Rep"].apply(clean_text)
    df["Hospital"] = ""
    df["Cliente"] = df.get("Cliente", "")

    hosp = df[df["Hospital"] != ""].copy()
    fed = df[df["Z.Rep"] == "16"].copy()
    resto = df[(df["Hospital"] == "") & (df["Z.Rep"] != "16")].copy()

    wb_out = Workbook()
    wb_out.remove(wb_out.active)

    COLUMNAS_BASE = [
        "Exp", "Hospital", "Población", "Dirección",
        "Consignatario", "Cliente", "Kgs",
        "Bultos", "Z.Rep"
    ]

    # METADATOS
    meta = pd.DataFrame({
        "Clave": ["Origen de datos", "CSV", "Reglas", "Generado"],
        "Valor": [
            origen,
            str(csv_path),
            str(reglas_path),
            _dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        ],
    })
    ws_meta = wb_out.create_sheet("METADATOS")
    for row in dataframe_to_rows(meta, index=False, header=True):
        ws_meta.append(row)
    style_sheet(ws_meta)

    # HOSPITALES
    ws_h = wb_out.create_sheet("HOSPITALES")
    for row in dataframe_to_rows(hosp[COLUMNAS_BASE], index=False, header=True):
        ws_h.append(row)
    style_sheet(ws_h)

    # FEDERACION
    ws_f = wb_out.create_sheet("FEDERACION")
    for row in dataframe_to_rows(fed[COLUMNAS_BASE], index=False, header=True):
        ws_f.append(row)
    style_sheet(ws_f)

    # ZREP
    for z, sub in resto.groupby("Z.Rep"):
        ws = wb_out.create_sheet(f"ZREP_{z}")
        for row in dataframe_to_rows(sub[COLUMNAS_BASE], index=False, header=True):
            ws.append(row)
        style_sheet(ws)

    # -------------------------------------------------
    # CREAR RESUMEN_UNICO AL FINAL
    # -------------------------------------------------

    operativas = []
    if "HOSPITALES" in wb_out.sheetnames:
        operativas.append("HOSPITALES")
    if "FEDERACION" in wb_out.sheetnames:
        operativas.append("FEDERACION")

    zrep_sheets = sorted([s for s in wb_out.sheetnames if s.startswith("ZREP_")])
    operativas.extend(zrep_sheets)

    ws_res = wb_out.create_sheet("RESUMEN_UNICO")
    ws_res.append(["Clave", "Expediciones", "Bultos", "Kilos"])

    for hoja in operativas:
        ws_res.append([
            hoja,
            f"=COUNTA('{hoja}'!A:A)-1",
            f"=SUM('{hoja}'!H:H)",
            f"=SUM('{hoja}'!G:G)"
        ])

    style_sheet(ws_res)
    set_widths(ws_res, [20, 15, 15, 15])

    # GUARDAR UNA SOLA VEZ
    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb_out.save(out_path)


def main():
    try:
        if WORKDIR.exists():
            os.chdir(WORKDIR)
    except Exception:
        pass

    parser = argparse.ArgumentParser()
    parser.add_argument("--csv")
    parser.add_argument("--reglas")
    parser.add_argument("--out")
    args = parser.parse_args()

    csv_p = Path(args.csv) if args.csv else Path(DEFAULT_CSV)
    reglas_p = Path(args.reglas) if args.reglas else Path(DEFAULT_REGLAS)
    out_p = Path(args.out) if args.out else Path(DEFAULT_OUT)

    origen = "LLEGADAS"
    run(csv_p, reglas_p, out_p, origen)
    print(f"OK: generado {out_p}")


if __name__ == "__main__":
    main()
