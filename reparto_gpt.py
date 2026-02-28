#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from __future__ import annotations
import argparse
import os
import datetime as _dt
import re
from pathlib import Path
from typing import List, Tuple

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows


# -------------------------
# Utilidades
# -------------------------

def clean_text(x) -> str:
    if pd.isna(x):
        return ""
    return re.sub(r"\s+", " ", str(x).strip())

def parse_kg(x) -> float:
    if pd.isna(x):
        return 0.0
    s = str(x).replace(",", ".")
    s = re.sub(r"[^0-9\.\-]", "", s)
    return float(s) if s else 0.0

def parse_int(x) -> int:
    if pd.isna(x):
        return 0
    s = re.sub(r"[^0-9\-]", "", str(x))
    return int(s) if s else 0

def norm(s: str) -> str:
    s = clean_text(s).upper()
    return re.sub(r"\s+", " ", s)


# -------------------------
# Excel helpers
# -------------------------

def style_sheet(ws):
    thin = Side(style="thin", color="D0D0D0")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    header_font = Font(bold=True)
    header_fill = PatternFill("solid", fgColor="F2F2F2")
    wrap = Alignment(wrap_text=True, vertical="top")

    for cell in ws[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.border = border

    for r in ws.iter_rows(min_row=2):
        for c in r:
            c.border = border
            c.alignment = wrap

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"

def add_df_sheet(wb: Workbook, name: str, df_sheet: pd.DataFrame):
    ws = wb.create_sheet(title=name)
    out = df_sheet.where(pd.notna(df_sheet), None)
    for row in dataframe_to_rows(out, index=False, header=True):
        ws.append(row)
    style_sheet(ws)


# -------------------------
# Core
# -------------------------

def load_csv(csv_path: Path) -> pd.DataFrame:
    df_raw = pd.read_csv(csv_path, sep=";", encoding="utf-8-sig", dtype=str, engine="python")

    df = pd.DataFrame({
        "Exp": df_raw["Exp"].apply(clean_text),
        "Kgs": df_raw["Kgs"].apply(parse_kg),
        "Bultos": df_raw.get("Btos.", 0).apply(parse_int) if "Btos." in df_raw else 0,
        "Consignatario": df_raw["Consignatario"].apply(clean_text),
        "Cliente": df_raw.get("Cliente", "").apply(clean_text) if "Cliente" in df_raw else "",
        "Población": df_raw["Población"].apply(clean_text),
        "Dirección": df_raw.get("Dir_OK", "").apply(clean_text) if "Dir_OK" in df_raw else "",
        "Z.Rep": df_raw.get("Z.Rep", "SIN_ZONA").apply(clean_text),
        "N_servicio": df_raw.get("N. servicio", "").apply(clean_text) if "N. servicio" in df_raw else "",
    })

    return df


def run(csv_path: Path, reglas_path: Path, out_path: Path, origen: str):

    df = load_csv(csv_path)

    wb_rules = load_workbook(reglas_path, data_only=True)
    rules_h = pd.DataFrame(wb_rules["REGLAS_HOSPITALES"].values)
    rules_f = pd.DataFrame(wb_rules["REGLAS_FEDERACION"].values)

    df["is_hospital"] = False
    df["is_fed"] = False

    # --- Separaciones simples (sin agrupación) ---
    hosp = df[df["is_hospital"]].copy()
    fed = df[df["is_fed"]].copy()
    resto = df[~(df["is_hospital"] | df["is_fed"])].copy()

    # --- Resúmenes ---
    resto_summary = (
        resto.groupby("Z.Rep")
        .agg(
            Paradas=("Población", "nunique"),
            Expediciones=("Exp", "nunique"),
            Bultos=("Bultos", "sum"),
            Kilos=("Kgs", "sum"),
        )
        .reset_index()
    )

    overview = pd.DataFrame(
        {
            "Bloque": ["HOSPITALES", "FEDERACION", "RESTO (todas rutas)"],
            "Paradas": [
                hosp["Población"].nunique(),
                fed["Población"].nunique(),
                resto["Población"].nunique(),
            ],
            "Expediciones": [
                hosp["Exp"].nunique(),
                fed["Exp"].nunique(),
                resto["Exp"].nunique(),
            ],
            "Kilos": [
                round(hosp["Kgs"].sum(), 1),
                round(fed["Kgs"].sum(), 1),
                round(resto["Kgs"].sum(), 1),
            ],
        }
    )

    resumen_unico_general = overview[overview["Bloque"] != "RESTO (todas rutas)"].copy()
    resumen_unico_general.insert(0, "Tipo", "GENERAL")
    resumen_unico_general = resumen_unico_general.rename(columns={"Bloque": "Clave"})
    resumen_unico_general["Bultos"] = None
    resumen_unico_general = resumen_unico_general[
        ["Tipo", "Clave", "Paradas", "Expediciones", "Bultos", "Kilos"]
    ]

    resumen_unico_resto = resto_summary.copy()
    resumen_unico_resto.insert(0, "Tipo", "RESTO")
    resumen_unico_resto = resumen_unico_resto.rename(columns={"Z.Rep": "Clave"})
    resumen_unico_resto = resumen_unico_resto[
        ["Tipo", "Clave", "Paradas", "Expediciones", "Bultos", "Kilos"]
    ]

    resumen_unico = pd.concat(
        [resumen_unico_general, resumen_unico_resto],
        ignore_index=True
    )

    # --- Excel ---
    wb_out = Workbook()
    wb_out.remove(wb_out.active)

    add_df_sheet(wb_out, "RESUMEN_GENERAL", overview)
    add_df_sheet(wb_out, "RESUMEN_UNICO", resumen_unico)
    add_df_sheet(wb_out, "HOSPITALES", hosp)
    add_df_sheet(wb_out, "FEDERACION", fed)

    for z, sub in resto.groupby("Z.Rep"):
        add_df_sheet(wb_out, f"ZREP_{z}", sub)

    wb_out.save(out_path)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--csv", required=True)
    parser.add_argument("--reglas", required=True)
    parser.add_argument("--out", required=True)
    args = parser.parse_args()

    run(Path(args.csv), Path(args.reglas), Path(args.out), "LLEGADAS")


if __name__ == "__main__":
    main()
