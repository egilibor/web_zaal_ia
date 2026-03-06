#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from __future__ import annotations

import argparse
import os
import datetime as _dt
import re
from pathlib import Path
from typing import List

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows


# -------------------------
# UTILIDADES
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
    try:
        return float(s)
    except:
        return 0.0


def parse_int(x) -> int:
    if pd.isna(x):
        return 0
    s = re.sub(r"[^0-9\-]", "", str(x))
    try:
        return int(s)
    except:
        return 0


def norm(s: str) -> str:
    s = clean_text(s).upper()
    trans = str.maketrans({
        "Á":"A","É":"E","Í":"I","Ó":"O","Ú":"U","Ü":"U","Ñ":"N","Ç":"C",
    })
    s = s.translate(trans)
    s = re.sub(r"[^\w\s]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def sheet_to_df(wb, name: str) -> pd.DataFrame:
    if name not in wb.sheetnames:
        return pd.DataFrame()
    ws = wb[name]
    data = list(ws.values)
    if not data:
        return pd.DataFrame()
    headers = list(data[0])
    rows = data[1:]
    return pd.DataFrame(rows, columns=headers)


def prepare_rules(df_rules: pd.DataFrame, pob_col: str, pat_col: str) -> pd.DataFrame:
    if df_rules.empty:
        return df_rules
    d = df_rules.copy()
    d["Pob_norm"] = d[pob_col].astype(str).apply(norm)
    d["Pat_norm"] = d[pat_col].astype(str).apply(norm)
    d = d[d["Pat_norm"].ne("")]
    return d


def match_rules(pob_norm: str, dir_norm: str, rules_df: pd.DataFrame, tag_field: str | None = None):
    for _, r in rules_df.iterrows():
        if r["Pob_norm"] and r["Pob_norm"] not in pob_norm:
            continue
        if r["Pat_norm"] and r["Pat_norm"] in dir_norm:
            if tag_field:
                return True, str(r.get(tag_field, "") or "")
            return True, ""
    return False, ""


def style_sheet(ws):
    thin = Side(style="thin")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    header_font = Font(bold=True)

    for cell in ws[1]:
        cell.font = header_font
        cell.border = border

    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.border = border

    ws.freeze_panes = "A2"


# -------------------------
# CORE
# -------------------------

def run(csv_path: Path, reglas_path: Path, out_path: Path, origen: str, delegacion: str) -> None:

    df = pd.read_csv(
        csv_path,
        sep=";",
        encoding="utf-8-sig",
        dtype=str,
        engine="python",
        on_bad_lines="skip"
    )

    df["Kgs"] = df["Kgs"].apply(parse_kg)
    df["Bultos"] = df["Btos."].apply(parse_int)
    df["Población"] = df["Población"].fillna("")
    df["Dirección"] = df["Dir_OK"].fillna("")
    df["Z.Rep"] = df["Z.Rep"].fillna("")
    df["Cliente"] = df.get("Cliente", "")

    # -------------------------
    # APLICAR REGLAS
    # -------------------------

    df["Pob_norm"] = df["Población"].apply(norm)
    df["Dir_norm"] = df["Dirección"].apply(norm)

    wb_rules = load_workbook(reglas_path, data_only=True)
    rules_h = sheet_to_df(wb_rules, "REGLAS_HOSPITALES")
    rules_f = sheet_to_df(wb_rules, "REGLAS_FEDERACION")

    rules_h_prep = prepare_rules(rules_h, "Población", "Patrón_dirección")
    rules_f_prep = prepare_rules(rules_f, "Población", "Patrón_dirección")

    df["is_hospital"] = False
    df["Hospital"] = ""

    for i, r in df.iterrows():
        ok, tag = match_rules(r["Pob_norm"], r["Dir_norm"], rules_h_prep, "Hospital_final")
        if ok:
            df.at[i, "is_hospital"] = True
            df.at[i, "Hospital"] = tag

    df["is_fed"] = False
    for i, r in df.iterrows():
        ok, _ = match_rules(r["Pob_norm"], r["Dir_norm"], rules_f_prep)
        if ok:
            df.at[i, "is_fed"] = True

    df["is_any_special"] = df["is_hospital"] | df["is_fed"]

    hosp = df[df["is_hospital"]].copy()
    fed = df[df["is_fed"]].copy()
    resto = df[~df["is_any_special"]].copy()
    resto["Z.Rep"] = (
        resto["Z.Rep"]
        .astype(str)
        .str.strip()
        .replace(".", "")
    )
    
    # -------------------------
    # EXCEL
    # -------------------------

    wb_out = Workbook()
    wb_out.remove(wb_out.active)

    COLUMNAS_BASE = [
        "Exp", "Hospital", "Población", "Dirección",
        "Consignatario", "Cliente", "Kgs",
        "Bultos", "Z.Rep", "N. servicio"
    ]

    # METADATOS
    meta = pd.DataFrame({
        "Clave": ["Delegación", "Origen de datos", "CSV", "Reglas", "Generado"],
        "Valor": [
            delegacion,
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
    existing = set(wb_out.sheetnames)

    for z, sub in resto.groupby("Z.Rep"):

        z = str(z).strip()
        if z == ".":
            z = ""

        nombre = f"ZREP_{z}"
        nombre = re.sub(r"[\\/*?:\[\]]", "_", nombre)[:31]

        base = nombre
        i = 1
        while nombre in existing:
            sufijo = f"_{i}"
            nombre = (base[:31 - len(sufijo)] + sufijo)
            i += 1

        existing.add(nombre)

        ws = wb_out.create_sheet(nombre)

        for row in dataframe_to_rows(sub[COLUMNAS_BASE], index=False, header=True):
            ws.append(row)

        style_sheet(ws)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb_out.save(out_path)


# -------------------------
# MAIN
# -------------------------

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--csv")
    parser.add_argument("--reglas")
    parser.add_argument("--out")
    parser.add_argument("--delegacion", default="castellon")

    args = parser.parse_args()

    csv_p = Path(args.csv)
    reglas_p = Path(args.reglas)
    out_p = Path(args.out)

    run(csv_p, reglas_p, out_p, "LLEGADAS", args.delegacion)

    print(f"OK: generado {out_p}")


if __name__ == "__main__":
    main()
