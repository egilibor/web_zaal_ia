#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from __future__ import annotations
import argparse
import re
from pathlib import Path
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils.dataframe import dataframe_to_rows


# =========================================================
# UTILIDADES
# =========================================================

def clean_text(x):
    if pd.isna(x):
        return ""
    return re.sub(r"\s+", " ", str(x).strip())


def parse_kg(x):
    if pd.isna(x) or str(x).strip() == "":
        return 0.0
    s = str(x).replace(",", ".")
    s = re.sub(r"[^0-9\.\-]", "", s)
    try:
        return float(s)
    except:
        return 0.0


def parse_int(x):
    if pd.isna(x):
        return 0
    s = re.sub(r"[^0-9\-]", "", str(x))
    try:
        return int(s)
    except:
        return 0


def norm(s):
    s = clean_text(s).upper()
    s = re.sub(r"[^\w\s]", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def safe_sheet_name(name, existing):
    name = re.sub(r"[\\/*?:\[\]]", " ", str(name))
    name = re.sub(r"\s+", " ", name).strip()
    base = name[:31]
    candidate = base
    i = 2
    while candidate in existing:
        suffix = f" {i}"
        candidate = (base[:31-len(suffix)] + suffix).strip()
        i += 1
    existing.add(candidate)
    return candidate


# =========================================================
# EXCEL HELPERS
# =========================================================

def style_sheet(ws):
    thin = Side(style="thin")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    header_font = Font(bold=True)
    header_fill = PatternFill("solid", fgColor="F2F2F2")

    for cell in ws[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.border = border

    for r in ws.iter_rows(min_row=2):
        for c in r:
            c.border = border

    ws.freeze_panes = "A2"


def add_df_sheet(wb, name, df):
    ws = wb.create_sheet(title=name)
    df = df.where(pd.notna(df), None)
    for row in dataframe_to_rows(df, index=False, header=True):
        ws.append(row)
    style_sheet(ws)


# =========================================================
# CARGA CSV
# =========================================================

def load_csv(csv_path):
    df_raw = pd.read_csv(csv_path, sep=";", dtype=str, encoding="utf-8-sig")

    df = pd.DataFrame({
        "Exp": df_raw["Exp"].apply(clean_text),
        "Kgs": df_raw["Kgs"].apply(parse_kg),
        "Bultos": df_raw.get("Btos.", 0),
        "Consignatario": df_raw["Consignatario"].apply(clean_text),
        "Cliente": df_raw.get("Cliente", ""),
        "Población": df_raw["Población"].apply(clean_text),
        "Dirección": df_raw.get("Dir_OK", "").apply(clean_text),
        "Z.Rep": df_raw.get("Z.Rep", "SIN_ZONA"),
        "N_servicio": df_raw.get("N. servicio", "")
    })

    df["Parada_key"] = df["Población"] + "||" + df["Dirección"]
    df["Pob_norm"] = df["Población"].apply(norm)
    df["Dir_norm"] = df["Dirección"].apply(norm)

    return df


# =========================================================
# CORE
# =========================================================

def run(csv_path, reglas_path, out_path, origen):

    df = load_csv(csv_path)

    wb_rules = load_workbook(reglas_path, data_only=True)

    rules_h = pd.DataFrame(wb_rules["REGLAS_HOSPITALES"].values)
    rules_h.columns = rules_h.iloc[0]
    rules_h = rules_h[1:]

    rules_f = pd.DataFrame(wb_rules["REGLAS_FEDERACION"].values)
    rules_f.columns = rules_f.iloc[0]
    rules_f = rules_f[1:]

    df["is_hospital"] = False
    df["Hospital"] = ""

    for i, r in df.iterrows():
        for _, rule in rules_h.iterrows():
            if norm(rule["Población"]) in r["Pob_norm"] and norm(rule["Patrón_dirección"]) in r["Dir_norm"]:
                df.at[i, "is_hospital"] = True
                df.at[i, "Hospital"] = rule.get("Hospital_final", "")
                break

    df["is_fed"] = False
    for i, r in df.iterrows():
        for _, rule in rules_f.iterrows():
            if norm(rule["Población"]) in r["Pob_norm"] and norm(rule["Patrón_dirección"]) in r["Dir_norm"]:
                df.at[i, "is_fed"] = True
                break

    df["is_any_special"] = df["is_hospital"] | df["is_fed"]

    hosp = df[df["is_hospital"]].copy()
    fed = df[df["is_fed"]].copy()
    resto = df[~df["is_any_special"]].copy()

    resto_summary = (
        resto.groupby("Z.Rep")
        .agg(
            Paradas=("Parada_key", "nunique"),
            Expediciones=("Exp", "nunique"),
            Bultos=("Bultos", "sum"),
            Kilos=("Kgs", "sum"),
        )
        .reset_index()
    )

    overview = pd.DataFrame({
        "Bloque": ["HOSPITALES", "FEDERACION", "RESTO (todas rutas)"],
        "Paradas": [
            hosp["Parada_key"].nunique(),
            fed["Parada_key"].nunique(),
            resto["Parada_key"].nunique()
        ],
        "Expediciones": [
            hosp["Exp"].nunique(),
            fed["Exp"].nunique(),
            resto["Exp"].nunique()
        ],
        "Kilos": [
            round(hosp["Kgs"].sum(), 1),
            round(fed["Kgs"].sum(), 1),
            round(resto["Kgs"].sum(), 1)
        ]
    })

    resumen_unico_general = overview.copy()
    resumen_unico_general.insert(0, "Tipo", "GENERAL")
    resumen_unico_general.rename(columns={"Bloque": "Clave"}, inplace=True)

    resumen_unico_resto = resto_summary.copy()
    resumen_unico_resto.insert(0, "Tipo", "RESTO")
    resumen_unico_resto.rename(columns={"Z.Rep": "Clave"}, inplace=True)

    resumen_unico = pd.concat([resumen_unico_general, resumen_unico_resto], ignore_index=True)

    wb_out = Workbook()
    wb_out.remove(wb_out.active)

    add_df_sheet(wb_out, "RESUMEN_GENERAL", overview)
    add_df_sheet(wb_out, "RESUMEN_UNICO", resumen_unico)
    add_df_sheet(wb_out, "HOSPITALES", hosp)
    add_df_sheet(wb_out, "FEDERACION", fed)

    existing = set(wb_out.sheetnames)

    for z, sub in resto.groupby("Z.Rep"):
        sheet_name = safe_sheet_name(f"ZREP_{z}", existing)
        add_df_sheet(wb_out, sheet_name, sub)

    wb_out.save(out_path)


# =========================================================
# MAIN
# =========================================================

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--csv", required=True)
    parser.add_argument("--reglas", required=True)
    parser.add_argument("--out", required=True)
    args = parser.parse_args()

    run(Path(args.csv), Path(args.reglas), Path(args.out), "LLEGADAS")


if __name__ == "__main__":
    main()
