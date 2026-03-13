#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AGENTE CASTELLÓN · SALIDA DIARIA
(HOSPITALES + FEDERACIÓN + RESTO POR Z.REP)

Versión v3

- Pregunta siempre por el origen de datos
- Si faltan rutas entra en modo interactivo
"""

from __future__ import annotations

import argparse
import os
import datetime as _dt
import re
import math
from pathlib import Path
from typing import List, Tuple

import pandas as pd
import difflib

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows


WORKDIR = Path(r"C:\REPARTO")
DEFAULT_CSV = str(WORKDIR / "llegadas.csv")
DEFAULT_REGLAS = str(WORKDIR / "Reglas_hospitales.xlsx")
DEFAULT_OUT = str(WORKDIR / "salida.xlsx")


# --------------------------------------------------
# Limpieza Excel
# --------------------------------------------------

def clean_excel_text(s):
    if pd.isna(s):
        return ""
    s = str(s)
    return re.sub(r"[\x00-\x1F\x7F]", "", s)


# --------------------------------------------------
# Callejero Castellón
# --------------------------------------------------

CALLES_CASTELLON = []

try:
    calles_path = Path(__file__).parent / "calles_castellon.csv"
    if calles_path.exists():
        df_calles = pd.read_csv(calles_path)
        if "nombre" in df_calles.columns:
            CALLES_CASTELLON = (
                df_calles["nombre"]
                .astype(str)
                .str.upper()
                .unique()
                .tolist()
            )
except Exception:
    CALLES_CASTELLON = []


def corregir_calle_castellon(poblacion: str, direccion: str) -> str:

    if not direccion:
        return direccion

    pob = norm(poblacion)

    if "CASTELL" not in pob:
        return direccion

    if not CALLES_CASTELLON:
        return direccion

    dir_norm = norm(direccion)

    match = difflib.get_close_matches(
        dir_norm,
        CALLES_CASTELLON,
        n=1,
        cutoff=0.80
    )

    if match:
        calle = match[0]
        return f"{calle} {direccion}"

    return direccion


# --------------------------------------------------
# Utilidades
# --------------------------------------------------

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
        return float(s) if s else 0.0
    except:
        return 0.0


def parse_int(x) -> int:

    if pd.isna(x):
        return 0

    s = re.sub(r"[^0-9\-]", "", str(x))

    try:
        return int(s) if s else 0
    except:
        return 0


def norm(s: str) -> str:

    s = clean_text(s).upper()

    trans = str.maketrans({
        "Á":"A","É":"E","Í":"I","Ó":"O","Ú":"U","Ü":"U","Ñ":"N","Ç":"C",
        "À":"A","È":"E","Ì":"I","Ò":"O","Ù":"U"
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


def pick_col(columns: List[str], candidates: List[str]) -> str | None:

    s = set(columns)

    for c in candidates:
        if c in s:
            return c

    return None


# --------------------------------------------------
# Excel helpers
# --------------------------------------------------

def style_sheet(ws):

    thin = Side(style="thin", color="D0D0D0")

    border = Border(
        left=thin,
        right=thin,
        top=thin,
        bottom=thin
    )

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

    for r in ws.iter_rows(
        min_row=2,
        max_row=ws.max_row,
        min_col=1,
        max_col=ws.max_column
    ):
        for c in r:
            c.border = border
            c.alignment = wrap

    ws.freeze_panes = "A2"


def set_widths(ws, widths: List[int]):

    for i, w in enumerate(widths, start=1):

        if i <= ws.max_column:
            ws.column_dimensions[get_column_letter(i)].width = w


def add_df_sheet(wb, name, df_sheet, widths):

    ws = wb.create_sheet(title=name)

    out = df_sheet.copy()

    for row in dataframe_to_rows(out, index=False, header=True):
        ws.append(row)

    style_sheet(ws)
    set_widths(ws, widths)


# --------------------------------------------------
# Core
# --------------------------------------------------

def load_csv(csv_path: Path) -> pd.DataFrame:

    df_raw = pd.read_csv(csv_path, sep=";", encoding="utf-8-sig", dtype=str, engine="python")

    cols = list(df_raw.columns)

    col_exp = pick_col(cols, ["Exp"])
    col_kg = pick_col(cols, ["Kgs"])
    col_pop = pick_col(cols, ["Población", "Pob_OK"])
    col_cons = pick_col(cols, ["Consignatario", "Cliente", "Nombre"])
    col_cli = pick_col(cols, ["Cliente"])
    col_dir_ok = pick_col(cols, ["Dir_OK"])
    col_dir_ent = pick_col(cols, ["Dir. entrega"])
    col_zrep = pick_col(cols, ["Z.Rep"])
    col_serv = pick_col(cols, ["N. servicio"])
    col_btos = pick_col(cols, ["Btos."])

    df = pd.DataFrame({

        "Exp": df_raw[col_exp].apply(clean_text),
        "Kgs": df_raw[col_kg].apply(parse_kg),
        "Bultos": df_raw[col_btos].apply(parse_int) if col_btos else 0,
        "Consignatario": df_raw[col_cons].apply(clean_text),
        "Cliente": df_raw[col_cli].apply(clean_text) if col_cli else "",
        "Población": df_raw[col_pop].apply(clean_text),
        "Dirección": (df_raw[col_dir_ok].apply(clean_text) if col_dir_ok else ""),
        "Z.Rep": df_raw[col_zrep].apply(clean_text) if col_zrep else "SIN_ZONA",
        "N_servicio": df_raw[col_serv].apply(clean_text) if col_serv else "",
    })

    if col_dir_ent:

        fb = df_raw[col_dir_ent].apply(clean_text)

        df.loc[df["Dirección"].eq(""), "Dirección"] = fb[df["Dirección"].eq("")]

    df["Consignatario"] = df["Consignatario"].apply(clean_excel_text)
    df["Dirección"] = df["Dirección"].apply(clean_excel_text)
    df["Cliente"] = df["Cliente"].apply(clean_excel_text)

    df["Dirección"] = df.apply(
        lambda r: corregir_calle_castellon(r["Población"], r["Dirección"]),
        axis=1
    )

    df["Dirección"] = df["Dirección"].str.upper()

    df["Parada_key"] = (df["Población"] + "||" + df["Dirección"]).str.strip("|")

    df["Pob_norm"] = df["Población"].apply(norm)
    df["Dir_norm"] = df["Dirección"].apply(norm)

    return df


def run(csv_path: Path, reglas_path: Path, out_path: Path, origen: str):

    df = load_csv(csv_path)

    hosp = df[df["Dir_norm"].str.contains("HOSPITAL", na=False)]
    fed = df[df["Dir_norm"].str.contains("FEDERACION", na=False)]
    resto = df[~df.index.isin(hosp.index) & ~df.index.isin(fed.index)]

    resto_grp = (
        resto.groupby(["Z.Rep", "Parada_key"])
        .agg(
            Población=("Población","first"),
            Dirección=("Dirección","first"),
            Expediciones=("Exp","nunique"),
            Bultos=("Bultos","sum"),
            Kilos=("Kgs","sum"),
        )
        .reset_index()
    )

    resto_summary = (
        resto_grp.groupby("Z.Rep")
        .agg(
            Paradas=("Parada_key","nunique"),
            Expediciones=("Expediciones","sum"),
            Bultos=("Bultos","sum"),
            Kilos=("Kilos","sum"),
        )
        .reset_index()
    )

    overview = pd.DataFrame({
        "Bloque":["HOSPITALES","FEDERACION","RESTO"],
        "Paradas":[
            hosp["Parada_key"].nunique(),
            fed["Parada_key"].nunique(),
            resto_grp["Parada_key"].nunique()
        ],
        "Expediciones":[
            hosp["Exp"].nunique(),
            fed["Exp"].nunique(),
            resto["Exp"].nunique()
        ],
        "Bultos":[
            hosp["Bultos"].sum(),
            fed["Bultos"].sum(),
            resto["Bultos"].sum()
        ],
        "Kilos":[
            hosp["Kgs"].sum(),
            fed["Kgs"].sum(),
            resto["Kgs"].sum()
        ]
    })

    resumen_unico_general = overview.copy()
    resumen_unico_general.insert(0,"Tipo","GENERAL")
    resumen_unico_general = resumen_unico_general.rename(columns={"Bloque":"Clave"})

    resumen_unico_resto = resto_summary.copy()
    resumen_unico_resto.insert(0,"Tipo","RESTO")
    resumen_unico_resto = resumen_unico_resto.rename(columns={"Z.Rep":"Clave"})

    resumen_unico = pd.concat(
        [resumen_unico_general,resumen_unico_resto],
        ignore_index=True
    )

    wb = Workbook()
    wb.remove(wb.active)

    add_df_sheet(wb,"RESUMEN_GENERAL",overview,[22,10,12,12,12])
    add_df_sheet(wb,"RESUMEN_UNICO",resumen_unico,[10,25,10,12,12,12])
    add_df_sheet(wb,"RESUMEN_RUTAS_RESTO",resto_summary,[15,10,12,12,12])

    wb.save(out_path)


def main():

    try:
        if WORKDIR.exists():
            os.chdir(WORKDIR)
    except:
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

    print("OK generado:", out_p)


if __name__ == "__main__":
    main()
