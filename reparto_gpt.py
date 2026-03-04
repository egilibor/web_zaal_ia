#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AGENTE CASTELLÓN · SALIDA DIARIA (HOSPITALES + FEDERACIÓN + RESTO POR Z.REP)

Versión v3 MULTI-DELEGACIÓN
"""

from __future__ import annotations

import argparse
import os
import datetime as _dt
import re
import math
from pathlib import Path
from typing import List, Tuple

WORKDIR = Path(r"C:\REPARTO")
DEFAULT_CSV = str(WORKDIR / "llegadas.csv")
DEFAULT_REGLAS = str(WORKDIR / "Reglas_hospitales.xlsx")
DEFAULT_OUT = str(WORKDIR / "salida.xlsx")

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows


# -------------------------------------------------
# UTILIDADES (INTACTAS)
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

def pick_col(columns: List[str], candidates: List[str]) -> str | None:
    s = set(columns)
    for c in candidates:
        if c in s:
            return c
    return None


# -------------------------------------------------
# COORDENADAS PARAMETRIZADAS
# -------------------------------------------------

def build_pueblo_coords(ruta_coordenadas: Path):
    df = pd.read_excel(ruta_coordenadas)
    coords = {}
    for _, r in df.iterrows():
        pueblo = norm(str(r["PUEBLO"]))
        lat = float(r["Latitud"])
        lon = float(r["Longitud"])
        coords[pueblo] = (lat, lon)
    return coords

def haversine(lat1, lon1, lat2, lon2):
    R = 6371
    phi1 = math.radians(lat1)
    phi2 = math.radians(lat2)
    dphi = math.radians(lat2 - lat1)
    dlambda = math.radians(lon2 - lon1)
    a = math.sin(dphi/2)**2 + math.cos(phi1)*math.cos(phi2)*math.sin(dlambda/2)**2
    return 2 * R * math.atan2(math.sqrt(a), math.sqrt(1-a))

def nearest_neighbor_route(pueblos, coords, lat_origen, lon_origen):
    restantes = pueblos.copy()
    ruta = []
    actual_lat = lat_origen
    actual_lon = lon_origen

    while restantes:
        mejor = None
        mejor_dist = float("inf")

        for p in restantes:
            if p not in coords:
                continue
            lat, lon = coords[p]
            dist = haversine(actual_lat, actual_lon, lat, lon)
            if mejor is None or dist < mejor_dist or (dist == mejor_dist and p < mejor):
                mejor_dist = dist
                mejor = p

        if mejor is None:
            ruta.extend(sorted(restantes))
            break

        ruta.append(mejor)
        actual_lat, actual_lon = coords[mejor]
        restantes.remove(mejor)

    return ruta


# -------------------------------------------------
# CORE ORIGINAL (NO TOCADO)
# -------------------------------------------------

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

    required = {"Exp": col_exp, "Kgs": col_kg, "Población": col_pop, "Consignatario": col_cons}
    missing = [k for k, v in required.items() if v is None]
    if missing:
        raise ValueError(f"Faltan columnas mínimas en CSV: {missing}.")

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

    for c in ["Consignatario", "Población", "Dirección", "Z.Rep"]:
        df.loc[df[c].eq(""), c] = f"SIN_{c.upper().replace('.', '')}"

    df["Parada_key"] = (df["Población"] + "||" + df["Dirección"]).str.strip("|")
    df["Pob_norm"] = df["Población"].apply(norm)
    df["Dir_norm"] = df["Dirección"].apply(norm)
    return df


# -------------------------------------------------
# RUN MULTI-DELEGACIÓN
# -------------------------------------------------

def run(csv_path: Path,
        reglas_path: Path,
        out_path: Path,
        origen: str,
        ruta_coordenadas: Path,
        lat_origen: float,
        lon_origen: float) -> None:

    df = load_csv(csv_path)

    wb_rules = load_workbook(reglas_path, data_only=True)
    rules_h = sheet_to_df(wb_rules, "REGLAS_HOSPITALES")
    rules_f = sheet_to_df(wb_rules, "REGLAS_FEDERACION")

    rules_h_prep = prepare_rules(rules_h, pob_col="Población", pat_col="Patrón_dirección")
    rules_h_prep["Hospital_final"] = rules_h_prep.get("Hospital_final", "").astype(str)

    rules_f_prep = prepare_rules(rules_f, pob_col="Población", pat_col="Patrón_dirección")

    df["is_hospital"] = False
    df["Hospital"] = ""
    for i, r in df.iterrows():
        ok, tag = match_rules(r["Pob_norm"], r["Dir_norm"], rules_h_prep, tag_field="Hospital_final")
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

    wb_out = Workbook()
    wb_out.remove(wb_out.active)

    COLUMNAS_BASE = [
        "Exp","Hospital","Población","Dirección",
        "Consignatario","Cliente","Kgs","Bultos","Z.Rep","N_servicio"
    ]

    add_df_sheet(wb_out, "HOSPITALES", hosp[COLUMNAS_BASE], [10]*10)
    add_df_sheet(wb_out, "FEDERACION", fed[COLUMNAS_BASE], [10]*10)

    coords = build_pueblo_coords(ruta_coordenadas)

    existing = set(wb_out.sheetnames)

    for z, sub in resto.groupby("Z.Rep"):
        sheet_name = safe_sheet_name(f"ZREP_{z}", existing)
        ws = wb_out.create_sheet(sheet_name)
        out = sub.copy()

        out["PUEBLO_NORM"] = out["Población"].apply(norm)
        pueblos_unicos = list(dict.fromkeys(out["PUEBLO_NORM"].tolist()))
        orden_pueblos = nearest_neighbor_route(
            pueblos_unicos,
            coords,
            lat_origen,
            lon_origen
        )
        ranking = {p: i for i, p in enumerate(orden_pueblos)}
        out["orden_pueblo"] = out["PUEBLO_NORM"].map(ranking).fillna(9999)

        out = out.sort_values(["orden_pueblo"], kind="stable").reset_index(drop=True)
        out = out.drop(columns=["PUEBLO_NORM","orden_pueblo"])
        out = out[COLUMNAS_BASE]

        for row in dataframe_to_rows(out, index=False, header=True):
            ws.append(row)

        style_sheet(ws)
        set_widths(ws, [10]*10)

    wb_out.save(out_path)


# -------------------------------------------------
# MAIN
# -------------------------------------------------

def main():

    parser = argparse.ArgumentParser(add_help=True)
    parser.add_argument("--csv", required=True)
    parser.add_argument("--reglas", required=True)
    parser.add_argument("--out", required=True)
    parser.add_argument("--coords", required=True)
    parser.add_argument("--lat", required=True, type=float)
    parser.add_argument("--lon", required=True, type=float)

    args = parser.parse_args()

    run(
        Path(args.csv),
        Path(args.reglas),
        Path(args.out),
        "LLEGADAS",
        Path(args.coords),
        args.lat,
        args.lon
    )

    print(f"OK generado {args.out}")


if __name__ == "__main__":
    main()
