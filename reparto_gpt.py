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


# ==========================================================
# UTILIDADES
# ==========================================================

def clean_text(x) -> str:
    if pd.isna(x):
        return ""
    return re.sub(r"\s+", " ", str(x).strip())


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
    except Exception:
        return 0.0


def parse_int(x) -> int:
    if pd.isna(x):
        return 0
    s = re.sub(r"[^0-9\-]", "", str(x))
    try:
        return int(s) if s else 0
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
    return re.sub(r"\s+", " ", s).strip()


def unique_join(values: List[str], sep: str = " / ") -> str:
    seen, out = set(), []
    for v in values:
        v = clean_text(v)
        if v and v not in seen:
            seen.add(v)
            out.append(v)
    return sep.join(out)


# ==========================================================
# CSV
# ==========================================================

def load_csv(csv_path: Path) -> pd.DataFrame:
    import csv
    df_raw = pd.read_csv(
        csv_path,
        sep=";",
        encoding="utf-8-sig",
        dtype=str,
        engine="python",
        quoting=csv.QUOTE_MINIMAL,
        on_bad_lines="warn"
    )

    df = pd.DataFrame({
        "Exp": df_raw["Exp"].apply(clean_text),
        "Kgs": df_raw["Kgs"].apply(parse_kg),
        "Bultos": df_raw.get("Btos.", 0).apply(parse_int) if "Btos." in df_raw else 0,
        "Consignatario": df_raw["Consignatario"].apply(clean_text),
        "Cliente": df_raw.get("Cliente", "").apply(clean_text),
        "Población": df_raw["Población"].apply(clean_text),
        "Dirección": df_raw.get("Dir_OK", "").apply(clean_text),
        "Z.Rep": df_raw.get("Z.Rep", "SIN_ZONA").apply(clean_text),
        "N_servicio": df_raw.get("N. servicio", "").apply(clean_text),
    })

    df["Pob_norm"] = df["Población"].apply(norm)
    df["Dir_norm"] = df["Dirección"].apply(norm)
    return df


# ==========================================================
# COORDENADAS
# ==========================================================

def build_pueblo_coords(ruta_coordenadas: Path):
    df = pd.read_excel(ruta_coordenadas)
    coords = {}
    for _, r in df.iterrows():
        pueblo = norm(str(r["PUEBLO"]))
        coords[pueblo] = (float(r["Latitud"]), float(r["Longitud"]))
    return coords


def haversine(lat1, lon1, lat2, lon2):
    R = 6371
    phi1, phi2 = map(math.radians, [lat1, lat2])
    dphi = math.radians(lat2 - lat1)
    dlambda = math.radians(lon2 - lon1)
    a = math.sin(dphi/2)**2 + math.cos(phi1)*math.cos(phi2)*math.sin(dlambda/2)**2
    return 2 * R * math.atan2(math.sqrt(a), math.sqrt(1-a))


def nearest_neighbor_route(pueblos, coords, lat0, lon0):
    restantes = pueblos.copy()
    ruta = []
    actual_lat, actual_lon = lat0, lon0

    while restantes:
        mejor, mejor_dist = None, float("inf")
        for p in restantes:
            if p not in coords:
                continue
            lat, lon = coords[p]
            dist = haversine(actual_lat, actual_lon, lat, lon)
            if mejor is None or dist < mejor_dist:
                mejor, mejor_dist = p, dist
        if mejor is None:
            ruta.extend(sorted(restantes))
            break
        ruta.append(mejor)
        actual_lat, actual_lon = coords[mejor]
        restantes.remove(mejor)
    return ruta


# ==========================================================
# CORE
# ==========================================================

def run(csv_path: Path,
        reglas_path: Path,
        out_path: Path,
        origen: str,
        ruta_coordenadas: Path,
        lat_origen: float,
        lon_origen: float):

    df = load_csv(csv_path)
    coords = build_pueblo_coords(ruta_coordenadas)

    wb_out = Workbook()
    wb_out.remove(wb_out.active)

    COLUMNAS = [
        "Exp", "Población", "Dirección",
        "Consignatario", "Cliente",
        "Kgs", "Bultos", "Z.Rep", "N_servicio"
    ]

    existing = set()

    for z, sub in df.groupby("Z.Rep"):
        sheet_name = f"ZREP_{z}"
        ws = wb_out.create_sheet(title=sheet_name)

        out = sub.copy()
        out["PUEBLO_NORM"] = out["Población"].apply(norm)

        pueblos_unicos = list(dict.fromkeys(out["PUEBLO_NORM"].tolist()))
        orden = nearest_neighbor_route(pueblos_unicos, coords, lat_origen, lon_origen)
        ranking = {p: i for i, p in enumerate(orden)}
        out["orden"] = out["PUEBLO_NORM"].map(ranking).fillna(9999)

        out = out.sort_values("orden", kind="stable")
        out = out[COLUMNAS]

        for row in dataframe_to_rows(out, index=False, header=True):
            ws.append(row)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb_out.save(out_path)


# ==========================================================
# MAIN
# ==========================================================

def main():

    parser = argparse.ArgumentParser()
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
