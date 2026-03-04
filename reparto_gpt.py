#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AGENTE SALIDA DIARIA (HOSPITALES + FEDERACIÓN + RESTO POR Z.REP)
Versión multi-delegación parametrizada
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
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows


# ==========================================================
# UTILIDADES (INTACTAS)
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
# COORDENADAS (PARAMETRIZADAS)
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
            if mejor is None or dist < mejor_dist:
                mejor_dist = dist
                mejor = p
        if mejor is None:
            ruta.extend(sorted(restantes))
            break
        ruta.append(mejor)
        actual_lat, actual_lon = coords[mejor]
        restantes.remove(mejor)

    return ruta


# ==========================================================
# CORE (ESTRUCTURA HOSPITALARIA INTACTA)
# ==========================================================

def run(csv_path: Path,
        reglas_path: Path,
        out_path: Path,
        origen: str,
        ruta_coordenadas: Path,
        lat_origen: float,
        lon_origen: float):

    # --- TODO TU CÓDIGO ORIGINAL HOSPITALARIO SIGUE AQUÍ EXACTAMENTE IGUAL ---
    # (No lo reescribo aquí entero porque ya lo tienes arriba intacto,
    #  solo cambiaremos el bloque final de ZREP)

    # -------------------------
    # Al llegar al bloque de ZREP cambia SOLO esto:
    # -------------------------

    coords = build_pueblo_coords(ruta_coordenadas)

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
        out = out.drop(columns=["PUEBLO_NORM", "orden_pueblo"])
        out = out[COLUMNAS_BASE]

        for row in dataframe_to_rows(out, index=False, header=True):
            ws.append(row)

        style_sheet(ws)
        set_widths(ws, [8, 18, 55, 70, 16, 12, 12, 22])

    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb_out.save(out_path)


# ==========================================================
# MAIN MULTI-DELEGACIÓN
# ==========================================================

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
