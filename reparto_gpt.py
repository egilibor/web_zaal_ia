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


# ============================================================
# CONFIG
# ============================================================

WORKDIR = Path(".")
DEFAULT_CSV = str(WORKDIR / "llegadas.csv")
DEFAULT_REGLAS = str(WORKDIR / "Reglas_hospitales.xlsx")
DEFAULT_OUT = str(WORKDIR / "salida.xlsx")


LIBRO_COORDS = Path(__file__).resolve().parent / "Libro_de_Servicio_Castellon_con_coordenadas.xlsx"

ORIGEN_LAT = 39.804106
ORIGEN_LON = -0.217351


# ============================================================
# UTILIDADES GENERALES
# ============================================================

def clean_text(x) -> str:
    if pd.isna(x):
        return ""
    s = str(x).strip()
    s = re.sub(r"\s+", " ", s)
    return s


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


def safe_sheet_name(name: str, existing: set) -> str:
    name = re.sub(r"[\\/*?:\[\]]", " ", str(name))
    name = re.sub(r"\s+", " ", name).strip()
    if not name:
        name = "SIN_NOMBRE"
    base = name[:31]
    candidate = base
    i = 2
    while candidate in existing:
        suffix = f" {i}"
        candidate = (base[:31-len(suffix)] + suffix).strip()
        i += 1
    existing.add(candidate)
    return candidate


# ============================================================
# COORDENADAS Y RUTA
# ============================================================

def haversine(lat1, lon1, lat2, lon2):
    R = 6371
    phi1, phi2 = math.radians(lat1), math.radians(lat2)
    dphi = math.radians(lat2 - lat1)
    dlambda = math.radians(lon2 - lon1)
    a = math.sin(dphi/2)**2 + math.cos(phi1)*math.cos(phi2)*math.sin(dlambda/2)**2
    return 2 * R * math.asin(math.sqrt(a))


def build_pueblo_coords(libro_path: Path) -> dict:
    df = pd.read_excel(libro_path)
    df = df.dropna(subset=["PUEBLO", "Latitud", "Longitud"])
    df["PUEBLO_NORM"] = df["PUEBLO"].astype(str).apply(norm)

    return (
        df.groupby("PUEBLO_NORM")
        .agg({"Latitud": "mean", "Longitud": "mean"})
        .to_dict("index")
    )


def nearest_neighbor_route(pueblos: list[str], coords: dict) -> list[str]:
    remaining = pueblos.copy()
    route = []
    current_lat, current_lon = ORIGEN_LAT, ORIGEN_LON

    while remaining:
        best = None
        best_dist = float("inf")

        for p in remaining:
            if p not in coords:
                continue
            lat = coords[p]["Latitud"]
            lon = coords[p]["Longitud"]
            dist = haversine(current_lat, current_lon, lat, lon)

            if dist < best_dist:
                best = p
                best_dist = dist

        if best is None:
            route.extend(remaining)
            break

        route.append(best)
        current_lat = coords[best]["Latitud"]
        current_lon = coords[best]["Longitud"]
        remaining.remove(best)

    return route


# ============================================================
# CORE
# ============================================================

def run(csv_path: Path, reglas_path: Path, out_path: Path, origen: str):

    df = pd.read_csv(csv_path, sep=";", encoding="utf-8-sig", dtype=str)

    df["Kgs"] = pd.to_numeric(df.get("Kgs", 0), errors="coerce").fillna(0)
    df["Bultos"] = pd.to_numeric(df.get("Btos.", 0), errors="coerce").fillna(0)

    resto_grp = (
        df.groupby(["Z.Rep", "Población", "Dir_OK"], dropna=False)
        .agg(
            Expediciones=("Exp", "nunique"),
            Bultos=("Bultos", "sum"),
            Kilos=("Kgs", "sum"),
        )
        .reset_index()
    )

    resto_grp["Kilos"] = resto_grp["Kilos"].round(1)

    wb_out = Workbook()
    wb_out.remove(wb_out.active)

    coords = build_pueblo_coords(LIBRO_COORDS)
    existing = set()

    for z, sub in resto_grp.groupby("Z.Rep"):

        sheet_name = safe_sheet_name(f"ZREP_{z}", existing)
        ws = wb_out.create_sheet(sheet_name)

        out = sub.copy()
        out["PUEBLO_NORM"] = out["Población"].apply(norm)

        pueblos_unicos = list(dict.fromkeys(out["PUEBLO_NORM"].tolist()))
        orden_pueblos = nearest_neighbor_route(pueblos_unicos, coords)
        ranking = {p: i for i, p in enumerate(orden_pueblos)}

        out["orden_pueblo"] = out["PUEBLO_NORM"].map(ranking).fillna(9999)
        out = out.sort_values(["orden_pueblo"], kind="stable").reset_index(drop=True)

        out.insert(0, "Parada", range(1, len(out) + 1))
        out = out.drop(columns=["PUEBLO_NORM", "orden_pueblo"])

        for row in dataframe_to_rows(out, index=False, header=True):
            ws.append(row)

    wb_out.save(out_path)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--csv")
    parser.add_argument("--reglas")
    parser.add_argument("--out")
    args = parser.parse_args()

    csv_p = Path(args.csv) if args.csv else Path(DEFAULT_CSV)
    reglas_p = Path(args.reglas) if args.reglas else Path(DEFAULT_REGLAS)
    out_p = Path(args.out) if args.out else Path(DEFAULT_OUT)

    run(csv_p, reglas_p, out_p, "LLEGADAS")


if __name__ == "__main__":
    main()
