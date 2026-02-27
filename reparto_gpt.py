#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AGENTE CASTELLÓN · SALIDA DIARIA (HOSPITALES + FEDERACIÓN + RESTO POR Z.REP)
Versión v3 + Optimización determinista por pueblos
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


WORKDIR = Path(".")
DEFAULT_CSV = str(WORKDIR / "llegadas.csv")
DEFAULT_REGLAS = str(WORKDIR / "Reglas_hospitales.xlsx")
DEFAULT_OUT = str(WORKDIR / "salida.xlsx")

LIBRO_COORDS = Path(__file__).resolve().parent / "Libro_de_Servicio_Castellon_con_coordenadas.xlsx"
ORIGEN_LAT = 39.804106
ORIGEN_LON = -0.217351


# =========================
# UTILIDADES GENERALES
# =========================

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


# =========================
# OPTIMIZACIÓN RUTAS
# =========================

def haversine(lat1, lon1, lat2, lon2):
    R = 6371
    phi1, phi2 = math.radians(lat1), math.radians(lat2)
    dphi = math.radians(lat2 - lat1)
    dlambda = math.radians(lon2 - lon1)
    a = math.sin(dphi/2)**2 + math.cos(phi1)*math.cos(phi2)*math.sin(dlambda/2)**2
    return 2 * R * math.asin(math.sqrt(a))


def build_pueblo_coords():
    df = pd.read_excel(LIBRO_COORDS)
    df = df.dropna(subset=["PUEBLO", "Latitud", "Longitud"])
    df["PUEBLO_NORM"] = df["PUEBLO"].astype(str).apply(norm)
    return (
        df.groupby("PUEBLO_NORM")
        .agg({"Latitud": "mean", "Longitud": "mean"})
        .to_dict("index")
    )


def nearest_neighbor_route(pueblos, coords):
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
