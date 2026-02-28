#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AGENTE CASTELLÓN · SALIDA DIARIA (HOSPITALES + FEDERACIÓN + RESTO POR Z.REP)

Versión v3
- Siempre pregunta por el "Origen de datos" (aunque pases argumentos).
- Si falta algún dato (csv/reglas/out), entra en modo interactivo y lo pregunta.

Uso rápido:
    python agente_castellon_salida_diaria_v3.py

Uso con argumentos:
    python agente_castellon_salida_diaria_v3.py --csv "llegadas.csv" --reglas "reglas.xlsx" --out "Salida.xlsx"
"""

from __future__ import annotations

import argparse
import os
import datetime as _dt
import re
from pathlib import Path

WORKDIR = Path(r"C:\REPARTO")
DEFAULT_CSV = str(WORKDIR / "llegadas.csv")
DEFAULT_REGLAS = str(WORKDIR / "Reglas_hospitales.xlsx")
DEFAULT_OUT = str(WORKDIR / "salida.xlsx")

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

def prompt_path(label: str, default: str = "") -> Path:
    while True:
        raw = input(f"{label}{' ['+default+']' if default else ''}: ").strip()
        if not raw and default:
            raw = default
        raw = raw.strip('"').strip("'")
        p = Path(raw)
        if p.exists():
            return p
        print(f"No existe: {p}")

def prompt_out_path(label: str, default: str = "") -> Path:
    while True:
        raw = input(f"{label}{' ['+default+']' if default else ''}: ").strip()
        if not raw and default:
            raw = default
        raw = raw.strip('"').strip("'")
        p = Path(raw)
        if p.suffix.lower() != ".xlsx":
            print("La salida debe terminar en .xlsx")
            continue
        return p

def prompt_origen() -> str:
    while True:
        origen = "LLEGADAS"
        if origen:
            return origen
        print("Escribe un nombre (no lo dejo vacío).")


# -------------------------
# Reglas
# -------------------------

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

def prepare_rules(df_rules: pd.DataFrame, pob_col: str, pat_col: str, active_col: str = "Activo") -> pd.DataFrame:
    if df_rules.empty:
        return df_rules
    d = df_rules.copy()
    if active_col in d.columns:
        d = d[d[active_col].astype(str).str.upper().eq("SI")].copy()
    d["Pob_norm"] = d[pob_col].astype(str).apply(norm)
    d["Pat_norm"] = d[pat_col].astype(str).apply(norm)
    d = d[d["Pat_norm"].ne("")].copy()
    return d

def match_rules(pob_norm: str, dir_norm: str, rules_df: pd.DataFrame, tag_field: str | None = None) -> Tuple[bool, str]:
    if rules_df.empty:
        return False, ""
    for _, r in rules_df.iterrows():
        rp = r.get("Pob_norm", "")
        if rp and rp not in pob_norm:
            continue
        pat = r.get("Pat_norm", "")
        if pat and pat in dir_norm:
            if tag_field:
                return True, str(r.get(tag_field, "") or "")
            return True, ""
    return False, ""


# -------------------------
# Excel writer
# -------------------------

def style_sheet(ws):
    thin = Side(style="thin", color="D0D0D0")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    header_font = Font(bold=True)
    header_fill = PatternFill("solid", fgColor="F2F2F2")
    wrap = Alignment(wrap_text=True, vertical="top")
    top = Alignment(vertical="top")

    if ws.max_row >= 1:
        for cell in ws[1]:
            cell.font = header_font
            cell.fill = header_fill
            cell.border = border
            cell.alignment = top

    for r in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for c in r:
            c.border = border
            c.alignment = wrap

    ws.freeze_panes = "A2"
    if ws.max_row >= 2 and ws.max_column >= 1:
        ws.auto_filter.ref = f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"

def set_widths(ws, widths: List[int]):
    for i, w in enumerate(widths, start=1):
        if i <= ws.max_column:
            ws.column_dimensions[get_column_letter(i)].width = w

def safe_sheet_name(name: str, existing: set) -> str:
    name = clean_text(name)
    name = re.sub(r"[\[\]\*:/\\\?]", " ", name)
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

def add_df_sheet(wb: Workbook, name: str, df_sheet: pd.DataFrame, widths: List[int]) -> None:
    ws = wb.create_sheet(title=name)
    out = df_sheet.copy()
    out.insert(0, "Nº", range(1, len(out) + 1))
    for row in dataframe_to_rows(out, index=False, header=True):
        ws.append(row)
    style_sheet(ws)
    set_widths(ws, widths)



import math

ORIGEN_LAT = 39.804106
ORIGEN_LON = -0.217351
COORD_FILE = Path(__file__).parent / "Libro_de_Servicio_Castellon_con_coordenadas.xlsx"

def build_pueblo_coords():
    df = pd.read_excel(COORD_FILE)
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

def nearest_neighbor_route(pueblos, coords):
    restantes = pueblos.copy()
    ruta = []
    actual_lat = ORIGEN_LAT
    actual_lon = ORIGEN_LON

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

# -------------------------
# Core
# -------------------------

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
    if col_dir_ent:
        fb = df_raw[col_dir_ent].apply(clean_text)
        df.loc[df["Dirección"].eq(""), "Dirección"] = fb[df["Dirección"].eq("")]

    for c in ["Consignatario", "Población", "Dirección", "Z.Rep"]:
        df.loc[df[c].eq(""), c] = f"SIN_{c.upper().replace('.', '')}"

    df["Parada_key"] = (df["Población"] + "||" + df["Dirección"]).str.strip("|")
    df["Pob_norm"] = df["Población"].apply(norm)
    df["Dir_norm"] = df["Dirección"].apply(norm)
    return df

def run(csv_path: Path, reglas_path: Path, out_path: Path, origen: str) -> None:
    df = load_csv(csv_path)

    wb_rules = load_workbook(reglas_path, data_only=True)
    rules_h = sheet_to_df(wb_rules, "REGLAS_HOSPITALES")
    rules_f = sheet_to_df(wb_rules, "REGLAS_FEDERACION")
    if rules_h.empty or rules_f.empty:
        raise ValueError("El Excel de reglas debe contener REGLAS_HOSPITALES y REGLAS_FEDERACION.")

    rules_h_prep = prepare_rules(rules_h, pob_col="Población", pat_col="Patrón_dirección")
    rules_h_prep["Hospital_final"] = rules_h_prep.get("Hospital_final", "").astype(str).replace({"None":"", "nan":""})

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
        ok, _ = match_rules(r["Pob_norm"], r["Dir_norm"], rules_f_prep, tag_field=None)
        if ok:
            df.at[i, "is_fed"] = True

    df["is_any_special"] = df["is_hospital"] | df["is_fed"]

    hosp = df[df["is_hospital"]].copy()
    hosp = hosp.sort_values(["Hospital", "Población", "Dirección"], kind="stable").reset_index(drop=True)

    fed = df[df["is_fed"]].copy()
    fed = fed.sort_values(["Población", "Dirección"], kind="stable").reset_index(drop=True)  

    resto = df[~df["is_any_special"]].copy()
    resto_grp = (
        resto.groupby(["Z.Rep", "Parada_key"], dropna=False)
        .agg(
            Población=("Población", "first"),
            Dirección=("Dirección", "first"),
            Consignatarios=("Consignatario", lambda s: unique_join(list(s))),
            Expediciones=("Exp", lambda s: s.nunique()),
            Bultos=("Bultos", "sum"),
            Kilos=("Kgs", "sum"),
            N_servicio=("N_servicio", lambda s: unique_join(list(s), sep=" | ")),
        )
        .reset_index()
    )
    resto_grp["Kilos"] = resto_grp["Kilos"].round(1)

    resto_summary = (
        resto_grp.groupby("Z.Rep")
        .agg(
            Paradas=("Parada_key", "nunique"),
            Expediciones=("Expediciones", "sum"),
            Bultos=("Bultos", "sum"),
            Kilos=("Kilos", "sum"),
        )
        .reset_index()
        .sort_values("Z.Rep")
    )
    resto_summary["Kilos"] = resto_summary["Kilos"].round(1)

    overview = pd.DataFrame(
        {
            "Bloque": ["HOSPITALES", "FEDERACION", "RESTO (todas rutas)"],
            "Paradas": [len(hosp), len(fed), resto_grp["Parada_key"].nunique()],
            "Expediciones": [
                int(df[df["is_hospital"]]["Exp"].nunique()),
                int(df[df["is_fed"]]["Exp"].nunique()),
                int(resto["Exp"].nunique()),
            ],
            "Kilos": [
                round(float(df[df["is_hospital"]]["Kgs"].sum()), 1),
                round(float(df[df["is_fed"]]["Kgs"].sum()), 1),
                round(float(resto["Kgs"].sum()), 1),
            ],
        }
    )

    # RESUMEN_UNICO = GENERAL + RESTO desglosado
    resumen_unico_general = overview[overview["Bloque"].ne("RESTO (todas rutas)")].copy()
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


    wb_out = Workbook()
    wb_out.remove(wb_out.active)
    COLUMNAS_BASE = [
        "Exp",
        "Hospital",
        "Población",
        "Dirección",
        "Consignatario",
        "Cliente",
        "Kgs",
        "Bultos",
        "Z.Rep",
        "N_servicio",
    ]
    meta = pd.DataFrame(
        {
            "Clave": ["Origen de datos", "CSV", "Reglas", "Generado"],
            "Valor": [
                origen,
                str(csv_path),
                str(reglas_path),
                _dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            ],
        }
    )
    add_df_sheet(wb_out, "METADATOS", meta, widths=[6, 22, 90])
    add_df_sheet(wb_out, "RESUMEN_GENERAL", overview, widths=[6, 22, 12, 14, 12])
    add_df_sheet(wb_out, "RESUMEN_UNICO", resumen_unico, widths=[6, 12, 28, 12, 14, 12, 12])
    add_df_sheet(
        wb_out,
        "HOSPITALES",
        hosp[COLUMNAS_BASE],
        widths=[10, 20, 18, 55, 25, 12, 10, 10],
    )
    add_df_sheet(
        wb_out,
        "FEDERACION",
        fed[COLUMNAS_BASE],
        widths=[6, 18, 55, 70, 16, 12],
    )
    add_df_sheet(wb_out, "RESUMEN_RUTAS_RESTO", resto_summary, widths=[6, 18, 10, 14, 12, 12])

    existing = set(wb_out.sheetnames)
    for z, sub in resto.groupby("Z.Rep"):
        sheet_name = safe_sheet_name(f"ZREP_{z}", existing)
        ws = wb_out.create_sheet(sheet_name)
        out = sub.copy()
        
        coords = build_pueblo_coords()
        out["PUEBLO_NORM"] = out["Población"].apply(norm)
        pueblos_unicos = list(dict.fromkeys(out["PUEBLO_NORM"].tolist()))
        orden_pueblos = nearest_neighbor_route(pueblos_unicos, coords)
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


def main():
    # Directorio de trabajo
    try:
        if WORKDIR.exists():
            os.chdir(WORKDIR)
    except Exception:
        pass

    parser = argparse.ArgumentParser(add_help=True)
    parser.add_argument("--csv", help="Ruta al CSV de LLEGADAS.")
    parser.add_argument("--reglas", help="Ruta al Excel de reglas (REGLAS_HOSPITALES + REGLAS_FEDERACION).")
    parser.add_argument("--out", help="Ruta de salida del Excel generado (.xlsx).")
    args = parser.parse_args()

    csv_p = Path(args.csv) if args.csv else None
    reglas_p = Path(args.reglas) if args.reglas else None
    out_p = Path(args.out) if args.out else None

    # Interactivo para rutas si faltan
    if csv_p is None or not csv_p.exists():
        csv_p = prompt_path("Ruta del CSV (LLEGADAS)", default=DEFAULT_CSV)
    if reglas_p is None or not reglas_p.exists():
        reglas_p = prompt_path("Ruta del Excel de reglas", default=DEFAULT_REGLAS)
    if out_p is None:
        default_out = str(Path.cwd() / "Salida_diaria.xlsx")
        out_p = prompt_out_path("Ruta de salida (.xlsx)", default=DEFAULT_OUT)

    # SIEMPRE preguntar origen (lo has pedido así)
    origen = prompt_origen()

    run(csv_p, reglas_p, out_p, origen)
    print(f"OK: generado {out_p}")

if __name__ == "__main__":
    main()
