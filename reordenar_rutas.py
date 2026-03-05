#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import math
from pathlib import Path
import pandas as pd


# -------------------------------------------------
# ORÍGENES POR DEFECTO
# -------------------------------------------------

LAT_CASTELLON = 39.804106
LON_CASTELLON = -0.217351

LAT_VALENCIA = 39.44069
LON_VALENCIA = -0.42589


# -------------------------------------------------
# COLUMNAS OBLIGATORIAS
# -------------------------------------------------

COLUMNAS_OBLIGATORIAS = [
    "Exp",
    "Hospital",
    "Población",
    "Dirección",
    "Consignatario",
    "Cliente",
    "Kgs",
    "Bultos",
    "Z.Rep",
    "N. servicio",
]


# -------------------------------------------------
# NORMALIZAR TEXTO
# -------------------------------------------------

def normalizar_texto(txt):

    if pd.isna(txt):
        return ""

    txt = str(txt).strip().upper()

    txt = (
        txt.replace("Á", "A")
        .replace("É", "E")
        .replace("Í", "I")
        .replace("Ó", "O")
        .replace("Ú", "U")
        .replace("Ü", "U")
        .replace("Ñ", "N")
    )

    txt = " ".join(txt.split())

    return txt


# -------------------------------------------------
# GOOGLE MAPS LINK
# -------------------------------------------------

def generar_link_pueblos(df_ruta, lat_origen, lon_origen):

    pueblos = df_ruta.drop_duplicates(subset="Población")

    puntos = [f"{lat_origen},{lon_origen}"]

    for _, row in pueblos.iterrows():

        lat = row.get("Latitud")
        lon = row.get("Longitud")

        if pd.notna(lat) and pd.notna(lon):

            puntos.append(f"{lat},{lon}")

    if len(puntos) < 2:
        return ""

    url = "https://www.google.com/maps/dir/" + "/".join(puntos)

    # Excel lo reconocerá siempre como enlace
    return f'=HYPERLINK("{url}","Abrir ruta en Google Maps")'


# -------------------------------------------------
# CARGAR COORDENADAS
# -------------------------------------------------

def cargar_coordenadas(ruta):

    df = pd.read_excel(ruta)

    df.columns = df.columns.str.strip().str.upper()

    columnas_necesarias = {"PUEBLO", "LATITUD", "LONGITUD"}

    if not columnas_necesarias.issubset(df.columns):

        raise ValueError(
            f"Columnas detectadas: {list(df.columns)}. "
            "Se esperaban: PUEBLO, LATITUD, LONGITUD."
        )

    coords = {}

    for _, row in df.iterrows():

        pueblo = normalizar_texto(row["PUEBLO"])
        lat = row["LATITUD"]
        lon = row["LONGITUD"]

        if pd.notna(pueblo) and pd.notna(lat) and pd.notna(lon):

            coords[pueblo] = (float(lat), float(lon))

    return coords


# -------------------------------------------------
# ORDENACIÓN ZREP
# -------------------------------------------------

def ordenar_dataframe_zrep(df, coords, lat_origen, lon_origen):

    for col in COLUMNAS_OBLIGATORIAS:
        if col not in df.columns:
            raise ValueError(f"Falta columna obligatoria: {col}")

    df = df.copy()

    df["Latitud"] = None
    df["Longitud"] = None

    filas_con_coord = []
    filas_sin_coord = []

    for idx, row in df.iterrows():

        pueblo_norm = normalizar_texto(row["Población"])

        if pueblo_norm in coords:

            lat, lon = coords[pueblo_norm]

            df.at[idx, "Latitud"] = lat
            df.at[idx, "Longitud"] = lon

            filas_con_coord.append((idx, lat, lon))

        else:

            filas_sin_coord.append(idx)

    visitados = []
    restantes = filas_con_coord.copy()

    lat_actual = lat_origen
    lon_actual = lon_origen

    while restantes:

        distancias = []

        for item in restantes:

            idx, lat, lon = item
            d = (lat - lat_actual) ** 2 + (lon - lon_actual) ** 2
            distancias.append((d, idx, lat, lon))

        distancias.sort(key=lambda x: (x[0], x[1]))

        _, idx_sel, lat_sel, lon_sel = distancias[0]

        visitados.append(idx_sel)

        lat_actual = lat_sel
        lon_actual = lon_sel

        restantes = [r for r in restantes if r[0] != idx_sel]

    orden_final = visitados + filas_sin_coord

    df_ordenado = df.loc[orden_final]

    return df_ordenado


# -------------------------------------------------
# ORDENAR HOSPITALES
# -------------------------------------------------

def ordenar_hospitales(df, coords, lat_origen, lon_origen):

    df = df.copy()

    df["_orden_original"] = range(len(df))

    df["_clave_parada"] = (
        df["Consignatario"].apply(normalizar_texto)
        + "|"
        + df["Dirección"].apply(normalizar_texto)
    )

    paradas = {}

    for idx, row in df.iterrows():

        clave = row["_clave_parada"]

        if clave not in paradas:

            paradas[clave] = {
                "indices": [],
                "poblacion": normalizar_texto(row["Población"]),
            }

        paradas[clave]["indices"].append(idx)

    paradas_con_coord = []
    paradas_sin_coord = []

    for clave, data in paradas.items():

        pueblo = data["poblacion"]

        if pueblo in coords:

            lat, lon = coords[pueblo]
            paradas_con_coord.append((clave, lat, lon))

        else:

            paradas_sin_coord.append(clave)

    orden_paradas = []

    lat_actual = lat_origen
    lon_actual = lon_origen

    restantes = paradas_con_coord.copy()

    while restantes:

        distancias = []

        for clave, lat, lon in restantes:

            d = (lat - lat_actual) ** 2 + (lon - lon_actual) ** 2
            distancias.append((d, clave, lat, lon))

        distancias.sort(key=lambda x: x[0])

        _, clave_sel, lat_sel, lon_sel = distancias[0]

        orden_paradas.append(clave_sel)

        lat_actual = lat_sel
        lon_actual = lon_sel

        restantes = [r for r in restantes if r[0] != clave_sel]

    orden_paradas.extend(paradas_sin_coord)

    orden_indices = []

    for clave in orden_paradas:

        indices = paradas[clave]["indices"]

        indices_ordenados = sorted(
            indices,
            key=lambda i: df.loc[i, "_orden_original"]
        )

        orden_indices.extend(indices_ordenados)

    df_final = df.loc[orden_indices].drop(
        columns=["_orden_original", "_clave_parada"]
    )

    return df_final


# -------------------------------------------------
# FUNCIÓN PRINCIPAL
# -------------------------------------------------

def reordenar_excel(
    input_path: Path,
    output_path: Path,
    ruta_coordenadas: Path,
    lat_origen: float = LAT_CASTELLON,
    lon_origen: float = LON_CASTELLON,
):

    hojas = pd.read_excel(input_path, sheet_name=None)

    coords = cargar_coordenadas(ruta_coordenadas)

    hojas_resultado = {}

    for nombre, df in hojas.items():

        if nombre.startswith("ZREP_"):

            df_ordenado = ordenar_dataframe_zrep(
                df,
                coords,
                lat_origen,
                lon_origen,
            )

            link = generar_link_pueblos(df_ordenado, lat_origen, lon_origen)

            df_ordenado.insert(0, "NAVEGACIÓN", "")

            if link:
                df_ordenado.loc[df_ordenado.index[0], "NAVEGACIÓN"] = link

            hojas_resultado[nombre] = df_ordenado

        elif nombre == "HOSPITALES":

            df_ordenado = ordenar_hospitales(
                df,
                coords,
                lat_origen,
                lon_origen,
            )

            hojas_resultado[nombre] = df_ordenado

        else:

            hojas_resultado[nombre] = df

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:

        for nombre, df in hojas_resultado.items():

            df.to_excel(writer, sheet_name=nombre, index=False)
