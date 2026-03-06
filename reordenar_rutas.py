#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from pathlib import Path
import pandas as pd


# -------------------------------------------------
# ORÍGENES
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
# DISTANCIA
# -------------------------------------------------

def distancia(a, b):
    return (a[0] - b[0]) ** 2 + (a[1] - b[1]) ** 2


# -------------------------------------------------
# 2-OPT
# -------------------------------------------------

def mejorar_ruta_2opt(coords):

    mejor = coords.copy()
    mejora = True

    while mejora:

        mejora = False

        for i in range(1, len(mejor) - 2):
            for j in range(i + 1, len(mejor)):

                if j - i == 1:
                    continue

                a = mejor[i - 1]
                b = mejor[i]
                c = mejor[j - 1]
                d = mejor[j % len(mejor)]

                actual = distancia(a, b) + distancia(c, d)
                nuevo = distancia(a, c) + distancia(b, d)

                if nuevo < actual:

                    mejor[i:j] = reversed(mejor[i:j])
                    mejora = True

    return mejor


# -------------------------------------------------
# GOOGLE MAPS LINK
# -------------------------------------------------

def generar_link_pueblos(df_ruta, lat_origen, lon_origen):

    puntos = [f"{lat_origen},{lon_origen}"]

    coords_vistas = set()

    for _, row in df_ruta.iterrows():

        lat = row.get("Latitud")
        lon = row.get("Longitud")

        if pd.notna(lat) and pd.notna(lon):

            clave = (round(float(lat), 5), round(float(lon), 5))

            if clave not in coords_vistas:

                coords_vistas.add(clave)
                puntos.append(f"{clave[0]},{clave[1]}")

    if len(puntos) < 2:
        return ""

    return "https://www.google.com/maps/dir/" + "/".join(puntos)


# -------------------------------------------------
# CARGAR COORDENADAS
# -------------------------------------------------

def cargar_coordenadas(ruta):

    df = pd.read_excel(ruta)

    df.columns = df.columns.str.strip().str.upper()

    if not {"PUEBLO", "LATITUD", "LONGITUD"}.issubset(df.columns):

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

        for idx, lat, lon in restantes:

            d = (lat - lat_actual) ** 2 + (lon - lon_actual) ** 2
            distancias.append((d, idx, lat, lon))

        distancias.sort(key=lambda x: (x[0], x[1]))

        _, idx_sel, lat_sel, lon_sel = distancias[0]

        visitados.append(idx_sel)

        lat_actual = lat_sel
        lon_actual = lon_sel

        restantes = [r for r in restantes if r[0] != idx_sel]

    coords_ruta = [(df.loc[i, "Latitud"], df.loc[i, "Longitud"]) for i in visitados]

    coords_mejoradas = mejorar_ruta_2opt(coords_ruta)

    visitados_nuevo = []
    usados = set()

    for c in coords_mejoradas:
        for i in visitados:
            if i not in usados and (df.loc[i, "Latitud"], df.loc[i, "Longitud"]) == c:
                visitados_nuevo.append(i)
                usados.add(i)
                break

    visitados = visitados_nuevo

    orden_final = visitados + filas_sin_coord

    return df.loc[orden_final]


# -------------------------------------------------
# FUNCIÓN PRINCIPAL
# -------------------------------------------------

def reordenar_excel(
    input_path: Path,
    output_path: Path,
    ruta_coordenadas: Path,
    lat_origen: float,
    lon_origen: float,
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

            #df_ordenado.insert(0, "NAVEGACIÓN", "")
            df_ordenado["NAVEGACIÓN"] = ""
            if link:
                df_ordenado.loc[df_ordenado.index[0], "NAVEGACIÓN"] = link

            hojas_resultado[nombre] = df_ordenado

        else:

            hojas_resultado[nombre] = df

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:

        for nombre, df in hojas_resultado.items():

            df.to_excel(writer, sheet_name=nombre, index=False)

