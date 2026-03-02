import math
from pathlib import Path
import pandas as pd


# -------------------------------------------------
# CONSTANTES
# -------------------------------------------------

LAT0 = 39.804106
LON0 = -0.217351

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
    "N_servicio",
]


# -------------------------------------------------
# UTILIDADES
# -------------------------------------------------

def normalizar_texto(txt: str) -> str:
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


def cargar_coordenadas(ruta: Path) -> dict:
    df = pd.read_excel(ruta)
    df.columns = df.columns.str.strip().str.upper()

    columnas_necesarias = {"PUEBLO", "LATITUD", "LONGITUD"}

    if not columnas_necesarias.issubset(set(df.columns)):
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
# DISTANCIA REAL (HAVERSINE)
# -------------------------------------------------

def haversine(lat1: float, lon1: float, lat2: float, lon2: float) -> float:
    R = 6371
    lat1, lon1, lat2, lon2 = map(math.radians, [lat1, lon1, lat2, lon2])

    dlat = lat2 - lat1
    dlon = lon2 - lon1

    a = (
        math.sin(dlat / 2) ** 2
        + math.cos(lat1) * math.cos(lat2) * math.sin(dlon / 2) ** 2
    )
    c = 2 * math.atan2(math.sqrt(a), math.sqrt(1 - a))

    return R * c


# -------------------------------------------------
# ORDENACIÓN POR HOJA ZREP
# -------------------------------------------------

def ordenar_dataframe_zrep(df: pd.DataFrame, coords: dict) -> pd.DataFrame:

    for col in COLUMNAS_OBLIGATORIAS:
        if col not in df.columns:
            raise ValueError(f"Falta columna obligatoria: {col}")

    df = df.copy()

    filas_con_coord = []
    filas_sin_coord = []

    for idx, row in df.iterrows():
        pueblo_norm = normalizar_texto(row["Población"])

        if pueblo_norm in coords:
            lat, lon = coords[pueblo_norm]
            filas_con_coord.append((idx, lat, lon))
        else:
            filas_sin_coord.append(idx)

    visitados = []
    restantes = filas_con_coord.copy()

    lat_actual = LAT0
    lon_actual = LON0

    while restantes:

        distancias = []
        for idx, lat, lon in restantes:
            d = haversine(lat_actual, lon_actual, lat, lon)
            distancias.append((d, idx, lat, lon))

        distancias.sort(key=lambda x: (x[0], x[1]))
        _, idx_sel, lat_sel, lon_sel = distancias[0]

        visitados.append(idx_sel)
        lat_actual = lat_sel
        lon_actual = lon_sel

        restantes = [r for r in restantes if r[0] != idx_sel]

    orden_final = visitados + filas_sin_coord

    df_ordenado = df.loc[orden_final].copy()

    # ---- KM ENTRE PARADAS ----
    kms = []
    lat_actual = LAT0
    lon_actual = LON0

    for _, row in df_ordenado.iterrows():
        pueblo_norm = normalizar_texto(row["Población"])

        if pueblo_norm in coords:
            lat, lon = coords[pueblo_norm]
            km = haversine(lat_actual, lon_actual, lat, lon)
            kms.append(round(km, 1))
            lat_actual = lat
            lon_actual = lon
        else:
            kms.append(None)

    df_ordenado["Km_desde_anterior"] = kms

    return df_ordenado


# -------------------------------------------------
# ORDENAR HOSPITALES
# -------------------------------------------------

def ordenar_hospitales(df: pd.DataFrame, coords: dict) -> pd.DataFrame:

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
    lat_actual = LAT0
    lon_actual = LON0
    restantes = paradas_con_coord.copy()

    while restantes:

        distancias = []
        for clave, lat, lon in restantes:
            d = haversine(lat_actual, lon_actual, lat, lon)
            distancias.append((d, clave, lat, lon))

        distancias.sort(key=lambda x: (x[0], x[1]))
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
    ).copy()

    # ---- KM ENTRE PARADAS ----
    kms = []
    lat_actual = LAT0
    lon_actual = LON0

    for _, row in df_final.iterrows():
        pueblo_norm = normalizar_texto(row["Población"])

        if pueblo_norm in coords:
            lat, lon = coords[pueblo_norm]
            km = haversine(lat_actual, lon_actual, lat, lon)
            kms.append(round(km, 1))
            lat_actual = lat
            lon_actual = lon
        else:
            kms.append(None)

    df_final["Km_desde_anterior"] = kms

    return df_final


# -------------------------------------------------
# FUNCIÓN PRINCIPAL
# -------------------------------------------------

def reordenar_excel(input_path: Path, output_path: Path, ruta_coordenadas: Path):

    hojas = pd.read_excel(input_path, sheet_name=None)
    coords = cargar_coordenadas(ruta_coordenadas)

    hojas_resultado = {}

 for nombre, df in hojas.items():

    if nombre.startswith("ZREP_"):
        hojas_resultado[nombre] = ordenar_dataframe_zrep(df, coords)

    elif nombre == "HOSPITALES":
        hojas_resultado[nombre] = ordenar_hospitales(df, coords)

    elif nombre == "FEDERACION":
        hojas_resultado[nombre] = ordenar_dataframe_zrep(df, coords)

    else:
        hojas_resultado[nombre] = df

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for nombre, df in hojas_resultado.items():
            df.to_excel(writer, sheet_name=nombre, index=False)


