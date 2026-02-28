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

    # Normalizar cabeceras: quitar espacios y pasar a mayúsculas
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

    for _, row in df.iterrows():
        pueblo = normalizar_texto(row["PUEBLO"])
        lat = row["LATITUD"]
        lon = row["LONGITUD"]

        if pd.notna(pueblo) and pd.notna(lat) and pd.notna(lon):
            coords[pueblo] = (float(lat), float(lon))

    return coords

    for _, row in df.iterrows():
        pueblo = normalizar_texto(row["PUEBLO"])
        lat = row["LATITUD"]
        lon = row["LONGITUD"]

        if pd.notna(pueblo) and pd.notna(lat) and pd.notna(lon):
            coords[pueblo] = (float(lat), float(lon))

    return coords


def calcular_angulo_distancia(lat: float, lon: float):
    dy = lat - LAT0
    dx = lon - LON0

    angulo = math.atan2(dy, dx)
    distancia = math.sqrt(dx ** 2 + dy ** 2)

    return angulo, distancia


# -------------------------------------------------
# ORDENACIÓN POR HOJA
# -------------------------------------------------

def ordenar_dataframe_zrep(df: pd.DataFrame, coords: dict) -> pd.DataFrame:

    for col in COLUMNAS_OBLIGATORIAS:
        if col not in df.columns:
            raise ValueError(f"Falta columna obligatoria: {col}")

    df = df.copy()

    # Separar con y sin coordenadas
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

        # Calcular distancias desde punto actual
        distancias = []
        for item in restantes:
            idx, lat, lon = item
            d = (lat - lat_actual) ** 2 + (lon - lon_actual) ** 2
            distancias.append((d, idx, lat, lon))

        # Elegir el más cercano
        distancias.sort(key=lambda x: (x[0], x[1]))
        d_min, idx_sel, lat_sel, lon_sel = distancias[0]

        visitados.append(idx_sel)

        # Actualizar punto actual
        lat_actual = lat_sel
        lon_actual = lon_sel

        # Quitar seleccionado
        restantes = [r for r in restantes if r[0] != idx_sel]

    # Añadir los sin coordenadas al final en orden original
    orden_final = visitados + filas_sin_coord

    df_ordenado = df.loc[orden_final]

    return df_ordenado


# -------------------------------------------------
# FUNCIÓN PRINCIPAL
# -------------------------------------------------

def reordenar_excel(input_path: Path, output_path: Path, ruta_coordenadas: Path):

    hojas = pd.read_excel(input_path, sheet_name=None)

    coords = cargar_coordenadas(ruta_coordenadas)

    hojas_resultado = {}

    for nombre, df in hojas.items():

        if nombre.startswith("ZREP_"):
            df_ordenado = ordenar_dataframe_zrep(df, coords)
            hojas_resultado[nombre] = df_ordenado
        else:
            hojas_resultado[nombre] = df

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for nombre, df in hojas_resultado.items():

            df.to_excel(writer, sheet_name=nombre, index=False)






