import math
from pathlib import Path
import pandas as pd


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
# DISTANCIA
# -------------------------------------------------

def haversine(lat1, lon1, lat2, lon2):
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
# ORDENACIÓN ZREP
# -------------------------------------------------

def ordenar_dataframe_zrep(df: pd.DataFrame, coords: dict, lat_origen: float, lon_origen: float) -> pd.DataFrame:

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

    lat_actual = lat_origen
    lon_actual = lon_origen

    while restantes:
        distancias = [
            (haversine(lat_actual, lon_actual, lat, lon), idx, lat, lon)
            for idx, lat, lon in restantes
        ]

        distancias.sort(key=lambda x: (x[0], x[1]))
        _, idx_sel, lat_sel, lon_sel = distancias[0]

        visitados.append(idx_sel)
        lat_actual = lat_sel
        lon_actual = lon_sel

        restantes = [r for r in restantes if r[0] != idx_sel]

    orden_final = visitados + filas_sin_coord
    df_ordenado = df.loc[orden_final].copy()

    return df_ordenado


# -------------------------------------------------
# ORDENACIÓN HOSPITALES
# -------------------------------------------------

def ordenar_hospitales(df: pd.DataFrame, coords: dict, lat_origen: float, lon_origen: float) -> pd.DataFrame:
    return ordenar_dataframe_zrep(df, coords, lat_origen, lon_origen)


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
            hojas_resultado[nombre] = ordenar_dataframe_zrep(df, coords, lat_origen, lon_origen)

        elif nombre in ("HOSPITALES", "FEDERACION"):
            hojas_resultado[nombre] = ordenar_hospitales(df, coords, lat_origen, lon_origen)

        else:
            hojas_resultado[nombre] = df.copy()

    # -------------------------------------------------
    # ORDEN FIJO DETERMINISTA
    # -------------------------------------------------

    orden_base = [
        "METADATOS",
        "RESUMEN_UNICO",
        "RESUMEN_GENERAL",
        "HOSPITALES",
        "FEDERACION",
        "RESUMEN_RUTAS_RESTO",
    ]

    zreps = sorted(
        [n for n in hojas_resultado if n.startswith("ZREP_")]
    )

    otras = [
        n for n in hojas_resultado
        if n not in orden_base and not n.startswith("ZREP_")
    ]

    orden_final = []

    for hoja in orden_base:
        if hoja in hojas_resultado:
            orden_final.append(hoja)

    orden_final.extend(zreps)
    orden_final.extend(otras)

    # -------------------------------------------------
    # ESCRITURA CONTROLADA
    # -------------------------------------------------

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for nombre in orden_final:
            hojas_resultado[nombre].to_excel(
                writer,
                sheet_name=nombre,
                index=False
            )
