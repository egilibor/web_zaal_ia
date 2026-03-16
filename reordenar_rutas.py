#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from pathlib import Path
from geocodificador import geocodificar
import pandas as pd
import re

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
    if a[0] is None or a[1] is None or b[0] is None or b[1] is None:
        return float('inf')
    return (a[0] - b[0]) ** 2 + (a[1] - b[1]) ** 2

# ----- PARADAS------
def extraer_calle_sin_numero(direccion: str) -> str:
    if pd.isna(direccion):
        return ""
    m = re.match(r"(.*?)[,\s]+\d+.*$", str(direccion).strip())
    if m:
        return m.group(1).strip()
    return str(direccion).strip()


def calcular_paradas_por_hoja(hojas_resultado: dict) -> dict:
    paradas = {}
    for nombre, df in hojas_resultado.items():
        if nombre in ("RESUMEN_UNICO", "METADATOS"):
            continue
        if "Población" not in df.columns or "Dirección" not in df.columns:
            continue
        calle_sin_num = df["Dirección"].apply(extraer_calle_sin_numero)
        clave = df["Población"].astype(str).str.strip().str.upper() + "|" + calle_sin_num.str.upper()
        paradas[nombre] = clave.nunique()
    return paradas
    
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

def generar_links_segmentos(df_ruta, lat_origen, lon_origen, tamanio=9):

    puntos = []
    coords_vistas = set()

    for _, row in df_ruta.iterrows():
        lat = row.get("Latitud")
        lon = row.get("Longitud")
        if pd.notna(lat) and pd.notna(lon):
            clave = (round(float(lat), 5), round(float(lon), 5))
            if clave not in coords_vistas:
                coords_vistas.add(clave)
                puntos.append(f"{clave[0]},{clave[1]}")

    if not puntos:
        return []

    origen = f"{lat_origen},{lon_origen}"
    links = []
    
    for i in range(0, len(puntos), tamanio):
        segmento = puntos[i:i + tamanio]
        # El segmento empieza donde terminó el anterior
        if i == 0:
            tramo = [origen] + segmento
        else:
            tramo = [puntos[i - 1]] + segmento
        links.append("https://www.google.com/maps/dir/" + "/".join(tramo))

    return links
    
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

def ordenar_dataframe_zrep(df, coords, lat_origen, lon_origen, api_key="", delegacion="castellon"):
    
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

        lat, lon = (None, None)

        # Primero intentar geocodificación por dirección completa
        if api_key:
            dir_limpia = str(row['Dirección']).strip()
            pob_limpia = str(row['Población']).strip()
            if dir_limpia.upper() not in ("NAN", "NONE", "") and pob_limpia.upper() not in ("NAN", "NONE", ""):
                provincia = "VALENCIA" if delegacion == "valencia" else "CASTELLON"
                direccion_completa = f"{dir_limpia}, {pob_limpia}, {provincia}, ESPAÑA"
                lat, lon = geocodificar(direccion_completa, api_key)
                print(f"Geocodificado: {direccion_completa} → {lat}, {lon}")

        # Fallback: libro de coordenadas por municipio
        if (lat is None or lon is None) and pueblo_norm in coords:
            lat, lon = coords[pueblo_norm]
            print(f"Fallback municipio: {pueblo_norm} → {lat}, {lon}")
        if pd.notna(lat) and pd.notna(lon):

            df.at[idx, "Latitud"] = lat
            df.at[idx, "Longitud"] = lon

            filas_con_coord.append((idx, float(lat), float(lon)))

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

        distancias.sort(key=lambda x: (x[0], str(x[1])))

        _, idx_sel, lat_sel, lon_sel = distancias[0]

        visitados.append(idx_sel)

        lat_actual = lat_sel
        lon_actual = lon_sel

        restantes = [r for r in restantes if r[0] != idx_sel]

    coords_ruta = [(df.loc[i, "Latitud"], df.loc[i, "Longitud"]) for i in visitados]
    coords_ruta = [(lat, lon) for lat, lon in coords_ruta if lat is not None and lon is not None]
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

    df_ordenado = df.loc[orden_final].copy()

    # Agrupar por calle dentro de cada población sin alterar orden de poblaciones
    df_ordenado["Calle_sin_num"] = df_ordenado["Dirección"].apply(extraer_calle_sin_numero)
    df_ordenado["Clave_parada"] = (
        df_ordenado["Población"].astype(str).str.strip().str.upper()
        + "|"
        + df_ordenado["Calle_sin_num"].str.upper()
    )

    # Asignar orden de población y orden de clave dentro de cada población
    poblaciones_orden = {p: i for i, p in enumerate(df_ordenado["Población"].unique())}
    df_ordenado["_ord_pob"] = df_ordenado["Población"].map(poblaciones_orden)

    claves_orden = {c: i for i, c in enumerate(df_ordenado["Clave_parada"].unique())}
    df_ordenado["_ord_clave"] = df_ordenado["Clave_parada"].map(claves_orden)

    df_ordenado = df_ordenado.sort_values(["_ord_pob", "_ord_clave"])
    df_ordenado = df_ordenado.drop(columns=["Calle_sin_num", "Clave_parada", "_ord_pob", "_ord_clave"])

    return df_ordenado


# -------------------------------------------------
# FUNCIÓN PRINCIPAL
# -------------------------------------------------
def reordenar_excel(
    input_path: Path,
    output_path: Path,
    ruta_coordenadas: Path,
    lat_origen: float,
    lon_origen: float,
    api_key: str = "",
    delegacion: str = "castellon",
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
                api_key=api_key,
                delegacion=delegacion,
            )
            
            link = generar_link_pueblos(df_ordenado, lat_origen, lon_origen)

            # Columna navegación al final (no rompe RESUMEN_UNICO)
            df_ordenado["NAVEGACIÓN"] = ""

            if link:
                df_ordenado.loc[df_ordenado.index[0], "NAVEGACIÓN"] = link

            hojas_resultado[nombre] = df_ordenado

        else:

            hojas_resultado[nombre] = df

    paradas_por_hoja = calcular_paradas_por_hoja(hojas_resultado)

    if "RESUMEN_UNICO" in hojas_resultado:
        df_res = hojas_resultado["RESUMEN_UNICO"].copy()
        if "Paradas" not in df_res.columns:
            df_res["Paradas"] = ""
        df_res["Paradas"] = df_res["Clave"].map(paradas_por_hoja).fillna(df_res["Paradas"])
        hojas_resultado["RESUMEN_UNICO"] = df_res

    # Construir enlaces de navegación por hoja
    hojas_navegacion = {}
    for nombre, df in hojas_resultado.items():
        if not nombre.startswith("ZREP_"):
            continue
        if "Latitud" not in df.columns or "Longitud" not in df.columns:
            continue
        link_completo = generar_link_pueblos(df, lat_origen, lon_origen)
        segmentos = generar_links_segmentos(df, lat_origen, lon_origen)
        hojas_navegacion[nombre] = {
            "link_completo": link_completo,
            "segmentos": segmentos
        }
        
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:

        for nombre, df in hojas_resultado.items():

            df.to_excel(writer, sheet_name=nombre, index=False)

    # Añadir filas de navegación al principio de cada hoja ZREP
    from openpyxl import load_workbook
    from openpyxl.styles import Font

    wb = load_workbook(output_path)

    for nombre, datos in hojas_navegacion.items():
        if nombre not in wb.sheetnames:
            continue
        ws = wb[nombre]
        ws.insert_rows(1, amount=len(datos) + 1)
        ws.cell(row=1, column=1).value = "RUTA COMPLETA"
        ws.cell(row=1, column=2).value = datos["link_completo"]
        ws.cell(row=1, column=2).font = Font(color="0000FF", underline="single")
        for i, link in enumerate(datos["segmentos"]):
            ws.cell(row=2 + i, column=1).value = f"SEGMENTO {i + 1}"
            ws.cell(row=2 + i, column=2).value = link
            ws.cell(row=2 + i, column=2).font = Font(color="0000FF", underline="single")

    wb.save(output_path)    
    return paradas_por_hoja
