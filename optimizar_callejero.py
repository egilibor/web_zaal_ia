import pandas as pd
import requests


def matriz_ors(coords, api_key):
    url = "https://api.openrouteservice.org/v2/matrix/driving-car"

    headers = {
        "Authorization": api_key,
        "Content-Type": "application/json"
    }

    body = {
        "locations": coords,
        "metrics": ["duration"]
    }

    r = requests.post(url, json=body, headers=headers, timeout=20)

    data = r.json()

    if "durations" not in data:
        raise Exception(f"Error ORS: {data}")

    return data["durations"]


def optimizar_rutas_callejero(input_excel, output_excel, api_key):

    xls = pd.ExcelFile(input_excel)
    writer = pd.ExcelWriter(output_excel)

    for sheet in xls.sheet_names:

        df = pd.read_excel(xls, sheet)

        # si no hay coordenadas se copia la hoja sin tocar
        if "Latitud" not in df.columns or "Longitud" not in df.columns:
            df.to_excel(writer, sheet_name=sheet, index=False)
            continue

        # eliminar filas sin coordenadas
        df = df.dropna(subset=["Latitud", "Longitud"])

        if len(df) < 2:
            df.to_excel(writer, sheet_name=sheet, index=False)
            continue

        # formato que exige ORS → [lon, lat]
        coords = df[["Longitud", "Latitud"]].astype(float).values.tolist()

        # límite ORS
        MAX_PUNTOS = 50
        coords = coords[:MAX_PUNTOS]
        df = df.iloc[:MAX_PUNTOS]

        matriz = matriz_ors(coords, api_key)

        # orden simple desde el primer punto
        orden = sorted(
            range(len(coords)),
            key=lambda i: matriz[0][i] if matriz[0][i] is not None else 999999
        )

        df = df.iloc[orden]

        df.to_excel(writer, sheet_name=sheet, index=False)

    writer.close()
