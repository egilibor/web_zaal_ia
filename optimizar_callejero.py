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

    r = requests.post(url, json=body, headers=headers)

    data = r.json()

    if "durations" not in data:
        raise Exception(f"Error ORS: {data}")

    return data["durations"]

def optimizar_rutas_callejero(input_excel, output_excel, api_key):

    xls = pd.ExcelFile(input_excel)
    writer = pd.ExcelWriter(output_excel)

    for sheet in xls.sheet_names:

        df = pd.read_excel(xls, sheet)

        # Si no hay columnas de coordenadas, copiar hoja
        if "Latitud" not in df.columns or "Longitud" not in df.columns:
            df.to_excel(writer, sheet_name=sheet, index=False)
            continue

        # --- LIMPIEZA ROBUSTA DE COORDENADAS ---
        lat = (
            df["Latitud"]
            .astype(str)
            .str.replace(" ", "")
            .str.replace(",", ".", regex=False)
            .str.replace("[^0-9\\.-]", "", regex=True)
        )

        lon = (
            df["Longitud"]
            .astype(str)
            .str.replace(" ", "")
            .str.replace(",", ".", regex=False)
            .str.replace("[^0-9\\.-]", "", regex=True)
        )

        df["Latitud"] = pd.to_numeric(lat, errors="coerce")
        df["Longitud"] = pd.to_numeric(lon, errors="coerce")

        df = df.dropna(subset=["Latitud", "Longitud"])

        if len(df) < 2:
            raise Exception("No hay suficientes coordenadas válidas")

        coords = list(zip(df["Latitud"], df["Longitud"]))

        # límite ORS
        MAX_PUNTOS = 50
        coords = coords[:MAX_PUNTOS]
        df = df.iloc[:MAX_PUNTOS]

        matriz = matriz_ors(coords, api_key)

        # orden simple desde el primer punto
        orden = sorted(range(len(coords)), key=lambda i: matriz[0][i] if matriz[0][i] is not None else 999999)
        df = df.iloc[orden]

        df.to_excel(writer, sheet_name=sheet, index=False)

    writer.close()






