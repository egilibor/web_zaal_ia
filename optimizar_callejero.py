import requests
import pandas as pd

def matriz_ors(coords, api_key):

    url = "https://api.openrouteservice.org/v2/matrix/driving-car"

    headers = {
        "Authorization": api_key,
        "Content-Type": "application/json"
    }

    locations = []

    for lat, lon in coords:

        try:
            lat = float(str(lat).replace(",", "."))
            lon = float(str(lon).replace(",", "."))
        except:
            continue

        locations.append([lon, lat])

    if len(locations) < 2:
        raise Exception("No hay suficientes coordenadas válidas")

    body = {
        "locations": locations,
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

        # si no hay coordenadas, copiar hoja sin tocar
        if "Latitud" not in df.columns or "Longitud" not in df.columns:
            df.to_excel(writer, sheet_name=sheet, index=False)
            continue

        # asegurar coordenadas válidas
        df["Latitud"] = pd.to_numeric(df["Latitud"], errors="coerce")
        df["Longitud"] = pd.to_numeric(df["Longitud"], errors="coerce")

        df = df.dropna(subset=["Latitud", "Longitud"])

        coords = list(zip(df["Latitud"], df["Longitud"]))

        MAX_PUNTOS = 50
        coords = coords[:MAX_PUNTOS]
        df = df.iloc[:MAX_PUNTOS]

        matriz = matriz_ors(coords, api_key)

        # orden simple inicial
        orden = sorted(range(len(coords)), key=lambda i: matriz[0][i])

        df = df.iloc[orden]

        df.to_excel(writer, sheet_name=sheet, index=False)

    writer.close()





