#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import sqlite3
import googlemaps
from pathlib import Path

DB_PATH = Path(__file__).parent / "geocache.db"

def _get_connection():
    conn = sqlite3.connect(DB_PATH)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS geocache (
            direccion TEXT PRIMARY KEY,
            latitud   REAL,
            longitud  REAL
        )
    """)
    conn.commit()
    return conn

def geocodificar(direccion: str, api_key: str) -> tuple:
    
    if not direccion or str(direccion).strip().upper() in ("NAN", "NONE", ""):
        return (None, None)

    direccion_norm = direccion.strip().upper()

    # Consultar caché
    conn = _get_connection()
    row = conn.execute(
        "SELECT latitud, longitud FROM geocache WHERE direccion = ?",
        (direccion_norm,)
    ).fetchone()

    if row:
        conn.close()
        return (row[0], row[1])

    # Llamar a la API
    try:
        gmaps = googlemaps.Client(key=api_key)
        result = gmaps.geocode(direccion_norm, region="es", language="es")

        if result:
            lat = result[0]["geometry"]["location"]["lat"]
            lon = result[0]["geometry"]["location"]["lng"]

            conn.execute(
                "INSERT INTO geocache (direccion, latitud, longitud) VALUES (?, ?, ?)",
                (direccion_norm, lat, lon)
            )
            conn.commit()
            conn.close()
            return (lat, lon)

    except Exception as e:
        print(f"Error geocodificando '{direccion}': {e}")

    conn.close()
    return (None, None)
