# WEB_ZAAL_IA

Aplicación Streamlit + Python para optimización de rutas de reparto de paquetería. Dos delegaciones: **Castellón** y **Valencia**.

## Arquitectura

- `app.py` — interfaz Streamlit con 5 fases + panel admin
- `reparto_gpt.py` — Fase 1: clasifica expediciones del CSV en hojas Excel por zona
- `reordenar_rutas.py` — Fase 3: geocodifica y ordena rutas con Google Routes API
- `geocodificador.py` — geocodificación con caché SQLite
- `add_resumen_unico.py` — genera hoja RESUMEN_UNICO en el Excel
- `modulo_valencia_gestores.py` — genera Excel por gestor de tráfico (solo Valencia)
- `auth.py` — autenticación y roles (admin / usuario estándar)

## Flujo de trabajo

1. **Fase 1** — sube CSV de llegadas → clasifica en HOSPITALES, FEDERACION, ZREP_* → geocodifica
2. **Fase 2** — ajuste manual: mover expediciones entre hojas o crear 2º reparto
3. **Fase 3** — ordena rutas con Google Routes API (dos pasadas: por CP y dentro de CP)
4. **Fase 4** — refino fino del orden: arrastrar expediciones o mover bloques
5. **Fase 5** — exportar KML por zona (solo admin)

## Algoritmo de ordenación (reordenar_rutas.py)

- **Primera pasada**: agrupa paradas por C.P., obtiene centroide desde archivo de referencia, ordena CPs con Routes API en modo ruta abierta (`circuito_cerrado=False`)
- **Segunda pasada**: dentro de cada CP ordena paradas individuales con Routes API en modo circuito cerrado
- Fallback euclidiano (nearest-neighbor) si la API falla o devuelve índices negativos
- Límite de 25 waypoints por llamada a la API

## Archivos de referencia

- `valencia_municipios_coordenadas.xlsx` — columnas: PUEBLO, CODPOS, LATITUD, LONGITUD
- `Libro_de_Servicio_Castellon_con_coordenadas.xlsx` — mismas columnas
- `Reglas_hospitales.xlsx` — hojas REGLAS_HOSPITALES y REGLAS_FEDERACION
- `gestor_zonas.xlsx` — asignación de zonas a gestores (solo Valencia)
- `calles_castellon.csv` — callejero para corrección de direcciones

## Stack

- Python 3.13, Streamlit, openpyxl, pandas, requests
- Google Routes API (`optimizeWaypointOrder: True`)
- Google Maps Geocoding API con caché SQLite
- Despliegue: Streamlit Cloud via GitHub (`egilibor/web_zaal_ia`, rama main)

## Convenciones

- No usar clases, solo funciones
- Rutas de archivo siempre con `pathlib.Path`
- Variables de Fase 4 con prefijo `_` para evitar colisiones
- API key en `st.secrets["GOOGLE_MAPS_API_KEY"]`
- Orígenes: Castellón `(39.804106, -0.217351)`, Valencia `(39.44069, -0.42589)`