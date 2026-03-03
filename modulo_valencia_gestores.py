# modulo_valencia_gestores.py

from datetime import datetime
from pathlib import Path
import pandas as pd
from openpyxl import load_workbook, Workbook


def _normalizar(texto: str) -> str:
    """Normaliza texto para cruces seguros sin alterar el original."""
    return " ".join(str(texto).strip().split())


def generar_libros_gestores(
    ruta_excel_final: str,
    ruta_asignacion: str,
    carpeta_salida: str
) -> dict:

    resultado = {
        "ok": False,
        "errores": [],
        "archivos_generados": {}
    }

    try:
        # -------------------------------------------------
        # 1. Validaciones básicas
        # -------------------------------------------------
        if not Path(ruta_excel_final).exists():
            resultado["errores"].append("No existe el Excel final validado.")
            return resultado

        if not Path(ruta_asignacion).exists():
            resultado["errores"].append("No existe el archivo gestor_zonas.xlsx.")
            return resultado

        Path(carpeta_salida).mkdir(parents=True, exist_ok=True)

        fecha_hoy = datetime.today().strftime("%Y-%m-%d")

        # -------------------------------------------------
        # 2. Detectar hojas territoriales
        # -------------------------------------------------
        wb_origen = load_workbook(ruta_excel_final, data_only=False)
        hojas_libro = wb_origen.sheetnames

        zonas_libro_raw = [h for h in hojas_libro if h.startswith("ZREP_")]

        if not zonas_libro_raw:
            resultado["errores"].append(
                "No se han encontrado hojas territoriales (ZREP_*) en el libro final."
            )
            return resultado

        # Mapa normalizado → nombre real hoja
        mapa_hojas = {_normalizar(z): z for z in zonas_libro_raw}
        zonas_libro_set = set(mapa_hojas.keys())

        # -------------------------------------------------
        # 3. Leer asignación
        # -------------------------------------------------
        df_asignacion = pd.read_excel(ruta_asignacion)

        if not {"ZONA_REP", "GESTOR"}.issubset(df_asignacion.columns):
            resultado["errores"].append(
                "El archivo gestor_zonas.xlsx debe contener columnas ZONA_REP y GESTOR."
            )
            return resultado

        df_asignacion["ZONA_REP"] = df_asignacion["ZONA_REP"].apply(_normalizar)
        df_asignacion["GESTOR"] = df_asignacion["GESTOR"].astype(str).str.strip()

        duplicadas = df_asignacion["ZONA_REP"][df_asignacion["ZONA_REP"].duplicated()]
        if not duplicadas.empty:
            resultado["errores"].append(
                f"Zonas duplicadas en gestor_zonas.xlsx: {list(duplicadas)}"
            )
            return resultado

        zonas_asignadas = set(df_asignacion["ZONA_REP"])
        mapa_zona_gestor = dict(zip(df_asignacion["ZONA_REP"], df_asignacion["GESTOR"]))
        gestores_detectados = sorted(df_asignacion["GESTOR"].unique())

        # -------------------------------------------------
        # 4. Validación cruzada
        # -------------------------------------------------
        zonas_sin_gestor = zonas_libro_set - zonas_asignadas
        if zonas_sin_gestor:
            resultado["errores"].append(
                f"Zonas sin asignación en gestor_zonas.xlsx: {list(zonas_sin_gestor)}"
            )
            return resultado

        zonas_inexistentes = zonas_asignadas - zonas_libro_set
        if zonas_inexistentes:
            resultado["errores"].append(
                f"Zonas asignadas que no existen en libro final: {list(zonas_inexistentes)}"
            )
            return resultado

        # -------------------------------------------------
        # 5. Generación por gestor
        # -------------------------------------------------
        for gestor in gestores_detectados:

            zonas_gestor_norm = [
                z for z in zonas_libro_set if mapa_zona_gestor[z] == gestor
            ]

            if not zonas_gestor_norm:
                continue

            wb_nuevo = Workbook()
            wb_nuevo.remove(wb_nuevo.active)

            dfs_todo = []

            for zona_norm in zonas_gestor_norm:

                zona_real = mapa_hojas[zona_norm]
                ws_origen = wb_origen[zona_real]
                ws_nueva = wb_nuevo.create_sheet(title=zona_real)

                datos = [list(row) for row in ws_origen.iter_rows(values_only=True)]
                if not datos:
                    continue

                encabezados = datos[0]
                filas = datos[1:]

                ws_nueva.append(encabezados)
                for fila in filas:
                    ws_nueva.append(fila)

                df_zona = pd.DataFrame(filas, columns=encabezados)
                df_zona["ZONA"] = zona_real
                dfs_todo.append(df_zona)

            # -------------------------------------------------
            # TODO
            # -------------------------------------------------
            df_todo = pd.concat(dfs_todo, ignore_index=True) if dfs_todo else pd.DataFrame()
            ws_todo = wb_nuevo.create_sheet(title="TODO")

            if not df_todo.empty:
                ws_todo.append(list(df_todo.columns))
                for _, row in df_todo.iterrows():
                    ws_todo.append(row.tolist())

            # -------------------------------------------------
            # RESUMEN_UNICO (Excel español)
            # -------------------------------------------------
            ws_resumen = wb_nuevo.create_sheet(title="RESUMEN_UNICO")

            ws_resumen["A1"] = "Total expediciones"
            ws_resumen["B1"] = "=CONTARA(TODO!A:A)-1"

            # Total Kgs
            if "Kgs" in df_todo.columns:
                col_kgs = df_todo.columns.get_loc("Kgs") + 1
                col_letter = ws_todo.cell(row=1, column=col_kgs).column_letter

                ws_resumen["A2"] = "Total Kgs"
                ws_resumen["B2"] = f"=SUMA(TODO!{col_letter}:{col_letter})"

            # Totales por zona
            ws_resumen["A4"] = "Totales por zona"

            if not df_todo.empty:
                col_zona = df_todo.columns.get_loc("ZONA") + 1
                col_zona_letter = ws_todo.cell(row=1, column=col_zona).column_letter

                zonas_unicas = sorted(df_todo["ZONA"].unique())
                fila_inicio = 5

                for i, zona in enumerate(zonas_unicas):
                    fila_actual = fila_inicio + i
                    ws_resumen[f"A{fila_actual}"] = zona
                    ws_resumen[f"B{fila_actual}"] = (
                        f'=CONTAR.SI(TODO!{col_zona_letter}:{col_zona_letter};"{zona}")'
                    )

            # -------------------------------------------------
            # Guardar archivo
            # -------------------------------------------------
            nombre_archivo = f"VALENCIA_{fecha_hoy}_{gestor}.xlsx"
            ruta_salida = Path(carpeta_salida) / nombre_archivo
            wb_nuevo.save(ruta_salida)

            resultado["archivos_generados"][gestor] = str(ruta_salida)

        resultado["ok"] = True
        return resultado

    except Exception as e:
        resultado["errores"].append(str(e))
        return resultado
