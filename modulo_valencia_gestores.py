# modulo_valencia_gestores.py

from datetime import datetime
from pathlib import Path
import pandas as pd
from openpyxl import load_workbook, Workbook


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
        zonas_libro = [h for h in hojas_libro if h.startswith("ZREP_")]

        if not zonas_libro:
            resultado["errores"].append(
                "No se han encontrado hojas territoriales (ZREP_*) en el libro final."
            )
            return resultado

        # -------------------------------------------------
        # 3. Leer asignación
        # -------------------------------------------------
        df_asignacion = pd.read_excel(ruta_asignacion)

        if not {"ZONA_REP", "GESTOR"}.issubset(df_asignacion.columns):
            resultado["errores"].append(
                "El archivo gestor_zonas.xlsx debe contener columnas ZONA_REP y GESTOR."
            )
            return resultado

        df_asignacion["ZONA_REP"] = df_asignacion["ZONA_REP"].astype(str).str.strip()
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
        zonas_libro_set = set(zonas_libro)

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

            zonas_gestor = [
                z for z in zonas_libro if mapa_zona_gestor[z] == gestor
            ]

            if not zonas_gestor:
                continue

            wb_nuevo = Workbook()
            wb_nuevo.remove(wb_nuevo.active)

            dfs_todo = []

            for zona in zonas_gestor:
                ws_origen = wb_origen[zona]
                ws_nueva = wb_nuevo.create_sheet(title=zona)

                datos = []
                for row in ws_origen.iter_rows(values_only=True):
                    datos.append(list(row))

                if not datos:
                    continue

                encabezados = datos[0]
                filas = datos[1:]

                # Escribir hoja zona
                ws_nueva.append(encabezados)
                for fila in filas:
                    ws_nueva.append(fila)

                # Construir dataframe para TODO
                df_zona = pd.DataFrame(filas, columns=encabezados)
                df_zona["ZONA"] = zona
                dfs_todo.append(df_zona)

            # -------------------------------------------------
            # Construir TODO
            # -------------------------------------------------
            if dfs_todo:
                df_todo = pd.concat(dfs_todo, ignore_index=True)
            else:
                df_todo = pd.DataFrame()

            ws_todo = wb_nuevo.create_sheet(title="TODO")

            if not df_todo.empty:
                ws_todo.append(list(df_todo.columns))
                for _, row in df_todo.iterrows():
                    ws_todo.append(row.tolist())

            # -------------------------------------------------
            # Construir RESUMEN_UNICO (con fórmulas)
            # -------------------------------------------------
            ws_resumen = wb_nuevo.create_sheet(title="RESUMEN_UNICO")

            ws_resumen["A1"] = "Total expediciones"
            ws_resumen["B1"] = "=COUNTA(TODO!A:A)-1"

            # Buscar columna Kgs
            if "Kgs" in df_todo.columns:
                col_kgs = df_todo.columns.get_loc("Kgs") + 1
                col_letter = ws_todo.cell(row=1, column=col_kgs).column_letter

                ws_resumen["A2"] = "Total Kgs"
                ws_resumen["B2"] = f"=SUM(TODO!{col_letter}:{col_letter})"

            # Totales por zona
            ws_resumen["A4"] = "Totales por zona"

            if not df_todo.empty:
                zonas_unicas = sorted(df_todo["ZONA"].unique())
                fila_inicio = 5

                for i, zona in enumerate(zonas_unicas):
                    fila_actual = fila_inicio + i
                    ws_resumen[f"A{fila_actual}"] = zona
                    ws_resumen[f"B{fila_actual}"] = (
                        f'=COUNTIF(TODO!{ws_todo["A1"].column_letter}:{ws_todo["A1"].column_letter},"*")'
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
