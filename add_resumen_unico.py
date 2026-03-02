from openpyxl import load_workbook
from openpyxl.styles import Font


def generar_resumen_unico(ruta_excel: str) -> None:
    wb = load_workbook(ruta_excel)

    # Reutilizar hoja si existe
    if "RESUMEN_UNICO" in wb.sheetnames:
        ws = wb["RESUMEN_UNICO"]
        ws.delete_rows(1, ws.max_row)
    else:
        ws = wb.create_sheet("RESUMEN_UNICO")

    # Detectar hojas operativas
    operativas = []

    if "HOSPITALES" in wb.sheetnames:
        operativas.append("HOSPITALES")

    if "FEDERACION" in wb.sheetnames:
        operativas.append("FEDERACION")

    zrep = sorted([s for s in wb.sheetnames if s.startswith("ZREP_")])
    operativas.extend(zrep)

    # Cabecera
    ws.append(["Clave", "Expediciones", "Bultos", "Kilos"])
    for cell in ws[1]:
        cell.font = Font(bold=True)

    # Filas dinámicas
    for hoja in operativas:
        ws.append([
            hoja,
            f"=COUNTA('{hoja}'!A:A)-1",
            f"=SUM('{hoja}'!H:H)",
            f"=SUM('{hoja}'!G:G)"
        ])

    # Ajuste ancho
    ws.column_dimensions["A"].width = 20
    ws.column_dimensions["B"].width = 15
    ws.column_dimensions["C"].width = 15
    ws.column_dimensions["D"].width = 15

        # Ajuste ancho
    ws.column_dimensions["A"].width = 20
    ws.column_dimensions["B"].width = 15
    ws.column_dimensions["C"].width = 15
    ws.column_dimensions["D"].width = 15

    # ---- FORZAR ORDEN DEFINITIVO ----
    orden_fijo = [
        "METADATOS",
        "RESUMEN_UNICO",
        "RESUMEN_GENERAL",
        "HOSPITALES",
        "FEDERACION",
        "RESUMEN_RUTAS_RESTO",
    ]

    # Añadir ZREP_* al final
    zrep = sorted([s for s in wb.sheetnames if s.startswith("ZREP_")])
    orden_fijo.extend(zrep)

    wb._sheets = [wb[name] for name in orden_fijo if name in wb.sheetnames]

    wb.save(ruta_excel)

