from openpyxl import load_workbook

def generar_resumen_unico(ruta_excel: str) -> None:
    wb = load_workbook(ruta_excel)

    # BORRAR SI EXISTE
    if "RESUMEN_UNICO" in wb.sheetnames:
        del wb["RESUMEN_UNICO"]

    # INSERTAR EN POSICIÓN 1 (segundo lugar, índice 1)
    wb.create_sheet("RESUMEN_UNICO", index=1)

    wb.save(ruta_excel)
