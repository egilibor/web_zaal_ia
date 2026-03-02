from openpyxl import load_workbook
from openpyxl.styles import Font


def generar_resumen_unico(ruta_excel: str) -> None:

    wb = load_workbook(ruta_excel)

    # -------------------------------------------------
    # OBTENER O CREAR HOJA SIN ROMPER ORDEN
    # -------------------------------------------------

    if "RESUMEN_UNICO" in wb.sheetnames:
        ws = wb["RESUMEN_UNICO"]
        ws.delete_rows(1, ws.max_row)  # Limpiar contenido
    else:
        # Si no existe, crearla en posición 2
        ws = wb.create_sheet("RESUMEN_UNICO", index=1)

    # -------------------------------------------------
    # DETECTAR HOJAS OPERATIVAS
    # -------------------------------------------------

    operativas = []

    if "HOSPITALES" in wb.sheetnames:
        operativas.append("HOSPITALES")

    if "FEDERACION" in wb.sheetnames:
        operativas.append("FEDERACION")

    zrep = sorted([s for s in wb.sheetnames if s.startswith("ZREP_")])
    operativas.extend(zrep)

    # -------------------------------------------------
    # CABECERA
    # -------------------------------------------------

    ws.append(["Clave", "Expediciones", "Bultos", "Kilos"])

    for cell in ws[1]:
        cell.font = Font(bold=True)

    # -------------------------------------------------
    # FILAS DINÁMICAS
    # -------------------------------------------------

    for hoja in operativas:
        ws.append([
            hoja,
            f"=COUNTA('{hoja}'!A:A)-1",
            f"=SUM('{hoja}'!H:H)",
            f"=SUM('{hoja}'!G:G)"
        ])

    # -------------------------------------------------
    # AJUSTE DE ANCHO
    # -------------------------------------------------

    ws.column_dimensions["A"].width = 20
    ws.column_dimensions["B"].width = 15
    ws.column_dimensions["C"].width = 15
    ws.column_dimensions["D"].width = 15

    wb.calculation.fullCalcOnLoad = True 
    wb.save(ruta_excel)

