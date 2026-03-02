#def generar_resumen_unico(path):
#    pass

from openpyxl import load_workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


def generar_resumen_unico(ruta_excel: str) -> None:
    wb = load_workbook(ruta_excel)

    # Eliminar hoja existente si existe
    if "RESUMEN_UNICO" in wb.sheetnames:
        del wb["RESUMEN_UNICO"]

    # Detectar hojas operativas
    operativas = []

    if "HOSPITALES" in wb.sheetnames:
        operativas.append("HOSPITALES")

    if "FEDERACION" in wb.sheetnames:
        operativas.append("FEDERACION")

    zrep = sorted([s for s in wb.sheetnames if s.startswith("ZREP_")])
    operativas.extend(zrep)

    # Crear hoja nueva
    ws = wb.create_sheet("RESUMEN_UNICO")

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

    # Ajuste simple de ancho
    ws.column_dimensions["A"].width = 20
    ws.column_dimensions["B"].width = 15
    ws.column_dimensions["C"].width = 15
    ws.column_dimensions["D"].width = 15

    wb.save(ruta_excel)






