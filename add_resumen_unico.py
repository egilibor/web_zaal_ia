from openpyxl import load_workbook
from openpyxl.styles import Font


def generar_resumen_unico(ruta_excel: str) -> None:
    wb = load_workbook(ruta_excel)

    # 1️⃣ Eliminar hoja si ya existe (sin reutilizar)
    if "RESUMEN_UNICO" in wb.sheetnames:
        del wb["RESUMEN_UNICO"]

    # 2️⃣ Crear justo después de METADATOS
    if "METADATOS" in wb.sheetnames:
        idx_meta = wb.sheetnames.index("METADATOS")
        ws = wb.create_sheet("RESUMEN_UNICO", index=idx_meta + 1)
    else:
        ws = wb.create_sheet("RESUMEN_UNICO")

    # 3️⃣ Detectar hojas operativas
    operativas = []

    if "HOSPITALES" in wb.sheetnames:
        operativas.append("HOSPITALES")

    if "FEDERACION" in wb.sheetnames:
        operativas.append("FEDERACION")

    zrep = sorted([s for s in wb.sheetnames if s.startswith("ZREP_")])
    operativas.extend(zrep)

    # 4️⃣ Cabecera
    ws.append(["Clave", "Expediciones", "Bultos", "Kilos"])
    for cell in ws[1]:
        cell.font = Font(bold=True)

    # 5️⃣ Filas dinámicas
    for hoja in operativas:
        ws.append([
            hoja,
            f"=COUNTA('{hoja}'!A:A)-1",
            f"=SUM('{hoja}'!H:H)",
            f"=SUM('{hoja}'!G:G)"
        ])

    # 6️⃣ Ajuste ancho columnas
    ws.column_dimensions["A"].width = 20
    ws.column_dimensions["B"].width = 15
    ws.column_dimensions["C"].width = 15
    ws.column_dimensions["D"].width = 15

    wb.save(ruta_excel)
