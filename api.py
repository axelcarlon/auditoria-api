import os
import tempfile
from datetime import datetime
from typing import List
from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
from xml_extractor import CFDIXMLExtractor
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# La variable DEBE llamarse 'app' para que coincida con el comando de Render
app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.post("/api/analizar")
async def analizar_facturas(files: List[UploadFile] = File(...)):
    wb = Workbook()
    ws = wb.active
    ws.title = "Dictamen Consolidado"
    ws.sheet_view.showGridLines = False

    fill_dark_blue = PatternFill(start_color="0F243E", end_color="0F243E", fill_type="solid")
    font_title = Font(color="FFFFFF", bold=True, size=16)
    align_center = Alignment(horizontal="center", vertical="center")
    thin_border = Border(
        left=Side(style='thin', color='BFBFBF'),
        right=Side(style='thin', color='BFBFBF'),
        top=Side(style='thin', color='BFBFBF'),
        bottom=Side(style='thin', color='BFBFBF')
    )

    ws.merge_cells('B2:J2')
    title_cell = ws['B2']
    title_cell.value = f"DICTAMEN EJECUTIVO DE AUDITORÍA MASIVA - {len(files)} CONCEPTOS"
    title_cell.fill = fill_dark_blue
    title_cell.font = font_title
    title_cell.alignment = align_center
    ws.row_dimensions[2].height = 30

    encabezados = [
        "Archivo", "UUID (Folio Fiscal)", "RFC Emisor", "Subtotal Base", 
        "IVA Retenido (Declarado)", "IVA Esperado (Cálculo)", 
        "ISR Retenido (Declarado)", "ISR Esperado (Cálculo)", "Dictamen Legal"
    ]
    
    start_row = 5
    for col_idx, titulo in enumerate(encabezados, start=2):
        cell = ws.cell(row=start_row, column=col_idx, value=titulo)
        cell.fill = fill_dark_blue
        cell.font = Font(color="FFFFFF", bold=True)
        cell.alignment = align_center
        cell.border = thin_border

    current_row = start_row + 1
    for file in files:
        if not file.filename.lower().endswith('.xml'):
            continue

        with tempfile.NamedTemporaryFile(delete=False, suffix=".xml") as tmp_xml:
            tmp_xml.write(await file.read())
            tmp_xml_path = tmp_xml.name

        try:
            extractor = CFDIXMLExtractor(tmp_xml_path)
            rfc, uuid, subtotal, iva_declarado, isr_declarado = extractor.extract_data()
            iva_esperado, isr_esperado = extractor.validate_taxes(subtotal)
            
            dif_iva = abs(iva_esperado - iva_declarado)
            dif_isr = abs(isr_esperado - isr_declarado)

            dictamen = "Sin discrepancias"
            if dif_iva > 0.10 or dif_isr > 0.10:
                dictamen = "RIESGO FISCAL DETECTADO"
            if isinstance(rfc, str) and "Error" in rfc:
                dictamen = "ERROR ESTRUCTURAL"
            elif subtotal == 0.0:
                dictamen = "ANOMALÍA: Sin subtotal"

            fila_data = [file.filename, uuid, rfc, subtotal, iva_declarado, iva_esperado, isr_declarado, isr_esperado, dictamen]

            for col_offset, valor in enumerate(fila_data):
                cell = ws.cell(row=current_row, column=col_offset + 2, value=valor)
                cell.border = thin_border
                cell.alignment = Alignment(vertical="center")
                if 5 <= (col_offset + 2) <= 9:
                    cell.number_format = '"$"#,##0.00_-'
                if (col_offset + 2) == 10:
                    cell.alignment = align_center
                    if "Sin discrepancias" in dictamen:
                        cell.fill = PatternFill(start_color="D9EAD3", end_color="D9EAD3", fill_type="solid")
                        cell.font = Font(color="274E13", bold=True)
                    elif "RIESGO" in dictamen:
                        cell.fill = PatternFill(start_color="F4CCCC", end_color="F4CCCC", fill_type="solid")
                        cell.font = Font(color="990000", bold=True)
                    else:
                        cell.fill = PatternFill(start_color="FCE5CD", end_color="FCE5CD", fill_type="solid")
                        cell.font = Font(color="B45F06", bold=True)
            current_row += 1
        finally:
            if os.path.exists(tmp_xml_path):
                os.remove(tmp_xml_path)

    for col_idx in range(2, 11):
        ws.column_dimensions[get_column_letter(col_idx)].width = 25
    ws.column_dimensions['A'].width = 3

    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_xlsx:
        output_name = tmp_xlsx.name
        wb.save(output_name)

    return FileResponse(output_name, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename=f"Auditoria_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx")
