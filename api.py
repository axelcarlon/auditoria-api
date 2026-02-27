import os
import tempfile
from datetime import datetime
from fastapi import FastAPI, File, UploadFile
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
from xml_extractor import CFDIXMLExtractor
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import PieChart3D, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.drawing.line import LineProperties
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.chart.marker import DataPoint
from openpyxl.drawing.text import Paragraph, ParagraphProperties, CharacterProperties
from openpyxl.chart.text import RichText
from openpyxl.drawing.colors import ColorChoice

app = FastAPI()

# Permite que páginas web externas (como tu sitio en Softr) envíen archivos a esta API
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.post("/api/analizar")
async def analizar_factura(file: UploadFile = File(...)):
    # 1. Guardar el XML recibido temporalmente en el servidor
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xml") as tmp_xml:
        tmp_xml.write(await file.read())
        tmp_xml_path = tmp_xml.name

    try:
        # 2. Extraer datos
        extractor = CFDIXMLExtractor(tmp_xml_path)
        rfc, uuid, subtotal, iva_declarado, isr_declarado = extractor.extract_data()

        dictamen = "Sin discrepancias"
        errores = 0
        riesgos = 0
        ok_count = 0
        iva_esperado = 0.0
        isr_esperado = 0.0
        total_riesgo_iva = 0.0
        total_riesgo_isr = 0.0

        if isinstance(rfc, str) and "Error" in rfc:
            dictamen = "ERROR ESTRUCTURAL: Archivo corrupto"
            errores = 1
        elif subtotal == 0.0:
            dictamen = "ANOMALÍA: Sin subtotal gravable"
            errores = 1
        else:
            iva_esperado, isr_esperado = extractor.validate_taxes(subtotal)
            dif_iva = abs(iva_esperado - iva_declarado)
            dif_isr = abs(isr_esperado - isr_declarado)

            if dif_iva > 0.10 or dif_isr > 0.10:
                dictamen = "RIESGO FISCAL DETECTADO"
                riesgos = 1
                total_riesgo_iva = dif_iva
                total_riesgo_isr = dif_isr
            else:
                ok_count = 1

        # 3. Construir Excel
        wb = Workbook()
        ws = wb.active
        ws.title = "Dictamen Individual"
        ws.sheet_view.showGridLines = False

        fill_dark_blue = PatternFill(start_color="0F243E", end_color="0F243E", fill_type="solid")
        fill_metric_box = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        font_title = Font(color="FFFFFF", bold=True, size=16)
        align_center = Alignment(horizontal="center", vertical="center")
        thin_border = Border(left=Side(style='thin', color='BFBFBF'),
                             right=Side(style='thin', color='BFBFBF'),
                             top=Side(style='thin', color='BFBFBF'),
                             bottom=Side(style='thin', color='BFBFBF'))

        ws.merge_cells('B2:I2')
        title_cell = ws['B2']
        title_cell.value = f"DICTAMEN EJECUTIVO DE AUDITORÍA - {file.filename}"
        title_cell.fill = fill_dark_blue
        title_cell.font = font_title
        title_cell.alignment = align_center
        ws.row_dimensions[2].height = 30

        # Tabla de Datos
        start_row = 5
        encabezados = [
            "Archivo", "UUID (Folio Fiscal)", "RFC Emisor", "Subtotal Base", 
            "IVA Retenido (Declarado)", "IVA Esperado (Cálculo)", 
            "ISR Retenido (Declarado)", "ISR Esperado (Cálculo)", "Dictamen Legal"
        ]
        
        for col_idx, titulo in enumerate(encabezados, start=2):
            cell = ws.cell(row=start_row, column=col_idx, value=titulo)
            cell.fill = fill_dark_blue
            cell.font = Font(color="FFFFFF", bold=True)
            cell.alignment = align_center
            cell.border = thin_border

        fila_data = [
            file.filename, uuid, rfc, subtotal if isinstance(subtotal, float) else 0.0,
            iva_declarado, iva_esperado, isr_declarado, isr_esperado, dictamen
        ]

        current_row = start_row + 1
        for col_idx, valor in enumerate(fila_data, start=2):
            cell = ws.cell(row=current_row, column=col_idx, value=valor)
            cell.border = thin_border
            cell.alignment = Alignment(vertical="center")
            if 5 <= col_idx <= 9:
                cell.number_format = '"$"#,##0.00_-'

        dictamen_cell = ws.cell(row=current_row, column=10)
        dictamen_cell.alignment = align_center
        if "Sin discrepancias" in dictamen:
            dictamen_cell.fill = PatternFill(start_color="D9EAD3", end_color="D9EAD3", fill_type="solid")
            dictamen_cell.font = Font(color="274E13", bold=True)
        elif "RIESGO" in dictamen:
            dictamen_cell.fill = PatternFill(start_color="F4CCCC", end_color="F4CCCC", fill_type="solid")
            dictamen_cell.font = Font(color="990000", bold=True)
        else:
            dictamen_cell.fill = PatternFill(start_color="FCE5CD", end_color="FCE5CD", fill_type="solid")
            dictamen_cell.font = Font(color="B45F06", bold=True)

        for col_idx in range(2, 11):
            column_letter = get_column_letter(col_idx)
            ws.column_dimensions[column_letter].width = 25
        ws.column_dimensions['A'].width = 3

        # 4. Guardar Excel temporalmente
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_xlsx:
            excel_filename = tmp_xlsx.name
            wb.save(excel_filename)

        # 5. Enviar respuesta al navegador
        timestamp = datetime.now().strftime("%d-%m-%Y_%H%Mhrs")
        return FileResponse(
            excel_filename, 
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", 
            filename=f"Dictamen_{timestamp}.xlsx"
        )

    finally:
        os.remove(tmp_xml_path)

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)