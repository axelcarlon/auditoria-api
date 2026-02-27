import os
import tempfile
from datetime import datetime
from typing import List
from fastapi import FastAPI, File, UploadFile
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
from xml_extractor import CFDIXMLExtractor
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import PieChart3D, Reference

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
    # 1. Preparación de datos
    resultados = []
    total_facturas = 0
    total_riesgo_monetario = 0.0
    total_errores = 0
    count_ok = 0
    count_riesgo = 0
    count_error = 0

    for file in files:
        if not file.filename.lower().endswith('.xml'):
            continue
        
        total_facturas += 1
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xml") as tmp_xml:
            tmp_xml.write(await file.read())
            tmp_xml_path = tmp_xml.name

        try:
            extractor = CFDIXMLExtractor(tmp_xml_path)
            rfc, uuid, subtotal, iva_d, isr_d = extractor.extract_data()
            iva_e, isr_e = extractor.validate_taxes(subtotal)
            
            dif_iva = abs(iva_e - iva_d)
            dif_isr = abs(isr_e - isr_d)
            discrepancia_total = dif_iva + dif_isr

            dictamen = "Sin discrepancias"
            if discrepancia_total > 0.10:
                dictamen = "RIESGO FISCAL DETECTADO"
                total_riesgo_monetario += discrepancia_total
                count_riesgo += 1
            elif isinstance(rfc, str) and "Error" in rfc:
                dictamen = "ERROR ESTRUCTURAL"
                total_errores += 1
                count_error += 1
            elif subtotal == 0.0:
                dictamen = "ANOMALÍA: Sin subtotal"
                total_errores += 1
                count_error += 1
            else:
                count_ok += 1

            resultados.append([
                file.filename, uuid, rfc, subtotal, iva_d, iva_e, isr_d, isr_e, dictamen
            ])
        finally:
            if os.path.exists(tmp_xml_path):
                os.remove(tmp_xml_path)

    # 2. Construcción del Excel (Diseño fiel al archivo subido)
    wb = Workbook()
    ws = wb.active
    ws.title = "Dashboard de Auditoría"
    ws.sheet_view.showGridLines = False

    # Estilos
    fill_dark_blue = PatternFill(start_color="0F243E", end_color="0F243E", fill_type="solid")
    fill_metric_box = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    font_white_bold = Font(color="FFFFFF", bold=True)
    font_title = Font(color="FFFFFF", bold=True, size=14)
    align_center = Alignment(horizontal="center", vertical="center")
    thin_border = Border(left=Side(style='thin', color='BFBFBF'), right=Side(style='thin', color='BFBFBF'),
                         top=Side(style='thin', color='BFBFBF'), bottom=Side(style='thin', color='BFBFBF'))

    # Título (Fila 2)
    ws.merge_cells('B2:J2')
    title_cell = ws['B2']
    title_cell.value = "DICTAMEN EJECUTIVO DE AUDITORÍA PREVENTIVA (ART. 30-B CFF)"
    title_cell.fill = fill_dark_blue
    title_cell.font = font_title
    title_cell.alignment = align_center
    ws.row_dimensions[2].height = 25

    # Métricas (Filas 4 y 5)
    metricas_headers = [("B4", "FACTURAS PROCESADAS"), ("E4", "RIESGO FISCAL TOTAL"), ("H4", "ERRORES ESTRUCTURALES")]
    for cell_ref, text in metricas_headers:
        ws[cell_ref] = text
        ws[cell_ref].fill = fill_dark_blue
        ws[cell_ref].font = font_white_bold
        ws[cell_ref].alignment = align_center

    ws['B5'] = total_facturas
    ws['E5'] = total_riesgo_monetario
    ws['E5'].number_format = '"$"#,##0.00_-'
    ws['H5'] = total_errores
    for cell_ref in ['B5', 'E5', 'H5']:
        ws[cell_ref].fill = fill_metric_box
        ws[cell_ref].font = font_white_bold
        ws[cell_ref].alignment = align_center

    # Encabezados de Tabla (Fila 8)
    headers = ["Archivo", "UUID (Folio Fiscal)", "RFC Emisor", "Subtotal Base", "IVA Ret. (Decl.)", "IVA Esp.", "ISR Ret. (Decl.)", "ISR Esp.", "Dictamen Legal"]
    for col_idx, text in enumerate(headers, start=2):
        cell = ws.cell(row=8, column=col_idx, value=text)
        cell.fill = fill_dark_blue
        cell.font = font_white_bold
        cell.alignment = align_center
        cell.border = thin_border

    # Datos (Fila 9+)
    for row_idx, data in enumerate(resultados, start=9):
        for col_idx, value in enumerate(data, start=2):
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            cell.border = thin_border
            if 5 <= col_idx <= 9:
                cell.number_format = '"$"#,##0.00_-'
            if col_idx == 10: # Columna Dictamen
                cell.alignment = align_center
                if "Sin discrepancias" in value:
                    cell.fill = PatternFill(start_color="D9EAD3", end_color="D9EAD3", fill_type="solid")
                elif "RIESGO" in value:
                    cell.fill = PatternFill(start_color="F4CCCC", end_color="F4CCCC", fill_type="solid")
                else:
                    cell.fill = PatternFill(start_color="FCE5CD", end_color="FCE5CD", fill_type="solid")

    # Hoja de Gráficos
    ws_data = wb.create_sheet("DatosGrafico")
    grafico_data = [["Estado", "Cantidad"], ["Sin discrepancias", count_ok], ["Riesgo Fiscal", count_riesgo], ["Error / Anomalía", count_error]]
    for r in grafico_data:
        ws_data.append(r)

    # Añadir Gráfico al Dashboard
    chart = PieChart3D()
    labels = Reference(ws_data, min_col=1, min_row=2, max_row=4)
    data_ref = Reference(ws_data, min_col=2, min_row=1, max_row=4)
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(labels)
    chart.title = "Distribución de Resultados"
    ws.add_chart(chart, "L4")

    # Ajustes de columna
    for col_idx in range(2, 11):
        ws.column_dimensions[get_column_letter(col_idx)].width = 22
    ws.column_dimensions['A'].width = 3

    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_xlsx:
        wb.save(tmp_xlsx.name)
        return FileResponse(tmp_xlsx.name, filename=f"Dictamen_AuditorIA_{datetime.now().strftime('%Y%m%d')}.xlsx")

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=int(os.environ.get("PORT", 8000)))
