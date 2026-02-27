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
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.legend import Legend

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

    wb = Workbook()
    ws = wb.active
    ws.title = "Dashboard de Auditoría"
    ws.sheet_view.showGridLines = False

    # Estilos Institucionales
    fill_dark_blue = PatternFill(start_color="0F243E", end_color="0F243E", fill_type="solid")
    font_white_bold = Font(color="FFFFFF", bold=True, size=11)
    font_title = Font(color="FFFFFF", bold=True, size=16)
    align_center = Alignment(horizontal="center", vertical="center")
    thin_border = Border(left=Side(style='thin', color='BFBFBF'), right=Side(style='thin', color='BFBFBF'),
                         top=Side(style='thin', color='BFBFBF'), bottom=Side(style='thin', color='BFBFBF'))

    # Título Principal
    ws.merge_cells('B2:J2')
    title_cell = ws['B2']
    title_cell.value = "DICTAMEN EJECUTIVO DE AUDITORÍA PREVENTIVA (ART. 30-B CFF)"
    title_cell.fill = fill_dark_blue
    title_cell.font = font_title
    title_cell.alignment = align_center
    ws.row_dimensions[2].height = 35

    # Métricas (Un solo color azul sólido)
    metricas = [
        ("B4", "B5", "FACTURAS PROCESADAS", total_facturas),
        ("E4", "E5", "RIESGO FISCAL TOTAL", total_riesgo_monetario),
        ("H4", "H5", "ERRORES ESTRUCTURALES", total_errores)
    ]
    for h_cell, v_cell, text, val in metricas:
        ws[h_cell] = text
        ws[h_cell].fill = fill_dark_blue
        ws[h_cell].font = font_white_bold
        ws[h_cell].alignment = align_center
        
        ws[v_cell] = val
        ws[v_cell].fill = fill_dark_blue # Mismo azul solicitado
        ws[v_cell].font = font_white_bold
        ws[v_cell].alignment = align_center
        if "RIESGO" in text:
            ws[v_cell].number_format = '"$"#,##0.00_-'

    # Encabezados de Tabla (Sin abreviaturas)
    headers = [
        "Archivo", "UUID (Folio Fiscal)", "RFC Emisor", "Subtotal Base", 
        "IVA Declarado", "IVA Esperado", "ISR Declarado", "ISR Esperado", "Dictamen Legal"
    ]
    for col_idx, text in enumerate(headers, start=2):
        cell = ws.cell(row=8, column=col_idx, value=text)
        cell.fill = fill_dark_blue
        cell.font = font_white_bold
        cell.alignment = align_center
        cell.border = thin_border

    # Datos
    for row_idx, data in enumerate(resultados, start=9):
        for col_idx, value in enumerate(data, start=2):
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            cell.border = thin_border
            if 5 <= col_idx <= 9:
                cell.number_format = '"$"#,##0.00_-'
            
            if col_idx == 10: # Estilo dinámico de la columna Dictamen
                cell.alignment = align_center
                if "Sin discrepancias" in value:
                    cell.fill = PatternFill(start_color="D9EAD3", end_color="D9EAD3", fill_type="solid")
                    cell.font = Font(color="274E13", bold=True) # Negritas y color verde
                elif "RIESGO" in value:
                    cell.fill = PatternFill(start_color="F4CCCC", end_color="F4CCCC", fill_type="solid")
                    cell.font = Font(color="990000", bold=True) # Negritas y color rojo
                else:
                    cell.fill = PatternFill(start_color="FCE5CD", end_color="FCE5CD", fill_type="solid")
                    cell.font = Font(color="B45F06", bold=True) # Negritas y color naranja

    # Activar Filtros
    ws.auto_filter.ref = f"B8:J{8 + len(resultados)}"

    # Hoja de Datos para Gráfico
    ws_data = wb.create_sheet("DatosGrafico")
    grafico_data = [
        ["Estado", "Cantidad"],
        ["Sin discrepancias", count_ok],
        ["Riesgo Fiscal", count_riesgo],
        ["Error / Anomalía", count_error]
    ]
    for r in grafico_data:
        ws_data.append(r)

    # Configuración de Gráfico 3D Ejecutivo
    chart = PieChart3D()
    labels = Reference(ws_data, min_col=1, min_row=2, max_row=4)
    data_ref = Reference(ws_data, min_col=2, min_row=1, max_row=4)
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(labels)
    
    chart.title = "Distribución de Resultados Fiscales"
    chart.title.tx.rich.p[0].r[0].rPr = Font(size=1400, b=True) # Título grande y negritas (14pt)
    
    chart.legend = Legend()
    chart.legend.position = 'b' # Leyenda en la parte inferior (bottom)
    
    chart.dataLabels = DataLabelList()
    chart.dataLabels.showVal = True
    chart.dataLabels.showPercent = True

    ws.add_chart(chart, "L4")

    # Ajuste de Columnas
    for col_idx in range(2, 11):
        ws.column_dimensions[get_column_letter(col_idx)].width = 25
    ws.column_dimensions['A'].width = 3

    # Nombre de archivo dinámico
    timestamp = datetime.now().strftime("%d-%m-%Y_%H%Mhrs")
    filename_final = f"Dictamen_Ejecutivo_Art30B_{timestamp}.xlsx"

    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_xlsx:
        wb.save(tmp_xlsx.name)
        return FileResponse(tmp_xlsx.name, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename=filename_final)

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=int(os.environ.get("PORT", 8000)))
