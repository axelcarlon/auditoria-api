import xml.etree.ElementTree as ET

class CFDIXMLExtractor:
    def __init__(self, xml_path):
        self.xml_path = xml_path
        self.ns = {
            'cfdi': 'http://www.sat.gob.mx/cfd/4',
            'tfd': 'http://www.sat.gob.mx/TimbreFiscalDigital'
        }

    def extract_data(self):
        try:
            tree = ET.parse(self.xml_path)
            root = tree.getroot()
            
            subtotal = float(root.attrib.get('SubTotal', 0.0))
            
            emisor = root.find('cfdi:Emisor', self.ns)
            rfc = emisor.attrib.get('Rfc', 'No encontrado') if emisor is not None else 'No encontrado'
            
            # Extracción del UUID (Folio Fiscal)
            uuid = 'No timbrado'
            complemento = root.find('cfdi:Complemento', self.ns)
            if complemento is not None:
                timbre = complemento.find('tfd:TimbreFiscalDigital', self.ns)
                if timbre is not None:
                    uuid = timbre.attrib.get('UUID', 'No timbrado')

            # Extracción de retenciones reales declaradas
            iva_declarado = 0.0
            isr_declarado = 0.0
            impuestos = root.find('cfdi:Impuestos', self.ns)
            if impuestos is not None:
                retenciones = impuestos.find('cfdi:Retenciones', self.ns)
                if retenciones is not None:
                    for retencion in retenciones.findall('cfdi:Retencion', self.ns):
                        impuesto_tipo = retencion.attrib.get('Impuesto')
                        importe = float(retencion.attrib.get('Importe', 0.0))
                        if impuesto_tipo == '002':  # Clave SAT para IVA
                            iva_declarado += importe
                        elif impuesto_tipo == '001': # Clave SAT para ISR
                            isr_declarado += importe
            
            return rfc, uuid, subtotal, iva_declarado, isr_declarado
        except Exception as e:
            return f"Error: {e}", "N/A", 0.0, 0.0, 0.0

    def validate_taxes(self, subtotal):
        iva_esperado = round(subtotal * 0.16, 2)
        isr_retenido_esperado = round(subtotal * 0.025, 2)
        return iva_esperado, isr_retenido_esperado