import requests
import logging
import json
import pytz
import base64
import re

from odoo import fields, models, api
from datetime import date
from odoo.exceptions import UserError

logger = logging.getLogger(__name__)

class Picking(models.Model):
    _inherit = "stock.picking"

    def get_config(self):
        config = self.env['elinvoice.configuration'].search([], limit=1)

        return {
            'api_url': config.api_url
        }

    def get_identification_document_type(self, document_type: str) -> str:

        documentType = ''

        if document_type == '36':
            documentType = 'NIT'
        
        if document_type == '13':
            documentType = 'DUI'
        
        if document_type == '37':
            documentType = 'Otro'

        if document_type == '03':
            documentType = 'Pasaporte'
        
        if document_type == '02':
            documentType = 'Carnet de residente'

        return documentType

    def format_phone(self, phone: str) -> str:
        phone = phone.replace('+503 ', '')
        phone = phone.replace(' ', '')

        return phone
    
    def agregar_adjunto_pdf(self, pdfContent, name):
        # pdf_base64 = base64.b64encode(pdfContent)
        # Crear el adjunto
        self.env['ir.attachment'].create({
            'name': name,              # Nombre del archivo
            'type': 'binary',               # Tipo de archivo
            'datas': pdfContent,            # Contenido codificado en Base64
            'res_model': 'sale.order',    # Modelo relacionado (facturas)
            'res_id': self.sale_id.id,              # ID del registro relacionado (factura actual)
            'mimetype': 'application/pdf',  # MIME type
        })

    def agregar_adjunto_json(self, jsonContent, name):
        # pdf_base64 = base64.b64encode(pdfContent)
        # Crear el adjunto
        self.env['ir.attachment'].create({
            'name': name,              # Nombre del archivo
            'type': 'binary',               # Tipo de archivo
            'datas': jsonContent,            # Contenido codificado en Base64
            'res_model': 'sale.order',    # Modelo relacionado (facturas)
            'res_id': self.sale_id.id,              # ID del registro relacionado (factura actual)
            'mimetype': 'application/json',  # MIME type
        })

    def create_confirmation(self):
        config = self.get_config()
        cf_confirmation = self.get_confirmation(config)
        data = cf_confirmation['data']

        if cf_confirmation:
            confirmation_data = {
                'sale_order_id': self.sale_id.id,
                'url_qr': data['dte_url'],
                'uid_de': data['uuid'],
                'received_seal': data['reception_seal'],
                'control_number': data['control_number'],
                'authorization_date': data['date'],
                'status': data['status'],
                'message': cf_confirmation['message'],
                'pdf_path': data['pdf_url'],
                'json_path': data['json_url'],
            }

            status = data['status']
            if status == 'PROCESADO':
                self.sale_id.is_processed = True # Si la repsuesta es "PROCESADO", cambiar a "true"
                pdfContent = data['data_pdf']
                if pdfContent != '-':
                    self.agregar_adjunto_pdf(pdfContent, data['control_number'] + '.pdf')
                    jsonContent = data['data_json']
                    if jsonContent != '-':
                        self.agregar_adjunto_json(jsonContent, data['control_number'] + '.json')

            new_confirmation = self.env['elinvoice.confirmation'].create(confirmation_data)
            return new_confirmation

        return None
    
    def get_confirmation(self, config):

        order = self.sale_id
        sequence = 'NRE/' + order.name
        doc_type = 'NRE'

        body = {}

        if doc_type == 'NRE': # Nota de remision
            body = self.makeNRE(sequence)

        logger.info(body)

        company = self.company_id
        url = config['api_url'] + '/dte/nr'
        headers = {
             'X-API-Key': company.api_key,
             'Content-Type': 'application/json'
        }
        
        response = requests.post(url=url, json=body, headers=headers)
        logger.info(url)
        logger.info(headers)
        logger.info(response)
        logger.info('### REQUEST DONE ###################')

        data = response.json()
        logger.info(data)
        return data

    def makeNRE(self, erp_code: str):
        
        fecha_actual = date.today()

        order = self.sale_id
        partner = order.partner_id
        lines = order.order_line

        items = []
        ivaRetention = 0.00

        for line in lines:
            if line.product_id.id > 0:

                if not line.product_id.default_code:
                    raise UserError('El producto: ' + str(line.product_id.name) + ', no tiene codigo interno')

                taxes = line.tax_id
                saleType = "EXENTA"

                if taxes:
                    for tax in taxes:
                        if 'IVA 13%' in tax.name:
                            saleType = "GRAVADA"

                        if 'RET 1%' in tax.name:
                            ivaRetention = ivaRetention + round(line.price_subtotal * 0.01, 2)

                discount = round(line.price_unit * line.product_uom_qty * (line.discount / 100), 2)
                items.append({
                    "type": "SERVICIOS",
                    "description": line.name,
                    "quantity": line.product_uom_qty,
                    "unitPrice": round(line.price_unit, 2), # round(line.price_unit - discount, 2),
                    "saleType": saleType,
                    "discountAmount": discount
                })

        documentType = self.get_identification_document_type(partner.tipodocumento)

        return {
            "erpCode": erp_code,
            "date": fecha_actual.strftime("%Y-%m-%d"),
            "paymentType": "CONTADO",
            "recipient": {
                "name": partner.name,
                "phone": self.format_phone(partner.phone),
                "email": partner.email,
                "address": {
                    "department": partner.departamento_id.code,
                    "municipality": partner.municipio_id.code,
                    "complement": partner.complemento,
                },
                "identificationDocument":{
                    "type": documentType,
                    "number": partner.numdocumento
                },
                "contributorType": partner.tipo_persona
            },
            "items": items,
            "retentionIva": ivaRetention,
            "retentionRenta": 0.00,
            "observation": order.note if order.note else ''
        }