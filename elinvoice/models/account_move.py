import requests
import logging
import json
import pytz
import base64
import re

from odoo import fields, models, api
from datetime import datetime
from odoo.exceptions import UserError

# contains all the supported document types (this is determined by the sequence)
accepted_doc_types = ['CCFE', 'FE', 'FEXE', 'FSEE', 'NCE', 'NDE', 'CCFEP']
# contains the doc numbers mapped to th document type
doc_type_values = {
    'CCFE': '03',
    'FE'  : '01',
    'FEXE': '11',
    'FSEE': '14',
    'NCE' : '05',
    'NDE' : '06',
    'CCFEP' : '07',
}
# contains the url mapped to the document type
doc_type_urls = {
    'CCFE': '/dte/ccf',
    'FE'  : '/dte/fc',
    'FEXE': '/dte/fex',
    'FSEE': '/dte/fse',
    'NCE' : '/dte/nc',
    'NDE' : '/dte/nd',
    'CCFEP' : '/dte/cr',
}

def get_current_date():
    date_utc = datetime.now(pytz.utc)
    timezone = pytz.timezone('America/Mexico_City')
    date = date_utc.astimezone(timezone)

    return date.strftime('%Y-%m-%dT%H:%M:%S.%fZ')

class AccountMove(models.Model):
    _inherit = "account.move"

    ###################################
    # DATABASE FIELDS
    ####################################

    confirmation_id = fields.One2many('elinvoice.confirmation', 'account_move_id', string='Confirmacion relacionada')

    purchase_doc_number = fields.Char(string="Documento de Compra", required=False)
    purchase_doc_resolution = fields.Char(string="Resolucion Documento Compra", required=False)
    purchase_doc_serie = fields.Char(string="Serie Documento Compra", required=False)

    purchase_doc_date = fields.Date(string="Fecha de documento de Compra", required=False)
    purchase_doc_generation = fields.Selection(selection=[('1', 'Físico'), 
                                                         ('2', 'Electrónico'),],
                                            string="Tipo de generación Compra", required=False)

    purchase_doc_type = fields.Selection(selection=[('03', 'Credito Fiscal'),
                                                    ('01', 'Factura'),
                                                    ('14', 'Factura de sujeto excluido'),],
                                        string="Tipo de documento de Compra",
                                        required=False)

    numero_control = fields.Char(string="Numero de control", required=False)
    codigo_generacion = fields.Char(string="Codigo de generacion", required=False)
    sello_recepcion = fields.Char(string="Sello de recepcion", required=False)
    tipo_documento = fields.Selection(
        selection=[
            ('1', 'Físico'),
            ('2', 'Electrónico')
        ],
        string="Tipo de documento",
        required=False
    )

    ###
    # Campo para saber si el documento ya fue procesado en Zaapdos en MH
    ###
    is_processed = fields.Boolean(string="Is Processed", default=False)

    ###
    # Campo para saber si el documento ya fue anulado en Zaapdos en MH
    ###
    is_anulated = fields.Boolean(string="Is Anulated", default=False)
    
    processed_status_mh = fields.Char(
        string="Estado de Procesamiento",
        compute="_compute_processed_status",
        store=True
    )

    @api.depends('is_processed', 'is_anulated')
    def _compute_processed_status(self):
        
        for record in self:

            mhStatus = 'No aplicado a MH'

            if record.is_processed:
                mhStatus = 'Procesado MH'

            if record.is_anulated:
                mhStatus = 'Anulado MH'

            record.processed_status_mh = mhStatus

    tipo_ingreso_renta_move = fields.Selection(
        selection=[
            ('0', 'N/A'),
            ('1', 'Profesiones, Artes y Oficios'),
            ('2', 'Actividades de Servicios'),
            ('3', 'Actividades Comerciales'),
            ('4', 'Actividades Industriales'),
            ('5', 'Actividades Agropecuarias'),
            ('6', 'Utilidades y Dividendos'),
            ('7', 'Exportaciones de bienes'),
            ('8', 'Servicios Realizados en el Exterior y Utilizados en El Salvador'),
            ('9', 'Exportaciones de servicios'),
            ('10', 'Otras Rentas Gravables'),
            ('12', 'Ingresos que ya fueron sujetos de retención en F910'),
            ('13', 'Sujetos pasivos excluidos (art. 6 LISR) e ingresos que no constituyen hecho generador del ISR')
        ],
        string="Tipo de ingreso (Renta)",
        required=False,
        default='0'
    )
    
    tipo_operacion_renta_move = fields.Selection(
        selection=[
            ('0', 'N/A'),
            ('1', 'Gravada'),
            ('2', 'No Gravada o Exenta'),
            ('3', 'Excluido o no Constituye Renta'),
            ('4', 'Mixta (Gravada y Exenta)'),
            ('12', 'Ingresos que ya fueron sujetos de retencion en F910'),
            ('13', 'Sujetos pasivos excluidos (art. 6 LISR) e ingresos que no constituyen hecho generador del ISR')
        ],
        string="Tipo de operacion (Renta)",
        required=False,
        default='0'
    )
    
    # @api.model
    # def create(self, vals):
    #     record = super(AccountMove, self).create(vals)

    #     tipoOperacionRenta = 0
    #     tipoIngresoRenta = 0

    #     # Determinar el "Tipo de operacion (Renta)"
    #     gravada_renta = 0
    #     exenta_renta = 0
    #     excluida_renta = 0

    #     lines = record.line_ids

    #     for line in lines:
    #         if line.product_id.default_code:

    #             # determinar si la operacion es gravada en renta
    #             if str(line.tipo_operacion_renta) == '1':
    #                 gravada_renta = 1
                
    #             # determinar si la operacion es exenta en renta
    #             if str(line.tipo_operacion_renta) == '2':
    #                 exenta_renta = 1
                
    #             # determinar si la operacion es excluida en renta
    #             if str(line.tipo_operacion_renta) == '3':
    #                 excluida_renta = 1

    #     # Determinar si es gravada en renta
    #     if gravada_renta == 1 and exenta_renta == 0 and excluida_renta == 0:
    #         tipoOperacionRenta = 1
        
    #     # Determinar si es exenta en renta
    #     if gravada_renta == 0 and exenta_renta == 1 and excluida_renta == 0:
    #         tipoOperacionRenta = 2
        
    #     # Determinar si es excluida en renta
    #     if gravada_renta == 0 and exenta_renta == 0 and excluida_renta == 1:
    #         tipoOperacionRenta = 3

    #     # Determinar si es mixta en renta
    #     if (gravada_renta == 1 and exenta_renta == 1) or (excluida_renta == 1 and (gravada_renta == 1 or exenta_renta == 1)):
    #         tipoOperacionRenta = 4

    #     record.tipo_operacion_renta_move= str(tipoOperacionRenta)

    #     # Determinar el "Tipo de ingreso (Renta)"

    #     # obtener el "tipo_ingreso_renta_conf" que se tiene configurado por defecto
    #     tipoIngresoRentaConf = record.company_id.tipo_ingreso_renta_conf

    #     sql = """
    #         SELECT 
    #         pt.tipo_ingreso_renta AS tipo_ingreso_renta,
    #         SUM(aml.quantity * aml.price_unit) AS amount
    #         FROM account_move_line AS aml
    #         INNER JOIN product_product AS pp ON aml.product_id = pp.id
    #         INNER JOIN product_template AS pt ON pp.product_tmpl_id = pt.id
    #         WHERE aml.move_id = %s
    #         GROUP BY pt.tipo_ingreso_renta
    #         ORDER BY SUM(aml.quantity * aml.price_unit) ASC
    #     """

    #     self.env.cr.execute(sql, (record.id,))
    #     result = self.env.cr.fetchall()

    #     if result:
    #         # obtener el ultimo elemento que siempre será el mayor
    #         ultimoElemento = result[(len(result) -1)]
    #         tipoIngresoRenta = ultimoElemento[0]

    #         for row in result:
    #             if row[1] == ultimoElemento[1]:
    #                 # Verificar si alguno de los elementos cumple con el de configuracion
    #                 if row[0] == tipoIngresoRentaConf:
    #                     tipoIngresoRenta = row[0]
    #                     break
    #                 else:
    #                     primerElemento = result[0]
    #                     tipoIngresoRenta = primerElemento[0]
    #             else:
    #                 break
        
    #     record.tipo_ingreso_renta_move= str(tipoIngresoRenta)

    #     return record
    
    # misma funcionalidad anterior pero en el evento "Confirmar"
    def action_post(self):        
        res = super(AccountMove, self).action_post()

        records = self

        if records:
            for record in records:
                
                record.ensure_one()

                tipoOperacionRenta = 0
                tipoIngresoRenta = 0

                # Determinar el "Tipo de operacion (Renta)"
                gravada_renta = 0
                exenta_renta = 0
                excluida_renta = 0

                lines = record.line_ids

                for line in lines:
                    if line.product_id.default_code:

                        # determinar si la operacion es gravada en renta
                        if str(line.tipo_operacion_renta) == '1':
                            gravada_renta = 1
                        
                        # determinar si la operacion es exenta en renta
                        if str(line.tipo_operacion_renta) == '2':
                            exenta_renta = 1
                        
                        # determinar si la operacion es excluida en renta
                        if str(line.tipo_operacion_renta) == '3':
                            excluida_renta = 1

                # Determinar si es gravada en renta
                if gravada_renta == 1 and exenta_renta == 0 and excluida_renta == 0:
                    tipoOperacionRenta = 1
                
                # Determinar si es exenta en renta
                if gravada_renta == 0 and exenta_renta == 1 and excluida_renta == 0:
                    tipoOperacionRenta = 2
                
                # Determinar si es excluida en renta
                if gravada_renta == 0 and exenta_renta == 0 and excluida_renta == 1:
                    tipoOperacionRenta = 3

                # Determinar si es mixta en renta
                if (gravada_renta == 1 and exenta_renta == 1) or (excluida_renta == 1 and (gravada_renta == 1 or exenta_renta == 1)):
                    tipoOperacionRenta = 4

                record.tipo_operacion_renta_move= str(tipoOperacionRenta)

                # Determinar el "Tipo de ingreso (Renta)"

                # obtener el "tipo_ingreso_renta_conf" que se tiene configurado por defecto
                tipoIngresoRentaConf = record.company_id.tipo_ingreso_renta_conf

                sql = """ SELECT 
                    pt.tipo_ingreso_renta AS tipo_ingreso_renta,
                    SUM(aml.quantity * aml.price_unit) AS amount
                    FROM account_move_line AS aml
                    INNER JOIN product_product AS pp ON aml.product_id = pp.id
                    INNER JOIN product_template AS pt ON pp.product_tmpl_id = pt.id
                    WHERE aml.move_id = %s
                    GROUP BY pt.tipo_ingreso_renta
                    ORDER BY SUM(aml.quantity * aml.price_unit) ASC """

                self.env.cr.execute(sql, (record.id,))
                result = self.env.cr.fetchall()

                if result:
                    # obtener el ultimo elemento que siempre será el mayor
                    ultimoElemento = result[(len(result) -1)]
                    tipoIngresoRenta = ultimoElemento[0]

                    for row in result:
                        if row[1] == ultimoElemento[1]:
                            # Verificar si alguno de los elementos cumple con el de configuracion
                            if row[0] == tipoIngresoRentaConf:
                                tipoIngresoRenta = row[0]
                                break
                            else:
                                primerElemento = result[0]
                                tipoIngresoRenta = primerElemento[0]
                        else:
                            break
                
                record.tipo_ingreso_renta_move= str(tipoIngresoRenta)
                
        return res
    ###################################
    # Integracion de Odoo-Zaapdos desde POS.
    ###################################

    # get_invoice_pdf_report_attachment - Odoo 17
    # action_invoice_download_pdf - Odoo 18
    def action_invoice_download_pdf(self):
        logger = logging.getLogger(__name__)
        
        # Ejecutar método original
        res = super(AccountMove, self).action_invoice_download_pdf()

        logger.info('##### APLICA A MH #####')
        apply = self.create_confirmation()

        return res
    
    ###################################
    # MISC
    ###################################
    def agregar_adjunto_pdf(self, pdfContent, name):
        # pdf_base64 = base64.b64encode(pdfContent)
        # Crear el adjunto
        self.env['ir.attachment'].create({
            'name': name,              # Nombre del archivo
            'type': 'binary',               # Tipo de archivo
            'datas': pdfContent,            # Contenido codificado en Base64
            'res_model': 'account.move',    # Modelo relacionado (facturas)
            'res_id': self.id,              # ID del registro relacionado (factura actual)
            'mimetype': 'application/pdf',  # MIME type
        })

    def agregar_adjunto_json(self, jsonContent, name):
        # pdf_base64 = base64.b64encode(pdfContent)
        # Crear el adjunto
        self.env['ir.attachment'].create({
            'name': name,              # Nombre del archivo
            'type': 'binary',               # Tipo de archivo
            'datas': jsonContent,            # Contenido codificado en Base64
            'res_model': 'account.move',    # Modelo relacionado (facturas)
            'res_id': self.id,              # ID del registro relacionado (factura actual)
            'mimetype': 'application/json',  # MIME type
        })

    def get_doc_type(self, sequence: str) -> str:
        data = sequence.split('/')
        return data[0]

    def get_config(self):
        config = self.env['elinvoice.configuration'].search([], limit=1)

        return {
            'api_url': config.api_url,
        }
    
    def format_phone(self, phone: str) -> str:
        phone = phone.replace('+503 ', '')
        phone = phone.replace(' ', '')

        return phone
    
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
    
    ###################################
    # CONFIRMATION PROCESS
    ###################################

    def create_confirmation(self):
        config = self.get_config()
        cf_confirmation = self.get_confirmation(config)
        data = cf_confirmation['data']

        if cf_confirmation:
            confirmation_data = {
                'account_move_id': self.id,
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
                self.is_processed = True # Si la repsuesta es "PROCESADO", cambiar a "true"
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
        logger = logging.getLogger(__name__)

        sequence = self.name
        doc_type = self.journal_id.code # self.get_doc_type(sequence)

        if doc_type not in accepted_doc_types:
            return False

        body = {}

        if doc_type == 'CCFE': # Creadito Fiscal
            body = self.makeCCFE(sequence)

        if doc_type == 'FE': # Factura Consumidor Final
            body = self.makeFE(sequence)

        if doc_type == 'FEXE': # Factura de exportacion
            body = self.makeFEXE(sequence)

        if doc_type == 'FSEE': # Factura sujeto excluido
            body = self.makeFSEE(sequence)

        if doc_type == 'NCE': # Nota de credito
            body = self.makeNCE(sequence)

        if doc_type == 'NDE': # Nota de debito
            body = self.makeNDE(sequence)

        if doc_type == 'CCFEP': # Comprobante de retencion
            body = self.makeCRE(sequence)

        logger.info(body)

        company = self.company_id
        url = config['api_url'] + doc_type_urls[doc_type]
        headers = {
            'X-API-Key': company.api_key,
            'Content-Type': 'application/json'
        }
        
        response = requests.post(url=url, json=body, headers=headers)
        logger.info('### REQUEST DONE ###################')

        data = response.json()
        logger.info(data)
        return data

    ###################################
    # DOCUMENT TYPES MAKERS
    ###################################

    def makeFE(self, erp_code: str):
        partner = self.partner_id
        items = []
        ivaRetention = round(0, 4)

        # resumen
        totalNoGravado = 0
        reteRenta = 0
        descuNoSuj = 0
        descuExenta = 0
        descuGravada = 0
        totalDescu = 0
        totalNoSuj = 0
        totalExenta = 0
        totalGravada = 0

        subTotal = 0
        subTotalVentas = 0
        montoTotalOperacion = 0
        totalPagar = 0
        totalIva = 0
        ivaRete1 = 0

        for line in self.line_ids:
            # Calculas los impuestos de Odoo
            if line.display_type == 'tax':

                if line.name == 'IVA 13%':
                    if line.balance < 0:
                        totalIva = round((line.balance * -1), 2)
                    else:
                        totalIva = round((line.balance), 2)

                if line.name == 'RET 1%':
                    if line.balance < 0:
                        ivaRete1 = round((line.balance * -1), 2)
                    else:
                        ivaRete1 = round((line.balance), 2)

            if line.product_id.id > 0 and line.display_type == 'product':

                if not line.product_id.default_code:
                    raise UserError('El producto: ' + str(line.product_id.name) + ', no tiene codigo interno')

                taxes = line.tax_ids
                exenta = True

                if taxes:
                    for tax in taxes:
                        if 'IVA 13%' in tax.name:
                            exenta = False

                a = '[' + line.product_id.default_code + '] ' + line.product_id.name # nombre del producto en el modelo "product_template"
                b = line.name # descripcion del producto en el modelo "account_move_line"
                c = b.replace(a, "").strip() # Tomar solo la descripcion del producto del campo "name" del modelo "account_move_line"

                items.append({
                    "type": "SERVICIOS",
                    "description": c if c else b, # mandar la descripcion del producto y la tiene y sino el nombre del producto
                    "quantity": round(line.quantity, 8),
                    "unitPrice": round(line.price_unit, 8),
                    "saleType": "EXENTA" if exenta else "GRAVADA",
                    "discountAmount": round(line.discount, 8)
                })

        documentType = self.get_identification_document_type(partner.tipodocumento)

        identificationDocument = {}

        if documentType and partner.numdocumento:
            identificationDocument = {
                "type": documentType,
                "number": partner.numdocumento
            }

        # resumen
        subTotal = self.amount_untaxed + totalIva
        subTotalVentas = self.amount_untaxed + totalIva
        montoTotalOperacion = self.amount_untaxed + totalIva
        totalPagar = self.amount_untaxed + totalIva - ivaRete1

        return {
            "erpCode": erp_code,
            "date": self.invoice_date.strftime("%Y-%m-%d"),
            "paymentType": self.invoice_payment_term_id.mh_code,
            "recipient": {
                "nrc": partner.nrc if partner.nrc else '' ,
                "economicActivity": partner.actividad_economica_id.code if partner.actividad_economica_id.code else '',
                "identificationDocument": identificationDocument,
                "name": partner.name,
                "phone": self.format_phone(partner.phone),
                "email": partner.email,
                "address": {
                    "department": partner.departamento_id.code,
                    "municipality": partner.municipio_id.code,
                    "complement": partner.complemento,
                },
                "contributorType": partner.tipo_persona
            },
            "items": items,
            "itemsNoGrav": [],
            "mhResumen": {
                "subTotal": round(subTotal, 2),
                "subTotalVentas": round(subTotalVentas, 2),
                "montoTotalOperacion": round(montoTotalOperacion, 2),
                "totalPagar": round(totalPagar, 2),
                "totalIva": round(totalIva, 2),
                "ivaRete1": round(ivaRete1, 2),
                # "totalNoGravado": round(totalNoGravado, 2),
                # "reteRenta": round(reteRenta, 2),
                # "descuNoSuj": round(descuNoSuj, 2),
                # "descuExenta": round(descuExenta, 2),
                # "descuGravada": round(descuGravada, 2),
                # "totalDescu": round(totalDescu, 2),
                # "totalNoSuj": round(totalNoSuj, 2),
                # "totalExenta": round(totalExenta, 2),
                # "totalGravada": round(totalGravada, 2),
            },
            "retentionIva": ivaRete1,
            "retentionRenta": 0.00,
            "observation": self.narration if self.narration else ''
        }

    def makeCCFE(self, erp_code: str):
        partner = self.partner_id
        items = []
        ivaRetention = round(0, 4)

        # resumen
        totalNoGravado = 0
        reteRenta = 0
        descuNoSuj = 0
        descuExenta = 0
        descuGravada = 0
        totalDescu = 0
        totalNoSuj = 0
        totalExenta = 0
        totalGravada = 0

        subTotal = 0
        subTotalVentas = 0
        montoTotalOperacion = 0
        totalPagar = 0
        totalIva = 0
        ivaRete1 = 0

        for line in self.line_ids:
            # Calculas los impuestos de Odoo
            if line.display_type == 'tax':

                if line.name == 'IVA 13%':
                    if line.balance < 0:
                        totalIva = round((line.balance * -1), 2)
                    else:
                        totalIva = round((line.balance), 2)

                if line.name == 'RET 1%':
                    if line.balance < 0:
                        ivaRete1 = round((line.balance * -1), 2)
                    else:
                        ivaRete1 = round((line.balance), 2)

            if line.product_id.id > 0 and line.display_type == 'product':

                if not line.product_id.default_code:
                    raise UserError('El producto: ' + str(line.product_id.name) + ', no tiene codigo interno')

                taxes = line.tax_ids
                saleType = "EXENTA"

                if taxes:
                    for tax in taxes:
                        if 'IVA 13%' in tax.name:
                            saleType = "GRAVADA"

                a = '[' + line.product_id.default_code + '] ' + line.product_id.name # nombre del producto en el modelo "product_template"
                b = line.name # descripcion del producto en el modelo "account_move_line"
                c = b.replace(a, "").strip() # Tomar solo la descripcion del producto del campo "name" del modelo "account_move_line"

                items.append({
                    "type": "SERVICIOS",
                    "description": c if c else b, # mandar la descripcion del producto y la tiene y sino el nombre del producto
                    "quantity": round(line.quantity, 8),
                    "unitPrice": round(line.price_unit, 8),
                    "saleType": saleType,
                    "discountAmount": round(line.discount, 8)
                })

        documentType = self.get_identification_document_type(partner.tipodocumento)

        # resumen
        subTotal = self.amount_untaxed
        subTotalVentas = self.amount_untaxed
        montoTotalOperacion = self.amount_untaxed + totalIva
        totalPagar = self.amount_untaxed + totalIva - ivaRete1

        return {
            "erpCode": erp_code,
            "date": self.invoice_date.strftime("%Y-%m-%d"),
            "paymentType": self.invoice_payment_term_id.mh_code,
            "recipient": {
                "nrc": partner.nrc,
                "economicActivity": partner.actividad_economica_id.code,
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
            "itemsNoGrav": [],
            "mhResumen": {
                "subTotal": round(subTotal, 2),
                "subTotalVentas": round(subTotalVentas, 2),
                "tributos": [
                    {
                        "valor": totalIva,
                        "codigo": "20",
                        "descripcion": "Impuesto al Valor Agregado 13%"
                    }
                ],
                "montoTotalOperacion": round(montoTotalOperacion, 2),
                "totalPagar": round(totalPagar, 2),
                "ivaRete1": round(ivaRete1, 2),
                "ivaPerci1": 0,
                # "totalNoGravado": round(totalNoGravado, 2),
                # "reteRenta": round(reteRenta, 2),
                # "descuNoSuj": round(descuNoSuj, 2),
                # "descuExenta": round(descuExenta, 2),
                # "descuGravada": round(descuGravada, 2),
                # "totalDescu": round(totalDescu, 2),
                # "totalNoSuj": round(totalNoSuj, 2),
                # "totalExenta": round(totalExenta, 2),
                # "totalGravada": round(totalGravada, 2),
            },
            "retentionIva": ivaRete1,
            "retentionRenta": 0.00,
            "observation": self.narration if self.narration else ''
        }

    def makeFEXE(self, erp_code: str):
        partner = self.partner_id
        items = []

        total_discount = round(0, 4)
        for line in self.line_ids:
            if line.product_id.id > 0 and line.display_type == 'product':

                if not line.product_id.default_code:
                    raise UserError('El producto: ' + str(line.product_id.name) + ', no tiene codigo interno')

                a = '[' + line.product_id.default_code + '] ' + line.product_id.name # nombre del producto en el modelo "product_template"
                b = line.name # descripcion del producto en el modelo "account_move_line"
                c = b.replace(a, "").strip() # Tomar solo la descripcion del producto del campo "name" del modelo "account_move_line"

                discount = round(line.price_unit * line.quantity * (line.discount / 100), 4)
                items.append({
                    "description": c if c else b, # mandar la descripcion del producto y la tiene y sino el nombre del producto
                    "quantity": round(line.quantity, 8),
                    "unitPrice": round(line.price_unit, 8),
                    "discountAmount": round(discount, 8)
                })

                total_discount = round((total_discount + discount), 4)

        documentType = self.get_identification_document_type(partner.tipodocumento)

        return {
            "erpCode": erp_code,
            "date": self.invoice_date.strftime("%Y-%m-%d"),
            "paymentType": self.invoice_payment_term_id.mh_code,
            "itemType": "SERVICIOS",
            "recipient": {
                "name": partner.name,
                "phone": self.format_phone(partner.phone),
                "email": partner.email,
                "country": partner.pais_id.code,
                "address": partner.complemento,
                "contributorType": partner.tipo_persona,
                "economicActivity": partner.actividad_economica_id.code,
                "identificationDocument":{
                    "type": documentType,
                    "number": partner.numdocumento
                }
            },
            "items": items,
            # "discountj": round(total_discount, 2),
            "retentionIva": 0.00,
            "retentionRenta": 0.00,
            "observation": self.narration if self.narration else ''
        }

    def makeFSEE(self, erp_code: str):
        partner = self.partner_id
        items = []
        retentionRenta = round(0, 4)

        total_discount = round(0, 4)
        for line in self.line_ids:
            if line.product_id.id > 0 and line.display_type == 'product':

                if not line.product_id.default_code:
                    raise UserError('El producto: ' + str(line.product_id.name) + ', no tiene codigo interno')

                taxes = line.tax_ids

                if taxes:
                    for tax in taxes:
                        if 'RENT 10%' in tax.name:
                            retentionRenta = round((retentionRenta + round(line.price_subtotal * 0.10, 4)), 4)

                a = '[' + line.product_id.default_code + '] ' + line.product_id.name # nombre del producto en el modelo "product_template"
                b = line.name # descripcion del producto en el modelo "account_move_line"
                c = b.replace(a, "").strip() # Tomar solo la descripcion del producto del campo "name" del modelo "account_move_line"

                discount = round(line.price_unit * line.quantity * (line.discount / 100), 4)
                items.append({
                    "type": "SERVICIOS",
                    "description": c if c else b, # mandar la descripcion del producto y la tiene y sino el nombre del producto
                    "quantity": round(line.quantity, 8),
                    "unitPrice": round(line.price_unit, 8),
                    # "saleType": "GRAVADA",
                    "discountAmount": round(discount, 8)
                })

                total_discount = round((total_discount + discount), 4)

        documentType = self.get_identification_document_type(partner.tipodocumento)

        return {
            "erpCode": erp_code,
            "date": self.invoice_date.strftime("%Y-%m-%d"),
            "paymentType": self.invoice_payment_term_id.mh_code,
            "recipient": {
                # "nrc": partner.nrc,                
                "name": partner.name,
                "phone": self.format_phone(partner.phone),
                "email": partner.email,
                "address": {
                    "department": partner.departamento_id.code,
                    "municipality": partner.municipio_id.code,
                    "complement": partner.complemento,
                },
                "economicActivity": partner.actividad_economica_id.code,
                "identificationDocument":{
                    "type": documentType,
                    "number": partner.numdocumento
                },
                "contributorType": partner.tipo_persona
            },
            "items": items,
            # "discount": round(total_discount, 2)
            "retentionIva": 0.00,
            "retentionRenta": retentionRenta,
            "observation": self.narration if self.narration else ''
        }

    def makeNCE(self, erp_code: str):
        logger = logging.getLogger(__name__)
        logger.info("###### NOTA DE CREDITO ######")

        partner = self.partner_id
        items = []
        ivaRetention = round(0, 4)

        for line in self.line_ids:
            if line.product_id.id > 0 and line.display_type == 'product':

                if not line.product_id.default_code:
                    raise UserError('El producto: ' + str(line.product_id.name) + ', no tiene codigo interno')

                taxes = line.tax_ids
                saleType = "EXENTA"

                if taxes:
                    for tax in taxes:
                        if 'IVA 13%' in tax.name:
                            saleType = "GRAVADA"

                        if 'RET 1%' in tax.name:
                            ivaRetention = round((ivaRetention + round(line.price_subtotal * 0.01, 4)), 4)

                a = '[' + line.product_id.default_code + '] ' + line.product_id.name # nombre del producto en el modelo "product_template"
                b = line.name # descripcion del producto en el modelo "account_move_line"
                c = b.replace(a, "").strip() # Tomar solo la descripcion del producto del campo "name" del modelo "account_move_line"

                discount = round(line.price_unit * line.quantity * (line.discount / 100), 4)
                items.append({
                    "type": "SERVICIOS",
                    "description": c if c else b, # mandar la descripcion del producto y la tiene y sino el nombre del producto
                    "quantity": round(line.quantity, 8),
                    "unitPrice": round(line.price_unit, 8), # round(line.price_unit - discount, 2),
                    "saleType": saleType,
                    "discountAmount": discount
                })

        documentType = self.get_identification_document_type(partner.tipodocumento)

        # obtener documento relacionado
        dte_uuid = ''
        for dte in self.reversed_entry_id.confirmation_id:
            if dte.status == 'PROCESADO':
                dte_uuid = dte.uid_de
                break

        return {
            "erpCode": erp_code,
            "date": self.invoice_date.strftime("%Y-%m-%d"),
            "paymentType": self.invoice_payment_term_id.mh_code,
            "recipient": {
                "nrc": partner.nrc,
                "economicActivity": partner.actividad_economica_id.code,
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
            "relatedTaxDocument": dte_uuid,
            "items": items,
            "retentionIva": ivaRetention,
            "retentionRenta": 0.00,
            "observation": self.narration if self.narration else ''
        }

    def makeNDE(self, erp_code: str):
        logger = logging.getLogger(__name__)
        logger.info("###### NOTA DE DEBITO ######")

        partner = self.partner_id
        items = []
        ivaRetention = round(0, 4)

        for line in self.line_ids:
            if line.product_id.id > 0 and line.display_type == 'product':

                if not line.product_id.default_code:
                    raise UserError('El producto: ' + str(line.product_id.name) + ', no tiene codigo interno')

                taxes = line.tax_ids
                saleType = "EXENTA"

                if taxes:
                    for tax in taxes:
                        if 'IVA 13%' in tax.name:
                            saleType = "GRAVADA"

                        if 'RET 1%' in tax.name:
                            ivaRetention = round((ivaRetention + round(line.price_subtotal * 0.01, 4)), 4)

                a = '[' + line.product_id.default_code + '] ' + line.product_id.name # nombre del producto en el modelo "product_template"
                b = line.name # descripcion del producto en el modelo "account_move_line"
                c = b.replace(a, "").strip() # Tomar solo la descripcion del producto del campo "name" del modelo "account_move_line"

                discount = round(line.price_unit * line.quantity * (line.discount / 100), 4)
                items.append({
                    "type": "SERVICIOS",
                    "description": c if c else b, # mandar la descripcion del producto y la tiene y sino el nombre del producto
                    "quantity": round(line.quantity, 8),
                    "unitPrice": round(line.price_unit, 8), # round(line.price_unit - discount, 2),
                    "saleType": saleType,
                    "discountAmount": discount
                })

        documentType = self.get_identification_document_type(partner.tipodocumento)

        # obtener documento relacionado
        dte_uuid = ''
        for dte in self.reversed_entry_id.confirmation_id:
            if dte.status == 'PROCESADO':
                dte_uuid = dte.uid_de
                break

        return {
            "erpCode": erp_code,
            "date": self.invoice_date.strftime("%Y-%m-%d"),
            "paymentType": self.invoice_payment_term_id.mh_code,
            "recipient": {
                "nrc": partner.nrc,
                "economicActivity": partner.actividad_economica_id.code,
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
            "relatedTaxDocument": dte_uuid,
            "items": items,
            "retentionIva": ivaRetention,
            "retentionRenta": 0.00,
            "observation": self.narration if self.narration else ''
        }

    def makeCRE(self, erp_code: str):
        partner = self.partner_id
        items = []
        for line in self.line_ids:
            if line.product_id.id > 0 and line.display_type == 'product':

                if not line.product_id.default_code:
                    raise UserError('El producto: ' + str(line.product_id.name) + ', no tiene codigo interno')

                a = '[' + line.product_id.default_code + '] ' + line.product_id.name # nombre del producto en el modelo "product_template"
                b = line.name # descripcion del producto en el modelo "account_move_line"
                c = b.replace(a, "").strip() # Tomar solo la descripcion del producto del campo "name" del modelo "account_move_line"
 
                # discount = round(line.price_unit * line.quantity * (line.discount / 100), 2)
                items.append({
                    # "documentNumber": self.purchase_doc_number,
                    "description": c if c else b, # mandar la descripcion del producto y la tiene y sino el nombre del producto
                    "amountSubject": round(line.price_unit, 4),
                    # "ivaRetained": round(line.price_unit * 0.1, 2),
                    "retentionCode": "RETENCION_IVA_1"
                })

        documentType = self.get_identification_document_type(partner.tipodocumento)

        # validar el tipo de documento de compra
        taxDocumentType = 'FC'
        if self.purchase_doc_type == '01':
            taxDocumentType = 'FC'
        
        if self.purchase_doc_type == '03':
            taxDocumentType = 'CCF'
        
        if self.purchase_doc_type == '14':
            taxDocumentType = 'FSE'

        # validar el tipo de generacio del documento de compra
        crDocumentType = 'FISICO'
        if self.purchase_doc_generation == '1':
            crDocumentType = 'FISICO'
        
        if self.purchase_doc_type == '2':
            crDocumentType = 'DIGITAL'
        
        return {
            "erpCode": erp_code,
            "date": self.invoice_date.strftime("%Y-%m-%d"),
            "paymentType": self.invoice_payment_term_id.mh_code,
            "recipient": {
                "nrc": partner.nrc,
                "economicActivity": partner.actividad_economica_id.code,
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
            "documentType": crDocumentType,
            "taxDocumentType": taxDocumentType,
            "taxDocumentNumber": self.purchase_doc_number,
            "taxDocumentGeneratedAt": self.purchase_doc_date.strftime("%Y-%m-%d"),
            "items": items,
            "retentionIva": 0.00,
            "retentionRenta": 0.00,
            "observation": self.narration if self.narration else ''
        }

    ###################################
    # ANULATION PROCESS
    ###################################

    def get_last_confirmation_uid(self):
        search_domain = [
            ('account_move_id', '=', self.id),
            ('status', '=', '2'),
        ]

        confirmation = self.env['elinvoice.confirmation'].search(search_domain, order='id desc', limit=1)
        return confirmation.uid_de
    
    def get_nce_related_doc(self, ref):
        pattern = r"CCFE/\d{4}/\d{5}"
        match = re.search(pattern, ref)

        sequence = match.group(0)
        secuence_domain = [
            ('name', '=', sequence),
        ]

        move = self.env['account.move'].search(secuence_domain, order='id desc', limit=1)

        confirm_domain = [
            ('account_move_id', '=', move.id),
            ('status', '=', '2'),
        ]

        confirmation = self.env['elinvoice.confirmation'].search(confirm_domain, order='id desc', limit=1)
        return {
            'control_number': confirmation.uid_de,
            'authorization_date': confirmation.authorization_date.split('T')[0]
        }

    def revert_document(self):
        config = self.get_config()
        cf_confirmation = self.cancel_document(config)
        data = cf_confirmation['data']

        if cf_confirmation:
            confirmation_data = {
                'account_move_id': self.id,
                'url_qr': '-',
                'uid_de': data['uuid'],
                'received_seal': data['reception_seal'],
                'control_number': data['control_number'],
                'authorization_date': data['date'],
                'status': data['status'],
                'message': cf_confirmation['message'],
            }

            status = data['status']
            if status == 'ANULADO':
                self.is_anulated = True # Si la repsuesta es "ANULADO", cambiar a "true"

            new_confirmation = self.env['elinvoice.confirmation'].create(confirmation_data)
            return new_confirmation

        return None

    def cancel_document(self, config):
        logger = logging.getLogger(__name__)
        erpCode = self.name

        #company = self.company_id
        # doc_type = self.get_doc_type(erpCode)

        # if doc_type not in accepted_doc_types:
        #     return False

        company = self.company_id
        url = config['api_url'] + '/dte/invalidate'
        headers = {
            'X-API-Key': company.api_key,
            'Content-Type': 'application/json'
        }

        body = {
            "erpCode": erpCode
        }

        logger.info(body)
        response = requests.post(url=url, json=body, headers=headers)
        logger.info('### ANULATION DONE ###################')

        data = response.json()
        logger.info(data)
        return data

MAX_NUMERO = 999999999999

UNIDADES = (
    'cero',
    'uno',
    'dos',
    'tres',
    'cuatro',
    'cinco',
    'seis',
    'siete',
    'ocho',
    'nueve'
)

DECENAS = (
    'diez',
    'once',
    'doce',
    'trece',
    'catorce',
    'quince',
    'dieciseis',
    'diecisiete',
    'dieciocho',
    'diecinueve'
)

DIEZ_DIEZ = (
    'cero',
    'diez',
    'veinte',
    'treinta',
    'cuarenta',
    'cincuenta',
    'sesenta',
    'setenta',
    'ochenta',
    'noventa'
)

CIENTOS = (
    '_',
    'ciento',
    'doscientos',
    'trescientos',
    'cuatroscientos',
    'quinientos',
    'seiscientos',
    'setecientos',
    'ochocientos',
    'novecientos'
)

def numero_a_letras(numero):
    numero_entero = int(numero)
    if numero_entero > MAX_NUMERO:
        raise OverflowError('Número demasiado alto')
    if numero_entero < 0:
        return 'menos %s' % numero_a_letras(abs(numero))
    letras_decimal = ''
    parte_decimal = int(round((abs(numero) - abs(numero_entero)) * 100))
    if parte_decimal > 9:
        letras_decimal = 'con %s/100' % parte_decimal
    elif parte_decimal < 9:
        letras_decimal = 'con 0%s/100' % parte_decimal
    if (numero_entero <= 99):
        resultado = leer_decenas(numero_entero)
    elif (numero_entero <= 999):
        resultado = leer_centenas(numero_entero)
    elif (numero_entero <= 999999):
        resultado = leer_miles(numero_entero)
    elif (numero_entero <= 999999999):
        resultado = leer_millones(numero_entero)
    else:
        resultado = leer_millardos(numero_entero)
    resultado = resultado.replace('uno mil', 'un mil')
    resultado = resultado.strip()
    resultado = resultado.replace(' _ ', ' ')
    resultado = resultado.replace('  ', ' ')

    resultado = '%s %s' % (resultado, letras_decimal)
    resultado = resultado.replace('ciento con', 'cien con')
    resultado = resultado.replace('veinticero', 'veinte')

    return resultado.upper()

def leer_decenas(numero):
    if numero < 10:
        return UNIDADES[numero]
    decena, unidad = divmod(numero, 10)
    if numero <= 19:
        resultado = DECENAS[unidad]
    elif numero <= 29:
        resultado = 'veinti%s' % UNIDADES[unidad]
    else:
        resultado = DIEZ_DIEZ[decena]
        if unidad > 0:
            resultado = '%s y %s' % (resultado, UNIDADES[unidad])
    return resultado

def leer_centenas(numero):
    centena, decena = divmod(numero, 100)
    if numero == 0:
        resultado = 'cien'
    else:
        resultado = CIENTOS[centena]
        if decena > 0:
            resultado = '%s %s' % (resultado, leer_decenas(decena))
    return resultado

def leer_miles(numero):
    millar, centena = divmod(numero, 1000)
    resultado = ''
    if (millar == 1):
        resultado = ''
    if (millar >= 2) and (millar <= 9):
        resultado = UNIDADES[millar]
    elif (millar >= 10) and (millar <= 99):
        resultado = leer_decenas(millar)
    elif (millar >= 100) and (millar <= 999):
        resultado = leer_centenas(millar)
    resultado = '%s mil' % resultado
    if centena > 0:
        resultado = '%s %s' % (resultado, leer_centenas(centena))
    return resultado

def leer_millones(numero):
    millon, millar = divmod(numero, 1000000)
    resultado = ''
    if (millon == 1):
        resultado = ' un millon '
    if (millon >= 2) and (millon <= 9):
        resultado = UNIDADES[millon]
    elif (millon >= 10) and (millon <= 99):
        resultado = leer_decenas(millon)
    elif (millon >= 100) and (millon <= 999):
        resultado = leer_centenas(millon)
    if millon > 1:
        resultado = '%s millones' % resultado
    if (millar > 0) and (millar <= 999):
        resultado = '%s %s' % (resultado, leer_centenas(millar))
    elif (millar >= 1000) and (millar <= 999999):
        resultado = '%s %s' % (resultado, leer_miles(millar))
    return resultado

def leer_millardos(numero):
    millardo, millon = divmod(numero, 1000000)
    return '%s millones %s' % (leer_miles(millardo), leer_millones(millon))
