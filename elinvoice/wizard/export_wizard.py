import base64
import logging
import xlsxwriter
import io

from io import BytesIO

from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.pagesizes import letter, landscape, A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors

from odoo import api, fields, models
from reportlab.platypus import Image
from datetime import datetime, timedelta

_logger = logging.getLogger(__name__)

class ExportWizard(models.TransientModel):
    _name = 'elinvoice.export_wizard'
    _description = 'Download CSV Wizard'

    # Define any fields needed for CSV generation
    report_type = fields.Selection(selection=[('CCFE', 'Ventas a Contribuyentes'),
                                                ('FE', 'Ventas a Consumidor Final'),
                                                ('CCFEP', 'Detalle de Compras'),
                                                ('CRE', 'Detalle de retencion 1%'),
                                                ('FSEE', 'Sujeto excluido'),
                                                ('FSEE', 'Sujeto excluido'),
                                                ('FEX', 'Detalle de FEX'),
                                                ('RET_IVA_1_162', 'Retencion IVA 1%'),
                                                ('PER_IVA_1_163', 'Percepcion IVA 1%'),
                                                ('IVA_2_161', 'Anticipo a cuenta IVA 2%'),],
                                    string="Tipo de reporte",
                                    required=True)

    file_format = fields.Selection(selection=[('CSV', 'CSV'),
                                                ('XLS', 'Excel'),
                                                ('PDF', 'PDF'),],
                                    string="Formato",
                                    required=True)

    date_from = fields.Date(string="Fecha inicio", required=True)
    date_to = fields.Date(string="Fecha fin", required=True)

    def get_report(self):
        csv_content = ""
        filename = 'NO_REPORT'

        if self.report_type == 'CCFE':
            filename = 'VENTAS CONTRIBUYENTES'
            if(self.file_format == 'XLS'):
                file = self.make_ccfe_xls(filename)
                return self.download_xls_file(file)
            elif(self.file_format == 'PDF'):
                file = self.make_ccfe_pdf(filename)
                return self.download_pdf_file(file)
            elif(self.file_format == 'CSV'):
                csv_content = self.make_ccfe_csv()
                return self.download_csv_file(csv_content, filename)

        if self.report_type == 'FE':
            filename = 'VENTAS CONSUMIDOR FINAL'
            if(self.file_format == 'XLS'):
                file = self.make_fe_xls(filename)
                return self.download_xls_file(file)
            elif(self.file_format == 'PDF'):
                file = self.make_fe_pdf(filename)
                return self.download_pdf_file(file)
            elif(self.file_format == 'CSV'):
                csv_content = self.make_fe_csv()
                return self.download_csv_file(csv_content, filename)

        if self.report_type == 'CCFEP':
            filename = 'COMPRAS'
            if(self.file_format == 'XLS'):
                file = self.make_ccfep_xls(filename)
                return self.download_xls_file(file)
            elif(self.file_format == 'PDF'):
                file = self.make_ccfep_pdf(filename)
                return self.download_pdf_file(file)
            elif(self.file_format == 'CSV'):
                csv_content = self.make_ccfep_csv()
                return self.download_csv_file(csv_content, filename)

        if self.report_type == 'CRE':
            filename = 'RETENCIONES IVA'
            if(self.file_format == 'XLS'):
                file = self.make_cre_xls(filename)
                return self.download_xls_file(file)
            elif(self.file_format == 'PDF'):
                file = self.make_cre_pdf(filename)
                return self.download_pdf_file(file)
            elif(self.file_format == 'CSV'):
                csv_content = self.make_cre_csv()
                return self.download_csv_file(csv_content, filename)

        if self.report_type == 'FSEE':
            filename = 'SUJETO EXCLUIDO'
            if(self.file_format == 'XLS'):
                file = self.make_fsee_xls(filename)
                return self.download_xls_file(file)
            elif(self.file_format == 'PDF'):
                file = self.make_fsee_pdf(filename)
                return self.download_pdf_file(file)
            elif(self.file_format == 'CSV'):
                csv_content = self.make_fsee_csv()
                return self.download_csv_file(csv_content, filename)
        
        # Detalle de FEX
        if self.report_type == 'FEX':
            filename = 'DETALLE DE FEX'
            if(self.file_format == 'XLS'):
                file = self.make_fex_xls(filename)
                return self.download_xls_file(file)
            elif(self.file_format == 'PDF'):
                file = self.make_fex_pdf(filename)
                return self.download_pdf_file(file)

        if self.report_type == 'RET_IVA_1_162':
            filename = 'RETENCION IVA 1%'
            if(self.file_format == 'XLS'):
                file = self.make_retencion_iva_xls(filename)
                return self.download_xls_file(file)
            elif(self.file_format == 'PDF'):
                file = self.make_retencion_iva_pdf(filename)
                return self.download_pdf_file(file)
            elif(self.file_format == 'CSV'):
                csv_content = self.make_retencion_iva_csv()
                return self.download_csv_file(csv_content, filename)
        
        if self.report_type == 'PER_IVA_1_163':
            filename = 'PERCEPCION IVA 1%'
            if(self.file_format == 'XLS'):
                file = self.make_percepcion_iva_xls(filename)
                return self.download_xls_file(file)
            elif(self.file_format == 'PDF'):
                file = self.make_percepcion_iva_pdf(filename)
                return self.download_pdf_file(file)
            elif(self.file_format == 'CSV'):
                csv_content = self.make_percepcion_iva_csv()
                return self.download_csv_file(csv_content, filename)
        
        if self.report_type == 'IVA_2_161':
            filename = 'ANTICIPO A CUENTA IVA 2%'
            if(self.file_format == 'XLS'):
                file = self.make_anticipo_iva_xls(filename)
                return self.download_xls_file(file)
            elif(self.file_format == 'PDF'):
                file = self.make_anticipo_iva_pdf(filename)
                return self.download_pdf_file(file)
            elif(self.file_format == 'CSV'):
                csv_content = self.make_anticipo_iva_csv()
                return self.download_csv_file(csv_content, filename)

        return self.download_csv_file(csv_content, filename)

    def download_csv_file(self, csv_content, filename):
        date = self.date_from.strftime('%m-%Y')

        filestore_record = self.env['ir.attachment'].create({
            'datas': base64.b64encode(csv_content.encode('utf-8')),  # Encode content as base64
            'name': filename  + ' ' + date + '.csv',
            'mimetype': 'text/csv',
        })

        return {
            'type': 'ir.actions.act_url',
            'url': '/web/content/?model=ir.attachment&id=%s&field=datas' % filestore_record.id,
            'target': 'self',
        }

    def download_xls_file(self, filename):
        with open(filename, 'rb') as f:
            file_content = f.read()

        filestore_record = self.env['ir.attachment'].create({
            'datas': base64.b64encode(file_content),  # Encode content as base64
            'name': filename,
            'type': 'binary',
            'mimetype': 'application/vnd.ms-excel',
        })

        return {
            'type': 'ir.actions.act_url',
            'url': '/web/content/?model=ir.attachment&id=%s&field=datas' % filestore_record.id,
            'target': 'self',
        }

    def download_pdf_file(self, filename):
        with open(filename, 'rb') as f:
            file_content = f.read()

        filestore_record = self.env['ir.attachment'].create({
            'datas': base64.b64encode(file_content),  # Encode content as base64
            'name': filename,
            'type': 'binary',
            'mimetype': 'application/vnd.pdf',
        })

        return {
            'type': 'ir.actions.act_url',
            'url': '/web/content/?model=ir.attachment&id=%s&field=datas' % filestore_record.id,
            'target': 'self',
        }

    ###################################
    # CSV DOCUMENT MAKERS
    ###################################

    def make_ccfe_csv(self):
        start_date = self.date_from
        end_date = self.date_to

        search_domain = [
            ('invoice_date', '>=', start_date.strftime('%Y-%m-%d')),
            ('invoice_date', '<=', end_date.strftime('%Y-%m-%d')),
            '|', '|',
            ('name', 'ilike', 'CCFE/'),
            ('name', 'ilike', 'RNCE/'),
            ('name', 'ilike', 'RNDE/'),
        ]

        data_to_export = self.env['account.move'].search(search_domain, order='date asc')

        csv_content = ""

        for record in data_to_export:
            confirmation = self.get_confirmation(record)

            if(confirmation['status'] == 0):
                continue

            partner = record.partner_id
            num_doc = partner.numdocumento if partner.numdocumento else ''

            total_exenta = 0
            total_gravada = 0

            for line in record.line_ids:
                if line.product_id.default_code:
                    total_gravada = round(total_gravada + line.price_subtotal, 2)

            total_iva = round(total_gravada * 0.13, 2)
            gran_total = round(total_gravada + total_exenta + total_iva, 2)

            nit = ""
            dui = ""

            if(partner.nit):
                nit = self.csv_str(partner.nit)
                dui = ""
            else:
                nit = ""
                dui = self.csv_str(num_doc)

            tipoOperacionRenta = 0
            if record.tipo_operacion_renta_move:
                tipoOperacionRenta = record.tipo_operacion_renta_move

            tipoIngresoRenta = 0
            if record.tipo_ingreso_renta_move:
                tipoIngresoRenta = record.tipo_ingreso_renta_move

            document_type = ''

            if record.journal_id.code == 'CCFE':
                document_type = '03'
            elif record.journal_id.code == 'NCE':
                document_type = '05'
            elif record.journal_id.code == 'NDE':
                document_type = '06'

            row = []
            row.append(record.date.strftime('%d/%m/%Y')) #A
            row.append('4') #B
            row.append(self.csv_str(document_type)) #C
            row.append(self.csv_str(confirmation['control_number'])) #D
            row.append(self.csv_str(confirmation['received_seal'])) #E
            row.append(self.csv_str(confirmation['uid_de'])) #F
            row.append(self.csv_str(record.name)) #G
            row.append(nit) #H
            row.append(self.csv_str(partner.name)) #I
            row.append(str(total_exenta)) #J
            row.append('0') #K
            row.append(str(total_gravada)) #L
            row.append(str(total_iva)) #M
            row.append('0') #N
            row.append('0') #O
            row.append(str(gran_total)) #P
            row.append(dui) #Q
            row.append(str(tipoOperacionRenta)) #R
            row.append(str(tipoIngresoRenta)) #S
            row.append('1') #T

            row = ";".join(row)
            csv_content = csv_content + row + "\n"

        # csv_content += self.make_nce_csv()

        return csv_content

    def make_nce_csv(self):
        start_date = self.date_from
        end_date = self.date_to

        search_domain = [
            ('name', '=like', 'NCE/%'),
            ('date', '>=', start_date.strftime('%Y-%m-%d')),
            ('date', '<=', end_date.strftime('%Y-%m-%d')),
        ]

        data_to_export = self.env['account.move'].search(search_domain, order='date asc')

        csv_content = ""

        for record in data_to_export:
            confirmation = self.get_confirmation(record)

            if(confirmation['status'] != 2):
                continue

            partner = record.partner_id
            num_doc = partner.numdocumento if partner.numdocumento else ''

            total_exenta = 0
            total_gravada = 0

            for line in record.line_ids:
                if line.product_id.default_code:
                    total_gravada = round(total_gravada + line.price_subtotal, 2)

            total_iva = round(total_gravada * 0.13, 2)
            gran_total = round(total_gravada + total_exenta + total_iva, 2)

            nit = ""
            dui = ""

            if(partner.nit):
                nit = self.csv_str(partner.nit)
                dui = ""
            else:
                nit = ""
                dui = self.csv_str(num_doc)

            tipoOperacionRenta = 0
            if record.tipo_operacion_renta_move:
                tipoOperacionRenta = record.tipo_operacion_renta_move

            tipoIngresoRenta = 0
            if record.tipo_ingreso_renta_move:
                tipoIngresoRenta = record.tipo_ingreso_renta_move

            row = []
            row.append(record.date.strftime('%d/%m/%Y')) #A
            row.append('4') #B
            row.append(self.csv_str('05')) #C
            row.append(self.csv_str(confirmation['control_number'])) #D
            row.append(self.csv_str(confirmation['received_seal'])) #E
            row.append(self.csv_str(confirmation['uid_de'])) #F
            row.append(self.csv_str(record.name)) #G
            row.append(nit) #H
            row.append(self.csv_str(partner.name)) #I
            row.append(str(total_exenta)) #J
            row.append('0') #K
            row.append(str(total_gravada)) #L
            row.append(str(total_iva)) #M
            row.append('0') #N
            row.append('0') #O
            row.append(str(gran_total)) #P
            row.append(dui) #Q
            row.append(str(tipoOperacionRenta)) #R
            row.append(str(tipoIngresoRenta)) #S
            row.append('1') #T

            row = ";".join(row)
            csv_content = csv_content + row + "\n"

        return csv_content

    def make_fe_csv(self):
        start_date = self.date_from
        end_date = self.date_to

        search_domain = [
            ('name', '=like', 'FE/%'),
            ('date', '>=', start_date.strftime('%Y-%m-%d')),
            ('date', '<=', end_date.strftime('%Y-%m-%d')),
        ]

        data_to_export = self.env['account.move'].search(search_domain, order='date asc')

        csv_content = ""

        for record in data_to_export:
            confirmation = self.get_confirmation(record)

            if(confirmation['status'] == 0):
                continue

            total_exenta = 0
            total_gravada = 0

            for line in record.line_ids:
                if line.product_id.default_code:

                    taxes = line.tax_ids
                    exenta = True
                    ivaItem = 0

                    if taxes:
                        for tax in taxes:
                            if 'IVA 13%' in tax.name or 'IVA 13% FAC' in tax.name:
                                exenta = False
                                ivaItem = line.price_subtotal * 0.13

                    if exenta:
                        total_exenta = round(total_exenta + line.price_subtotal, 2)
                    else:
                        total_gravada = round(total_gravada + (line.price_subtotal + ivaItem), 2)

            tipoOperacionRenta = 0
            if record.tipo_operacion_renta_move:
                tipoOperacionRenta = record.tipo_operacion_renta_move

            tipoIngresoRenta = 0
            if record.tipo_ingreso_renta_move:
                tipoIngresoRenta = record.tipo_ingreso_renta_move

            row = []
            row.append(record.date.strftime('%d/%m/%Y')) #A
            row.append('4') #B
            row.append(self.csv_str('01')) #C
            row.append(self.csv_str(confirmation['control_number'])) #D
            row.append(self.csv_str(confirmation['received_seal'])) #E
            row.append(self.csv_str(confirmation['uid_de'])) #F
            row.append(self.csv_str(record.name)) #G
            row.append(self.csv_str(confirmation['uid_de'])) #H
            row.append(self.csv_str(confirmation['uid_de'])) #I
            row.append('') #J
            row.append(str(total_exenta)) #K
            row.append('0') #L
            row.append('0') #M
            row.append(str(total_gravada)) #N
            row.append('0') #O
            row.append('0') #P
            row.append('0') #Q
            row.append('0') #R
            row.append('0') #S
            row.append(str(record.amount_total)) #T
            row.append(str(tipoOperacionRenta)) #U
            row.append(str(tipoIngresoRenta)) #V
            row.append('2') #W

            row = ";".join(row)
            csv_content = csv_content + row + "\n"

        csv_content += self.make_fexe_csv()

        return csv_content

    def make_fexe_csv(self):
        start_date = self.date_from
        end_date = self.date_to

        search_domain = [
            ('name', '=like', 'FEXE/%'),
            ('date', '>=', start_date.strftime('%Y-%m-%d')),
            ('date', '<=', end_date.strftime('%Y-%m-%d')),
        ]

        data_to_export = self.env['account.move'].search(search_domain, order='date asc')

        csv_content = ""

        for record in data_to_export:
            confirmation = self.get_confirmation(record)

            if(confirmation['status'] == 0):
                continue

            total_exenta = 0
            total_gravada = 0

            for line in record.line_ids:
                if line.product_id.default_code:
                    
                    taxes = line.tax_ids
                    exenta = True

                    if taxes:
                        for tax in taxes:
                            if 'IVA 13%' in tax.name:
                                exenta = False

                    if exenta:
                        total = line.price_subtotal
                        total_exenta = round(total_exenta + total, 2)
                    else:
                        total = round(line.price_subtotal * 1.13, 2)
                        total_gravada = round(total_gravada + total, 2)

            tipoOperacionRenta = 0
            if record.tipo_operacion_renta_move:
                tipoOperacionRenta = record.tipo_operacion_renta_move

            tipoIngresoRenta = 0
            if record.tipo_ingreso_renta_move:
                tipoIngresoRenta = record.tipo_ingreso_renta_move

            row = []
            row.append(record.date.strftime('%d/%m/%Y')) #A
            row.append('4') #B
            row.append(self.csv_str('11')) #C
            row.append(self.csv_str(confirmation['control_number'])) #D
            row.append(self.csv_str(confirmation['received_seal'])) #E
            row.append(self.csv_str(confirmation['uid_de'])) #F
            row.append(self.csv_str(record.name)) #G
            row.append(self.csv_str(confirmation['uid_de'])) #H
            row.append(self.csv_str(confirmation['uid_de'])) #I
            row.append('') #J
            row.append('0') #K
            row.append('0') #L
            row.append('0') #M
            row.append('0') #N
            row.append('0') #O
            row.append('0') #P
            row.append(str(record.amount_untaxed)) #Q
            row.append('0') #R
            row.append('0') #S
            row.append(str(record.amount_total)) #T
            row.append(str(tipoOperacionRenta)) #U
            row.append(str(tipoIngresoRenta)) #V
            row.append('2') #W

            row = ";".join(row)
            csv_content = csv_content + row + "\n"

        return csv_content

    def make_ccfep_csv(self):
        start_date = self.date_from
        end_date = self.date_to

        search_domain = [
            '|',  # OR lógico
            ('name', 'like', 'CCFEP%'),
            ('name', 'like', 'RNCEP%'),
            ('date', '>=', start_date.strftime('%Y-%m-%d')),
            ('date', '<=', end_date.strftime('%Y-%m-%d')),
        ]

        data_to_export = self.env['account.move'].search(search_domain, order='date asc')

        csv_content = ""
        for record in data_to_export:
            partner = record.partner_id
            num_doc = partner.numdocumento if partner.numdocumento else ''

            is_internal = partner.is_internal if partner.is_internal else 'SI'

            total_exenta = 0
            total_gravada = 0

            for line in record.line_ids:
                taxes = line.tax_ids
                exenta = True
                if taxes:
                    for tax in taxes:
                        if 'IVA 13%' in tax.name:
                            exenta = False

                if exenta:
                    total = line.price_subtotal
                    total_exenta = round(total_exenta + total, 2)
                else:
                    total = line.price_subtotal
                    total_gravada = round(total_gravada + total, 2)

            total_iva = round(total_gravada * 0.13, 2)
            gran_total = round(total_gravada + total_exenta + total_iva, 2)

            nit = ""
            dui = ""

            if(partner.nit):
                nit = self.csv_str(partner.nit)
                dui = ""
            else:
                nit = ""
                dui = self.csv_str(num_doc)

            categoria = partner.categoria if partner.categoria else '1'

            row = []
            row.append(record.date.strftime('%d/%m/%Y')) #A

            num_tax_doc = record.codigo_generacion if record.codigo_generacion else ''

            if(is_internal == 'SI'):
                generation = record.tipo_documento if record.tipo_documento else '4'

                if (generation == '2'):
                    generation = '4'

                row.append(self.csv_str(generation)) #B
                row.append(self.csv_str('03')) #C
            else:
                row.append(self.csv_str('3')) #B
                row.append(self.csv_str('13')) #C

            row.append(self.csv_str(num_tax_doc)) #D
            row.append(nit) #E
            row.append(self.csv_str(partner.name)) #F
            row.append(str(total_exenta)) #G
            row.append('0') #H
            row.append('0') #I

            if(is_internal == 'SI'):
                row.append(str(total_gravada)) #J
            else:
                row.append(str('0')) #J

            row.append('0') #K
            row.append('0') #L

            if(is_internal == 'SI'):
                row.append('0') #M
            else:
                row.append(str(total_gravada)) #M

            row.append(str(total_iva)) #N
            row.append(str(gran_total)) #O
            row.append(dui) #P
            row.append('1') #Q
            row.append(self.csv_str(categoria)) #R costo = 1, gasto = 2
            row.append('4') #S

            if categoria == '1': # si en R es costo = 5, si es gasto = 2
                row.append('5') #T
            else:
                row.append('2') #T

            row.append('3') #U

            row = ";".join(row)
            csv_content = csv_content + row + "\n"

        return csv_content

    def make_cre_csv(self):
        start_date = self.date_from
        end_date = self.date_to

        search_domain = [
            ('name', '=like', 'CCFEP%'),
            ('date', '>=', start_date.strftime('%Y-%m-%d')),
            ('date', '<=', end_date.strftime('%Y-%m-%d')),
        ]

        data_to_export = self.env['account.move'].search(search_domain, order='date asc')

        csv_content = ""
        for record in data_to_export:
            retencion = False
            monto_sujeto = 0

            for line in record.line_ids:
                taxes = line.tax_ids
                if taxes:
                    for tax in taxes:
                        if tax.name == 'RET 1%':
                            retencion = True
                            monto_sujeto = monto_sujeto + line.price_subtotal

            if not retencion:
                continue

            partner = record.partner_id
            num_doc = partner.numdocumento if partner.numdocumento else ''

            nit = ""
            dui = ""

            if(partner.nit):
                nit = self.csv_str(partner.nit)
                dui = ""
            else:
                nit = ""
                dui = self.csv_str(num_doc)

            row = []
            row.append(nit) #A
            row.append(record.date.strftime('%d/%m/%Y')) #B
            row.append(self.csv_str('07')) #C
            row.append(self.csv_str(record.purchase_doc_resolution)) #D
            row.append(self.csv_str(record.purchase_doc_serie)) #E
            row.append(self.csv_str(record.purchase_doc_number)) #F
            row.append(str(monto_sujeto)) #G
            row.append(str(round(monto_sujeto * 0.01, 2))) #H
            row.append(dui) #I
            row.append('10') #J

            _logger.info(row)

            row = ";".join(row)
            csv_content = csv_content + row + "\n"

        return csv_content

    def make_fsee_csv(self):
        start_date = self.date_from
        end_date = self.date_to

        search_domain = [
            ('name', '=like', 'FSEE/%'),
            ('date', '>=', start_date.strftime('%Y-%m-%d')),
            ('date', '<=', end_date.strftime('%Y-%m-%d')),
        ]

        data_to_export = self.env['account.move'].search(search_domain, order='date asc')

        csv_content = ""
        for record in data_to_export:
            confirmation = self.get_confirmation(record)

            if(confirmation['status'] != 2):
                continue

            partner = record.partner_id
            num_doc = partner.numdocumento if partner.numdocumento else ''

            # total_exenta = 0
            # total_gravada = 0

            # for line in record.line_ids:
            #     taxes = line.tax_ids
            #     exenta = True
            #     if taxes:
            #         for tax in taxes:
            #             if 'IVA 13%' in tax.name:
            #                 exenta = False

            #     if exenta:
            #         total = line.price_subtotal
            #         total_exenta = round(total_exenta + total, 2)
            #     else:
            #         total = line.price_subtotal
            #         total_gravada = round(total_gravada + total, 2)

            # total_iva = round(total_gravada * 0.13, 2)
            # gran_total = round(total_gravada + total_exenta + total_iva, 2)

            row = []

            row.append('2') #A
            row.append(self.csv_str(num_doc)) #B
            row.append(self.csv_str(partner.name)) #C
            row.append(record.date.strftime('%d/%m/%Y')) #D
            row.append(self.csv_str(confirmation['received_seal'])) #E
            row.append(self.csv_str(confirmation['uid_de'])) #F
            row.append(str(abs(record.amount_untaxed))) #G
            row.append('0') #H
            row.append('1') #I
            row.append('2') #J
            row.append('4') #K
            row.append('2') #L
            row.append('5') #M

            row = ";".join(row)
            csv_content = csv_content + row + "\n"

        return csv_content

    def make_retencion_iva_csv(self):
        start_date = self.date_from
        end_date = self.date_to

        _logger.info('***** CompanyID: ' + str(self.env.company.id) + ' *****')

        search_domain = [
            ('name', 'like', 'CREP/%'),
            ('date', '>=', start_date.strftime('%Y-%m-%d')),
            ('date', '<=', end_date.strftime('%Y-%m-%d')),
            ('company_id', '=', self.env.company.id)
        ]

        data_to_export = self.env['account.move'].search(search_domain, order='date asc')

        csv_content = ""

        for record in data_to_export:
            
            _logger.info('***** RecordID: ' + str(record.id) + ' *****')
            _logger.info('***** RelatedDocument: ' + str(record.related_document_cre) + ' *****')

            confirmations = self.env['elinvoice.confirmation'].search([('uid_de', '=', record.related_document_cre)])

            confirmation = next(
                (c for c in confirmations if c.account_move_id.company_id.id == self.env.company.id),
                None
            )

            if confirmation:
                _logger.info('#### Si existe confirmacion ####')

                move = confirmation.account_move_id
                
                if move.is_processed and not move.is_anulated:
                    _logger.info('#### Si existe move ####')

                    partner = record.partner_id
                    num_doc = partner.numdocumento if partner.numdocumento else ''

                    nit = ""
                    dui = ""

                    if(partner.nit):
                        nit = self.csv_str(partner.nit)
                        dui = ""
                    else:
                        nit = ""
                        dui = self.csv_str(num_doc)

                    row = []
                    row.append(self.csv_str(nit)) # NIT
                    row.append(self.csv_str(record.date.strftime('%d/%m/%Y'))) # FECHA
                    row.append(self.csv_str('07')) # TIPO DOC
                    row.append(self.csv_str(record.purchase_doc_serie)) # SERIE
                    row.append(self.csv_str(self.formatWithoutDashes(record.purchase_doc_number))) # NUM DOC
                    row.append(self.csv_str(format(move.amount_untaxed, '.2f'))) # MONTO SUJETO
                    row.append(self.csv_str(format(record.amount_untaxed, '.2f'))) # MONTO RETENCION
                    row.append(self.csv_str(num_doc)) # DUI
                    row.append(self.csv_str('7')) # ANEXO

                    row = ";".join(row)
                    csv_content = csv_content + row + "\n"
            else:
                _logger.info('#### No existe confirmacion ####')

        return csv_content

    def make_percepcion_iva_csv(self):
        start_date = self.date_from
        end_date = self.date_to

        search_domain = [
            '&',
                '|', '|',
                    ('name', '=like', 'CCFEP/%'),
                    ('name', '=like', 'RNCEP/%'),
                    ('name', '=like', 'RNDEP/%'),
            '&',
                ('invoice_date', '>=', start_date.strftime('%Y-%m-%d')),
                ('invoice_date', '<=', end_date.strftime('%Y-%m-%d')),
        ]

        data_to_export = self.env['account.move'].search(search_domain, order='invoice_date asc, id asc')

        csv_content = ""

        global_sujeto = 0
        global_percepcion = 0

        for record in data_to_export:
            
            partner = record.partner_id
            num_doc = partner.numdocumento if partner.numdocumento else ''

            monto_sujeto = 0
            monto_percepcion = 0

            for line in record.line_ids:

                taxes = line.tax_ids

                if taxes:

                    for tax in taxes:

                        if 'PER 1%' in tax.name:
                            monto_sujeto += line.price_subtotal
                            monto_percepcion += line.price_subtotal * 0.01

            nit = ""
            dui = ""

            if(partner.nit):
                nit = self.csv_str(partner.nit)
                dui = ""
            else:
                nit = ""
                dui = self.csv_str(num_doc)
            
            tipo_doc = '00'
            # Determinar el tipo de documento
            if record.journal_id.code == 'CCFEP':
                tipo_doc = '03'
            elif record.journal_id.code == 'NCEP':
                tipo_doc = '05'
            elif record.journal_id.code == 'NDEP':
                tipo_doc = '06'

            row = []
            row.append(self.csv_str(nit)) # NIT
            row.append(self.csv_str(record.date.strftime('%d/%m/%Y'))) # FECHA
            row.append(self.csv_str(tipo_doc)) # TIPO DOC
            row.append(self.csv_str(record.sello_recepcion)) # SERIE
            row.append(self.csv_str(self.formatWithoutDashes(record.codigo_generacion))) # NUM DOC
            row.append(self.csv_str(format(monto_sujeto, '.2f'))) # MONTO SUJETO
            row.append(self.csv_str(format(monto_percepcion, '.2f'))) # MONTO PERCEPCION
            row.append(self.csv_str(dui)) # DUI
            row.append(self.csv_str('8')) # ANEXO

            row = ";".join(row)
            csv_content = csv_content + row + "\n"

        csv_content += self.make_nce_csv()

        return csv_content

    def make_anticipo_iva_csv(self):
        start_date = self.date_from
        end_date = self.date_to

        search_domain = [
            ('name', '=like', 'DCLEP/%'),
            ('invoice_date', '>=', start_date.strftime('%Y-%m-%d')),
            ('invoice_date', '<=', end_date.strftime('%Y-%m-%d')),
        ]

        data_to_export = self.env['account.move'].search(search_domain, order='invoice_date asc, id asc')

        csv_content = ""

        global_sujeto = 0
        global_anticipo = 0

        for record in data_to_export:
            
            partner = record.partner_id
            num_doc = partner.numdocumento if partner.numdocumento else ''

            monto_anticipo = 0
            monto_sujeto = 0

            for line in record.line_ids:

                monto_anticipo += line.price_unit * line.quantity
            
            monto_sujeto = monto_anticipo / 0.02

            nit = ""
            dui = ""

            if(partner.nit):
                nit = self.csv_str(partner.nit)
                dui = ""
            else:
                nit = ""
                dui = self.csv_str(num_doc)
            
            row = []
            row.append(self.csv_str(nit)) # NIT
            row.append(self.csv_str(record.date.strftime('%d/%m/%Y'))) # FECHA
            row.append(self.csv_str(record.sello_recepcion)) # SERIE
            row.append(self.csv_str(self.formatWithoutDashes(record.codigo_generacion))) # NUM DOC
            row.append(self.csv_str(format(monto_sujeto, '.2f'))) # MONTO SUJETO
            row.append(self.csv_str(format(monto_anticipo, '.2f'))) # MONTO ANTICIPO
            row.append(self.csv_str(dui)) # DUI
            row.append(self.csv_str('6')) # ANEXO

            row = ";".join(row)
            csv_content = csv_content + row + "\n"

        csv_content += self.make_nce_csv()

        return csv_content
    
    ###################################
    # XLS DOCUMENT MAKERS
    ###################################
    def make_retencion_iva_xls(self, filename):
        date = self.date_from.strftime('%m-%Y')
        filename = filename + ' ' + date + '.xls'

        row = 0

        workbook = xlsxwriter.Workbook(filename)
        
        bold = workbook.add_format({
            'bold': True, 
            'align': 'center', 
            'valign': 'vcenter', 
            'text_wrap': True
        })

        bold_left = workbook.add_format({
            'bold': True, 
            'align': 'left', 
            'valign': 'vcenter', 
            'text_wrap': True
        })

        center = workbook.add_format({
            'bold': False, 
            'align': 'center', 
            'valign': 'vcenter', 
            'text_wrap': True
        })

        left = workbook.add_format({
            'bold': False, 
            'align': 'left', 
            'valign': 'vcenter', 
            'text_wrap': True
        })

        right = workbook.add_format({
            'bold': False, 
            'align': 'right', 
            'valign': 'vcenter', 
            'text_wrap': True
        })

        worksheet = workbook.add_worksheet()

        # Document header

        # add the company logo to the excel file if it exists
        company = self.env['res.company'].search([], limit=1)
        if company.logo:
            logo_stream = BytesIO(base64.b64decode(company.logo))
            worksheet.insert_image('K1', 'logo.png', {'image_data': logo_stream})

        worksheet.write(row, 0, 'RETENCION IVA 1%', bold)
        row += 2

        worksheet.write(row, 0, company.name, bold)
        row += 1

        worksheet.write(row, 0, 'ACTIVIDAD ECONOMICA:')
        worksheet.write(row, 1, company.descactividad)
        row += 1

        worksheet.write(row, 0, 'NIT:')
        worksheet.write(row, 1, company.nit)
        row += 1

        worksheet.write(row, 0, 'NRC:')
        worksheet.write(row, 1, company.nrc)
        row += 1

        worksheet.write(row, 0, 'MES:')
        worksheet.write(row, 1, MESES[int(self.date_from.strftime('%m')) - 1])
        row += 1

        worksheet.write(row, 0, 'AÑO:')
        worksheet.write(row, 1, self.date_from.strftime('%Y'))
        row += 1

        # Document body
        row += 2

        worksheet.set_column(0, 0, 20) # NIT
        worksheet.set_column(1, 1, 15) # FECHA
        worksheet.set_column(2, 2, 10) # TIPO DOC
        worksheet.set_column(3, 3, 50) # SERIE
        worksheet.set_column(4, 4, 50) # NUM DOC
        worksheet.set_column(5, 5, 14) # MONTO SUJETO
        worksheet.set_column(6, 6, 14) # MONTO RETENCION
        worksheet.set_column(7, 7, 20) # DUI
        worksheet.set_column(8, 8, 10) # ANEXO
        
        # Headers
        worksheet.write(row, 0, 'NIT', bold)
        worksheet.write(row, 1, 'FECHA', bold)
        worksheet.write(row, 2, 'TIPO DOC', bold)
        worksheet.write(row, 3, 'SERIE', bold)
        worksheet.write(row, 4, 'NUM DOC', bold)
        worksheet.write(row, 5, 'MONTO\nSUJETO', bold)
        worksheet.write(row, 6, 'MONTO\nRETENCION 1%', bold)
        worksheet.write(row, 7, 'DUI', bold)
        worksheet.write(row, 8, 'ANEXO', bold)

        start_date = self.date_from
        end_date = self.date_to

        row += 1 # skip headers

        _logger.info('***** CompanyID: ' + str(self.env.company.id) + ' *****')

        search_domain = [
            ('name', 'like', 'CREP/%'),
            ('date', '>=', start_date.strftime('%Y-%m-%d')),
            ('date', '<=', end_date.strftime('%Y-%m-%d')),
            ('company_id', '=', self.env.company.id)
        ]

        data_to_export = self.env['account.move'].search(search_domain, order='date asc')

        global_sujeto = 0
        global_retencion = 0

        for record in data_to_export:

            _logger.info('***** RecordID: ' + str(record.id) + ' *****')
            _logger.info('***** RelatedDocument: ' + str(record.related_document_cre) + ' *****')

            confirmations = self.env['elinvoice.confirmation'].search([('uid_de', '=', record.related_document_cre)])

            confirmation = next(
                (c for c in confirmations if c.account_move_id.company_id.id == self.env.company.id),
                None
            )

            if confirmation:
                _logger.info('#### Si existe confirmacion ####')

                move = confirmation.account_move_id
                
                if move.is_processed and not move.is_anulated:
                    _logger.info('#### Si existe move ####')

                    partner = record.partner_id
                    num_doc = partner.numdocumento if partner.numdocumento else ''

                    nit = ""
                    dui = ""

                    if(partner.nit):
                        nit = partner.nit
                        dui = ""
                    else:
                        nit = ""
                        dui = num_doc
                    
                    worksheet.write(row, 0, nit, center)
                    worksheet.write(row, 1, record.date.strftime('%d/%m/%Y'), center)
                    worksheet.write(row, 2, '07', center)
                    worksheet.write(row, 3, record.purchase_doc_serie, center)
                    worksheet.write(row, 4, self.formatWithoutDashes(record.purchase_doc_number), center)
                    worksheet.write(row, 5, format(move.amount_untaxed, '.2f'), center)
                    worksheet.write(row, 6, format(record.amount_untaxed, '.2f'), center)
                    worksheet.write(row, 7, num_doc, center)
                    worksheet.write(row, 8, '7', center)

                    row += 1

                    global_sujeto += move.amount_untaxed
                    global_retencion += record.amount_untaxed
            else:
                _logger.info('#### No existe confirmacion ####')
        

        # Document footer
        worksheet.write(row, 4, 'TOTALES', center)
        worksheet.write(row, 5, format(global_sujeto, '.2f'), center)
        worksheet.write(row, 6, format(global_retencion, '.2f'), center)

        workbook.close()

        return filename

    def make_percepcion_iva_xls(self, filename):
        date = self.date_from.strftime('%m-%Y')
        filename = filename + ' ' + date + '.xls'

        row = 0

        workbook = xlsxwriter.Workbook(filename)
        
        bold = workbook.add_format({
            'bold': True, 
            'align': 'center', 
            'valign': 'vcenter', 
            'text_wrap': True
        })

        bold_left = workbook.add_format({
            'bold': True, 
            'align': 'left', 
            'valign': 'vcenter', 
            'text_wrap': True
        })

        center = workbook.add_format({
            'bold': False, 
            'align': 'center', 
            'valign': 'vcenter', 
            'text_wrap': True
        })

        left = workbook.add_format({
            'bold': False, 
            'align': 'left', 
            'valign': 'vcenter', 
            'text_wrap': True
        })

        right = workbook.add_format({
            'bold': False, 
            'align': 'right', 
            'valign': 'vcenter', 
            'text_wrap': True
        })

        worksheet = workbook.add_worksheet()

        # Document header

        # add the company logo to the excel file if it exists
        company = self.env['res.company'].search([], limit=1)
        if company.logo:
            logo_stream = BytesIO(base64.b64decode(company.logo))
            worksheet.insert_image('K1', 'logo.png', {'image_data': logo_stream})

        worksheet.write(row, 0, 'PERCEPCION IVA 1%', bold)
        row += 2

        worksheet.write(row, 0, company.name, bold)
        row += 1

        worksheet.write(row, 0, 'ACTIVIDAD ECONOMICA:')
        worksheet.write(row, 1, company.descactividad)
        row += 1

        worksheet.write(row, 0, 'NIT:')
        worksheet.write(row, 1, company.nit)
        row += 1

        worksheet.write(row, 0, 'NRC:')
        worksheet.write(row, 1, company.nrc)
        row += 1

        worksheet.write(row, 0, 'MES:')
        worksheet.write(row, 1, MESES[int(self.date_from.strftime('%m')) - 1])
        row += 1

        worksheet.write(row, 0, 'AÑO:')
        worksheet.write(row, 1, self.date_from.strftime('%Y'))
        row += 1

        # Document body
        row += 2

        worksheet.set_column(0, 0, 20) # NIT
        worksheet.set_column(1, 1, 15) # FECHA
        worksheet.set_column(2, 2, 10) # TIPO DOC
        worksheet.set_column(3, 3, 50) # SERIE
        worksheet.set_column(4, 4, 50) # NUM DOC
        worksheet.set_column(5, 5, 14) # MONTO SUJETO
        worksheet.set_column(6, 6, 14) # MONTO RETENCION
        worksheet.set_column(7, 7, 20) # DUI
        worksheet.set_column(8, 8, 10) # ANEXO
        
        # Headers
        worksheet.write(row, 0, 'NIT', bold)
        worksheet.write(row, 1, 'FECHA', bold)
        worksheet.write(row, 2, 'TIPO DOC', bold)
        worksheet.write(row, 3, 'SERIE', bold)
        worksheet.write(row, 4, 'NUM DOC', bold)
        worksheet.write(row, 5, 'MONTO\nSUJETO', bold)
        worksheet.write(row, 6, 'MONTO\nPERCEPCION 1%', bold)
        worksheet.write(row, 7, 'DUI', bold)
        worksheet.write(row, 8, 'ANEXO', bold)

        start_date = self.date_from
        end_date = self.date_to

        row += 1 # skip headers

        search_domain = [
            '&',
                '|', '|',
                    ('name', '=like', 'CCFEP/%'),
                    ('name', '=like', 'RNCEP/%'),
                    ('name', '=like', 'RNDEP/%'),
            '&',
                ('invoice_date', '>=', start_date.strftime('%Y-%m-%d')),
                ('invoice_date', '<=', end_date.strftime('%Y-%m-%d')),
        ]

        data_to_export = self.env['account.move'].search(search_domain, order='invoice_date asc, id asc')

        global_sujeto = 0
        global_percepcion = 0

        for record in data_to_export:

            partner = record.partner_id
            num_doc = partner.numdocumento if partner.numdocumento else ''

            monto_sujeto = 0
            monto_percepcion = 0

            for line in record.line_ids:
                if line.product_id.default_code:
                        
                    taxes = line.tax_ids

                    if taxes:

                        for tax in taxes:

                            if 'PER 1%' in tax.name:
                                monto_sujeto += line.price_subtotal
                                monto_percepcion += line.price_subtotal * 0.01

            nit = ""
            dui = ""

            if(partner.nit):
                nit = partner.nit
                dui = ""
            else:
                nit = ""
                dui = num_doc
            
            tipo_doc = '00'
            # Determinar el tipo de documento
            if record.journal_id.code == 'CCFEP':
                tipo_doc = '03'
            elif record.journal_id.code == 'NCEP':
                tipo_doc = '05'
            elif record.journal_id.code == 'NDEP':
                tipo_doc = '06'

            worksheet.write(row, 0, nit, center)
            worksheet.write(row, 1, record.date.strftime('%d/%m/%Y'), center)
            worksheet.write(row, 2, tipo_doc, center)
            worksheet.write(row, 3, record.sello_recepcion, center)
            worksheet.write(row, 4, self.formatWithoutDashes(record.codigo_generacion), center)
            worksheet.write(row, 5, format(monto_sujeto, '.2f'), center)
            worksheet.write(row, 6, format(monto_percepcion, '.2f'), center)
            worksheet.write(row, 7, dui, center)
            worksheet.write(row, 8, '8', center)

            row += 1

            global_sujeto += monto_sujeto
            global_percepcion += monto_percepcion

        # Document footer
        worksheet.write(row, 4, 'TOTALES', bold)
        worksheet.write(row, 5, format(global_sujeto, '.2f'), bold)
        worksheet.write(row, 6, format(global_percepcion, '.2f'), bold)

        workbook.close()

        return filename

    def make_anticipo_iva_xls(self, filename):
        date = self.date_from.strftime('%m-%Y')
        filename = filename + ' ' + date + '.xls'

        row = 0

        workbook = xlsxwriter.Workbook(filename)
        
        bold = workbook.add_format({
            'bold': True, 
            'align': 'center', 
            'valign': 'vcenter', 
            'text_wrap': True
        })

        bold_left = workbook.add_format({
            'bold': True, 
            'align': 'left', 
            'valign': 'vcenter', 
            'text_wrap': True
        })

        center = workbook.add_format({
            'bold': False, 
            'align': 'center', 
            'valign': 'vcenter', 
            'text_wrap': True
        })

        left = workbook.add_format({
            'bold': False, 
            'align': 'left', 
            'valign': 'vcenter', 
            'text_wrap': True
        })

        right = workbook.add_format({
            'bold': False, 
            'align': 'right', 
            'valign': 'vcenter', 
            'text_wrap': True
        })

        worksheet = workbook.add_worksheet()

        # Document header

        # add the company logo to the excel file if it exists
        company = self.env['res.company'].search([], limit=1)
        if company.logo:
            logo_stream = BytesIO(base64.b64decode(company.logo))
            worksheet.insert_image('K1', 'logo.png', {'image_data': logo_stream})

        worksheet.write(row, 0, 'ANTICIPO A CUENTA IVA 2%', bold)
        row += 2

        worksheet.write(row, 0, company.name, bold)
        row += 1

        worksheet.write(row, 0, 'ACTIVIDAD ECONOMICA:')
        worksheet.write(row, 1, company.descactividad)
        row += 1

        worksheet.write(row, 0, 'NIT:')
        worksheet.write(row, 1, company.nit)
        row += 1

        worksheet.write(row, 0, 'NRC:')
        worksheet.write(row, 1, company.nrc)
        row += 1

        worksheet.write(row, 0, 'MES:')
        worksheet.write(row, 1, MESES[int(self.date_from.strftime('%m')) - 1])
        row += 1

        worksheet.write(row, 0, 'AÑO:')
        worksheet.write(row, 1, self.date_from.strftime('%Y'))
        row += 1

        # Document body
        row += 2

        worksheet.set_column(0, 0, 20) # NIT
        worksheet.set_column(1, 1, 15) # FECHA
        worksheet.set_column(3, 2, 50) # SERIE
        worksheet.set_column(4, 3, 50) # NUM DOC
        worksheet.set_column(5, 4, 14) # MONTO SUJETO
        worksheet.set_column(6, 5, 16) # MONTO ANTICIPO
        worksheet.set_column(7, 6, 20) # DUI
        worksheet.set_column(8, 7, 10) # ANEXO
        
        # Headers
        worksheet.write(row, 0, 'NIT', bold)
        worksheet.write(row, 1, 'FECHA', bold)
        worksheet.write(row, 2, 'SERIE', bold)
        worksheet.write(row, 3, 'NUM DOC', bold)
        worksheet.write(row, 4, 'MONTO\nSUJETO', bold)
        worksheet.write(row, 5, 'MONTO\nANTICIPO IVA 2%', bold)
        worksheet.write(row, 6, 'DUI', bold)
        worksheet.write(row, 7, 'ANEXO', bold)

        start_date = self.date_from
        end_date = self.date_to

        row += 1 # skip headers

        search_domain = [
            ('name', '=like', 'DCLEP/%'),
            ('invoice_date', '>=', start_date.strftime('%Y-%m-%d')),
            ('invoice_date', '<=', end_date.strftime('%Y-%m-%d')),
        ]

        data_to_export = self.env['account.move'].search(search_domain, order='invoice_date asc, id asc')

        global_sujeto = 0
        global_anticipo = 0

        for record in data_to_export:

            partner = record.partner_id
            num_doc = partner.numdocumento if partner.numdocumento else ''

            monto_anticipo = 0
            monto_sujeto = 0

            for line in record.line_ids:

                monto_anticipo += line.price_unit * line.quantity
            
            monto_sujeto = monto_anticipo / 0.02

            nit = ""
            dui = ""

            if(partner.nit):
                nit = partner.nit
                dui = ""
            else:
                nit = ""
                dui = num_doc

            worksheet.write(row, 0, nit, center)
            worksheet.write(row, 1, record.date.strftime('%d/%m/%Y'), center)
            worksheet.write(row, 2, record.sello_recepcion, center)
            worksheet.write(row, 3, self.formatWithoutDashes(record.codigo_generacion), center)
            worksheet.write(row, 4, format(monto_sujeto, '.2f'), center)
            worksheet.write(row, 5, format(monto_anticipo, '.2f'), center)
            worksheet.write(row, 6, dui, center)
            worksheet.write(row, 7, '6', center)

            row += 1

            global_sujeto += monto_sujeto
            global_anticipo += monto_anticipo

        # Document footer
        worksheet.write(row, 3, 'TOTALES', bold)
        worksheet.write(row, 4, format(global_sujeto, '.2f'), bold)
        worksheet.write(row, 5, format(global_anticipo, '.2f'), bold)

        workbook.close()

        return filename
    
    def make_fex_xls(self, filename):
        date = self.date_from.strftime('%m-%Y')
        filename = filename + ' ' + date + '.xls'

        row = 0

        workbook = xlsxwriter.Workbook(filename)
        bold = workbook.add_format({'bold': True})
        worksheet = workbook.add_worksheet()

        # Document header

        # add the company logo to the excel file if it exists
        company = self.env.company
        if company.logo:
            logo_stream = BytesIO(base64.b64decode(company.logo))
            worksheet.insert_image('K1', 'logo.png', {'image_data': logo_stream})

        worksheet.write(row, 0, 'DETALLE DE FACTURAS DE EXPORTACION', bold)
        row += 2

        worksheet.write(row, 0, company.name, bold)
        row += 1

        worksheet.write(row, 0, 'ACTIVIDAD ECONOMICA:')
        worksheet.write(row, 1, company.descactividad)
        row += 1

        worksheet.write(row, 0, 'NIT:')
        worksheet.write(row, 1, company.nit)
        row += 1

        worksheet.write(row, 0, 'NRC:')
        worksheet.write(row, 1, company.nrc)
        row += 1

        worksheet.write(row, 0, 'MES:')
        worksheet.write(row, 1, MESES[int(self.date_from.strftime('%m')) - 1])
        row += 1

        worksheet.write(row, 0, 'AÑO:')
        worksheet.write(row, 1, self.date_from.strftime('%Y'))
        row += 1

        # Document body
        row += 2

        # Headers
        worksheet.write(row, 0, 'No.', bold)
        worksheet.write(row, 1, 'FECHA', bold)
        worksheet.write(row, 2, 'NUMERO DE CONTROL INTERNO (DEL)', bold)
        worksheet.write(row, 3, 'NUMERO DE CONTROL INTERNO (AL)', bold)
        worksheet.write(row, 4, 'NUMERO DE GENERACION (DEL)', bold)
        worksheet.write(row, 5, 'NUMERO DE GENERACION (AL)', bold)
        worksheet.write(row, 6, 'VENTAS EXENTAS', bold)
        worksheet.write(row, 7, 'VENTAS NO SUJETAS', bold)
        worksheet.write(row, 8, 'VENTAS GRAVADAS', bold)
        worksheet.write(row, 9, 'VENTA A CUENTA DE TERCEROS', bold)
        worksheet.write(row, 10, 'EXPORTACIONES DE SERVICIOS', bold)
        worksheet.write(row, 11, 'IVA RETENIDO', bold)
        worksheet.write(row, 12, 'TOTAL VENTAS', bold)

        start_date = self.date_from
        end_date = self.date_to

        row += 1 # skip headers
        n = 1

        global_exenta = 0
        global_gravada = 0
        global_iva_retenido = 0
        global_total = 0
        global_iva = 0
        global_exportacion = 0

        # FEX
        # recorrer los dias del rango de fechas seleccionado
        for i in range((end_date - start_date).days + 1):

            # dia
            day = start_date + timedelta(days=i)

            # obtener todos los movimientos de FC de este dia
            search_domain = [
                '&',
                    ('name', '=like', 'FEXE/%'),
                    ('invoice_date', '=', day.strftime('%Y-%m-%d')),
            ]

            data_to_export = self.env['account.move'].search(search_domain, order='invoice_date asc, id asc')

            total_exenta = 0
            total_no_suj = 0
            total_gravada = 0
            total_terceros = 0
            total_exportacion = 0
            total_iva_retenido = 0
            total = 0
            
            resultados = []
            rango_inicio = None
            rango_fin = None
            rango_total = 0

            for record in data_to_export:

                confirmation = self.get_confirmation_for_tax_books(record)

                IsAnulated = False

                if confirmation['status'] == 0 and confirmation['uid_de'] == '-':
                    continue
                elif confirmation['status'] == 0 and confirmation['uid_de'] != '-':
                    IsAnulated = True

                for line in record.line_ids:
                    if line.product_id.default_code:
                        total_exportacion += line.price_subtotal if not IsAnulated else 0
                                        
                total = round((total_exportacion), 2) if not IsAnulated else 0

                # if IsAnulated:
                if False:

                    # Si había un rango abierto de procesados, lo cerramos primero
                    if rango_inicio:
                        resultados.append({
                            # "n": n,
                            "fecha": day.strftime('%d-%m-%Y'),
                            "numero_control_del": rango_inicio['control_number'],
                            "numero_control_al": rango_fin['control_number'],
                            "codigo_generacion_del": rango_inicio['generation_code'],
                            "codigo_generacion_al": rango_fin['generation_code'],
                            "exentas": total_exenta,
                            "no_sujetas": total_no_suj,
                            "gravadas": total_gravada,
                            "a_terceros": total_terceros,
                            "exportacion_servicios": total_exportacion,
                            "iva_retenido": total_iva_retenido,
                            "total": rango_total,
                        })
                        rango_inicio = rango_fin = None
                        rango_total = 0

                    # Agregamos el anulado solito
                    resultados.append({
                        # "n": n,
                        "fecha": day.strftime('%d-%m-%Y'),
                        "numero_control_del": confirmation['control_number'],
                        "numero_control_al": confirmation['control_number'],
                        "codigo_generacion_del": confirmation['uid_de'],
                        "codigo_generacion_al": confirmation['uid_de'],
                        "exentas": 0,
                        "no_sujetas": 0,
                        "gravadas": 0,
                        "a_terceros": 0,
                        "exportacion_servicios": 0,
                        "iva_retenido": 0,
                        "total": 0,
                    })

                    # reiniciamos acumuladores
                    total_exenta = 0
                    total_gravada = 0
                    total_iva_retenido = 0
                    total = 0
                else:
                    # Procesado: acumular
                    if not rango_inicio:
                        rango_inicio = {'record': record, 'generation_code': confirmation['uid_de'], 'control_number': confirmation['control_number']}
                    rango_fin = {'record': record, 'generation_code': confirmation['uid_de'], 'control_number': confirmation['control_number']}
                    rango_total = total

            # Al terminar el loop, guardar último bloque abierto
            if rango_inicio:
                resultados.append({
                    # "n": n,
                    "fecha": day.strftime('%d-%m-%Y'),
                    "numero_control_del": rango_inicio['control_number'],
                    "numero_control_al": rango_fin['control_number'],
                    "codigo_generacion_del": rango_inicio['generation_code'],
                    "codigo_generacion_al": rango_fin['generation_code'],
                    "exentas": total_exenta,
                    "no_sujetas": total_no_suj,
                    "gravadas": total_gravada,
                    "a_terceros": total_terceros,
                    "exportacion_servicios": total_exportacion,
                    "iva_retenido": total_iva_retenido,
                    "total": rango_total,
                })

            # Mostrar en logs
            for r in resultados:

                worksheet.write(row, 0, n)
                worksheet.write(row, 1, r['fecha'])
                worksheet.write(row, 2, r['numero_control_del'])
                worksheet.write(row, 3, r['numero_control_al'])
                worksheet.write(row, 4, r['codigo_generacion_del'])
                worksheet.write(row, 5, r['codigo_generacion_al'])
                worksheet.write(row, 6, format(r['exentas'], '.2f'))
                worksheet.write(row, 7, format(r['no_sujetas'], '.2f'))
                worksheet.write(row, 8, format(r['gravadas'], '.2f'))
                worksheet.write(row, 9, format(r['a_terceros'], '.2f'))
                worksheet.write(row, 10, format(r['exportacion_servicios'], '.2f'))
                worksheet.write(row, 11, format(r['iva_retenido'], '.2f'))
                worksheet.write(row, 12, format(r['total'], '.2f'))
                
                row += 1

                global_exenta += r['exentas']
                global_gravada += r['gravadas']
                global_iva_retenido += r['iva_retenido']
                global_total += r['total']
                global_exportacion += r['exportacion_servicios']
        
                n += 1

        # Document footer
        worksheet.write(row, 5, 'TOTALES', bold)
        worksheet.write(row, 6, format(global_exenta, '.2f'), bold)
        worksheet.write(row, 7, format(0, '.2f'), bold)
        worksheet.write(row, 8, format(global_gravada, '.2f'), bold)
        worksheet.write(row, 9, format(0, '.2f'), bold)
        worksheet.write(row, 10, format(global_exportacion, '.2f'), bold)
        worksheet.write(row, 11, format(global_iva_retenido, '.2f'), bold)
        worksheet.write(row, 12, format(global_total, '.2f'), bold)

        row += 4

        # Document footer
        worksheet.merge_range('F'+str(row)+':F'+str((row + 1)), 'RESUMEN DE OPERACIONES', bold)
        worksheet.merge_range('G'+str(row)+':I'+str(row), 'VENTAS PROPIAS', bold)
        worksheet.merge_range('J'+str(row)+':L'+str(row), 'A CUENTA DE TERCEROS', bold)

        row += 1

        worksheet.write((row - 1), 6, 'Valor \nneto', bold)
        worksheet.write((row - 1), 7, 'IVA \ndebito', bold)
        worksheet.write((row - 1), 8, 'Valor \ntotal', bold)
        worksheet.write((row - 1), 9, 'Valor \nneto', bold)
        worksheet.write((row - 1), 10, 'IVA \ndebito', bold)
        worksheet.write((row - 1), 11, 'Valor \ntotal', bold)

        worksheet.write((row), 5, 'Exportaciones segun FEX')
        worksheet.write((row), 6, global_exportacion)
        worksheet.write((row), 7, 0)
        worksheet.write((row), 8, (global_exportacion))
        worksheet.write((row), 9, 0)
        worksheet.write((row), 10, 0)
        worksheet.write((row), 11, 0)

        row += 2

        worksheet.write((row), 5, 'Firma:', bold)

        workbook.close()

        return filename
    
    def make_fe_xls(self, filename):
        date = self.date_from.strftime('%m-%Y')
        filename = filename + ' ' + date + '.xls'

        row = 0

        workbook = xlsxwriter.Workbook(filename)
        bold = workbook.add_format({'bold': True})
        worksheet = workbook.add_worksheet()

        # Document header

        # add the company logo to the excel file if it exists
        company = self.env.company
        if company.logo:
            logo_stream = BytesIO(base64.b64decode(company.logo))
            worksheet.insert_image('K1', 'logo.png', {'image_data': logo_stream})

        worksheet.write(row, 0, 'LIBRO DE VENTAS CONSUMIDOR FINAL', bold)
        row += 2

        worksheet.write(row, 0, company.name, bold)
        row += 1

        worksheet.write(row, 0, 'ACTIVIDAD ECONOMICA:')
        worksheet.write(row, 1, company.descactividad)
        row += 1

        worksheet.write(row, 0, 'NIT:')
        worksheet.write(row, 1, company.nit)
        row += 1

        worksheet.write(row, 0, 'NRC:')
        worksheet.write(row, 1, company.nrc)
        row += 1

        worksheet.write(row, 0, 'MES:')
        worksheet.write(row, 1, MESES[int(self.date_from.strftime('%m')) - 1])
        row += 1

        worksheet.write(row, 0, 'AÑO:')
        worksheet.write(row, 1, self.date_from.strftime('%Y'))
        row += 1

        # Document body
        row += 2

        # Headers
        worksheet.write(row, 0, 'No.', bold)
        worksheet.write(row, 1, 'FECHA', bold)
        worksheet.write(row, 2, 'NUMERO DE CONTROL INTERNO (DEL)', bold)
        worksheet.write(row, 3, 'NUMERO DE CONTROL INTERNO (AL)', bold)
        worksheet.write(row, 4, 'NUMERO DE GENERACION (DEL)', bold)
        worksheet.write(row, 5, 'NUMERO DE GENERACION (AL)', bold)
        worksheet.write(row, 6, 'VENTAS EXENTAS', bold)
        worksheet.write(row, 7, 'VENTAS NO SUJETAS', bold)
        worksheet.write(row, 8, 'VENTAS GRAVADAS', bold)
        worksheet.write(row, 9, 'VENTA A CUENTA DE TERCEROS', bold)
        worksheet.write(row, 10, 'EXPORTACIONES DE SERVICIOS', bold)
        worksheet.write(row, 11, 'IVA RETENIDO', bold)
        worksheet.write(row, 12, 'TOTAL VENTAS', bold)

        start_date = self.date_from
        end_date = self.date_to

        row += 1 # skip headers
        n = 1

        global_exenta = 0
        global_gravada = 0
        global_iva_retenido = 0
        global_total = 0
        global_iva = 0
        global_exportacion = 0

        search_domain = [
            ('name', '=like', 'FE/%'),
            ('date', '>=', start_date.strftime('%Y-%m-%d')),
            ('date', '<=', end_date.strftime('%Y-%m-%d')),
        ]

        n = 1

        # recorrer los dias del rango de fechas seleccionado
        for i in range((end_date - start_date).days + 1):

            # dia
            day = start_date + timedelta(days=i)

            # obtener todos los movimientos de FC de este dia
            search_domain = [
                '&',
                    ('name', '=like', 'FE/%'),
                    ('invoice_date', '=', day.strftime('%Y-%m-%d')),
            ]

            data_to_export = self.env['account.move'].search(search_domain, order='invoice_date asc, id asc')

            total_exenta = 0
            total_no_suj = 0
            total_gravada = 0
            total_terceros = 0
            total_exportacion = 0
            total_iva_retenido = 0
            total = 0
            
            resultados = []
            rango_inicio = None
            rango_fin = None
            rango_total = 0

            for record in data_to_export:

                confirmation = self.get_confirmation_for_tax_books(record)

                IsAnulated = False

                if confirmation['status'] == 0 and confirmation['uid_de'] == '-':
                    continue
                elif confirmation['status'] == 0 and confirmation['uid_de'] != '-':
                    IsAnulated = True

                for line in record.line_ids:
                    if line.product_id.default_code:

                        taxes = line.tax_ids
                        exenta = True
                        ivaItem = 0

                        if taxes:
                            for tax in taxes:
                                if 'IVA 13%' in tax.name:
                                    exenta = False
                                    ivaItem = line.price_subtotal * 0.13 if not IsAnulated else 0
                                    global_iva += ivaItem
                                
                                if 'RET 1%' in tax.name:
                                    total_iva_retenido += line.price_subtotal * 0.01 if not IsAnulated else 0

                        if exenta:
                            total_exenta += line.price_subtotal if not IsAnulated else 0
                        else:
                            total_gravada += line.price_subtotal + ivaItem if not IsAnulated else 0
                
                total = round((total_exenta + total_gravada), 2) if not IsAnulated else 0

                # if IsAnulated:
                if False:

                    # Si había un rango abierto de procesados, lo cerramos primero
                    if rango_inicio:
                        resultados.append({
                            # "n": n,
                            "fecha": day.strftime('%d-%m-%Y'),
                            "numero_control_del": rango_inicio['control_number'],
                            "numero_control_al": rango_fin['control_number'],
                            "codigo_generacion_del": rango_inicio['generation_code'],
                            "codigo_generacion_al": rango_fin['generation_code'],
                            "exentas": total_exenta,
                            "no_sujetas": total_no_suj,
                            "gravadas": total_gravada,
                            "a_terceros": total_terceros,
                            "exportacion_servicios": total_exportacion,
                            "iva_retenido": total_iva_retenido,
                            "total": rango_total,
                        })
                        rango_inicio = rango_fin = None
                        rango_total = 0

                    # Agregamos el anulado solito
                    resultados.append({
                        # "n": n,
                        "fecha": day.strftime('%d-%m-%Y'),
                        "numero_control_del": confirmation['control_number'],
                        "numero_control_al": confirmation['control_number'],
                        "codigo_generacion_del": confirmation['uid_de'],
                        "codigo_generacion_al": confirmation['uid_de'],
                        "exentas": 0,
                        "no_sujetas": 0,
                        "gravadas": 0,
                        "a_terceros": 0,
                        "exportacion_servicios": 0,
                        "iva_retenido": 0,
                        "total": 0,
                    })

                    # reiniciamos acumuladores
                    total_exenta = 0
                    total_gravada = 0
                    total_iva_retenido = 0
                    total = 0
                else:
                    # Procesado: acumular
                    if not rango_inicio:
                        rango_inicio = {'record': record, 'generation_code': confirmation['uid_de'], 'control_number': confirmation['control_number']}
                    rango_fin = {'record': record, 'generation_code': confirmation['uid_de'], 'control_number': confirmation['control_number']}
                    rango_total = total

            # Al terminar el loop, guardar último bloque abierto
            if rango_inicio:
                resultados.append({
                    # "n": n,
                    "fecha": day.strftime('%d-%m-%Y'),
                    "numero_control_del": rango_inicio['control_number'],
                    "numero_control_al": rango_fin['control_number'],
                    "codigo_generacion_del": rango_inicio['generation_code'],
                    "codigo_generacion_al": rango_fin['generation_code'],
                    "exentas": total_exenta,
                    "no_sujetas": total_no_suj,
                    "gravadas": total_gravada,
                    "a_terceros": total_terceros,
                    "exportacion_servicios": total_exportacion,
                    "iva_retenido": total_iva_retenido,
                    "total": rango_total,
                })

            # Mostrar en logs
            for r in resultados:
                
                worksheet.write(row, 0, n)
                worksheet.write(row, 1, r['fecha'])
                worksheet.write(row, 2, r['numero_control_del'])
                worksheet.write(row, 3, r['numero_control_al'])
                worksheet.write(row, 4, r['codigo_generacion_del'])
                worksheet.write(row, 5, r['codigo_generacion_al'])
                worksheet.write(row, 6, format(r['exentas'], '.2f'))
                worksheet.write(row, 7, format(r['no_sujetas'], '.2f'))
                worksheet.write(row, 8, format(r['gravadas'], '.2f'))
                worksheet.write(row, 9, format(r['a_terceros'], '.2f'))
                worksheet.write(row, 10, format(r['exportacion_servicios'], '.2f'))
                worksheet.write(row, 11, format(r['iva_retenido'], '.2f'))
                worksheet.write(row, 12, format(r['total'], '.2f'))
                
                row += 1

                global_exenta += r['exentas']
                global_gravada += r['gravadas']
                global_iva_retenido += r['iva_retenido']
                global_total += r['total']
        
                n += 1

        # Document footer
        worksheet.write(row, 5, 'TOTALES', bold)
        worksheet.write(row, 6, format(global_exenta, '.2f'), bold)
        worksheet.write(row, 7, format(0, '.2f'), bold)
        worksheet.write(row, 8, format(global_gravada, '.2f'), bold)
        worksheet.write(row, 9, format(0, '.2f'), bold)
        worksheet.write(row, 10, format(global_exportacion, '.2f'), bold)
        worksheet.write(row, 11, format(global_iva_retenido, '.2f'), bold)
        worksheet.write(row, 12, format(global_total, '.2f'), bold)

        row += 4

        # Document footer
        worksheet.merge_range('F'+str(row)+':F'+str((row + 1)), 'RESUMEN DE OPERACIONES', bold)
        worksheet.merge_range('G'+str(row)+':I'+str(row), 'VENTAS PROPIAS', bold)
        worksheet.merge_range('J'+str(row)+':L'+str(row), 'A CUENTA DE TERCEROS', bold)

        row += 1

        worksheet.write((row - 1), 6, 'Valor \nneto', bold)
        worksheet.write((row - 1), 7, 'IVA \ndebito', bold)
        worksheet.write((row - 1), 8, 'Valor \ntotal', bold)
        worksheet.write((row - 1), 9, 'Valor \nneto', bold)
        worksheet.write((row - 1), 10, 'IVA \ndebito', bold)
        worksheet.write((row - 1), 11, 'Valor \ntotal', bold)

        worksheet.write((row), 5, 'Ventas netas gravadas a consumidores')
        worksheet.write((row), 6, global_gravada - global_iva)
        worksheet.write((row), 7, global_iva)
        worksheet.write((row), 8, (global_gravada))
        worksheet.write((row), 9, 0)
        worksheet.write((row), 10, 0)
        worksheet.write((row), 11, 0)

        row += 1

        worksheet.write((row), 5, 'Ventas netas exentas a consumidores')
        worksheet.write((row), 6, global_exenta)
        worksheet.write((row), 7, 0)
        worksheet.write((row), 8, (global_exenta))
        worksheet.write((row), 9, 0)
        worksheet.write((row), 10, 0)
        worksheet.write((row), 11, 0)

        row += 2

        worksheet.write((row), 5, 'Firma:', bold)

        workbook.close()

        return filename

    def make_ccfe_xls(self, filename):
        date = self.date_from.strftime('%m-%Y')
        filename = filename + ' ' + date + '.xls'

        row = 0

        workbook = xlsxwriter.Workbook(filename)

        bold = workbook.add_format({
            'bold': True, 
            'align': 'center', 
            'valign': 'vcenter', 
            'text_wrap': True
        })

        bold_left = workbook.add_format({
            'bold': True, 
            'align': 'left', 
            'valign': 'vcenter', 
            'text_wrap': True
        })

        center = workbook.add_format({
            'bold': False, 
            'align': 'center', 
            'valign': 'vcenter', 
            'text_wrap': True
        })

        left = workbook.add_format({
            'bold': False, 
            'align': 'left', 
            'valign': 'vcenter', 
            'text_wrap': True
        })

        right = workbook.add_format({
            'bold': False, 
            'align': 'right', 
            'valign': 'vcenter', 
            'text_wrap': True
        })

        worksheet = workbook.add_worksheet()

        # Document header

        # add the company logo to the excel file if it exists
        company = self.env.company
        if company.logo:
            logo_stream = BytesIO(base64.b64decode(company.logo))
            worksheet.insert_image('K1', 'logo.png', {'image_data': logo_stream})

        worksheet.merge_range('A1:D1', 'LIBRO DE VENTAS A CONTRIBUYENTES', bold)
        row += 2

        worksheet.write(row, 0, company.name)
        row += 1

        worksheet.write(row, 0, 'ACTIVIDAD ECONOMICA:')
        worksheet.write(row, 1, company.descactividad)
        row += 1

        worksheet.write(row, 0, 'NIT:')
        worksheet.write(row, 1, company.nit)
        row += 1

        worksheet.write(row, 0, 'NRC:')
        worksheet.write(row, 1, company.nrc)
        row += 1

        worksheet.write(row, 0, 'MES:')
        worksheet.write(row, 1, MESES[int(self.date_from.strftime('%m')) - 1])
        row += 1

        worksheet.write(row, 0, 'AÑO:')
        worksheet.write(row, 1, self.date_from.strftime('%Y'))
        row += 1

        # Document body
        row += 3

        worksheet.set_column(0, 0, 8)   # No.
        worksheet.set_column(1, 1, 15)   # Fecha
        worksheet.set_column(2, 2, 50)   # Num. doc.
        worksheet.set_column(3, 3, 15)   # NRC
        worksheet.set_column(4, 4, 20)   # NIT
        worksheet.set_column(5, 5, 60)   # NOMBRE CLIENTE
        worksheet.set_column(6, 6, 10)   # EXENTAS
        worksheet.set_column(7, 7, 10)   # NO SUJETAS
        worksheet.set_column(8, 8, 10)   # GRAVADAS
        worksheet.set_column(9, 9, 10)   # DEBITO FSICAL
        worksheet.set_column(10, 10, 10)   # GRAVADAS
        worksheet.set_column(11, 11, 10)   # IVA
        worksheet.set_column(12, 12, 10)   # IVA RETENIDO
        worksheet.set_column(13, 13, 10)   # TOTAL
        
        # Headers
        worksheet.merge_range('A11:A12', 'No.', bold)
        worksheet.merge_range('B11:B12', 'Fecha', bold)
        worksheet.merge_range('C11:C12', 'NUMERO DE DOCUMENTO', bold)
        worksheet.merge_range('D11:D12', 'NRC.', bold)
        worksheet.merge_range('E11:E12', 'NIT', bold)
        worksheet.merge_range('F11:F12', 'NOMBRE CLIENTE', bold)
        worksheet.merge_range('G11:J11', 'VENTAS PROPIAS', bold)
        worksheet.merge_range('K11:L11', 'VENTAS A TERCEROS', bold)
        worksheet.merge_range('M11:M12', 'IVA\nRETENIDO', bold)
        worksheet.merge_range('N11:N12', 'TOTAL', bold)

        worksheet.write(row, 6, 'EXENT.', bold)
        worksheet.write(row, 7, 'NO SUJ.', bold)
        worksheet.write(row, 8, 'GRAV', bold)
        worksheet.write(row, 9, 'DEBITO\nFISCAL', bold)
        worksheet.write(row, 10, 'GRAV', bold)
        worksheet.write(row, 11, 'IVA', bold)

        start_date = self.date_from
        end_date = self.date_to

        row += 1 # skip headers
        n = 1

        global_exenta = 0
        global_gravada = 0
        grand_total = 0
        global_iva = 0
        global_iva_retenido = 0
        global_total = 0

        search_domain = [
            ('invoice_date', '>=', start_date.strftime('%Y-%m-%d')),
            ('invoice_date', '<=', end_date.strftime('%Y-%m-%d')),
            '|', '|',
            ('name', 'ilike', 'CCFE/'),
            ('name', 'ilike', 'RNCE/'),
            ('name', 'ilike', 'RNDE/'),
        ]

        data_to_export = self.env['account.move'].search(search_domain, order='invoice_date asc, id asc')
        
        numeroDocumento = 0

        for record in data_to_export:

            confirmation = self.get_confirmation_for_tax_books(record)

            isAnulated = False

            if(confirmation['status'] == 0 and confirmation['uid_de'] == '-'):
                continue
            elif(confirmation['status'] == 0 and confirmation['uid_de'] != '-'):
                isAnulated = True

            numeroDocumento = str(confirmation['uid_de'])

            partner = record.partner_id
            num_doc = partner.numdocumento if partner.numdocumento else ''

            total_exenta = 0
            total_gravada = 0
            total_gravada_terceros = 0
            total_iva_tercerros = 0
            total_iva_retenido = 0
            total_no_suj = 0

            for line in record.line_ids:
                if line.product_id.default_code:
                        
                    taxes = line.tax_ids
                    exenta = True

                    if taxes:
                        for tax in taxes:
                            if 'IVA 13%' in tax.name:
                                exenta = False

                            if 'RET 1%' in tax.name:
                                total_iva_retenido += round((line.price_subtotal * 0.01), 2)

                    if exenta:
                        total = line.price_subtotal
                        total_exenta = round(total_exenta + total, 2)
                    else:
                        total = line.price_subtotal
                        total_gravada = round(total_gravada + total, 2)

            total_iva = round(total_gravada * 0.13, 2)
            gran_total = round(total_gravada + total_exenta + total_iva, 2)

            nit = ""
            dui = ""

            if(partner.nit):
                nit = partner.nit
                dui = ""
            else:
                nit = ""
                dui = num_doc
            
            if record.name.startswith('RNCE/') and not isAnulated:
                total_exenta = total_exenta * -1 if total_exenta > 0 else 0
                total_no_suj = total_no_suj * -1 if total_no_suj > 0 else 0
                total_gravada = total_gravada * -1 if total_gravada > 0 else 0
                total_iva = total_iva * -1 if total_iva > 0 else 0
                total_iva_retenido = total_iva_retenido * -1 if total_iva_retenido > 0 else 0
                gran_total = gran_total * -1 if gran_total > 0 else 0
            
            if isAnulated:
                total_exenta = 0
                total_no_suj = 0
                total_gravada = 0
                total_iva = 0
                total_iva_retenido = 0
                gran_total = 0

            nrc = partner.nrc if partner.nrc else ''
            name = partner.name if partner and not isAnulated else 'INVÁLIDO'

            worksheet.write(row, 0, n, center)
            worksheet.write(row, 1, record.invoice_date.strftime('%d/%m/%Y'), center)
            worksheet.write(row, 2, numeroDocumento, center)
            worksheet.write(row, 3, nrc, center)
            worksheet.write(row, 4, nit, center)
            worksheet.write(row, 5, name, left)
            worksheet.write(row, 6, total_exenta, center)
            worksheet.write(row, 7, total_no_suj, center)
            worksheet.write(row, 8, total_gravada, center)
            worksheet.write(row, 9, total_iva, center)
            worksheet.write(row, 10, total_gravada_terceros, center)
            worksheet.write(row, 11, total_iva_tercerros, center)
            worksheet.write(row, 12, total_iva_retenido, center)
            worksheet.write(row, 13, gran_total, center)

            global_total += gran_total
            global_exenta += total_exenta
            global_gravada += total_gravada
            grand_total += gran_total
            global_iva += total_iva
            global_iva_retenido += total_iva_retenido

            n += 1
            row += 1

        # Document footer
        worksheet.write(row, 5, 'TOTALES', bold_left)
        worksheet.write(row, 6, global_exenta, bold)
        worksheet.write(row, 7, 0, bold)
        worksheet.write(row, 8, global_gravada, bold)
        worksheet.write(row, 9, global_iva, bold)
        worksheet.write(row, 10, 0, bold)
        worksheet.write(row, 11, 0, bold)
        worksheet.write(row, 12, global_iva_retenido, bold)
        worksheet.write(row, 13, global_total, bold)

        row += 4

        # Document footer
        worksheet.merge_range('F'+str(row)+':F'+str((row + 1)), 'RESUMEN DE OPERACIONES', bold)
        worksheet.merge_range('G'+str(row)+':I'+str(row), 'VENTAS PROPIAS', bold)
        worksheet.merge_range('J'+str(row)+':L'+str(row), 'A CUENTA DE TERCEROS', bold)

        row += 1

        worksheet.write((row - 1), 6, 'Valor \nneto', bold)
        worksheet.write((row - 1), 7, 'IVA \ndebito', bold)
        worksheet.write((row - 1), 8, 'Valor \ntotal', bold)
        worksheet.write((row - 1), 9, 'Valor \nneto', bold)
        worksheet.write((row - 1), 10, 'IVA \ndebito', bold)
        worksheet.write((row - 1), 11, 'Valor \ntotal', bold)

        worksheet.write((row), 5, 'Ventas netas gravadas a contribuyentes', left)
        worksheet.write((row), 6, global_gravada, center)
        worksheet.write((row), 7, global_iva, center)
        worksheet.write((row), 8, (global_gravada + global_iva), center)
        worksheet.write((row), 9, 0, center)
        worksheet.write((row), 10, 0, center)
        worksheet.write((row), 11, 0, center)

        row += 1

        worksheet.write((row), 5, 'Ventas netas exentas a contribuyentes', left)
        worksheet.write((row), 6, global_exenta, center)
        worksheet.write((row), 7, 0, center)
        worksheet.write((row), 8, (global_exenta), center)
        worksheet.write((row), 9, 0, center)
        worksheet.write((row), 10, 0, center)
        worksheet.write((row), 11, 0, center)
        
        row += 2

        worksheet.write((row), 5, 'Firma:', bold_left)

        workbook.close()

        return filename

    def make_ccfep_xls(self, filename):
        date = self.date_from.strftime('%m-%Y')
        filename = filename + ' ' + date + '.xls'

        row = 0

        workbook = xlsxwriter.Workbook(filename)
        bold = workbook.add_format({'bold': True})
        worksheet = workbook.add_worksheet()

        # Document header

        # add the company logo to the excel file if it exists
        company = self.env['res.company'].search([], limit=1)
        if company.logo:
            logo_stream = BytesIO(base64.b64decode(company.logo))
            worksheet.insert_image('K1', 'logo.png', {'image_data': logo_stream})

        worksheet.write(row, 0, 'LIBRO DE COMPRAS', bold)
        row += 2

        worksheet.write(row, 0, company.name, bold)
        row += 1

        worksheet.write(row, 0, 'ACTIVIDAD ECONOMICA:')
        worksheet.write(row, 1, company.descactividad)
        row += 1

        worksheet.write(row, 0, 'NIT:')
        worksheet.write(row, 1, company.nit)
        row += 1

        worksheet.write(row, 0, 'NRC:')
        worksheet.write(row, 1, company.nrc)
        row += 1

        worksheet.write(row, 0, 'MES:')
        worksheet.write(row, 1, MESES[int(self.date_from.strftime('%m')) - 1])
        row += 1

        worksheet.write(row, 0, 'AÑO:')
        worksheet.write(row, 1, self.date_from.strftime('%Y'))
        row += 1

        # Document body
        row += 2

        # Headers
        worksheet.write(row, 0, 'No.', bold)
        worksheet.write(row, 1, 'FECHA', bold)
        worksheet.write(row, 2, 'CLASE DE DOCUMENTO', bold)
        worksheet.write(row, 3, 'TIPO DE DOCUMENTO', bold)
        worksheet.write(row, 4, 'NUMERO DE DOCUMENTO', bold)
        worksheet.write(row, 5, 'NIT O NRC PROVEEDOR', bold)
        worksheet.write(row, 6, 'NOMBRE PROVEEDOR', bold)
        worksheet.write(row, 7, 'COMPRAS INTERNAS EXENTAS', bold)
        worksheet.write(row, 8, 'INTERNACIONES EXENTAS Y/ NO SUJETAS', bold)
        worksheet.write(row, 9, 'IMPORTACIONES EXENTAS Y/O NO SUJETAS', bold)
        worksheet.write(row, 10, 'COMPRAS INTERNAS GRAVADAS', bold)
        worksheet.write(row, 11, 'INTERNACIONES GRAVADAS DE BIENES', bold)
        worksheet.write(row, 12, 'IMPORTACIONES GRAVADAS DE BIENES', bold)
        worksheet.write(row, 13, 'IMPORTACIONES GRAVADAS DE SERVICIOS', bold)
        worksheet.write(row, 14, 'CREDITO FISCAL', bold)
        worksheet.write(row, 15, 'PERCEPCION', bold)
        worksheet.write(row, 16, 'TOTAL DE COMPRAS', bold)

        start_date = self.date_from
        end_date = self.date_to

        search_domain = [
            '|',  # OR lógico
            ('name', 'like', 'CCFEP%'),
            ('name', 'like', 'RNCEP%'),
            ('date', '>=', start_date.strftime('%Y-%m-%d')),
            ('date', '<=', end_date.strftime('%Y-%m-%d')),
        ]

        data_to_export = self.env['account.move'].search(search_domain, order='date asc')

        row += 1 # skip headers
        n = 1

        global_exenta = 0
        global_gravada = 0
        gran_total = 0
        global_iva = 0
        global_importaciones = 0
        global_percepcion = 0
        global_total = 0

        for record in data_to_export:
            partner = record.partner_id
            num_doc = partner.numdocumento if partner.numdocumento else ''

            is_internal = partner.is_internal if partner.is_internal else 'SI'

            total_exenta = 0
            total_gravada = 0
            total_percepcion = 0

            for line in record.line_ids:
                taxes = line.tax_ids
                exenta = True
                if taxes:
                    for tax in taxes:
                        if 'IVA 13%' in tax.name:
                            exenta = False

                        if 'PER 1%' in tax.name:
                            total_percepcion = round((total_percepcion + round(line.price_subtotal * 0.01, 4)), 4) # Obtener la percepcion del 1%

                if exenta:
                    total = line.price_subtotal
                    total_exenta = round(total_exenta + total, 2)
                else:
                    total = line.price_subtotal
                    total_gravada = round(total_gravada + total, 2)

            total_iva = round(total_gravada * 0.13, 2)
            gran_total = round(total_gravada + total_exenta + total_iva + total_percepcion, 2)

            nit = ""
            dui = ""

            if(partner.nit):
                nit = partner.nit
                dui = ""
            else:
                nit = ""
                dui = num_doc

            num_tax_doc = record.codigo_generacion if record.codigo_generacion else ''

            worksheet.write(row, 0, n)
            worksheet.write(row, 1, record.date.strftime('%d/%m/%Y'))
            worksheet.write(row, 2, '4')
            worksheet.write(row, 3, '03')
            worksheet.write(row, 4, num_tax_doc)
            worksheet.write(row, 5, nit)
            worksheet.write(row, 6, partner.name)
            worksheet.write(row, 7, total_exenta)
            worksheet.write(row, 8, 0)
            worksheet.write(row, 9, 0)

            if(is_internal == 'SI'):
                global_gravada += total_gravada
                worksheet.write(row, 10, total_gravada)
            else:
                worksheet.write(row, 10, 0)

            worksheet.write(row, 11, 0)
            worksheet.write(row, 12, 0)

            if(is_internal == 'SI'):
                worksheet.write(row, 13, 0)
            else:
                global_importaciones += total_gravada
                worksheet.write(row, 13, total_gravada)


            worksheet.write(row, 14, total_iva)
            worksheet.write(row, 15, total_percepcion)
            worksheet.write(row, 16, gran_total)

            global_exenta += total_exenta
            # gran_total += gran_total
            global_iva += total_iva
            global_percepcion += total_percepcion
            global_total += gran_total

            n += 1
            row += 1

        # Document footer
        worksheet.write(row, 6, 'TOTALES', bold)
        worksheet.write(row, 7, global_exenta, bold)
        worksheet.write(row, 10, global_gravada, bold)
        worksheet.write(row, 13, global_importaciones, bold)
        worksheet.write(row, 14, global_iva, bold)
        worksheet.write(row, 15, global_percepcion, bold)
        worksheet.write(row, 16, global_total, bold)

        workbook.close()

        return filename

    def make_cre_xls(self, filename):
        date = self.date_from.strftime('%m-%Y')
        filename = filename + ' ' + date + '.xls'

        row = 0

        workbook = xlsxwriter.Workbook(filename)
        bold = workbook.add_format({'bold': True})
        worksheet = workbook.add_worksheet()

        # Document header

        # add the company logo to the excel file if it exists
        company = self.env['res.company'].search([], limit=1)
        if company.logo:
            logo_stream = BytesIO(base64.b64decode(company.logo))
            worksheet.insert_image('K1', 'logo.png', {'image_data': logo_stream})

        worksheet.write(row, 0, 'LIBRO DE RETENCIONES 1%', bold)
        row += 2

        worksheet.write(row, 0, company.name, bold)
        row += 1

        worksheet.write(row, 0, 'ACTIVIDAD ECONOMICA:')
        worksheet.write(row, 1, company.descactividad)
        row += 1

        worksheet.write(row, 0, 'NIT:')
        worksheet.write(row, 1, company.nit)
        row += 1

        worksheet.write(row, 0, 'NRC:')
        worksheet.write(row, 1, company.nrc)
        row += 1

        worksheet.write(row, 0, 'MES:')
        worksheet.write(row, 1, MESES[int(self.date_from.strftime('%m')) - 1])
        row += 1

        worksheet.write(row, 0, 'AÑO:')
        worksheet.write(row, 1, self.date_from.strftime('%Y'))
        row += 1

        # Document body
        row += 2

        # Headers
        worksheet.write(row, 0, 'No.', bold)
        worksheet.write(row, 1, 'NIT DEL SUJETO', bold)
        worksheet.write(row, 2, 'FECHA', bold)
        worksheet.write(row, 3, 'TIPO DE DOCUMENTO', bold)
        worksheet.write(row, 4, 'NUMERO DE DOCUMENTO', bold)
        worksheet.write(row, 5, 'MONTO SUJETO', bold)
        worksheet.write(row, 6, 'MONTO DE LA RETENCION 1%', bold)
        worksheet.write(row, 7, 'DUI SUJETO', bold)

        start_date = self.date_from
        end_date = self.date_to

        search_domain = [
            ('name', '=like', 'CCFEP%'),
            ('date', '>=', start_date.strftime('%Y-%m-%d')),
            ('date', '<=', end_date.strftime('%Y-%m-%d')),
        ]

        data_to_export = self.env['account.move'].search(search_domain, order='date asc')

        row += 1 # skip headers
        n = 1

        global_sujeto = 0
        global_retencion = 0

        for record in data_to_export:
            retencion = False
            monto_sujeto = 0

            for line in record.line_ids:
                taxes = line.tax_ids
                if taxes:
                    for tax in taxes:
                        if tax.name == 'RET 1%':
                            retencion = True
                            monto_sujeto = monto_sujeto + line.price_subtotal

            if not retencion:
                continue

            partner = record.partner_id
            num_doc = partner.numdocumento if partner.numdocumento else ''

            nit = ""
            dui = ""

            if(partner.nit):
                nit = partner.nit
                dui = ""
            else:
                nit = ""
                dui = num_doc

            worksheet.write(row, 0, n)
            worksheet.write(row, 1, nit)
            worksheet.write(row, 2, record.date.strftime('%d/%m/%Y'))
            worksheet.write(row, 3, '07')
            worksheet.write(row, 4, record.purchase_doc_number)
            worksheet.write(row, 5, monto_sujeto)
            worksheet.write(row, 6, round(monto_sujeto * 0.01, 2))
            worksheet.write(row, 7, dui)

            global_sujeto += monto_sujeto
            global_retencion += round(monto_sujeto * 0.01, 2)

            n += 1
            row += 1

        # Document footer
        worksheet.write(row, 4, 'TOTALES', bold)
        worksheet.write(row, 5, global_sujeto, bold)
        worksheet.write(row, 6, global_retencion, bold)

        workbook.close()

        return filename

    def make_fsee_xls(self, filename):
        date = self.date_from.strftime('%m-%Y')
        filename = filename + ' ' + date + '.xls'

        row = 0

        workbook = xlsxwriter.Workbook(filename)
        bold = workbook.add_format({'bold': True})
        worksheet = workbook.add_worksheet()

        # Document header

        # add the company logo to the excel file if it exists
        company = self.env['res.company'].search([], limit=1)
        if company.logo:
            logo_stream = BytesIO(base64.b64decode(company.logo))
            worksheet.insert_image('K1', 'logo.png', {'image_data': logo_stream})

        worksheet.write(row, 0, 'LIBRO DE COMPRAS A SUJETO EXCLUIDO', bold)
        row += 2

        worksheet.write(row, 0, company.name, bold)
        row += 1

        worksheet.write(row, 0, 'ACTIVIDAD ECONOMICA:')
        worksheet.write(row, 1, company.descactividad)
        row += 1

        worksheet.write(row, 0, 'NIT:')
        worksheet.write(row, 1, company.nit)
        row += 1

        worksheet.write(row, 0, 'NRC:')
        worksheet.write(row, 1, company.nrc)
        row += 1

        worksheet.write(row, 0, 'MES:')
        worksheet.write(row, 1, MESES[int(self.date_from.strftime('%m')) - 1])
        row += 1

        worksheet.write(row, 0, 'AÑO:')
        worksheet.write(row, 1, self.date_from.strftime('%Y'))
        row += 1

        # Document body
        row += 2

        # Headers
        worksheet.write(row, 0, 'No.', bold)
        worksheet.write(row, 1, 'TIPO DE DOCUMENTO', bold)
        worksheet.write(row, 2, 'NUMERO DE DOCUMENTO', bold)
        worksheet.write(row, 3, 'NOMBRE, RAZON SOCIAL O DENOMINACION DEL SUJETO EXCLUIDO', bold)
        worksheet.write(row, 4, 'FECHA DE EMISION DEL DOCUMENTO', bold)
        worksheet.write(row, 5, 'NUMERO DE DOCUMENTO', bold)
        worksheet.write(row, 6, 'MONTO DE LA OPERACIÓN', bold)
        worksheet.write(row, 7, 'MONTO DE LA RETENCION 10%', bold)

        start_date = self.date_from
        end_date = self.date_to

        search_domain = [
            ('name', '=like', 'FSEE/%'),
            ('date', '>=', start_date.strftime('%Y-%m-%d')),
            ('date', '<=', end_date.strftime('%Y-%m-%d')),
        ]

        data_to_export = self.env['account.move'].search(search_domain, order='date asc')

        row += 1 # skip headers
        n = 1

        global_monto = 0
        global_retencion = 0

        for record in data_to_export:
            confirmation = self.get_confirmation(record)

            if(confirmation['status'] != 2):
                continue

            partner = record.partner_id
            num_doc = partner.numdocumento if partner.numdocumento else ''

            total_exenta = 0
            total_gravada = 0

            for line in record.line_ids:
                taxes = line.tax_ids
                exenta = True
                if taxes:
                    for tax in taxes:
                        if 'IVA 13%' in tax.name:
                            exenta = False

                if exenta:
                    total = line.price_subtotal
                    total_exenta = round(total_exenta + total, 2)
                else:
                    total = line.price_subtotal
                    total_gravada = round(total_gravada + total, 2)

            worksheet.write(row, 0, n)
            worksheet.write(row, 1, '2')
            worksheet.write(row, 2, num_doc)
            worksheet.write(row, 3, partner.name)
            worksheet.write(row, 4, record.date.strftime('%d/%m/%Y'))
            worksheet.write(row, 5, confirmation['uid_de'])
            worksheet.write(row, 6, abs(record.amount_untaxed))
            worksheet.write(row, 7, 0)

            global_monto += abs(record.amount_untaxed)
            global_retencion += 0

            n += 1
            row += 1

        # Document footer
        worksheet.write(row, 5, 'TOTALES', bold)
        worksheet.write(row, 6, global_monto, bold)
        worksheet.write(row, 7, global_retencion, bold)

        workbook.close()

        return filename

    ###################################
    # PDF DOCUMENT MAKERS
    ###################################
    def make_retencion_iva_pdf(self, filename):
        # Retencion IVA 1% - Casilla 162

        date = self.date_from.strftime('%m-%Y')
        filename = filename + ' ' + date + '.pdf'

        doc = SimpleDocTemplate(filename, pagesize=landscape(A4))
        styles = getSampleStyleSheet()
        elements = []

        # adding the company logo to the pdf document in the top right corner
        company = self.env['res.company'].search([], limit=1)
        if company.logo:
            elements.append(Image(io.BytesIO(base64.b64decode(company.logo)), LOGO_WITH, LOGO_HEIGHT))

        # Document header
        elements.append(Paragraph('RETENCION IVA 1%', styles['Title']))
        elements.append(Spacer(1, 12))

        elements.append(Paragraph(company.name, styles['Normal']))
        elements.append(Spacer(1, 6))

        elements.append(Paragraph('ACTIVIDAD ECONOMICA: ' + str(company.descactividad), styles['Normal']))
        elements.append(Spacer(1, 6))

        elements.append(Paragraph('NIT: ' + str(company.nit), styles['Normal']))
        elements.append(Spacer(1, 6))

        elements.append(Paragraph('NRC: ' + str(company.nrc), styles['Normal']))
        elements.append(Spacer(1, 6))

        elements.append(Paragraph('MES: ' + MESES[int(self.date_from.strftime('%m')) - 1], styles['Normal']))
        elements.append(Spacer(1, 6))

        elements.append(Paragraph('AÑO: ' + self.date_from.strftime('%Y'), styles['Normal']))
        elements.append(Spacer(1, 12))

        # Document body
        data = [
            ['NIT', 'FECHA', 'TIPO DOC', 'SERIE', 'NUM. DOC.', 'MONTO\nSUJETO', 'MONTO\nRETENCION 1%', 'DUI', 'ANEXO'],
        ]

        start_date = self.date_from
        end_date = self.date_to

        n = 1

        _logger.info('***** CompanyID: ' + str(self.env.company.id) + ' *****')

        search_domain = [
            ('name', 'like', 'CREP/%'),
            ('date', '>=', start_date.strftime('%Y-%m-%d')),
            ('date', '<=', end_date.strftime('%Y-%m-%d')),
            ('company_id', '=', self.env.company.id)
        ]

        data_to_export = self.env['account.move'].search(search_domain, order='date asc, id asc')

        global_sujeto = 0
        global_retencion = 0

        for record in data_to_export:
            
            _logger.info('***** RecordID: ' + str(record.id) + ' *****')
            _logger.info('***** RelatedDocument: ' + str(record.related_document_cre) + ' *****')

            confirmations = self.env['elinvoice.confirmation'].search([('uid_de', '=', record.related_document_cre)])

            confirmation = next(
                (c for c in confirmations if c.account_move_id.company_id.id == self.env.company.id),
                None
            )

            if confirmation:
                _logger.info('#### Si existe confirmacion ####')

                move = confirmation.account_move_id
                
                if move.is_processed and not move.is_anulated:
                    _logger.info('#### Si existe move ####')

                    partner = record.partner_id
                    num_doc = partner.numdocumento if partner.numdocumento else ''

                    nit = ""
                    dui = ""

                    if(partner.nit):
                        nit = partner.nit
                        dui = ""
                    else:
                        nit = ""
                        dui = num_doc

                    data.append([
                        nit,
                        record.date.strftime('%d/%m/%Y'),
                        '07',
                        record.purchase_doc_serie,
                        self.formatWithoutDashes(record.purchase_doc_number),
                        format(move.amount_untaxed, '.2f'),
                        format(record.amount_untaxed, '.2f'),
                        num_doc,
                        '7'
                    ])

                    global_sujeto += move.amount_untaxed
                    global_retencion += record.amount_untaxed

                    n += 1
            else:
                _logger.info('#### No existe confirmacion ####')

        # Document footer
        data.append([
            '',
            '',
            '',
            '',
            'TOTALES',
            format(global_sujeto, '.2f'),
            format(global_retencion, '.2f'),
            '',
            ''
        ])

        table = Table(data, colWidths=[60, 50, 40, 155, 155, 50, 50, 60, 40])
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (9, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (9, 0), colors.whitesmoke),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('FONTSIZE', (0, 0), (-1, -1), 6),
            
            # set max column size and wrap the words inside it so rows can be displayed correctly
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('INNERGRID', (0, 0), (-1, -1), 0.25, colors.black),
            ('BOX', (0, 0), (-1, -1), 0.25, colors.black),
        ]))

        elements.append(table)

        doc.leftMargin = DOC_MARGIN
        doc.rightMargin = DOC_MARGIN
        doc.topMargin = DOC_MARGIN
        doc.bottomMargin = DOC_MARGIN

        doc.build(elements)

        return filename

    def make_percepcion_iva_pdf(self, filename):
        # Percepcion IVA 1% - Casilla 163

        date = self.date_from.strftime('%m-%Y')
        filename = filename + ' ' + date + '.pdf'

        doc = SimpleDocTemplate(filename, pagesize=landscape(A4))
        styles = getSampleStyleSheet()
        elements = []

        # adding the company logo to the pdf document in the top right corner
        company = self.env['res.company'].search([], limit=1)
        if company.logo:
            elements.append(Image(io.BytesIO(base64.b64decode(company.logo)), LOGO_WITH, LOGO_HEIGHT))

        # Document header
        elements.append(Paragraph('PERCEPCION IVA 1%', styles['Title']))
        elements.append(Spacer(1, 12))

        elements.append(Paragraph(company.name, styles['Normal']))
        elements.append(Spacer(1, 6))

        elements.append(Paragraph('ACTIVIDAD ECONOMICA: ' + str(company.descactividad), styles['Normal']))
        elements.append(Spacer(1, 6))

        elements.append(Paragraph('NIT: ' + str(company.nit), styles['Normal']))
        elements.append(Spacer(1, 6))

        elements.append(Paragraph('NRC: ' + str(company.nrc), styles['Normal']))
        elements.append(Spacer(1, 6))

        elements.append(Paragraph('MES: ' + MESES[int(self.date_from.strftime('%m')) - 1], styles['Normal']))
        elements.append(Spacer(1, 6))

        elements.append(Paragraph('AÑO: ' + self.date_from.strftime('%Y'), styles['Normal']))
        elements.append(Spacer(1, 12))

        # Document body
        data = [
            ['NIT', 'FECHA', 'TIPO DOC', 'SERIE', 'NUM. DOC.', 'MONTO\nSUJETO', 'MONTO\nPERCEPCION 1%', 'DUI', 'ANEXO'],
        ]

        start_date = self.date_from
        end_date = self.date_to

        n = 1

        search_domain = [
            '&',
                '|', '|',
                    ('name', '=like', 'CCFEP/%'),
                    ('name', '=like', 'RNCEP/%'),
                    ('name', '=like', 'RNDEP/%'),
            '&',
                ('invoice_date', '>=', start_date.strftime('%Y-%m-%d')),
                ('invoice_date', '<=', end_date.strftime('%Y-%m-%d')),
        ]

        data_to_export = self.env['account.move'].search(search_domain, order='invoice_date asc, id asc')

        global_sujeto = 0
        global_percepcion = 0

        for record in data_to_export:

            partner = record.partner_id
            num_doc = partner.numdocumento if partner.numdocumento else ''

            monto_sujeto = 0
            monto_percepcion = 0

            for line in record.line_ids:

                taxes = line.tax_ids

                if taxes:

                    for tax in taxes:

                        if 'PER 1%' in tax.name:
                            monto_sujeto += line.price_subtotal
                            monto_percepcion += line.price_subtotal * 0.01

            nit = ""
            dui = ""

            if(partner.nit):
                nit = partner.nit
                dui = ""
            else:
                nit = ""
                dui = num_doc
            
            tipo_doc = '00'
            # Determinar el tipo de documento
            if record.journal_id.code == 'CCFEP':
                tipo_doc = '03'
            elif record.journal_id.code == 'NCEP':
                tipo_doc = '05'
            elif record.journal_id.code == 'NDEP':
                tipo_doc = '06'

            data.append([
                nit,
                record.date.strftime('%d/%m/%Y'),
                tipo_doc,
                record.sello_recepcion,
                self.formatWithoutDashes(record.codigo_generacion),
                format(monto_sujeto, '.2f'),
                format(monto_percepcion, '.2f'),
                dui,
                '8'
            ])

            n += 1

            global_sujeto += monto_sujeto
            global_percepcion += monto_percepcion

        # Document footer
        data.append([
            '',
            '',
            '',
            '',
            'TOTALES',
            format(global_sujeto, '.2f'),
            format(global_percepcion, '.2f'),
            '',
            ''
        ])

        table = Table(data, colWidths=[60, 50, 40, 155, 155, 50, 60, 60, 40])
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (9, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (9, 0), colors.whitesmoke),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('FONTSIZE', (0, 0), (-1, -1), 6),
            
            # set max column size and wrap the words inside it so rows can be displayed correctly
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('INNERGRID', (0, 0), (-1, -1), 0.25, colors.black),
            ('BOX', (0, 0), (-1, -1), 0.25, colors.black),
        ]))

        elements.append(table)

        doc.leftMargin = DOC_MARGIN
        doc.rightMargin = DOC_MARGIN
        doc.topMargin = DOC_MARGIN
        doc.bottomMargin = DOC_MARGIN

        doc.build(elements)

        return filename

    def make_anticipo_iva_pdf(self, filename):
        # Anticipo a cuenta IVA 2% - Casilla 161

        date = self.date_from.strftime('%m-%Y')
        filename = filename + ' ' + date + '.pdf'

        doc = SimpleDocTemplate(filename, pagesize=landscape(A4))
        styles = getSampleStyleSheet()
        elements = []

        # adding the company logo to the pdf document in the top right corner
        company = self.env['res.company'].search([], limit=1)
        if company.logo:
            elements.append(Image(io.BytesIO(base64.b64decode(company.logo)), LOGO_WITH, LOGO_HEIGHT))

        # Document header
        elements.append(Paragraph('ANTICIPO A CUENTA IVA 2%', styles['Title']))
        elements.append(Spacer(1, 12))

        elements.append(Paragraph(company.name, styles['Normal']))
        elements.append(Spacer(1, 6))

        elements.append(Paragraph('ACTIVIDAD ECONOMICA: ' + str(company.descactividad), styles['Normal']))
        elements.append(Spacer(1, 6))

        elements.append(Paragraph('NIT: ' + str(company.nit), styles['Normal']))
        elements.append(Spacer(1, 6))

        elements.append(Paragraph('NRC: ' + str(company.nrc), styles['Normal']))
        elements.append(Spacer(1, 6))

        elements.append(Paragraph('MES: ' + MESES[int(self.date_from.strftime('%m')) - 1], styles['Normal']))
        elements.append(Spacer(1, 6))

        elements.append(Paragraph('AÑO: ' + self.date_from.strftime('%Y'), styles['Normal']))
        elements.append(Spacer(1, 12))

        # Document body
        data = [
            ['NIT', 'FECHA', 'SERIE', 'NUM. DOC.', 'MONTO\nSUJETO', 'MONTO\nANTICIPO IVA 2%', 'DUI', 'ANEXO'],
        ]

        start_date = self.date_from
        end_date = self.date_to

        n = 1

        search_domain = [
            ('name', '=like', 'DCLEP/%'),
            ('invoice_date', '>=', start_date.strftime('%Y-%m-%d')),
            ('invoice_date', '<=', end_date.strftime('%Y-%m-%d')),
        ]

        data_to_export = self.env['account.move'].search(search_domain, order='invoice_date asc, id asc')

        global_sujeto = 0
        global_anticipo = 0

        for record in data_to_export:

            partner = record.partner_id
            num_doc = partner.numdocumento if partner.numdocumento else ''

            monto_anticipo = 0
            monto_sujeto = 0

            for line in record.line_ids:

                monto_anticipo += line.price_unit * line.quantity
            
            monto_sujeto = monto_anticipo / 0.02

            nit = ""
            dui = ""

            if(partner.nit):
                nit = partner.nit
                dui = ""
            else:
                nit = ""
                dui = num_doc

            data.append([
                nit,
                record.date.strftime('%d/%m/%Y'),
                record.sello_recepcion,
                self.formatWithoutDashes(record.codigo_generacion),
                format(monto_sujeto, '.2f'),
                format(monto_anticipo, '.2f'),
                dui,
                '6'
            ])

            n += 1

            global_sujeto += monto_sujeto
            global_anticipo += monto_anticipo

        # Document footer
        data.append([
            '',
            '',
            '',
            'TOTALES',
            format(global_sujeto, '.2f'),
            format(global_anticipo, '.2f'),
            '',
            ''
        ])

        table = Table(data, colWidths=[60, 50, 155, 155, 50, 60, 60, 40])
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (8, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (8, 0), colors.whitesmoke),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('FONTSIZE', (0, 0), (-1, -1), 6),
            
            # set max column size and wrap the words inside it so rows can be displayed correctly
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('INNERGRID', (0, 0), (-1, -1), 0.25, colors.black),
            ('BOX', (0, 0), (-1, -1), 0.25, colors.black),
        ]))

        elements.append(table)

        doc.leftMargin = DOC_MARGIN
        doc.rightMargin = DOC_MARGIN
        doc.topMargin = DOC_MARGIN
        doc.bottomMargin = DOC_MARGIN

        doc.build(elements)

        return filename
    
    def make_fex_pdf(self, filename):
        date = self.date_from.strftime('%m-%Y')
        filename = filename + ' ' + date + '.pdf'

        doc = SimpleDocTemplate(filename, pagesize=landscape(A4))
        styles = getSampleStyleSheet()
        elements = []

        # adding the company logo to the pdf document in the top right corner
        company = self.env.company
        if company.logo:
            elements.append(Image(io.BytesIO(base64.b64decode(company.logo)), LOGO_WITH, LOGO_HEIGHT))

        # Document header
        elements.append(Paragraph('DETALLE DE FACTURAS DE EXPORTACION', styles['Title']))
        elements.append(Spacer(1, 12))

        elements.append(Paragraph(company.name, styles['Normal']))
        elements.append(Spacer(1, 6))

        elements.append(Paragraph('ACTIVIDAD ECONOMICA: ' + str(company.descactividad), styles['Normal']))
        elements.append(Spacer(1, 6))

        elements.append(Paragraph('NIT: ' + str(company.nit), styles['Normal']))
        elements.append(Spacer(1, 6))

        elements.append(Paragraph('NRC: ' + str(company.nrc), styles['Normal']))
        elements.append(Spacer(1, 6))

        elements.append(Paragraph('MES: ' + MESES[int(self.date_from.strftime('%m')) - 1], styles['Normal']))
        elements.append(Spacer(1, 6))

        elements.append(Paragraph('AÑO: ' + self.date_from.strftime('%Y'), styles['Normal']))
        elements.append(Spacer(1, 12))

        # Document body
        data = []

        data.append([
            'No.',
            'FECHA',
            'NUMERO DE CONTROL \nINTERNO DEL',
            'NUMERO DE CONTROL \nINTERNO AL',
            'CODIGO DE \nGENERACION DEL',
            'CODIGO DE \nGENERACION DEL',
            'VENTAS \nEXENTAS',
            'VENTAS \nNO SUJ.',
            'VENTAS \nGRAVADAS',
            'VTAS. A \nTERCEROS',
            'EXPORT. DE \nSERVICIOS',
            'IVA \nRET.',
            'TOTAL',
        ])

        start_date = self.date_from
        end_date = self.date_to

        global_exenta = 0
        global_gravada = 0
        global_iva_retenido = 0
        global_total = 0
        global_iva = 0
        global_exportacion = 0

        n = 1

        # FEX
        # recorrer los dias del rango de fechas seleccionado
        for i in range((end_date - start_date).days + 1):

            # dia
            day = start_date + timedelta(days=i)

            # obtener todos los movimientos de FC de este dia
            search_domain = [
                '&',
                    ('name', '=like', 'FEXE/%'),
                    ('invoice_date', '=', day.strftime('%Y-%m-%d')),
            ]

            data_to_export = self.env['account.move'].search(search_domain, order='invoice_date asc, id asc')

            total_exenta = 0
            total_no_suj = 0
            total_gravada = 0
            total_terceros = 0
            total_exportacion = 0
            total_iva_retenido = 0
            total = 0
            
            resultados = []
            rango_inicio = None
            rango_fin = None
            rango_total = 0

            for record in data_to_export:

                confirmation = self.get_confirmation_for_tax_books(record)

                IsAnulated = False

                if confirmation['status'] == 0 and confirmation['uid_de'] == '-':
                    continue
                elif confirmation['status'] == 0 and confirmation['uid_de'] != '-':
                    IsAnulated = True

                for line in record.line_ids:
                    if line.product_id.default_code:
                        total_exportacion += line.price_subtotal if not IsAnulated else 0
                                        
                total = round((total_exportacion), 2) if not IsAnulated else 0

                # if IsAnulated:
                if False:

                    # Si había un rango abierto de procesados, lo cerramos primero
                    if rango_inicio:
                        resultados.append({
                            # "n": n,
                            "fecha": day.strftime('%d-%m-%Y'),
                            "numero_control_del": rango_inicio['control_number'],
                            "numero_control_al": rango_fin['control_number'],
                            "codigo_generacion_del": rango_inicio['generation_code'],
                            "codigo_generacion_al": rango_fin['generation_code'],
                            "exentas": total_exenta,
                            "no_sujetas": total_no_suj,
                            "gravadas": total_gravada,
                            "a_terceros": total_terceros,
                            "exportacion_servicios": total_exportacion,
                            "iva_retenido": total_iva_retenido,
                            "total": rango_total,
                        })
                        rango_inicio = rango_fin = None
                        rango_total = 0

                    # Agregamos el anulado solito
                    resultados.append({
                        # "n": n,
                        "fecha": day.strftime('%d-%m-%Y'),
                        "numero_control_del": confirmation['control_number'],
                        "numero_control_al": confirmation['control_number'],
                        "codigo_generacion_del": confirmation['uid_de'],
                        "codigo_generacion_al": confirmation['uid_de'],
                        "exentas": 0,
                        "no_sujetas": 0,
                        "gravadas": 0,
                        "a_terceros": 0,
                        "exportacion_servicios": 0,
                        "iva_retenido": 0,
                        "total": 0,
                    })

                    # reiniciamos acumuladores
                    total_exenta = 0
                    total_gravada = 0
                    total_iva_retenido = 0
                    total = 0
                else:
                    # Procesado: acumular
                    if not rango_inicio:
                        rango_inicio = {'record': record, 'generation_code': confirmation['uid_de'], 'control_number': confirmation['control_number']}
                    rango_fin = {'record': record, 'generation_code': confirmation['uid_de'], 'control_number': confirmation['control_number']}
                    rango_total = total

            # Al terminar el loop, guardar último bloque abierto
            if rango_inicio:
                resultados.append({
                    # "n": n,
                    "fecha": day.strftime('%d-%m-%Y'),
                    "numero_control_del": rango_inicio['control_number'],
                    "numero_control_al": rango_fin['control_number'],
                    "codigo_generacion_del": rango_inicio['generation_code'],
                    "codigo_generacion_al": rango_fin['generation_code'],
                    "exentas": total_exenta,
                    "no_sujetas": total_no_suj,
                    "gravadas": total_gravada,
                    "a_terceros": total_terceros,
                    "exportacion_servicios": total_exportacion,
                    "iva_retenido": total_iva_retenido,
                    "total": rango_total,
                })

            # Mostrar en logs
            for r in resultados:

                data.append([
                    n,
                    r['fecha'],
                    r['numero_control_del'],
                    r['numero_control_al'],
                    r['codigo_generacion_del'],
                    r['codigo_generacion_al'],
                    format(r['exentas'], '.2f'),
                    format(r['no_sujetas'], '.2f'),
                    format(r['gravadas'], '.2f'),
                    format(r['a_terceros'], '.2f'),
                    format(r['exportacion_servicios'], '.2f'),
                    format(r['iva_retenido'], '.2f'),
                    format(r['total'], '.2f'),
                ])

                global_exenta += r['exentas']
                global_gravada += r['gravadas']
                global_iva_retenido += r['iva_retenido']
                global_total += r['total']
                global_exportacion += r['exportacion_servicios']
        
                n += 1

        # Totales
        data.append([
            '',
            '',
            '',
            '',
            '',
            'TOTALES',
            format(global_exenta, '.2f'),
            format(0, '.2f'),
            format(global_gravada, '.2f'),
            format(0, '.2f'),
            format(global_exportacion, '.2f'),
            format(global_iva_retenido, '.2f'),
            format(global_total, '.2f'),
        ])

        table = Table(data, colWidths=[20, 40, 115, 115, 130, 130, 35, 35, 40, 40, 45, 30, 35])
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('FONTSIZE', (0, 0), (-1, -1), 6),
            # set max column size and wrap the words inside it so rows can be displayed correctly
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('INNERGRID', (0, 0), (-1, -1), 0.25, colors.black),
            ('BOX', (0, 0), (-1, -1), 0.25, colors.black),
        ]))

        elements.append(table)
        
        elements.append(Spacer(1, 20))
        
        # Document footer
        data2 = [
            ['Resumen de operaciones', 'Ventas propias', '', '', 'A cuenta de terceros', '', ''],
            ['', 'Valor neto', 'IVA Debito', 'Valor total', 'Valor neto', 'IVA debito', 'Valor total'],
        ]

        data2.append([
            'Exportaciones segun FEX',
            format(global_exportacion, '.2f'),
            format(0, '.2f'),
            format(global_exportacion, '.2f'),
            format(0, '.2f'),
            format(0, '.2f'),
            format(0, '.2f'),
        ])

        table2 = Table(data2)

        table2.setStyle(TableStyle([
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('FONTSIZE', (0, 0), (-1, -1), 6),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),

            ('SPAN', (0, 0), (0, 1)),  # Resumen
            ('SPAN', (1, 0), (3, 0)),  # Ventas propias
            ('SPAN', (4, 0), (6, 0)),  # A cuenta de terceros
        ]))

        elements.append(table2)

        elements.append(Spacer(1, 20))

        # Document footer
        data1 = []
        data1.append([
            'F. _______________________________',
        ])

        table1 = Table(data1)

        table1.setStyle(TableStyle([
            ('FONTSIZE', (0, 0), (-1, -1), 10),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ]))

        elements.append(table1)

        doc.leftMargin = DOC_MARGIN
        doc.rightMargin = DOC_MARGIN
        doc.topMargin = DOC_MARGIN
        doc.bottomMargin = DOC_MARGIN

        doc.build(elements)

        return filename

    def make_fe_pdf(self, filename):
        date = self.date_from.strftime('%m-%Y')
        filename = filename + ' ' + date + '.pdf'

        doc = SimpleDocTemplate(filename, pagesize=landscape(A4))
        styles = getSampleStyleSheet()
        elements = []

        # adding the company logo to the pdf document in the top right corner
        company = self.env.company
        if company.logo:
            elements.append(Image(io.BytesIO(base64.b64decode(company.logo)), LOGO_WITH, LOGO_HEIGHT))

        # Document header
        elements.append(Paragraph('LIBRO DE VENTAS CONSUMIDOR FINAL', styles['Title']))
        elements.append(Spacer(1, 12))

        elements.append(Paragraph(company.name, styles['Normal']))
        elements.append(Spacer(1, 6))

        elements.append(Paragraph('ACTIVIDAD ECONOMICA: ' + str(company.descactividad), styles['Normal']))
        elements.append(Spacer(1, 6))

        elements.append(Paragraph('NIT: ' + str(company.nit), styles['Normal']))
        elements.append(Spacer(1, 6))

        elements.append(Paragraph('NRC: ' + str(company.nrc), styles['Normal']))
        elements.append(Spacer(1, 6))

        elements.append(Paragraph('MES: ' + MESES[int(self.date_from.strftime('%m')) - 1], styles['Normal']))
        elements.append(Spacer(1, 6))

        elements.append(Paragraph('AÑO: ' + self.date_from.strftime('%Y'), styles['Normal']))
        elements.append(Spacer(1, 12))

        # Document body
        data = []

        data.append([
            'No.',
            'FECHA',
            'NUMERO DE CONTROL \nINTERNO DEL',
            'NUMERO DE CONTROL \nINTERNO AL',
            'CODIGO DE \nGENERACION DEL',
            'CODIGO DE \nGENERACION DEL',
            'VENTAS \nEXENTAS',
            'VENTAS \nNO SUJ.',
            'VENTAS \nGRAVADAS',
            'VTAS. A \nTERCEROS',
            'EXPORT. DE \nSERVICIOS',
            'IVA \nRET.',
            'TOTAL',
        ])

        start_date = self.date_from
        end_date = self.date_to

        global_exenta = 0
        global_gravada = 0
        global_iva_retenido = 0
        global_total = 0
        global_iva = 0
        global_exportacion = 0

        n = 1
        
        # recorrer los dias del rango de fechas seleccionado
        for i in range((end_date - start_date).days + 1):

            # dia
            day = start_date + timedelta(days=i)

            # obtener todos los movimientos de FC de este dia
            search_domain = [
                '&',
                    ('name', '=like', 'FE/%'),
                    ('invoice_date', '=', day.strftime('%Y-%m-%d')),
            ]

            data_to_export = self.env['account.move'].search(search_domain, order='invoice_date asc, id asc')

            total_exenta = 0
            total_no_suj = 0
            total_gravada = 0
            total_terceros = 0
            total_exportacion = 0
            total_iva_retenido = 0
            total = 0
            
            resultados = []
            rango_inicio = None
            rango_fin = None
            rango_total = 0

            for record in data_to_export:

                confirmation = self.get_confirmation_for_tax_books(record)

                IsAnulated = False

                if confirmation['status'] == 0 and confirmation['uid_de'] == '-':
                    continue
                elif confirmation['status'] == 0 and confirmation['uid_de'] != '-':
                    IsAnulated = True

                for line in record.line_ids:
                    if line.product_id.default_code:

                        taxes = line.tax_ids
                        exenta = True
                        ivaItem = 0

                        if taxes:
                            for tax in taxes:
                                if 'IVA 13%' in tax.name:
                                    exenta = False
                                    ivaItem = line.price_subtotal * 0.13 if not IsAnulated else 0
                                    global_iva += ivaItem
                                
                                if 'RET 1%' in tax.name:
                                    total_iva_retenido += line.price_subtotal * 0.01 if not IsAnulated else 0

                        if exenta:
                            total_exenta += line.price_subtotal if not IsAnulated else 0
                        else:
                            total_gravada += line.price_subtotal + ivaItem if not IsAnulated else 0
                
                total = round((total_exenta + total_gravada), 2) if not IsAnulated else 0

                # if IsAnulated:
                if False:                    

                    # Si había un rango abierto de procesados, lo cerramos primero
                    if rango_inicio:
                        resultados.append({
                            # "n": n,
                            "fecha": day.strftime('%d-%m-%Y'),
                            "numero_control_del": rango_inicio['control_number'],
                            "numero_control_al": rango_fin['control_number'],
                            "codigo_generacion_del": rango_inicio['generation_code'],
                            "codigo_generacion_al": rango_fin['generation_code'],
                            "exentas": total_exenta,
                            "no_sujetas": total_no_suj,
                            "gravadas": total_gravada,
                            "a_terceros": total_terceros,
                            "exportacion_servicios": total_exportacion,
                            "iva_retenido": total_iva_retenido,
                            "total": rango_total,
                        })
                        rango_inicio = rango_fin = None
                        rango_total = 0

                    # Agregamos el anulado solito
                    resultados.append({
                        # "n": n,
                        "fecha": day.strftime('%d-%m-%Y'),
                        "numero_control_del": confirmation['control_number'],
                        "numero_control_al": confirmation['control_number'],
                        "codigo_generacion_del": confirmation['uid_de'],
                        "codigo_generacion_al": confirmation['uid_de'],
                        "exentas": 0,
                        "no_sujetas": 0,
                        "gravadas": 0,
                        "a_terceros": 0,
                        "exportacion_servicios": 0,
                        "iva_retenido": 0,
                        "total": 0,
                    })

                    # reiniciamos acumuladores
                    total_exenta = 0
                    total_gravada = 0
                    total_iva_retenido = 0
                    total = 0
                else:
                    # Procesado: acumular
                    if not rango_inicio:
                        rango_inicio = {'record': record, 'generation_code': confirmation['uid_de'], 'control_number': confirmation['control_number']}
                    rango_fin = {'record': record, 'generation_code': confirmation['uid_de'], 'control_number': confirmation['control_number']}
                    rango_total = total

            # Al terminar el loop, guardar último bloque abierto
            if rango_inicio:
                resultados.append({
                    # "n": n,
                    "fecha": day.strftime('%d-%m-%Y'),
                    "numero_control_del": rango_inicio['control_number'],
                    "numero_control_al": rango_fin['control_number'],
                    "codigo_generacion_del": rango_inicio['generation_code'],
                    "codigo_generacion_al": rango_fin['generation_code'],
                    "exentas": total_exenta,
                    "no_sujetas": total_no_suj,
                    "gravadas": total_gravada,
                    "a_terceros": total_terceros,
                    "exportacion_servicios": total_exportacion,
                    "iva_retenido": total_iva_retenido,
                    "total": rango_total,
                })

            # Mostrar en logs
            for r in resultados:

                data.append([
                    n,
                    r['fecha'],
                    r['numero_control_del'],
                    r['numero_control_al'],
                    r['codigo_generacion_del'],
                    r['codigo_generacion_al'],
                    format(r['exentas'], '.2f'),
                    format(r['no_sujetas'], '.2f'),
                    format(r['gravadas'], '.2f'),
                    format(r['a_terceros'], '.2f'),
                    format(r['exportacion_servicios'], '.2f'),
                    format(r['iva_retenido'], '.2f'),
                    format(r['total'], '.2f'),
                ])

                global_exenta += r['exentas']
                global_gravada += r['gravadas']
                global_iva_retenido += r['iva_retenido']
                global_total += r['total']
        
                n += 1

        # Totales
        data.append([
            '',
            '',
            '',
            '',
            '',
            'TOTALES',
            format(global_exenta, '.2f'),
            format(0, '.2f'),
            format(global_gravada, '.2f'),
            format(0, '.2f'),
            format(global_exportacion, '.2f'),
            format(global_iva_retenido, '.2f'),
            format(global_total, '.2f'),
        ])

        table = Table(data, colWidths=[20, 40, 115, 115, 130, 130, 35, 35, 40, 40, 45, 30, 35])
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('FONTSIZE', (0, 0), (-1, -1), 6),
            # set max column size and wrap the words inside it so rows can be displayed correctly
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('INNERGRID', (0, 0), (-1, -1), 0.25, colors.black),
            ('BOX', (0, 0), (-1, -1), 0.25, colors.black),
        ]))

        elements.append(table)
        
        elements.append(Spacer(1, 20))
        
        # Document footer
        data2 = [
            ['Resumen de operaciones', 'Ventas propias', '', '', 'A cuenta de terceros', '', ''],
            ['', 'Valor neto', 'IVA Debito', 'Valor total', 'Valor neto', 'IVA debito', 'Valor total'],
        ]
        
        data2.append([
            'Ventas netas gravadas a consumidor',
            format(global_gravada - global_iva, '.2f'),
            format(global_iva, '.2f'),
            format(global_gravada, '.2f'),
            format(0, '.2f'),
            format(0, '.2f'),
            format(0, '.2f'),
        ])

        data2.append([
            'Ventas netas exentas a consumidor',
            format(global_exenta, '.2f'),
            format(0, '.2f'),
            format(global_exenta, '.2f'),
            format(0, '.2f'),
            format(0, '.2f'),
            format(0, '.2f'),
        ])

        table2 = Table(data2)

        table2.setStyle(TableStyle([
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('FONTSIZE', (0, 0), (-1, -1), 6),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),

            ('SPAN', (0, 0), (0, 1)),  # Resumen
            ('SPAN', (1, 0), (3, 0)),  # Ventas propias
            ('SPAN', (4, 0), (6, 0)),  # A cuenta de terceros
        ]))

        elements.append(table2)

        elements.append(Spacer(1, 20))

        # Document footer
        data1 = []
        data1.append([
            'F. _______________________________',
        ])

        table1 = Table(data1)

        table1.setStyle(TableStyle([
            ('FONTSIZE', (0, 0), (-1, -1), 10),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ]))

        elements.append(table1)

        doc.leftMargin = DOC_MARGIN
        doc.rightMargin = DOC_MARGIN
        doc.topMargin = DOC_MARGIN
        doc.bottomMargin = DOC_MARGIN

        doc.build(elements)

        return filename

    def make_ccfe_pdf(self, filename):
        date = self.date_from.strftime('%m-%Y')
        filename = filename + ' ' + date + '.pdf'

        doc = SimpleDocTemplate(filename, pagesize=landscape(A4))
        styles = getSampleStyleSheet()
        elements = []

        # adding the company logo to the pdf document in the top right corner
        #companyy = self.env['res.company'].search([], limit=1)
        company = self.env.company
        if company.logo:
            elements.append(Image(io.BytesIO(base64.b64decode(company.logo)), LOGO_WITH, LOGO_HEIGHT))

        # Document header
        elements.append(Paragraph('LIBRO DE VENTAS A CONTRIBUYENTES', styles['Title']))
        elements.append(Spacer(1, 12))

        elements.append(Paragraph(company.name, styles['Normal']))
        elements.append(Spacer(1, 6))

        elements.append(Paragraph('ACTIVIDAD ECONOMICA: ' + str(company.descactividad), styles['Normal']))
        elements.append(Spacer(1, 6))

        elements.append(Paragraph('NIT: ' + str(company.nit), styles['Normal']))
        elements.append(Spacer(1, 6))

        elements.append(Paragraph('NRC: ' + str(company.nrc), styles['Normal']))
        elements.append(Spacer(1, 6))

        elements.append(Paragraph('MES: ' + MESES[int(self.date_from.strftime('%m')) - 1], styles['Normal']))
        elements.append(Spacer(1, 6))

        elements.append(Paragraph('AÑO: ' + self.date_from.strftime('%Y'), styles['Normal']))
        elements.append(Spacer(1, 12))

        # Document body
        data = [
            ['No.', 'FECHA', 'NUMERO DE DOCUMENTO', 'NRC', 'NIT', 'NOMBRE CLIENTE', 'VENTAS PROPIAS', '', '', '', 'VENTAS A TERCEROS', '', 'IVA \nRETENIDO', 'TOTAL'],
            ['', '', '', '', '', '', 'EXENT.', 'NO SUJ.', 'GRAV', 'DEBITO\nFISCAL', 'GRAV.', 'IVA', '', '']
        ]

        start_date = self.date_from
        end_date = self.date_to

        n = 1
        global_exenta = 0
        global_gravada = 0
        grand_total = 0
        global_iva = 0
        global_iva_retenido = 0
        global_total = 0

        search_domain = [
            ('invoice_date', '>=', start_date.strftime('%Y-%m-%d')),
            ('invoice_date', '<=', end_date.strftime('%Y-%m-%d')),
            '|', '|',
            ('name', 'ilike', 'CCFE/'),
            ('name', 'ilike', 'RNCE/'),
            ('name', 'ilike', 'RNDE/'),
        ]

        data_to_export = self.env['account.move'].search(search_domain, order='invoice_date asc, id asc')
        
        numeroDocumento = 0

        for record in data_to_export:

            confirmation = self.get_confirmation_for_tax_books(record)

            isAnulated = False

            if(confirmation['status'] == 0 and confirmation['uid_de'] == '-'):
                continue
            elif(confirmation['status'] == 0 and confirmation['uid_de'] != '-'):
                isAnulated = True

            numeroDocumento = str(confirmation['uid_de'])
            
            partner = record.partner_id
            num_doc = partner.numdocumento if partner.numdocumento else ''

            total_exenta = 0
            total_gravada = 0
            total_gravada_terceros = 0
            total_iva_tercerros = 0
            total_iva_retenido = 0
            total_no_suj = 0

            for line in record.line_ids:
                taxes = line.tax_ids
                exenta = True
                if taxes:
                    for tax in taxes:
                        if 'IVA 13%' in tax.name:
                            exenta = False

                        if 'RET 1%' in tax.name:
                            total_iva_retenido += round((line.price_subtotal * 0.01), 2)

                if exenta:
                    total = line.price_subtotal
                    total_exenta = round(total_exenta + total, 2)
                else:
                    total = line.price_subtotal
                    total_gravada = round(total_gravada + total, 2)

            total_iva = round(total_gravada * 0.13, 2)
            gran_total = round(total_gravada + total_exenta + total_iva, 2)

            nit = ""
            dui = ""

            if(partner.nit):
                nit = partner.nit
                dui = ""
            else:
                nit = ""
                dui = num_doc
            
            if record.name.startswith('RNCE/') and not isAnulated:
                total_exenta = total_exenta * -1 if total_exenta > 0 else 0
                total_no_suj = total_no_suj * -1 if total_no_suj > 0 else 0
                total_gravada = total_gravada * -1 if total_gravada > 0 else 0
                total_iva = total_iva * -1 if total_iva > 0 else 0
                total_iva_retenido = total_iva_retenido * -1 if total_iva_retenido > 0 else 0
                gran_total = gran_total * -1 if gran_total > 0 else 0

            if isAnulated:
                total_exenta = 0
                total_no_suj = 0
                total_gravada = 0
                total_iva = 0
                total_iva_retenido = 0
                gran_total = 0

            data.append([
                n,
                record.invoice_date.strftime('%d/%m/%Y'),
                numeroDocumento,
                partner.nrc if partner.nrc else '',
                nit,
                partner.name[:40] if partner and not isAnulated else 'INVÁLIDO',
                format(total_exenta, '.2f'),
                format(total_no_suj, '.2f'),
                format(total_gravada, '.2f'),
                format(total_iva, '.2f'),
                format(total_gravada_terceros, '.2f'),
                format(total_iva_tercerros, '.2f'),
                format(total_iva_retenido, '.2f'),
                format(gran_total, '.2f'),
            ])

            global_total += gran_total
            global_exenta += total_exenta
            global_gravada += total_gravada
            grand_total += gran_total
            global_iva += total_iva
            global_iva_retenido += total_iva_retenido

            n += 1

        # Document footer
        data.append([
            '',
            '',
            '',
            '',
            '',
            'TOTALES',
            format(global_exenta, '.2f'),
            format(0, '.2f'),
            format(global_gravada, '.2f'),
            format(global_iva, '.2f'),
            0,
            0,
            format(global_iva_retenido, '.2f'),
            format(global_total, '.2f'),
        ])

        table = Table(data, colWidths=[20, 45, 150, 60, 60, 160, 35, 35, 35, 35, 38, 38, 40, 35])
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (13, 1), colors.grey),
            ('TEXTCOLOR', (0, 0), (13, 1), colors.whitesmoke),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('FONTSIZE', (0, 0), (-1, -1), 6),
            
            # set max column size and wrap the words inside it so rows can be displayed correctly
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('INNERGRID', (0, 0), (-1, -1), 0.25, colors.black),
            ('BOX', (0, 0), (-1, -1), 0.25, colors.black),

            ('ALIGN', (5, 2), (5, -1), 'LEFT'),

            ('SPAN', (0, 0), (0, 1)),  # N°
            ('SPAN', (1, 0), (1, 1)),  # Fecha
            ('SPAN', (2, 0), (2, 1)),  # Numero de Documento
            ('SPAN', (3, 0), (3, 1)),  # NRC
            ('SPAN', (4, 0), (4, 1)),  # NIT
            ('SPAN', (5, 0), (5, 1)),  # Nombre Cliente
            ('SPAN', (6, 0), (9, 0)),  # Ventas Propias
            ('SPAN', (10, 0), (11, 0)),  # Ventas a Terceros
            ('SPAN', (12, 0), (12, 1)),  # IVA Retenido
            ('SPAN', (13, 0), (13, 1)),  # Total

        ]))

        elements.append(table)
        
        elements.append(Spacer(1, 20))
        
        # Document footer
        data2 = [
            ['Resumen de operaciones', 'Ventas propias', '', '', 'A cuenta de terceros', '', ''],
            ['', 'Valor neto', 'IVA Debito', 'Valor total', 'Valor neto', 'IVA debito', 'Valor total'],
        ]
        
        data2.append([
            'Ventas netas gravadas a contribuyentes',
            format(global_gravada, '.2f'),
            format(global_iva, '.2f'),
            format(global_gravada + global_iva, '.2f'),
            format(0, '.2f'),
            format(0, '.2f'),
            format(0, '.2f'),
        ])

        data2.append([
            'Ventas netas exentas a contribuyentes',
            format(global_exenta, '.2f'),
            format(0, '.2f'),
            format(global_exenta, '.2f'),
            format(0, '.2f'),
            format(0, '.2f'),
            format(0, '.2f'),
        ])

        table2 = Table(data2)

        table2.setStyle(TableStyle([
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('FONTSIZE', (0, 0), (-1, -1), 6),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),

            ('SPAN', (0, 0), (0, 1)),  # Resumen
            ('SPAN', (1, 0), (3, 0)),  # Ventas propias
            ('SPAN', (4, 0), (6, 0)),  # A cuenta de terceros
        ]))

        elements.append(table2)

        elements.append(Spacer(1, 20))

        # Document footer
        data1 = []
        data1.append([
            'F. _______________________________',
        ])

        table1 = Table(data1)

        table1.setStyle(TableStyle([
            ('FONTSIZE', (0, 0), (-1, -1), 10),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ]))

        elements.append(table1)

        doc.leftMargin = DOC_MARGIN
        doc.rightMargin = DOC_MARGIN
        doc.topMargin = DOC_MARGIN
        doc.bottomMargin = DOC_MARGIN

        doc.build(elements)

        return filename

    def make_ccfep_pdf(self, filename):
        date = self.date_from.strftime('%m-%Y')
        filename = filename + ' ' + date + '.pdf'

        doc = SimpleDocTemplate(filename, pagesize=landscape(A4))
        styles = getSampleStyleSheet()
        elements = []

        # adding the company logo to the pdf document in the top right corner
        company = self.env['res.company'].search([], limit=1)
        if company.logo:
            elements.append(Image(io.BytesIO(base64.b64decode(company.logo)), LOGO_WITH, LOGO_HEIGHT))

        # Document header
        elements.append(Paragraph('LIBRO DE COMPRAS', styles['Title']))
        elements.append(Spacer(1, 12))

        elements.append(Paragraph(company.name, styles['Normal']))
        elements.append(Spacer(1, 6))

        elements.append(Paragraph('ACTIVIDAD ECONOMICA: ' + str(company.descactividad), styles['Normal']))
        elements.append(Spacer(1, 6))

        elements.append(Paragraph('NIT: ' + str(company.nit), styles['Normal']))
        elements.append(Spacer(1, 6))

        elements.append(Paragraph('NRC: ' + str(company.nrc), styles['Normal']))
        elements.append(Spacer(1, 6))

        elements.append(Paragraph('MES: ' + MESES[int(self.date_from.strftime('%m')) - 1], styles['Normal']))
        elements.append(Spacer(1, 6))

        elements.append(Paragraph('AÑO: ' + self.date_from.strftime('%Y'), styles['Normal']))
        elements.append(Spacer(1, 12))

        # Document body
        data = []

        data.append([
            'No.',
            'FECHA',
            'NUMERO DE DOCUMENTO',
            'NIT O NRC PROVEEDOR',
            'NOMBRE PROVEEDOR',
            'COMPRAS \n INTERNAS \n EXENTAS',
            'INTERNACIONES \n EXENTAS Y/O \n NO SUJETAS',
            'IMPORTACIONES \n EXENTAS Y/O \n NO SUJETAS',
            'COMPRAS \n INTERNAS \n GRAVADAS',
            'IMPORTACIONES \n GRAVADAS \n DE SERVICIOS',
            'CREDITO \n FISCAL',
            'PERCEPCION',
            'TOTAL',
        ])

        start_date = self.date_from
        end_date = self.date_to

        search_domain = [
            '|',  # OR lógico
            ('name', 'like', 'CCFEP%'),
            ('name', 'like', 'RNCEP%'),
            ('date', '>=', start_date.strftime('%Y-%m-%d')),
            ('date', '<=', end_date.strftime('%Y-%m-%d')),
        ]

        data_to_export = self.env['account.move'].search(search_domain, order='date asc')

        n = 1
        global_exenta = 0
        global_gravada = 0
        gran_total = 0
        global_iva = 0
        global_importaciones = 0
        global_percepcion = 0
        global_total = 0
        
        for record in data_to_export:
            partner = record.partner_id
            num_doc = partner.numdocumento if partner.numdocumento else ''

            is_internal = partner.is_internal if partner.is_internal else 'SI'

            total_exenta = 0
            total_gravada = 0
            total_percepcion = 0

            for line in record.line_ids:
                taxes = line.tax_ids
                exenta = True
                if taxes:
                    for tax in taxes:
                        if 'IVA 13%' in tax.name:
                            exenta = False

                        if 'PER 1%' in tax.name:
                            total_percepcion = round((total_percepcion + round(line.price_subtotal * 0.01, 4)), 4) # Obtener la percepcion del 1%

                if exenta:
                    total = line.price_subtotal
                    total_exenta = round(total_exenta + total, 2)
                else:
                    total = line.price_subtotal
                    total_gravada = round(total_gravada + total, 2)

            total_iva = round(total_gravada * 0.13, 2)
            gran_total = round(total_gravada + total_exenta + total_iva + total_percepcion, 2)

            nit = ""
            dui = ""

            if(partner.nit):
                nit = partner.nit
                dui = ""
            else:
                nit = ""
                dui = num_doc

            num_tax_doc = record.codigo_generacion if record.codigo_generacion else ''

            if(is_internal == 'SI'):
                data.append([
                    n,
                    record.date.strftime('%d/%m/%Y'),
                    num_tax_doc,
                    nit,
                    partner.name,
                    format(total_exenta, '.2f'),
                    format(0, '.2f'),
                    format(0, '.2f'),
                    format(total_gravada, '.2f'),
                    format(0, '.2f'),
                    format(total_iva, '.2f'),
                    format(total_percepcion, '.2f'),
                    format(gran_total, '.2f'),
                ])
                global_gravada += total_gravada
            else:
                data.append([
                    n,
                    record.date.strftime('%d/%m/%Y'),
                    num_tax_doc,
                    nit,
                    partner.name,
                    format(total_exenta, '.2f'),
                    format(0, '.2f'),
                    format(0, '.2f'),
                    format(0, '.2f'),
                    format(total_gravada, '.2f'),
                    format(total_iva, '.2f'),
                    format(total_percepcion, '.2f'),
                    format(gran_total, '.2f'),
                ])
                global_importaciones += total_gravada

            global_exenta += total_exenta
            # gran_total += gran_total
            global_iva += total_iva
            global_percepcion += total_percepcion
            global_total += gran_total

            n += 1

        # Document footer
        data.append([
            '',
            '',
            '',
            '',
            'TOTALES',
            format(global_exenta, '.2f'),
            format(0, '.2f'),
            format(0, '.2f'),
            format(global_gravada, '.2f'),
            format(global_importaciones, '.2f'),
            format(global_iva, '.2f'),
            format(global_percepcion, '.2f'),
            format(global_total, '.2f'),
        ])

        table = Table(data)

        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('FONTSIZE', (0, 0), (-1, -1), 5),
            # set max column size and wrap the words inside it so rows can be displayed correctly
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('INNERGRID', (0, 0), (-1, -1), 0.25, colors.black),
            ('BOX', (0, 0), (-1, -1), 0.25, colors.black),
        ]))

        elements.append(table)

        doc.leftMargin = DOC_MARGIN
        doc.rightMargin = DOC_MARGIN
        doc.topMargin = DOC_MARGIN
        doc.bottomMargin = DOC_MARGIN

        doc.build(elements)

        return filename

    def make_cre_pdf(self, filename):
        date = self.date_from.strftime('%m-%Y')
        filename = filename + ' ' + date + '.pdf'

        doc = SimpleDocTemplate(filename, pagesize=landscape(A4))
        styles = getSampleStyleSheet()
        elements = []

        # adding the company logo to the pdf document in the top right corner
        company = self.env['res.company'].search([], limit=1)
        if company.logo:
            elements.append(Image(io.BytesIO(base64.b64decode(company.logo)), LOGO_WITH, LOGO_HEIGHT))

        # Document header
        elements.append(Paragraph('LIBRO DE RETENCIONES DE IVA 1%', styles['Title']))
        elements.append(Spacer(1, 12))

        elements.append(Paragraph(company.name, styles['Normal']))
        elements.append(Spacer(1, 6))

        elements.append(Paragraph('ACTIVIDAD ECONOMICA: ' + str(company.descactividad), styles['Normal']))
        elements.append(Spacer(1, 6))

        elements.append(Paragraph('NIT: ' + str(company.nit), styles['Normal']))
        elements.append(Spacer(1, 6))

        elements.append(Paragraph('NRC: ' + str(company.nrc), styles['Normal']))
        elements.append(Spacer(1, 6))

        elements.append(Paragraph('MES: ' + MESES[int(self.date_from.strftime('%m')) - 1], styles['Normal']))
        elements.append(Spacer(1, 6))

        elements.append(Paragraph('AÑO: ' + self.date_from.strftime('%Y'), styles['Normal']))
        elements.append(Spacer(1, 12))

        # Document body
        data = []

        data.append([
            'No.',
            'NIT DEL SUJETO',
            'FECHA',
            'TIPO DE DOCUMENTO',
            'NUMERO DE DOCUMENTO',
            'MONTO SUJETO',
            'MONTO DE LA \n RETENCION 1%',
            'DUI SUJETO',
        ])

        start_date = self.date_from
        end_date = self.date_to

        search_domain = [
            ('name', '=like', 'CCFEP%'),
            ('date', '>=', start_date.strftime('%Y-%m-%d')),
            ('date', '<=', end_date.strftime('%Y-%m-%d')),
        ]

        data_to_export = self.env['account.move'].search(search_domain, order='date asc')

        n = 1
        global_sujeto = 0
        global_retencion = 0

        for record in data_to_export:
            retencion = False
            monto_sujeto = 0

            for line in record.line_ids:
                taxes = line.tax_ids
                if taxes:
                    for tax in taxes:
                        if tax.name == 'RET 1%':
                            retencion = True
                            monto_sujeto = monto_sujeto + line.price_subtotal

            if not retencion:
                continue

            partner = record.partner_id
            num_doc = partner.numdocumento if partner.numdocumento else ''

            nit = ""
            dui = ""

            if(partner.nit):
                nit = partner.nit
                dui = ""
            else:
                nit = ""
                dui = num_doc

            data.append([
                n,
                nit,
                record.date.strftime('%d/%m/%Y'),
                '07',
                record.purchase_doc_number,
                format(monto_sujeto, '.2f'),
                format(monto_sujeto * 0.01, '.2f'),
                num_doc,
            ])

            global_sujeto += monto_sujeto
            global_retencion += monto_sujeto * 0.01

            n += 1

        # Document footer
        data.append([
            '',
            '',
            '',
            '',
            'TOTALES',
            format(global_sujeto, '.2f'),
            format(global_retencion, '.2f'),
            '',
        ])

        table = Table(data)

        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('FONTSIZE', (0, 0), (-1, -1), 6),
            # set max column size and wrap the words inside it so rows can be displayed correctly
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('INNERGRID', (0, 0), (-1, -1), 0.25, colors.black),
            ('BOX', (0, 0), (-1, -1), 0.25, colors.black),
        ]))

        elements.append(table)

        doc.leftMargin = DOC_MARGIN
        doc.rightMargin = DOC_MARGIN
        doc.topMargin = DOC_MARGIN
        doc.bottomMargin = DOC_MARGIN

        doc.build(elements)

        return filename

    def make_fsee_pdf(self, filename):
        date = self.date_from.strftime('%m-%Y')
        filename = filename + ' ' + date + '.pdf'

        doc = SimpleDocTemplate(filename, pagesize=landscape(A4))
        styles = getSampleStyleSheet()
        elements = []

        # adding the company logo to the pdf document in the top right corner
        company = self.env['res.company'].search([], limit=1)
        if company.logo:
            elements.append(Image(io.BytesIO(base64.b64decode(company.logo)), LOGO_WITH, LOGO_HEIGHT))

        # Document header
        elements.append(Paragraph('LIBRO DE COMPRAS A SUJETO EXCLUIDO', styles['Title']))
        elements.append(Spacer(1, 12))

        elements.append(Paragraph(company.name, styles['Normal']))
        elements.append(Spacer(1, 6))

        elements.append(Paragraph('ACTIVIDAD ECONOMICA: ' + str(company.descactividad), styles['Normal']))
        elements.append(Spacer(1, 6))

        elements.append(Paragraph('NIT: ' + str(company.nit), styles['Normal']))
        elements.append(Spacer(1, 6))

        elements.append(Paragraph('NRC: ' + str(company.nrc), styles['Normal']))
        elements.append(Spacer(1, 6))

        elements.append(Paragraph('MES: ' + MESES[int(self.date_from.strftime('%m')) - 1], styles['Normal']))
        elements.append(Spacer(1, 6))

        elements.append(Paragraph('AÑO: ' + self.date_from.strftime('%Y'), styles['Normal']))
        elements.append(Spacer(1, 12))

        # Document body
        data = []

        data.append([
            'No.',
            'TIPO DE DOCUMENTO',
            'NUMERO DE DOCUMENTO',
            'NOMBRE, RAZON SOCIAL O \n DENOMINACION DEL SUJETO EXCLUIDO',
            'FECHA DE EMISION DEL DOCUMENTO',
            'NUMERO DE DOCUMENTO',
            'MONTO DE LA \n OPERACIÓN',
            'MONTO DE LA \n RETENCION 10%',
            'DUI SUJETO',
        ])

        start_date = self.date_from
        end_date = self.date_to

        search_domain = [
            ('name', '=like', 'FSEE/%'),
            ('date', '>=', start_date.strftime('%Y-%m-%d')),
            ('date', '<=', end_date.strftime('%Y-%m-%d')),
        ]

        data_to_export = self.env['account.move'].search(search_domain, order='date asc')

        n = 1
        global_monto = 0
        global_retencion = 0

        for record in data_to_export:
            confirmation = self.get_confirmation(record)

            if(confirmation['status'] != 2):
                continue

            partner = record.partner_id
            num_doc = partner.numdocumento if partner.numdocumento else ''

            total_exenta = 0
            total_gravada = 0

            for line in record.line_ids:
                taxes = line.tax_ids
                exenta = True
                if taxes:
                    for tax in taxes:
                        if 'IVA 13%' in tax.name:
                            exenta = False

                if exenta:
                    total = line.price_subtotal
                    total_exenta = round(total_exenta + total, 2)
                else:
                    total = line.price_subtotal
                    total_gravada = round(total_gravada + total, 2)

            total_retencion = total_exenta * 0.1
            data.append([
                n,
                '2',
                num_doc,
                partner.name,
                record.date.strftime('%d/%m/%Y'),
                confirmation['uid_de'],
                format(record.amount_untaxed, '.2f'),
                format(total_retencion, '.2f'),
                num_doc,
            ])

            global_monto += abs(record.amount_untaxed)
            global_retencion += total_retencion

            n += 1

        # Document footer
        data.append([
            '',
            '',
            '',
            '',
            '',
            'TOTALES',
            format(global_monto, '.2f'),
            format(global_retencion, '.2f'),
            '',
        ])

        table = Table(data)

        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('FONTSIZE', (0, 0), (-1, -1), 6),
            # set max column size and wrap the words inside it so rows can be displayed correctly
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('INNERGRID', (0, 0), (-1, -1), 0.25, colors.black),
            ('BOX', (0, 0), (-1, -1), 0.25, colors.black),
        ]))

        elements.append(table)

        doc.leftMargin = DOC_MARGIN
        doc.rightMargin = DOC_MARGIN
        doc.topMargin = DOC_MARGIN
        doc.bottomMargin = DOC_MARGIN

        doc.build(elements)

        return filename

    ###################################
    # MISC
    ###################################

    def csv_str(self, value: str) -> str :
        return str(value)

    def get_confirmation(self, move):
        confirmation = {
            'url_qr': '-',
            'uid_de': '-',
            'received_seal': '-',
            'control_number': '-',
            'authorization_date': '-',
            'status': 0,
        }

        anulado = False
        for conf in move.confirmation_id:
            try:
                # if conf.status == 2 and 'DTE Autorizado satisfactoriamente' in conf.errors and not anulado:
                if (conf.status == 2 and 'DTE Autorizado satisfactoriamente' in conf.errors and not anulado) or conf.status == 'PROCESADO':
                    confirmation = {
                        'url_qr': conf.url_qr,
                        'uid_de': conf.uid_de,
                        'received_seal': conf.received_seal,
                        'control_number': conf.control_number,
                        'authorization_date': conf.authorization_date,
                        # 'status': conf.status,
                        'status': 2,
                    }

                # if conf.status == 2 and 'DTE Autorizado satisfactoriamente' not in conf.errors:
                if (conf.status == 2 and 'DTE Autorizado satisfactoriamente' not in conf.errors) or conf.status == 'ANULADO':
                    anulado = True
                    confirmation = {
                        'url_qr': '-',
                        'uid_de': '-',
                        'received_seal': '-',
                        'control_number': '-',
                        'authorization_date': '-',
                        'status': 0,
                    }

            except Exception as e:
                confirmation = {
                    'url_qr': '-',
                    'uid_de': '-',
                    'received_seal': '-',
                    'control_number': '-',
                    'authorization_date': '-',
                    'status': 0,
                }

        _logger.info(confirmation)
        return confirmation

    def get_confirmation_for_tax_books(self, move):
        confirmation = {
            'url_qr': '-',
            'uid_de': '-',
            'received_seal': '-',
            'control_number': '-',
            'authorization_date': '-',
            'status': 0,
        }

        anulado = False
        for conf in move.confirmation_id:
            try:
                # if conf.status == 2 and 'DTE Autorizado satisfactoriamente' in conf.errors and not anulado:
                # if (conf.status == 2 and 'DTE Autorizado satisfactoriamente' in conf.errors and not anulado) or conf.status == 'PROCESADO':
                if (conf.status == '1' and anulado == False) or conf.status == 'PROCESADO':
                    
                    confirmation = {
                        'url_qr': conf.url_qr,
                        'uid_de': conf.uid_de,
                        'received_seal': conf.received_seal,
                        'control_number': conf.control_number,
                        'authorization_date': conf.authorization_date,
                        # 'status': conf.status,
                        'status': 1,
                    }

                # if conf.status == 2 and 'DTE Autorizado satisfactoriamente' not in conf.errors:
                # if (conf.status == 2 and 'DTE Autorizado satisfactoriamente' not in conf.errors) or conf.status == 'ANULADO':
                if conf.status == '2' or conf.status == 'ANULADO':
                    anulado = True
                    
                    # Obtener el registro del DTE pero en estado PROCESADO para obtener el codigo de generacion, sello de recepcion y numero de control
                    dte_procesado = self.env['elinvoice.confirmation'].search([
                        ('account_move_id', '=', move.id),
                        ('status', '=', 'PROCESADO')
                    ], limit=1)

                    confirmation = {
                        'url_qr': dte_procesado['url_qr'],
                        'uid_de': dte_procesado['uid_de'],
                        'received_seal': dte_procesado['received_seal'],
                        'control_number': dte_procesado['control_number'],
                        'authorization_date': dte_procesado['authorization_date'],
                        'status': 0,
                    }

            except Exception as e:
                confirmation = {
                    'url_qr': '-',
                    'uid_de': '-',
                    'received_seal': '-',
                    'control_number': '-',
                    'authorization_date': '-',
                    'status': 0,
                }

        _logger.info(confirmation)
        return confirmation

MESES = [
    'Enero',
    'Febrero',
    'Marzo',
    'Abril',
    'Mayo',
    'Junio',
    'Julio',
    'Agosto',
    'Septiembre',
    'Octubre',
    'Noviembre',
    'Diciembre'
]

LOGO_WITH = 100
LOGO_HEIGHT = 50
DOC_MARGIN = 16
