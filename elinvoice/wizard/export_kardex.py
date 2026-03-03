import base64
import logging
import xlsxwriter
import io

from io import BytesIO

from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
from reportlab.lib.pagesizes import letter, landscape, A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT

from odoo import api, fields, models
from reportlab.platypus import Image
from decimal import Decimal, ROUND_HALF_UP

_logger = logging.getLogger(__name__)


class ExportKardex(models.TransientModel):
    _name = 'elinvoice.export_kardex'
    _description = 'Download CSV Wizard'

    # fields 
    product_id = fields.Many2one(
        comodel_name='product.product',
        string='Producto',
        required=False,
        domain=[('product_tmpl_id.type', '=', 'consu')],
    )

    file_format = fields.Selection(
        selection=[
            ('PDF', 'PDF'),
            ('XLS', 'EXCEL'),
        ],
        string="Formato",
        required=True,
        default='PDF'
    )
    date_from = fields.Date(string="Fecha inicio", required=True, default=fields.Date.today)
    date_to = fields.Date(string="Fecha fin", required=True, default=fields.Date.today)

    # functions
    def get_report(self):
        filename = 'kardex'
        if(self.file_format == 'PDF'):
            file = self.make_kardex_pdf(filename)
            return self.download_kardex_pdf_file(file)
        elif(self.file_format == 'XLS'):
            file = self.make_kardex_xls(filename)
            return self.download_kardex_xls_file(file)

    def download_kardex_pdf_file(self, filename):
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

    def download_kardex_xls_file(self, filename):
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

    ###################################
    # PDF DOCUMENT MAKERS
    ###################################
    def make_kardex_pdf(self, filename):
        filename = filename + '.pdf'

        producto = self.product_id
        _productId = producto.id
        _from = self.date_from
        _to = self.date_to

        # Determinar si es un producto o todos
        if self.product_id:
            productos = self.product_id
        else:
            productos = self.env['product.product'].search([
                ('product_tmpl_id.type', '=', 'consu')
            ])

        doc = SimpleDocTemplate(filename, pagesize=landscape(A4))
        styles = getSampleStyleSheet()
        elements = []

        for producto in productos:

            producto = producto
            _productId = producto.id

            # Document header
            elements.append(Paragraph('<u>REGISTRO DE CONTROL DE INVENTARIOS</u>', styles['title']))
            elements.append(Spacer(1, 12))

            # Codigo del producto
            elements.append(Paragraph('<b>Artículo:</b> ' + str(producto.default_code), styles['Normal']))
            elements.append(Spacer(1, 6))

            # Descripcion del producto
            elements.append(Paragraph('<b>Descripción del artículo:</b> ' + str(producto.name), styles['Normal']))
            elements.append(Spacer(1, 6))

            # Presentacion del producto
            elements.append(Paragraph('<b>Unidad de medida:</b> ' + str(producto.uom_id.name), styles['Normal']))
            elements.append(Spacer(1, 6))

            # Desde - Hasta
            elements.append(Paragraph(
                '<b>Desde:</b> ' + str(_from.strftime('%Y-%m-%d')) + ' | <b>Hasta:</b> ' + str(_to.strftime('%Y-%m-%d')),
                styles['Normal']))
            elements.append(Spacer(1, 6))

            # Table body
            data = []

            # Table header
            data = [
                ['N°', 'FECHA', 'N° DOC.', 'TIPO DE\nDOCUMENTO', 'TIPO DE\nOPERACION', 'NOMBRE PROVEEDOR/\nNOMBRE CLIENTE', 'NACIONALIDAD\nPROVEEDOR', 'COSTO', 'ENTRADAS', '', '', 'SALIDAS', '', '', 'SALDO', '', ''],
                ['', '', '', '', '', '', '', '', 'CANT.', 'PRECIO\nUNITARIO', 'TOTAL', 'CANT.', 'PRECIO\nUNITARIO', 'TOTAL', 'CANT.', 'COSTO U.\nPROMEDIO', 'TOTAL', ]
            ]

            n = 1

            tieneSaldoAnterior = False

            _totalInvertido = 0
            _costoPromedio = 0

            _totalEntradas = 0
            _totalSalidas = 0

            cantidad_saldo = 0
            precio_saldo = 0
            total_saldo = 0

            lines = self.env['stock.valuation.layer'].search([
                ('product_id', '=', _productId),
                ('quantity', '!=', 0),
                ('create_date', '<', _from.strftime('%Y-%m-%d')),
            ])

            for line in lines:
                if line:
                    tieneSaldoAnterior = True
                    cantidad_entrada = 0
                    precio_entrada = 0
                    total_entrada = 0
                    cantidad_salida = 0
                    precio_salida = 0
                    total_salida = 0

                    if line.quantity > 0:
                        cantidad_entrada = round(line.quantity, 2)
                        precio_entrada = round(line.unit_cost, 2)
                        total_entrada = round(line.value, 2)

                        cantidad_saldo = round((cantidad_saldo + cantidad_entrada), 2)
                        _totalEntradas = _totalEntradas + cantidad_entrada
                    else:
                        cantidad_salida = round((line.quantity * -1), 2)
                        precio_salida = round(line.unit_cost, 3)  # viene desde Odoo
                        total_salida = round((line.value * -1), 3)  # viende desde Odoo
                        # precio_salida = round(_costoPromedio, 2) # costo calculado
                        # total_salida = round((cantidad_salida * precio_salida), 2) # total calculado

                        cantidad_saldo = round((cantidad_saldo - cantidad_salida), 2)
                        _totalSalidas = _totalSalidas + cantidad_salida

                    _totalInvertido = round((_totalInvertido + total_entrada - total_salida), 2)
                    if cantidad_saldo > 0:
                        _costoPromedio = round((_totalInvertido / cantidad_saldo), 2)
                    else:
                        _costoPromedio = 0
                    precio_saldo = round(_costoPromedio, 2)
                    total_saldo = round((cantidad_saldo * precio_saldo), 2)

            _cantidad_saldo_anterior = cantidad_saldo
            _prercio_saldo_anterior = precio_saldo
            _total_saldo_anterior = total_saldo

            _totalInvertido = 0
            _costoPromedio = 0

            # _totalEntradas = 0
            # _totalSalidas = 0

            cantidad_saldo = 0
            precio_saldo = 0
            total_saldo = 0

            purchase_invoice = ''
            purchase_tax_document = ''
            supplier = ''

            sale_invoice = ''
            sale_tax_document = ''
            customer = ''

            lines = self.env['stock.valuation.layer'].search([
                ('product_id', '=', _productId),
                ('quantity', '!=', 0),
                ('create_date', '>=', _from.strftime('%Y-%m-%d')),
                ('create_date', '<=', _to.strftime('%Y-%m-%d')),
            ])

            if tieneSaldoAnterior:
                data.append([
                    '',
                    '',
                    '',
                    '',
                    '',
                    'SALDO ANTERIOR',
                    '',
                    format(0, '.2f'),
                    format(0, '.2f'),
                    format(0, '.2f'),
                    format(0, '.2f'),
                    format(0, '.2f'),
                    format(0, '.2f'),
                    format(0, '.2f'),
                    format(_cantidad_saldo_anterior, '.2f'),
                    format(_prercio_saldo_anterior, '.2f'),
                    format(_total_saldo_anterior, '.2f')
                ])

                cantidad_saldo = cantidad_saldo + _cantidad_saldo_anterior
                _totalInvertido = _totalInvertido + _total_saldo_anterior

                n = n + 1

            for line in lines:
                purchase = False
                sale = False
                if line:
                    stockValuationLayer = line
                    stockMove = stockValuationLayer.stock_move_id

                    purchaseOrderLine = stockMove.purchase_line_id
                    purchaseOrder = purchaseOrderLine.order_id

                    saleOrderLine = stockMove.sale_line_id
                    saleOrder = saleOrderLine.order_id

                    if purchaseOrder:
                        _logger.info('purchaseOrderId: ' + str(purchaseOrder.id))
                        invoiceIds = purchaseOrder.invoice_ids
                        if invoiceIds:
                            purchase = True
                            _logger.info('Si existen facturas para esta orden de compra')
                            for invoiceId in invoiceIds:
                                _logger.info('invoiceId: ' + str(invoiceId.id))
                                purchase_invoice = invoiceId
                                purchase_tax_document = purchase_invoice.journal_id
                                supplier = purchase_invoice.partner_id

                    if saleOrder:
                        _logger.info('saleOrderId: ' + str(saleOrder.id))
                        invoiceIds = saleOrder.invoice_ids
                        if invoiceIds:
                            sale = True
                            _logger.info('Si existen facturas para esta orden de venta')
                            for invoiceId in invoiceIds:
                                _logger.info('invoiceId: ' + str(invoiceId.id))
                                sale_invoice = invoiceId
                                sale_tax_document = sale_invoice.journal_id
                                customer = sale_invoice.partner_id

                    cantidad_entrada = 0
                    precio_entrada = 0
                    total_entrada = 0
                    cantidad_salida = 0
                    precio_salida = 0
                    total_salida = 0

                    if line.quantity > 0:
                        cantidad_entrada = round(line.quantity, 2)
                        precio_entrada = round(line.unit_cost, 2)
                        total_entrada = round(line.value, 2)

                        cantidad_saldo = round((cantidad_saldo + cantidad_entrada), 2)
                        _totalEntradas = _totalEntradas + cantidad_entrada
                    else:
                        cantidad_salida = round((line.quantity * -1), 2)
                        precio_salida = round(line.unit_cost, 3)  # viene desde Odoo
                        total_salida = round((line.value * -1), 3)  # viende desde Odoo
                        # precio_salida = round(_costoPromedio, 2) # costo calculado
                        # total_salida = round((cantidad_salida * precio_salida), 2) # total calculado

                        cantidad_saldo = round((cantidad_saldo - cantidad_salida), 2)
                        _totalSalidas = _totalSalidas + cantidad_salida

                    _totalInvertido = round((_totalInvertido + total_entrada - total_salida), 2)
                    if cantidad_saldo > 0:
                        _costoPromedio = round((_totalInvertido / cantidad_saldo), 2)
                    else:
                        _costoPromedio = 0
                    precio_saldo = round(_costoPromedio, 2)
                    total_saldo = round((cantidad_saldo * precio_saldo), 2)

                    if not sale and not purchase:
                        date = ''
                        document_number = ''
                        tax_document_type = ''
                        move_type = ''
                        owner_name = 'MOVIMIENTO DE INVENTARIO'
                        owner_country = ''
                    else:
                        if purchase:
                            date = purchase_invoice.date.strftime('%d-%m-%Y')
                            document_number = purchase_invoice.sequence_number
                            tax_document_type = purchase_tax_document.code
                            move_type = 'COMPRA'
                            owner_name = supplier.name[:30]
                            owner_country = supplier.pais_id.name
                        else:
                            date = sale_invoice.date.strftime('%d-%m-%Y')
                            document_number = sale_invoice.sequence_number
                            tax_document_type = sale_tax_document.code
                            move_type = 'VENTA'
                            owner_name = customer.name[:30]
                            owner_country = customer.pais_id.name

                    # Table body
                    data.append([
                        str(n),
                        date, # purchase_invoice.date.strftime('%d-%m-%Y') if purchase else sale_invoice.date.strftime('%d-%m-%Y'),
                        document_number, # purchase_invoice.sequence_number if purchase else sale_invoice.sequence_number,
                        tax_document_type, # purchase_tax_document.code if purchase else sale_tax_document.code,
                        move_type, #'COMPRA' if purchase else 'VENTA',
                        owner_name, # supplier.name[:30] if purchase else customer.name[:30],
                        owner_country, # supplier.pais_id.name if purchase else customer.pais_id.name,
                        format(_costoPromedio, '.2f'),
                        format(cantidad_entrada, '.2f'),
                        format(precio_entrada, '.2f'),
                        format(total_entrada, '.2f'),
                        format(cantidad_salida, '.2f'),
                        format(precio_salida, '.2f'),
                        format(total_salida, '.2f'),
                        format(cantidad_saldo, '.2f'),
                        format(precio_saldo, '.2f'),
                        format(total_saldo, '.2f')
                    ])

                n = n + 1

            table = Table(data, colWidths=[25, 50, 38, 58, 54, 140, 68, 38, 33, 50, 33, 33, 50, 33, 33, 50, 33])
            table.setStyle(TableStyle([

                ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
                ('FONTSIZE', (0, 0), (-1, -1), 8),

                ('SPAN', (0, 0), (0, 1)),  # N°
                ('SPAN', (1, 0), (1, 1)),  # Fecha
                ('SPAN', (2, 0), (2, 1)),  # N° doc.
                ('SPAN', (3, 0), (3, 1)),  # Tipo de documento
                ('SPAN', (4, 0), (4, 1)),  # Tipo de operacion
                ('SPAN', (5, 0), (5, 1)),  # Nombre proveedor/cliente
                ('SPAN', (6, 0), (6, 1)),  # Nacionalidad
                ('SPAN', (7, 0), (7, 1)),  # Costo
                ('SPAN', (8, 0), (10, 0)),  # Entradas
                ('SPAN', (11, 0), (13, 0)),  # Saidas
                ('SPAN', (14, 0), (16, 0)),  # Saldos

                ('ALIGN', (0, 0), (16, 1), 'CENTER'),
                ('GRID', (0, 0), (16, 1), 0.5, colors.black),

                ('ALIGN', (0, 2), (4, n), 'CENTER'),
                ('ALIGN', (5, 2), (5, n), 'LEFT'),
                ('ALIGN', (6, 2), (16, n), 'CENTER'),

                ('LINEBELOW', (0, n), (16, n), 0.5, colors.black),

            ]))

            elements.append(table)

            # Table
            data = [
                ['', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', ''],
                ['TOTAL UNIDADES ENTRADAS:', '', '', '', '', '', '', '', format(_totalEntradas, '.2f'), '', '', '', '', '',
                '', '', ''],
                ['TOTAL UNIDADES SALIDAS:', '', '', '', '', '', '', '', format(_totalSalidas, '.2f'), '', '', '', '', '',
                '', '', ''],
                ['EXISTENCIAS:', '', '', '', '', '', '', '', format(cantidad_saldo, '.2f'), '', '', '', '', '', '', '', ''],
            ]

            table = Table(data, colWidths=[30, 50, 38, 58, 54, 100, 100, 38, 33, 51, 33, 33, 51, 33, 33, 51, 33])
            table.setStyle(TableStyle([
                ('SPAN', (0, 0), (16, 0)),  # Fila 1

                ('SPAN', (0, 1), (7, 1)),  # Fila 2
                ('SPAN', (8, 1), (16, 1)),  # Fila 2

                ('SPAN', (0, 2), (7, 2)),  # Fila 3
                ('SPAN', (8, 2), (16, 2)),  # Fila 3

                ('SPAN', (0, 3), (7, 3)),  # Fila 4
                ('SPAN', (8, 3), (16, 3)),  # Fila 4

                # Estilos para toda la tabla
                ('ALIGN', (0, 1), (0, 1), 'RIGHT'),
                ('ALIGN', (0, 2), (0, 2), 'RIGHT'),
                ('ALIGN', (0, 3), (0, 3), 'RIGHT'),

                ('ALIGN', (1, 1), (1, 1), 'LEFT'),
                ('ALIGN', (1, 2), (1, 2), 'LEFT'),
                ('ALIGN', (1, 3), (1, 3), 'LEFT'),

                ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
                ('FONTSIZE', (0, 0), (-1, -1), 8),
            ]))

            elements.append(table)

            elements.append(PageBreak())

        doc.leftMargin = 20
        doc.rightMargin = 20
        doc.topMargin = 10
        doc.bottomMargin = 10

        doc.build(elements)

        return filename

    def make_kardex_xls(self, filename):
        filename = filename + '.xls'

        producto = self.product_id
        _productId = producto.id
        _from = self.date_from
        _to = self.date_to

        # Determinar si es un producto o todos
        if self.product_id:
            productos = self.product_id
        else:
            productos = self.env['product.product'].search([
                ('product_tmpl_id.type', '=', 'consu')
            ])
            
        workbook = xlsxwriter.Workbook(filename)

        center_format = workbook.add_format({
            'align': 'center',
            'valign': 'vcenter',
            'bold': True,
            'text_wrap': True,
            'font_size': 8,
            'font_name': 'Arial'
        })

        left_format = workbook.add_format({
            'align': 'left',
            'valign': 'vcenter',
            'bold': True,
            'font_size': 8,
            'font_name': 'Arial'
        })

        center_format_regular = workbook.add_format({
            'align': 'center',
            'valign': 'vcenter',
            'bold': False,
            'font_size': 8,
            'font_name': 'Arial'
        })

        left_format_regular = workbook.add_format({
            'align': 'left',
            'valign': 'vcenter',
            'bold': False,
            'font_size': 8,
            'font_name': 'Arial'
        })

        worksheet = workbook.add_worksheet()

        # Document header
        # worksheet.merge_range('A1:Q1', 'REGISTRO DE CONTROL DE INVENTARIOS', center_format)
        # worksheet.merge_range('A2:Q2', 'Articulo: ' + str(producto.default_code), left_format)
        # worksheet.merge_range('A3:Q3', 'Descripcion del articulo: ' + str(producto.name), left_format)
        # worksheet.merge_range('A4:Q4', 'Unidad de medida: Unidad', left_format)
        # worksheet.merge_range('A5:Q5', 'Desde: ' + str(_from.strftime('%Y-%m-%d')) + ' | Hasta: ' + str(_to.strftime('%Y-%m-%d')), left_format)

        worksheet.merge_range('A6:A7', 'N°', center_format)
        worksheet.merge_range('B6:B7', 'CODIGO\nARTICULO', center_format)
        worksheet.merge_range('C6:C7', 'ARTICULO', center_format)
        worksheet.merge_range('D6:D7', 'UNIDAD\nMEDIDA', center_format)
        worksheet.merge_range('E6:E7', 'DESDE', center_format)
        worksheet.merge_range('F6:F7', 'HASTA', center_format)
        worksheet.merge_range('G6:G7', 'FECHA\nDOCUMENTO', center_format)
        worksheet.merge_range('H6:H7', 'N°DOC.', center_format)
        worksheet.merge_range('I6:I7', 'TIPO DE\nDOCUMENTO', center_format)
        worksheet.merge_range('J6:J7', 'TIPO DE\nOPERACION', center_format)
        worksheet.merge_range('K6:K7', 'NOMBRE PROVEEDOR\nNOMBRE CLIENTE', center_format)
        worksheet.merge_range('L6:L7', 'NACIONALIDAD\nPROVEEDOR', center_format)
        worksheet.merge_range('M6:M7', 'COSTO', center_format)
        worksheet.merge_range('N6:P6', 'ENTRADAS', center_format)
        worksheet.merge_range('Q6:S6', 'SALIDAS', center_format)
        worksheet.merge_range('T6:V6', 'SALDOS', center_format)

        worksheet.write(6, 13, 'CANTIDAD', center_format)
        worksheet.write(6, 14, 'PRECIO', center_format)
        worksheet.write(6, 15, 'TOTAL', center_format)
        worksheet.write(6, 16, 'CANTIDAD', center_format)
        worksheet.write(6, 17, 'PRECIO', center_format)
        worksheet.write(6, 18, 'TOTAL', center_format)
        worksheet.write(6, 19, 'CANTIDAD', center_format)
        worksheet.write(6, 20, 'PRECIO', center_format)
        worksheet.write(6, 21, 'TOTAL', center_format)

        worksheet.set_column(0, 0, 5)
        worksheet.set_column(1, 1, 10)
        worksheet.set_column(2, 2, 20)
        worksheet.set_column(3, 3, 10)
        worksheet.set_column(4, 4, 10)
        worksheet.set_column(5, 5, 10)
        worksheet.set_column(6, 6, 10)
        worksheet.set_column(7, 7, 10)
        worksheet.set_column(8, 8, 10)
        worksheet.set_column(9, 9, 10)
        worksheet.set_column(10, 10, 20)
        worksheet.set_column(11, 11, 15)
        worksheet.set_column(12, 12, 10)
        worksheet.set_column(13, 13, 10)
        worksheet.set_column(14, 14, 10)
        worksheet.set_column(15, 15, 10)
        worksheet.set_column(16, 16, 10)
        worksheet.set_column(17, 17, 10)
        worksheet.set_column(18, 18, 10)
        worksheet.set_column(19, 19, 10)
        worksheet.set_column(20, 20, 10)
        worksheet.set_column(21, 21, 10)

        row = 7

        for producto in productos:
        
            producto = producto
            _productId = producto.id

            tieneSaldoAnterior = False

            _totalInvertido = 0
            _costoPromedio = 0

            _totalEntradas = 0
            _totalSalidas = 0

            cantidad_saldo = 0
            precio_saldo = 0
            total_saldo = 0

            lines = self.env['stock.valuation.layer'].search([
                ('product_id', '=', _productId),
                ('quantity', '!=', 0),
                ('create_date', '<', _from.strftime('%Y-%m-%d')),
            ])

            for line in lines:
                if line:
                    tieneSaldoAnterior = True
                    cantidad_entrada = 0
                    precio_entrada = 0
                    total_entrada = 0
                    cantidad_salida = 0
                    precio_salida = 0
                    total_salida = 0

                    if line.quantity > 0:
                        cantidad_entrada = round(line.quantity, 2)
                        precio_entrada = round(line.unit_cost, 2)
                        total_entrada = round(line.value, 2)

                        cantidad_saldo = round((cantidad_saldo + cantidad_entrada), 2)
                        _totalEntradas = _totalEntradas + cantidad_entrada
                    else:
                        cantidad_salida = round((line.quantity * -1), 2)
                        precio_salida = round(line.unit_cost, 3)  # viene desde Odoo
                        total_salida = round((line.value * -1), 3)  # viende desde Odoo
                        # precio_salida = round(_costoPromedio, 2) # costo calculado
                        # total_salida = round((cantidad_salida * precio_salida), 2) # total calculado

                        cantidad_saldo = round((cantidad_saldo - cantidad_salida), 2)
                        _totalSalidas = _totalSalidas + cantidad_salida

                    _totalInvertido = round((_totalInvertido + total_entrada - total_salida), 2)
                    if cantidad_saldo > 0:
                        _costoPromedio = round((_totalInvertido / cantidad_saldo), 2)
                    else:
                        _costoPromedio = 0
                    precio_saldo = round(_costoPromedio, 2)
                    total_saldo = round((cantidad_saldo * precio_saldo), 2)

            _cantidad_saldo_anterior = cantidad_saldo
            _prercio_saldo_anterior = precio_saldo
            _total_saldo_anterior = total_saldo

            _totalInvertido = 0
            _costoPromedio = 0

            # _totalEntradas = 0
            # _totalSalidas = 0

            cantidad_saldo = 0
            precio_saldo = 0
            total_saldo = 0

            purchase_invoice = ''
            purchase_tax_document = ''
            supplier = ''

            sale_invoice = ''
            sale_tax_document = ''
            customer = ''

            lines = self.env['stock.valuation.layer'].search([
                ('product_id', '=', _productId),
                ('quantity', '!=', 0),
                ('create_date', '>=', _from.strftime('%Y-%m-%d')),
                ('create_date', '<=', _to.strftime('%Y-%m-%d')),
            ])

            if tieneSaldoAnterior:
                worksheet.write(row, 0, '', center_format_regular)
                worksheet.write(row, 1, producto.default_code or '', center_format_regular)
                worksheet.write(row, 2, producto.name, center_format_regular)
                worksheet.write(row, 3, producto.uom_id.name or '', center_format_regular)
                worksheet.write(row, 4, '', center_format_regular)
                worksheet.write(row, 5, '', center_format_regular)
                worksheet.write(row, 6, '', center_format_regular)
                worksheet.write(row, 7, '', center_format_regular)
                worksheet.write(row, 8, '', center_format_regular)
                worksheet.write(row, 9, '', center_format_regular)
                worksheet.write(row, 10, 'SALDO ANTERIOR', left_format_regular)
                worksheet.write(row, 11, '', center_format_regular)
                worksheet.write(row, 12, '0.00', center_format_regular)
                worksheet.write(row, 13, '0.00', center_format_regular)
                worksheet.write(row, 14, '0.00', center_format_regular)
                worksheet.write(row, 15, '0.00', center_format_regular)
                worksheet.write(row, 16, '0.00', center_format_regular)
                worksheet.write(row, 17, '0.00', center_format_regular)
                worksheet.write(row, 18, '0.00', center_format_regular)
                worksheet.write(row, 19, format(_cantidad_saldo_anterior, '.2f'), center_format_regular)
                worksheet.write(row, 20, format(_prercio_saldo_anterior, '.2f'), center_format_regular)
                worksheet.write(row, 21, format(_total_saldo_anterior, '.2f'), center_format_regular)

                cantidad_saldo = cantidad_saldo + _cantidad_saldo_anterior
                _totalInvertido = _totalInvertido + _total_saldo_anterior

                row = row + 1

            n = 1

            for line in lines:
                purchase = False
                sale = False
                if line:
                    stockValuationLayer = line
                    stockMove = stockValuationLayer.stock_move_id

                    purchaseOrderLine = stockMove.purchase_line_id
                    purchaseOrder = purchaseOrderLine.order_id

                    saleOrderLine = stockMove.sale_line_id
                    saleOrder = saleOrderLine.order_id

                    if purchaseOrder:
                        _logger.info('purchaseOrderId: ' + str(purchaseOrder.id))
                        invoiceIds = purchaseOrder.invoice_ids
                        if invoiceIds:
                            purchase = True
                            _logger.info('Si existen facturas para esta orden de compra')
                            for invoiceId in invoiceIds:
                                _logger.info('invoiceId: ' + str(invoiceId.id))
                                purchase_invoice = invoiceId
                                purchase_tax_document = purchase_invoice.journal_id
                                supplier = purchase_invoice.partner_id

                    if saleOrder:
                        _logger.info('saleOrderId: ' + str(saleOrder.id))
                        invoiceIds = saleOrder.invoice_ids
                        if invoiceIds:
                            sale = True
                            _logger.info('Si existen facturas para esta orden de venta')
                            for invoiceId in invoiceIds:
                                _logger.info('invoiceId: ' + str(invoiceId.id))
                                sale_invoice = invoiceId
                                sale_tax_document = sale_invoice.journal_id
                                customer = sale_invoice.partner_id

                    cantidad_entrada = 0
                    precio_entrada = 0
                    total_entrada = 0
                    cantidad_salida = 0
                    precio_salida = 0
                    total_salida = 0

                    if line.quantity > 0:
                        cantidad_entrada = round(line.quantity, 2)
                        precio_entrada = round(line.unit_cost, 2)
                        total_entrada = round(line.value, 2)

                        cantidad_saldo = round((cantidad_saldo + cantidad_entrada), 2)
                        _totalEntradas = _totalEntradas + cantidad_entrada
                    else:
                        cantidad_salida = round((line.quantity * -1), 2)
                        precio_salida = round(line.unit_cost, 3)  # viene desde Odoo
                        total_salida = round((line.value * -1), 3)  # viende desde Odoo
                        # precio_salida = round(_costoPromedio, 2) # costo calculado
                        # total_salida = round((cantidad_salida * precio_salida), 2) # total calculado

                        cantidad_saldo = round((cantidad_saldo - cantidad_salida), 2)
                        _totalSalidas = _totalSalidas + cantidad_salida

                    _totalInvertido = round((_totalInvertido + total_entrada - total_salida), 2)
                    if cantidad_saldo > 0:
                        _costoPromedio = round((_totalInvertido / cantidad_saldo), 2)
                    else:
                        _costoPromedio = 0
                    precio_saldo = round(_costoPromedio, 2)
                    total_saldo = round((cantidad_saldo * precio_saldo), 2)

                    # Determinar numeros de documento
                    dte_uuid = ''
                    if not sale and not purchase:
                        dte_uuid = ''
                    else:
                        if purchase:
                            dte_uuid = purchase_invoice.codigo_generacion
                        else:
                            for dte in sale_invoice.confirmation_id:
                                if dte.status == 'PROCESADO':
                                    dte_uuid = dte.uid_de
                                    break
                    
                    if not sale and not purchase:
                        date = ''
                        tax_document_type = ''
                        move_type = ''
                        owner_name = 'MOVIMIENTO DE INVENTARIO'
                        owner_country = ''
                    else:
                        if purchase:
                            date = purchase_invoice.date.strftime('%d-%m-%Y')
                            tax_document_type = purchase_tax_document.code
                            move_type = 'COMPRA'
                            owner_name = supplier.name[:30]
                            owner_country = supplier.pais_id.name
                        else:
                            date = sale_invoice.date.strftime('%d-%m-%Y')
                            tax_document_type = sale_tax_document.code
                            move_type = 'VENTA'
                            owner_name = customer.name[:30]
                            owner_country = customer.pais_id.name

                    # Table body
                    worksheet.write(row, 0, str(n), center_format_regular)
                    worksheet.write(row, 1, producto.default_code or '', center_format_regular)
                    worksheet.write(row, 2, producto.name, left_format_regular)
                    worksheet.write(row, 3, producto.uom_id.name or '', center_format_regular)
                    worksheet.write(row, 4, _from.strftime('%Y-%m-%d'), center_format_regular)
                    worksheet.write(row, 5, _to.strftime('%Y-%m-%d'), center_format_regular)
                    worksheet.write(row, 6, date, center_format_regular)
                    worksheet.write(row, 7, dte_uuid, center_format_regular)
                    worksheet.write(row, 8, tax_document_type, center_format_regular)
                    worksheet.write(row, 9, move_type, center_format_regular)
                    worksheet.write(row, 10, owner_name, left_format_regular)
                    worksheet.write(row, 11, owner_country, center_format_regular)
                    worksheet.write(row, 12, format(_costoPromedio, '.2f'), center_format_regular)
                    worksheet.write(row, 13, format(cantidad_entrada, '.2f'), center_format_regular)
                    worksheet.write(row, 14, format(precio_entrada, '.2f'), center_format_regular)
                    worksheet.write(row, 15, format(total_entrada, '.2f'),center_format_regular)
                    worksheet.write(row, 16, format(cantidad_salida, '.2f'), center_format_regular)
                    worksheet.write(row, 17, format(precio_salida, '.2f'), center_format_regular)
                    worksheet.write(row, 18, format(total_salida, '.2f'), center_format_regular)
                    worksheet.write(row, 19, format(cantidad_saldo, '.2f'), center_format_regular)
                    worksheet.write(row, 20, format(precio_saldo, '.2f'), center_format_regular)
                    worksheet.write(row, 21, format(total_saldo, '.2f'), center_format_regular)

                row = row + 1
                n = n + 1

            if _totalEntradas > 0:
                row = row + 1

                worksheet.merge_range('G'+str(row)+':H'+str(row), 'TOTAL UNIDADES ENTRADAS', left_format)
                worksheet.write((row - 1), 8, format(_totalEntradas, '.2f'), center_format)
                row = row + 1

                worksheet.merge_range('G'+str(row)+':H'+str(row), 'TOTAL UNIDADES SALIDAS', left_format)
                worksheet.write((row - 1), 8, format(_totalSalidas, '.2f'), center_format)
                row = row + 1

                worksheet.merge_range('G'+str(row)+':H'+str(row), 'EXISTENCIAS', left_format)
                worksheet.write((row - 1), 8, format(cantidad_saldo, '.2f'), center_format)

                row += 2

        workbook.close()

        return filename
