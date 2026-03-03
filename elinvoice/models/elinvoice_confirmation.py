from odoo import fields, models

class Confirmation(models.Model):
    _name = "elinvoice.confirmation"
    _description = "Electronic Invoicing Confirmation"

    # Field to link to the account.move model
    account_move_id = fields.Many2one('account.move', string='Factura', required=False)

    # Field to link to the sale.order model
    sale_order_id = fields.Many2one('sale.order', string='Nota de remision', required=False)

    dte_path = fields.Char(string="DTE", required=False)
    pdf_path = fields.Char(string="PDF", required=False)
    json_path = fields.Char(string="JSON", required=False)

    url_qr = fields.Char(string="URL", required=True)
    uid_de = fields.Char(string="UID", required=True)
    received_seal = fields.Char(string="Sello de recibido", required=True)
    control_number = fields.Char(string="Numero de control", required=True)
    authorization_date = fields.Char(string="Fecha de autorizacion", required=True)
    status = fields.Char(string="Estado", required=True)
    message = fields.Char(string="Mensage", required=True)
