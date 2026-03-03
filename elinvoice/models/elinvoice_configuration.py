from odoo import fields, models

class Configuration(models.Model):
    _name = "elinvoice.configuration"
    _description = "Electronic Invoicing Configuration"

    api_url = fields.Char(string="URL del Facturador", required=True)
