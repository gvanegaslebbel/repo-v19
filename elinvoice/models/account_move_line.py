import logging

from odoo import fields, models, api

_logger = logging.getLogger(__name__)

class AccountMoveLine(models.Model):
    _inherit = "account.move.line"

    tipo_operacion_renta = fields.Selection(
        selection=[
            ('0', 'N/A'),
            ('1', 'Gravada'),
            ('2', 'No Gravada o Exenta'),
            ('3', 'Excluido o no Constituye Renta')
        ],
        string="Tipo de operacion (Renta)",
        default='1',
        required=False)
