from odoo import api, fields, models

class Company(models.Model):
    _inherit = 'res.company'

    nit = fields.Char(string="NIT", required=False)
    nrc = fields.Char(string="NRC", required=False)

    tipoestablecimiento = fields.Char(string="Tipo de establecimiento", required=False)
    codestablemh = fields.Char(string="Codigo de establecimiento MH", required=False)
    codestable = fields.Char(string="Codigo de establecimiento", required=False)

    codpuntoventamh = fields.Char(string="Codigo punto de venta MH", required=False)
    codpuntoventa = fields.Char(string="Codigo punto de venta", required=False)

    codactividad = fields.Char(string="Codigo de la Actividad", required=False)
    descactividad = fields.Char(string="Actividad", required=False)

    api_key = fields.Char(string="API Key del facturador", required=False)

    tipo_ingreso_renta_conf = fields.Selection(
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
        default='0')
