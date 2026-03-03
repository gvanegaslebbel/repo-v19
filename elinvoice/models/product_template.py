from odoo import fields, models

class ProductTemplate(models.Model):
    _inherit = "product.template"

    tipo_ingreso_renta = fields.Selection(
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
