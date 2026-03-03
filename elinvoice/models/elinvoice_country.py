from odoo import fields, models

class Country(models.Model):
    _name = 'elinvoice.country'
    _description = "País"

    code = fields.Char(string="Código", required=False)
    name = fields.Char(string="Nombre", required=False)
    states = fields.One2many('elinvoice.state', 'country_id', string='Departamentos')

    def _compute_display_name(self):
        for record in self:
            if record.code:
                record.display_name = f"{record.code} - {record.name}"
            else:
                record.display_name = record.name
