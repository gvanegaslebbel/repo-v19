from odoo import api, fields, models

class State(models.Model):
    _name = 'elinvoice.state'
    _description = "Departamento"

    code = fields.Char(string="Código", required=False)
    name = fields.Char(string="Nombre", required=False)
    country_id = fields.Many2one('elinvoice.country', string='País')
    cities = fields.One2many('elinvoice.city', 'state_id', string='Municipios')

    def _compute_display_name(self):
        for record in self:
            if record.code:
                record.display_name = f"{record.code} - {record.name}"
            else:
                record.display_name = record.name

