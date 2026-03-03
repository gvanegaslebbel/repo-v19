from odoo import api, fields, models

class City(models.Model):
    _name = 'elinvoice.city'
    _description = "Municipio"

    code = fields.Char(string="Código", required=False)
    name = fields.Char(string="Nombre", required=False)
    state_id = fields.Many2one('elinvoice.state', string='Departamento')

    def _compute_display_name(self):
        for record in self:
            if record.code:
                record.display_name = f"{record.code} - {record.name}"
            else:
                record.display_name = record.name
