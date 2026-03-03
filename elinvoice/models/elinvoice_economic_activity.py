from odoo import fields, models

class EconomicActivity(models.Model):
    _name = 'elinvoice.economic_activity'
    _description = "Actividad económica"

    code = fields.Char(string="Código", required=False)
    name = fields.Char(string="Nombre", required=False)

    def _compute_display_name(self):
        for record in self:
            if record.code:
                record.display_name = f"{record.code} - {record.name}"
            else:
                record.display_name = record.name
