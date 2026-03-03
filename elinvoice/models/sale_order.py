from odoo import fields, models, api

class SaleOrder(models.Model):
    _inherit = 'sale.order'

    ###
    # Confirmacion de MH
    ###
    confirmation_id = fields.One2many('elinvoice.confirmation', 'sale_order_id', string='Confirmacion relacionada')

    ###
    # Campo para saber si el documento ya fue procesado en Zaapdos en MH
    ###
    is_processed = fields.Boolean(string="Is Processed", default=False)

    ###
    # Campo para saber si el documento ya fue anulado en Zaapdos en MH
    ###
    is_anulated = fields.Boolean(string="Is Anulated", default=False)

    processed_status_mh = fields.Char(
        string="Estado de Procesamiento",
        compute="_compute_processed_status",
        store=True
    )

    @api.depends('is_processed', 'is_anulated')
    def _compute_processed_status(self):
        
        for record in self:

            mhStatus = 'No aplicado a MH'

            if record.is_processed:
                mhStatus = 'Procesado MH'

            if record.is_anulated:
                mhStatus = 'Anulado MH'

            record.processed_status_mh = mhStatus