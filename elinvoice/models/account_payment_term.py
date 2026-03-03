# -*- coding: utf-8 -*-

from odoo import api, fields, models


class AccountPaymentTerm(models.Model):
    _inherit = "account.payment.term"

    mh_code = fields.Char(string="Código", required=False)