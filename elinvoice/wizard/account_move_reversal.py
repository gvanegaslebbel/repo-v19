# -*- coding: utf-8 -*-
from odoo import models, fields, api
from odoo.tools.translate import _
from odoo.exceptions import UserError
import logging

class AccountMoveReversal(models.TransientModel):
    _inherit = 'account.move.reversal'

    def refund_moves(self):
        res = super(AccountMoveReversal, self).refund_moves()
        moves = self.new_move_ids

        logger = logging.getLogger(__name__)
        logger.info('############ REVERSE ############')
        logger.info(moves)
        
        for move in moves:
            move.is_processed = False
            move.is_anulated = False

        return res

    def modify_moves(self):
        res = super(AccountMoveReversal, self).modify_moves()
        moves = self.new_move_ids

        logger = logging.getLogger(__name__)
        logger.info('############ REVERSE ############')
        logger.info(moves)

        for move in moves:
            move.is_processed = False
            move.is_anulated = False

        return res
