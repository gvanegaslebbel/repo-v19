from odoo import api, fields, models

class Partner(models.Model):
    _inherit = 'res.partner'

    tipodocumento = fields.Selection(selection=[('36', 'NIT'),
                                                ('13', 'DUI'),
                                                ('37', 'Otro'),
                                                ('03', 'Pasaporte'),
                                                ('02', 'Carnet de Residente'),],
                                    string="Tipo de documento",
                                    required=False)

    numdocumento = fields.Char(string="Numero de documento", required=False)

    nit = fields.Char(string="NIT", required=False)
    nrc = fields.Char(string="NRC", required=False)

    codactividad = fields.Char(string="Codigo de la Actividad", required=False)
    descactividad = fields.Char(string="Actividad", required=False)

    departamento = fields.Char(string="Codigo departamento", required=False)
    municipio = fields.Char(string="Codigo municipio", required=False)

    codpais = fields.Char(string="Codigo pais", required=False)
    nombrepais = fields.Char(string="Nombre pais", required=False)

    is_internal = fields.Selection(selection=[('SI', 'Interno'),
                                            ('NO', 'Externo'),],
                                    string="Local/Externo",
                                    required=False)

    categoria = fields.Selection(selection=[('1', 'Costo'),
                                            ('2', 'Gasto'),],
                                    string="Categoria",
                                    required=False)

    
    pais_id = fields.Many2one('elinvoice.country', string='País', required=False)
    departamento_id = fields.Many2one('elinvoice.state', string='Departamento', required=False) #domain="[('country_id', '=', country_id)]"
    municipio_id = fields.Many2one('elinvoice.city', string='Municipio', required=False) #domain="[('state_id', '=', state_id)]"
    complemento = fields.Char(string="Complemento", required=False)
    actividad_economica_id = fields.Many2one('elinvoice.economic_activity', string='Actividad económica', required=False)
    tipo_persona = fields.Selection(selection=[('NATURAL', 'Persona natural'),
                                               ('JURIDICA', 'Persona juridica')],
                                        string="Tipo de persona",
                                        required=False)