# -*- coding: utf-8 -*-

{
    'name': 'Facturación Electrónica',
    'version': '1.0',
    'summary': 'Integración con MH SV',
    'description': 'Este módulo permite conectar facturación electrónica con Odoo.',
    'author': 'Lebbel',
    'license': 'OPL-1',
    'website': 'https://github.com/raulhndz-lebbel/seper-sv',
    'application': True,
    'installable': True,
    'auto_install': False,
    'category': 'Accounting/Accounting',
    'depends': ['account_accountant', 'stock', 'sale_management'],
    'data': [
        'security/ir.model.access.csv',
        'data/elinvoice_data.xml',
        'data/elinvoice_data_locations.xml',
        'data/elinvoice_data_economic_activities.xml',
        'data/elinvoice_data_economic_activities_complement.xml',

        'views/elinvoice_configuration_view.xml',
        'views/account_move_view_inherit.xml',
        'views/res_partner_view_inherit.xml',
        'views/res_company_view_inherit.xml',
        'views/product_template_views_inherit.xml',
        'views/stock_picking_view_inherit.xml',
        'views/sale_order_view_inherit.xml',
        'views/account_payment_term_views_inherit.xml',

        'wizard/export_view.xml',
        'wizard/export_kardex.xml',

        'views/elinvoice_menuitems.xml',

        'views/elinvoice_country_search.xml',
        'views/elinvoice_state_search.xml',
        'views/elinvoice_city_search.xml',
        'views/elinvoice_economic_activity_search.xml',
    ]
}
