{
    'name': 'Stock Movement Report, Inventory Balance & Tracking Report, Stock Balance Report',
    'summary': 'Advanced Stock Movement Audit Report with Dynamic Opening Balance and Excel XLSX Export. Stock movement report, inventory balance, and advanced stock tracking',
    'version': '18.0.1.0.0',
    'category': 'Inventory/Reporting',
    'author': 'Advantix Solutions',
    'license': 'OPL-1',
    'price': 17.35,
    'currency': 'EUR',
    'depends': [
        'base',
        'stock',

    ],

    'data': [
        'security/ir.model.access.csv',
        'views/stock_tracking_report_wizard.xml',
    ],
    'assets': {},
    'images': ['static/description/banner.png'],
    'installable': True,
    'application': False,
    'auto_install': False,
}
