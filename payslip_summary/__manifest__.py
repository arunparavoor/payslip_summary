# -*- coding: utf-8 -*-
{
    'name': 'Payslip Summary Report',
	'category': 'Human Resources',
    'author':'Arun Reghu Kumar',
    'version': '0.1',
    'license': "AGPL-3",
    'description': """    
    Payslip Summary Report - Generate summary of payslips for the specified period. This will highlight all the changes from previous month.

    """,
    'maintainer': 'Arun Reghu Kumar',
    'depends': [
        'hr_contract',        
        'hr_payroll'
    ],
    'data': [       
        'wizard/pay_slip_summary_view.xml'
    ],	
    'installable': True,
    'auto_install': False,
}
