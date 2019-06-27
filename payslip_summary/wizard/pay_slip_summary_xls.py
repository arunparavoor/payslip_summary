# -*- coding: utf-8 -*-

import xlwt
import base64
import calendar
import datetime
from io import BytesIO
from xlsxwriter.workbook import Workbook
from odoo import models, fields, api, _
from odoo.exceptions import UserError, ValidationError, Warning
from datetime import date



class PaySlipSummaryReport(models.TransientModel):
    _name = "pay.slip.summary.report"
    
    start_date = fields.Date(string='Start Date', required=True, default=date.today().replace(day=1))
    end_date = fields.Date(string='End Date', required=True, default=date.today().replace(day=calendar.monthrange(date.today().year, date.today().month)[1]))
    pay_slip_summary_data = fields.Char('Name', size=256)
    file_name = fields.Binary('Pay Slip Summary Report', readonly=True)
    state = fields.Selection([('choose', 'choose'), ('get', 'get')],
                             default='choose')

    _sql_constraints = [
            ('check','CHECK((start_date <= end_date))',"End date must be greater then sPFrt date")
    ]

    @api.multi
    def action_pay_slip_summary_report(self):
        file_name = 'Pay Slip Summary.xls'
        query="""select 
                    hr_employee.name as EMP_NAME,         
                    to_char(hr_payslip.date_from,'Mon-YYYY') as month,
                    hr_employee.id as employee_id,
                    hr_payslip.note as note,
                    (select total from hr_payslip_line as plineinner where plineinner.slip_id=hr_payslip.id and code like '%%BASIC%%') basic_total
                    ,(select total from hr_payslip_line as plineinner where plineinner.slip_id=hr_payslip.id and code like '%%HRA%%') hra_total  
                    ,(select total from hr_payslip_line as plineinner where plineinner.slip_id=hr_payslip.id and code like '%%CA%%') CA_total
                    ,(select total from hr_payslip_line as plineinner where plineinner.slip_id=hr_payslip.id and code like '%%PF%%') PF_total    
                    ,(select total from hr_payslip_line as plineinner where plineinner.slip_id=hr_payslip.id and code='GROSS') gross_total   
                    ,(select total from hr_payslip_line as plineinner where plineinner.slip_id=hr_payslip.id and code='NET') net_total 
                    ,CASE
                    WHEN (
                    COALESCE((select plineinner.total from hr_payslip_line as plineinner where plineinner.slip_id=hr_payslip.id  and code like '%%BASIC%%'),0)=COALESCE((select plineinner.total from  hr_payslip payslipinner left join hr_payslip_line as plineinner on plineinner.slip_id=payslipinner.id where payslipinner.employee_id=hr_payslip.employee_id and payslipinner.date_from >= hr_payslip.date_from - interval '1 month' AND payslipinner.date_to <= hr_payslip.date_from - interval '1 days' and code like '%%BASIC%%'),0)
                    AND 
                    COALESCE((select plineinner.total from hr_payslip_line as plineinner where plineinner.slip_id=hr_payslip.id and code like '%%HRA%%'  ),0)=COALESCE((select plineinner.total from  hr_payslip payslipinner left join hr_payslip_line as plineinner on plineinner.slip_id=payslipinner.id where payslipinner.employee_id=hr_payslip.employee_id and payslipinner.date_from >= hr_payslip.date_from - interval '1 month' AND payslipinner.date_to <= hr_payslip.date_from - interval '1 days' and  code like '%%HRA%%' ),0)
                    AND 
                    COALESCE((select plineinner.total from hr_payslip_line as plineinner where plineinner.slip_id=hr_payslip.id and code like '%%CA%%'  ),0)=COALESCE((select plineinner.total from  hr_payslip payslipinner left join hr_payslip_line as plineinner on plineinner.slip_id=payslipinner.id where payslipinner.employee_id=hr_payslip.employee_id and payslipinner.date_from >= hr_payslip.date_from - interval '1 month' AND payslipinner.date_to <= hr_payslip.date_from - interval '1 days' and  code like '%%CA%%' ),0)
                    AND 
                    COALESCE((select plineinner.total from hr_payslip_line as plineinner where plineinner.slip_id=hr_payslip.id and code like '%%PF%%'  ),0)=COALESCE((select plineinner.total from  hr_payslip payslipinner left join hr_payslip_line as plineinner on plineinner.slip_id=payslipinner.id where payslipinner.employee_id=hr_payslip.employee_id and payslipinner.date_from >= hr_payslip.date_from - interval '1 month' AND payslipinner.date_to <= hr_payslip.date_from - interval '1 days' and  code like '%%PF%%' ),0)
                    AND
                    COALESCE((select plineinner.total from hr_payslip_line as plineinner where plineinner.slip_id=hr_payslip.id and code='GROSS'  ),0)=COALESCE((select plineinner.total from  hr_payslip payslipinner left join hr_payslip_line as plineinner on plineinner.slip_id=payslipinner.id where payslipinner.employee_id=hr_payslip.employee_id and payslipinner.date_from >= hr_payslip.date_from - interval '1 month' AND payslipinner.date_to <= hr_payslip.date_from - interval '1 days' and  code='GROSS' ),0)
                    AND 
                    COALESCE((select plineinner.total from hr_payslip_line as plineinner where plineinner.slip_id=hr_payslip.id and code like 'NET'  ),0)=COALESCE((select plineinner.total from  hr_payslip payslipinner left join hr_payslip_line as plineinner on plineinner.slip_id=payslipinner.id where payslipinner.employee_id=hr_payslip.employee_id and payslipinner.date_from >= hr_payslip.date_from - interval '1 month' AND payslipinner.date_to <= hr_payslip.date_from - interval '1 days' and  code like 'NET' ),0)
                    ) THEN 'No'    
                    ELSE 'Yes'
                    END AS ISCHANGE,
                    CASE
                    WHEN (
                    COALESCE((select plineinner.total from hr_payslip_line as plineinner where plineinner.slip_id=hr_payslip.id  and code like '%%BASIC%%'),0)=COALESCE((select plineinner.total from  hr_payslip payslipinner left join hr_payslip_line as plineinner on plineinner.slip_id=payslipinner.id where payslipinner.employee_id=hr_payslip.employee_id and payslipinner.date_from >= hr_payslip.date_from - interval '1 month' AND payslipinner.date_to <= hr_payslip.date_from - interval '1 days' and code like '%%BASIC%%'),0)
                    ) THEN 'No'    
                    ELSE 'Yes'
                    END AS ISBASIC,
                    CASE
                    WHEN (
                    COALESCE((select plineinner.total from hr_payslip_line as plineinner where plineinner.slip_id=hr_payslip.id and code like '%%HRA%%'  ),0)=COALESCE((select plineinner.total from  hr_payslip payslipinner left join hr_payslip_line as plineinner on plineinner.slip_id=payslipinner.id where payslipinner.employee_id=hr_payslip.employee_id and payslipinner.date_from >= hr_payslip.date_from - interval '1 month' AND payslipinner.date_to <= hr_payslip.date_from - interval '1 days' and  code like '%%HRA%%' ),0)
                    ) THEN 'No'    
                    ELSE 'Yes'
                    END AS ISHRA,
                    CASE
                    WHEN ( 
                    COALESCE((select plineinner.total from hr_payslip_line as plineinner where plineinner.slip_id=hr_payslip.id and code like '%%CA%%'  ),0)=COALESCE((select plineinner.total from  hr_payslip payslipinner left join hr_payslip_line as plineinner on plineinner.slip_id=payslipinner.id where payslipinner.employee_id=hr_payslip.employee_id and payslipinner.date_from >= hr_payslip.date_from - interval '1 month' AND payslipinner.date_to <= hr_payslip.date_from - interval '1 days' and  code like '%%CA%%' ),0)
                    ) THEN 'No'    
                    ELSE 'Yes'
                    END AS ISCA,  
                    CASE
                    WHEN ( 
                    COALESCE((select plineinner.total from hr_payslip_line as plineinner where plineinner.slip_id=hr_payslip.id and code like '%%PF%%'  ),0)=COALESCE((select plineinner.total from  hr_payslip payslipinner left join hr_payslip_line as plineinner on plineinner.slip_id=payslipinner.id where payslipinner.employee_id=hr_payslip.employee_id and payslipinner.date_from >= hr_payslip.date_from - interval '1 month' AND payslipinner.date_to <= hr_payslip.date_from - interval '1 days' and  code like '%%PF%%' ),0)
                    ) THEN 'No'    
                    ELSE 'Yes'
                    END AS ISPF,   
                    CASE
                    WHEN (
                    COALESCE((select plineinner.total from hr_payslip_line as plineinner where plineinner.slip_id=hr_payslip.id and code='GROSS'  ),0)=COALESCE((select plineinner.total from  hr_payslip payslipinner left join hr_payslip_line as plineinner on plineinner.slip_id=payslipinner.id where payslipinner.employee_id=hr_payslip.employee_id and payslipinner.date_from >= hr_payslip.date_from - interval '1 month' AND payslipinner.date_to <= hr_payslip.date_from - interval '1 days' and  code='GROSS' ),0)
                    ) THEN 'No'    
                    ELSE 'Yes'
                    END AS ISGROSS,   
                    CASE
                    WHEN (
                    COALESCE((select plineinner.total from hr_payslip_line as plineinner where plineinner.slip_id=hr_payslip.id and code like 'NET'  ),0)=COALESCE((select plineinner.total from  hr_payslip payslipinner left join hr_payslip_line as plineinner on plineinner.slip_id=payslipinner.id where payslipinner.employee_id=hr_payslip.employee_id and payslipinner.date_from >= hr_payslip.date_from - interval '1 month' AND payslipinner.date_to <= hr_payslip.date_from - interval '1 days' and  code like 'NET' ),0)
                    ) THEN 'No'    
                    ELSE 'Yes'
                    END AS ISNET
                    FROM hr_payslip  
                    INNER JOIN hr_employee ON hr_payslip.employee_id = hr_employee.id 
                    where hr_payslip.date_from::date >= %s AND hr_payslip.date_to::date <= %s       
            """  
        params = (self.start_date, self.end_date,)      
        self.env.cr.execute(query,params)
        
        workbook = xlwt.Workbook(encoding="UTF-8")
        format0 = xlwt.easyxf('font:height 500,bold True;pattern: pattern solid, fore_colour gray25;align: horiz center')
        formathead2 = xlwt.easyxf('font:height 250,bold True;pattern: pattern solid, fore_colour gray25;align: horiz center')
        format1 = xlwt.easyxf('font:bold True;pattern: pattern solid, fore_colour gray25;align: horiz left')
        format2 = xlwt.easyxf('font:bold True;align: horiz left')
        format3 = xlwt.easyxf('align: horiz left')
        format4 = xlwt.easyxf('align: horiz right')
        format5 = xlwt.easyxf('font:bold True;align: horiz right')
        format6 = xlwt.easyxf('font:bold True;pattern: pattern solid, fore_colour gray25;align: horiz right')
        format6yellow = xlwt.easyxf('font:bold True;pattern: pattern solid, fore_colour yellow;align: horiz right')
        format6yellowleft = xlwt.easyxf('font:bold True;pattern: pattern solid, fore_colour yellow;align: horiz left')
        format7 = xlwt.easyxf('font:bold True;borders:top thick;align: horiz right')
        format8 = xlwt.easyxf('font:bold True;borders:top thick;pattern: pattern solid, fore_colour gray25;align: horiz left')
        sheet = workbook.add_sheet("Payslip Summary Report")
        sheet.col(0).width = int(7*260)
        sheet.col(1).width = int(30*260)
        sheet.col(2).width = int(18*260)
        sheet.col(3).width = int(18*260)
        sheet.col(4).width = int(10*260)
        sheet.col(5).width = int(10*260)
        sheet.col(6).width = int(10*260)
        sheet.col(7).width = int(10*260)
        sheet.col(8).width = int(10*260)
        sheet.col(9).width = int(10*260)
        sheet.col(10).width = int(30*260)
        sheet.write_merge(0, 0, 0,10, 'Payslip Summary Report',format0) 
        sheet.write_merge(1, 1, 0,10, 'Payslip:'+str(self.start_date)+' To '+str(self.end_date),formathead2)        
        sheet.write(2, 0, 'Sl.No#', format1)
        sheet.write(2, 1, 'Employee', format1)
        sheet.write(2, 2, 'Month', format1)        
        sheet.write(2, 3, 'Note', format1)
        sheet.write(2, 4, 'Basic', format6)
        sheet.write(2, 5, 'HRA', format6)
        sheet.write(2, 6, 'CA', format6)
        sheet.write(2, 7, 'PF', format6)
        sheet.write(2, 8, 'Gross', format6)
        sheet.write(2, 9, 'Net', format6)
        sheet.write(2, 10, 'Is Change(Compare with Previous month?)', format1)
        i=3
        for row in self.env.cr.fetchall():         
            sheet.write(i, 0, i-2, format3)
            sheet.write(i, 1, row[0], format3)
            sheet.write(i, 2, row[1], format3)
            sheet.write(i, 3, row[3], format3)
            if row[11]=='No':
                sheet.write(i, 4, row[4], format4)
            else:
                sheet.write(i, 4, row[4], format6yellow)
            if row[12]=='No':
                sheet.write(i, 5, row[5], format4)
            else:
                sheet.write(i, 5, row[5], format6yellow)
            if row[13]=='No':
                sheet.write(i, 6, row[6], format4)
            else:
                sheet.write(i, 6, row[6], format6yellow)
            if row[14]=='No':
                sheet.write(i, 7, row[7], format4)
            else:
                sheet.write(i, 7, row[7], format6yellow)
            if row[15]=='No':          
                sheet.write(i, 8, row[8], format4)
            else:
                sheet.write(i, 8, row[8], format6yellow)
            if row[16]=='No':
               sheet.write(i, 9, row[9], format4)
            else:
                sheet.write(i, 9, row[9], format6yellow)
            if row[10]=='No':
                sheet.write(i, 10, row[10], format3)
            else:
                sheet.write(i, 10, row[10], format6yellowleft)
            i=i+1
                
        fp = BytesIO()
        workbook.save(fp)
        self.write({'state': 'get', 'file_name': base64.encodestring(fp.getvalue()), 'pay_slip_summary_data': file_name})
        fp.close()
        return {
           'type': 'ir.actions.act_window',
           'res_model': 'pay.slip.summary.report',
           'view_mode': 'form',
           'view_type': 'form',
           'res_id': self.id,
           'target': 'new',
        }