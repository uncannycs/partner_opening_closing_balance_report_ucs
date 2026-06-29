from odoo import models, fields, api
import datetime
from datetime import datetime, timedelta
import calendar
from dateutil.relativedelta import relativedelta
import base64
from io import BytesIO
import xlsxwriter
import logging
import pytz

_logger = logging.getLogger(__name__)


class PartnerOpeningClosingBalance(models.TransientModel):
    _name = 'partner.opening.closing.balance'
    _description = 'Partner Opening Closing Balance Report'

    date_from = fields.Date(string='Start Date', required=True)
    date_to = fields.Date(string='End Date', required=True)
    partner_ids = fields.Many2many('res.partner', string='Partner')
    result_selection = fields.Selection([
        ('customer', 'Receivable Accounts'),
        ('supplier', 'Payable Accounts'),
    ], string="Partner's", required=True, default='customer')
    company_id = fields.Many2one('res.company', string='Company', required=True, readonly=True,
                                 default=lambda self: self.env.company)
    
    balance_file = fields.Binary()
    balance_name = fields.Char()
    is_zero_account = fields.Boolean(string="Is Zero Account")
    partner_domain = fields.Json(default=lambda self: [["customer_rank", ">", 0]])

    @api.onchange("result_selection")
    def _onchange_selection_id(self):
        if self.result_selection == "customer":
            self.partner_domain = [["customer_rank", ">", 0]]
        else:
            self.partner_domain = [["supplier_rank", ">", 0]]
    



    def action_print_report(self):
        # ------------------------- Excel _________________________
        filename = " Opening - Closing Balance - " + self.date_from.strftime("%B - %Y")
        fp = BytesIO()
        partner_list = []
        if not self.partner_ids:
            if self.result_selection == 'customer':
                partner_ids = self.env['res.partner'].search([('customer_rank', '>', 0)])
            else:
                partner_ids = self.env['res.partner'].search([('supplier_rank', '>', 0)])
        else:
            partner_ids = self.partner_ids
        
        for partner in partner_ids:
            opening_balance = debit = credit = closing_balance = 0
            
            cr = self.env.cr
            query = """
                SELECT
                    SUM(l.debit) - SUM(l.credit) AS balance
                FROM
                    account_move_line l
                JOIN
                    account_move m ON l.move_id = m.id
                JOIN
                    res_partner p ON l.partner_id = p.id
                JOIN
                    account_account a ON l.account_id = a.id
                WHERE
                    l.date < %s
                    AND m.state = 'posted'
                    AND a.account_type IN ('asset_receivable', 'liability_payable')
                    AND l.partner_id = %s
                    AND l.company_id = %s
            """
            params = [self.date_from, partner.id, self.company_id.id]
            
            if self.result_selection == 'customer':
                query += " AND p.customer_rank > 0"
            else:
                query += " AND p.supplier_rank > 0"
            query += " GROUP BY p.name"
            cr.execute(query, params)
            results = cr.fetchall()
            for data in results:
                opening_balance += data[0]
            
            cr = self.env.cr
            query = """
                SELECT
                   p.name AS partner_name,
                   SUM(l.debit) AS debit,
                   SUM(l.credit) AS credit,
                   SUM(l.debit) - SUM(l.credit) AS balance
                FROM
                   account_move_line l
                JOIN
                   account_move m ON l.move_id = m.id
                JOIN
                   res_partner p ON l.partner_id = p.id
                JOIN
                    account_account a ON l.account_id = a.id
                WHERE
                   l.date BETWEEN %s AND %s
                   AND m.state = 'posted'
                   AND a.account_type IN ('asset_receivable', 'liability_payable')
                   AND l.partner_id = %s
                   AND l.company_id = %s
            """
            params = [self.date_from, self.date_to, partner.id, self.company_id.id]
            
            if self.result_selection == 'customer':
                query += " AND p.customer_rank > 0"
            else:
                query += " AND p.supplier_rank > 0"
            query += " GROUP BY p.name"
            cr.execute(query, params)
            results = cr.fetchall()
            for data in results:
                debit += data[1]
                credit += data[2]
            
            closing_balance = (opening_balance + debit) - credit
            values = [opening_balance, debit, credit, closing_balance]
            if not self.is_zero_account:
                if any(val != 0 for val in values):
                    partner_list.append((partner.name, opening_balance, debit, credit, closing_balance))
            else:
                partner_list.append((partner.name, opening_balance, debit, credit, closing_balance))
        
        
        workbook = xlsxwriter.Workbook(fp)
        worksheet = workbook.add_worksheet('Balance')
        worksheet.set_column('A:A', 60)
        worksheet.set_column('B:E', 20)
        
        heading_format = workbook.add_format({
            'bold':True,
            'align':'center',
            'valign':'vcenter',
            'font_size':18,
        })
        
        date_format = workbook.add_format({
            'bold':True,
            'align':'left',
            'valign':'vcenter',
        })
        heading_text = "Partner Trial Balance"
        worksheet.merge_range('A1:E3',heading_text,heading_format)
        worksheet.merge_range('A4:E5',f"Date Range : {self.date_from.strftime('%d/%m/%y')} - {self.date_to.strftime('%d/%m/%y')}",date_format)
        style_center = workbook.add_format({
            'align': 'center',
            'border':1,
            'bold':True
            })
        worksheet.write(5, 0, 'Name', style_center)
        worksheet.write(5, 1, 'Opening Balance', style_center)
        worksheet.write(5, 2, 'Debit', style_center)
        worksheet.write(5, 3, 'Credit', style_center)
        worksheet.write(5, 4, 'Closing', style_center)
        
        row = 6
        if partner_list:
            style_left = workbook.add_format({'align':'left'})
            for data in partner_list:
                worksheet.write(row, 0, data[0], style_left)
                worksheet.write(row, 1, data[1], style_left)
                worksheet.write(row, 2, data[2], style_left)
                worksheet.write(row, 3, data[3], style_left)
                worksheet.write(row, 4, data[4], style_left)
                row += 1
        
        workbook.close()
        self.balance_file = base64.b64encode(fp.getvalue())
        fp.close()
        self.balance_name = filename
        return {
            'type': 'ir.actions.act_url',
            'url': 'web/content/?model=partner.opening.closing.balance&download=true&field=balance_file&id={}&filename={}'.format(
                self.id, self.balance_name
            ),
            'target': 'new',
        }