# -*- coding: utf-8 -*-
##############################################################################
#
#    ODOO Open Source Management Solution
#
#    ODOO Addon module by Uncanny Consulting Services LLP
#    Copyright (C) 2023 Uncanny Consulting Services LLP (<https://uncannycs.com>).
#
##############################################################################
{
    "name": "Partner Opening Closing Balance Report Ucs",
    "summary": "The Partner Opening Closing Balance Report module provides a comprehensive Excel-based Trial Balance report for partners (Customers or Vendors) within the Accounting module.",
    "version": "19.0.1.0.0",
    "category": "Accounting",
    "description": """ Key Features:
- Generate a Partner Trial Balance report in Excel (.xlsx) format
- Filter by Receivable Accounts (Customers) or Payable Accounts (Vendors)
- Select a custom date range (Start Date to End Date)
- Optionally filter by specific partners or run for all partners at once
- Displays Opening Balance (net balance before the start date), Debit, 
  Credit, and Closing Balance for each partner
- Option to include or exclude partners with zero balances
- Automatically scoped to the current company
- Report is accessible from Accounting → Reporting → Trial Balance Report""",
    "website": "https://uncannycs.com",
    "author": "Uncanny Consulting Services LLP",
    "maintainers": "Uncanny Consulting Services LLP",
    "license": "Other proprietary",
    "external_dependencies": {
        "python": ["xlsxwriter"],
    },
    "depends": ['base', 'account'],
    "data": [
        "security/ir.model.access.csv",
        "wizard/partner_opening_closing_balance.xml",
    ],
    "application": False,
    "auto_install": False,
    "installable": True,
}
