# Copyright 2020 Netkia (https://netkia.es)
# License AGPL-3.0 or later (https://www.gnu.org/licenses/agpl).

from odoo import fields, models


class HrEmployee(models.Model):
    _inherit = "hr.employee"

    # Analytic tag applied to this employee's payroll cost distribution.
    account_analytic_tag_id = fields.Many2one("account.analytic.tag")

    # Unique employee code used to match rows in the payroll Excel import.
    payroll_code = fields.Char(size=9)

    _sql_constraints = [
        (
            "payroll_code_unique",
            "unique(payroll_code)",
            "Payroll code already exists!",
        )
    ]
