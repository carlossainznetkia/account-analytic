# Copyright 2020 Netkia (https://netkia.es)
# License AGPL-3.0 or later (https://www.gnu.org/licenses/agpl).

from odoo import api, fields, models


class ResConfigSettings(models.TransientModel):
    _inherit = "res.config.settings"

    # Global payroll distribution tag; its analytic_distribution is updated
    # by the wizard.
    employee_account_analytic_tag_id = fields.Many2one("account.analytic.tag")

    # Expected column header names in the payroll Excel file (cols A, B, C).
    payroll_col_code = fields.Char(default="")
    payroll_col_name = fields.Char(default="EMPLOYEE NAME")
    payroll_col_cost = fields.Char(default="COMPANY COST")

    @api.model
    def get_values(self):
        res = super().get_values()
        param = self.env["ir.config_parameter"].sudo()
        tag = param.get_param(
            "hr_payroll_analytic_tag.employee_account_analytic_tag_id"
        )
        if tag:
            res.update({"employee_account_analytic_tag_id": int(tag)})
        res.update(
            {
                "payroll_col_code": param.get_param(
                    "hr_payroll_analytic_tag.col_header_code", ""
                ),
                "payroll_col_name": param.get_param(
                    "hr_payroll_analytic_tag.col_header_name", "EMPLOYEE NAME"
                ),
                "payroll_col_cost": param.get_param(
                    "hr_payroll_analytic_tag.col_header_cost", "COMPANY COST"
                ),
            }
        )
        return res

    def set_values(self):
        res = super().set_values()
        param = self.env["ir.config_parameter"].sudo()
        tag_id = (
            self.employee_account_analytic_tag_id.id
            if self.employee_account_analytic_tag_id
            else False
        )
        param.set_param(
            "hr_payroll_analytic_tag.employee_account_analytic_tag_id",
            tag_id,
        )
        param.set_param(
            "hr_payroll_analytic_tag.col_header_code",
            self.payroll_col_code if self.payroll_col_code is not False else "",
        )
        param.set_param(
            "hr_payroll_analytic_tag.col_header_name",
            self.payroll_col_name if self.payroll_col_name is not False else "",
        )
        param.set_param(
            "hr_payroll_analytic_tag.col_header_cost",
            self.payroll_col_cost if self.payroll_col_cost is not False else "",
        )
        return res
