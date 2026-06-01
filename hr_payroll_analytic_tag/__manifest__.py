# Copyright 2020 Netkia (https://netkia.es)
# License AGPL-3.0 or later (https://www.gnu.org/licenses/agpl).

{
    "name": "HR Payroll Analytic Tag",
    "summary": """
        Distribute payroll costs across analytic accounts using employee
        analytic tags and an Excel import.
    """,
    "author": "Netkia, Odoo Community Association (OCA)",
    "license": "AGPL-3",
    "website": "https://github.com/OCA/account-analytic",
    "category": "Accounting & Finance",
    "version": "18.0.1.0.0",
    "depends": [
        "hr",
        "account_analytic_tag",
    ],
    "data": [
        "data/ir_config_parameter_data.xml",
        "security/ir.model.access.csv",
        "views/res_config_settings_views.xml",
        "views/hr_employee_views.xml",
        "wizard/payroll_distribution_process_wiz_views.xml",
    ],
    "installable": True,
    "application": False,
    "maintainers": ["carlossainznetkia", "hugomartineznetkia"],
}
