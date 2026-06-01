# Copyright 2020 Netkia (https://netkia.es)
# License AGPL-3.0 or later (https://www.gnu.org/licenses/agpl).

import base64
import io
import logging

import xlrd
import xlwt

from odoo import _, fields, models
from odoo.exceptions import ValidationError

_logger = logging.getLogger(__name__)


class PayrollDistributionProcessWiz(models.TransientModel):
    _name = "payroll.distribution.process.wiz"
    _description = "Payroll Distribution Process"

    upload_file = fields.Binary(required=True)
    file_name = fields.Char()

    def _parse_excel(self):
        """Parse the uploaded .xls file by locating columns through their header names.

        Reads all headers from row index 3, finds the position of each configured
        column name, and extracts data rows from index 4 onwards.
        Returns a list of dicts with fixed keys: "code", "name", "cost".
        Raises ValidationError if any required header is not found.
        """
        workbook = xlrd.open_workbook(
            filename="", file_contents=base64.b64decode(self.upload_file)
        )
        worksheet = workbook.sheet_by_index(0)
        header_row = [worksheet.cell_value(3, col) for col in range(worksheet.ncols)]

        expected_code, expected_name, expected_cost = self._get_col_headers()

        def _col_has_data(col):
            # True if the column holds at least one non-empty value in the data
            # rows (row index 4 onwards).
            return any(
                worksheet.cell_value(row, col) not in ("", None)
                for row in range(4, worksheet.nrows)
            )

        def _find_col(header_name):
            if header_name == "":
                # A3 exports use an empty header for the employee code column,
                # but they also include blank spacer columns whose header is
                # empty too. Pick the empty-header column that actually has data
                # below it, so the spacer column is not mistaken for the code.
                empty_cols = [c for c, v in enumerate(header_row) if v == ""]
                for col in empty_cols:
                    if _col_has_data(col):
                        return col
                if empty_cols:
                    return empty_cols[0]
                raise ValidationError(
                    _("Required column %s not found in the header row (row 4).")
                    % repr(header_name)
                )
            try:
                return header_row.index(header_name)
            except ValueError:
                label = repr(header_name)  # repr shows '' for empty strings
                raise ValidationError(
                    _("Required column %s not found in the header row (row 4).") % label
                ) from None

        idx_code = _find_col(expected_code)
        idx_name = _find_col(expected_name)
        idx_cost = _find_col(expected_cost)

        excel_rows = [
            {
                "code": worksheet.cell_value(row, idx_code),
                "name": worksheet.cell_value(row, idx_name),
                "cost": worksheet.cell_value(row, idx_cost),
            }
            for row in range(4, worksheet.nrows)
        ]
        return excel_rows

    def _get_config_tag(self):
        """Return the global payroll distribution tag from settings, or False."""
        tag_id = (
            self.env["ir.config_parameter"]
            .sudo()
            .get_param("hr_payroll_analytic_tag.employee_account_analytic_tag_id")
        )
        if not tag_id:
            return False
        return self.env["account.analytic.tag"].browse(int(tag_id))

    def _get_col_headers(self):
        """Return the configured expected column header names as a
        (code, name, cost) tuple.

        Falls back to English defaults when the parameters have never been set.
        An empty string ("") is a valid header name — the original A3 payroll
        export uses an empty header for the employee code column.
        """
        param = self.env["ir.config_parameter"].sudo()
        return (
            param.get_param("hr_payroll_analytic_tag.col_header_code", ""),
            param.get_param("hr_payroll_analytic_tag.col_header_name", "EMPLOYEE NAME"),
            param.get_param("hr_payroll_analytic_tag.col_header_cost", "COMPANY COST"),
        )

    def check_data(self):
        """Validate the uploaded Excel without applying any changes.

        Verifies column structure, employee codes, analytic tag assignments,
        and consistency between employee and global distribution tag accounts.
        Returns a success notification action; raises ValidationError on failure.
        """
        employee_obj = self.env["hr.employee"]

        if not self.upload_file:
            raise ValidationError(_("You must upload a file."))

        excel_rows = self._parse_excel()

        config_tag = self._get_config_tag()
        if not config_tag or not config_tag.active_analytic_distribution:
            raise ValidationError(
                _(
                    "The payroll distribution tag is not configured. "
                    "Please set it in Configuration."
                )
            )

        config_account_ids = list(config_tag.analytic_distribution.keys())

        for row in excel_rows:
            employee_code = row.get("code")
            if not employee_code:
                continue
            employee = employee_obj.search([("payroll_code", "=", employee_code)])
            if not employee:
                raise ValidationError(
                    _("No employee found for code %s.") % str(employee_code)
                )
            if not employee.account_analytic_tag_id:
                raise ValidationError(
                    _("Employee %s does not have a payroll analytic tag assigned.")
                    % employee.name
                )
            for (
                account_id
            ) in employee.account_analytic_tag_id.analytic_distribution.keys():
                if account_id not in config_account_ids:
                    raise ValidationError(
                        _(
                            "The analytic accounts of the employees do not match "
                            "those defined in the global payroll distribution tag."
                        )
                    )

        return {
            "type": "ir.actions.client",
            "tag": "display_notification",
            "params": {
                "type": "success",
                "message": _("File format is valid."),
                "sticky": False,
            },
        }

    def _compute_distribution(self, excel_rows):
        """Compute the weighted analytic distribution and apply it to the global tag.

        For each data row, weights each analytic account percentage by the employee's
        company cost. Percentages are normalised to sum exactly to 100, with any
        rounding error absorbed by the account with the max or min percentage.
        Returns the updated config tag record.
        """
        employee_obj = self.env["hr.employee"]
        config_tag = self._get_config_tag()

        total_cost = 0.0
        weighted_amounts = {}  # {account_id_str: cumulative weighted amount}

        for idx, row in enumerate(excel_rows):
            # The A3 export ends with a totals row whose employee code is empty.
            # Skip any row without a code (totals row or blank line); do not skip
            # by position, since files may or may not include a totals row.
            code = row.get("code")
            if not code:
                _logger.debug("Row %d has no payroll code — skipping", idx)
                continue
            employee = employee_obj.search([("payroll_code", "=", code)], limit=1)
            employee_cost = row.get("cost") or 0.0
            if not employee:
                _logger.debug(
                    "No employee found for payroll code %r — skipping row",
                    code,
                )
                continue
            distribution = employee.account_analytic_tag_id.analytic_distribution
            if not isinstance(distribution, dict):
                _logger.debug(
                    "Employee %s (code %r) has no analytic distribution — skipping",
                    employee.name,
                    code,
                )
                continue
            for (
                account_id,
                percentage,
            ) in distribution.items():
                current = weighted_amounts.get(account_id, 0.0)
                if percentage and employee_cost:
                    current += (employee_cost * percentage) / 100
                weighted_amounts[account_id] = current
            total_cost += employee_cost

        if not total_cost:
            raise ValidationError(
                _(
                    "No se han encontrado distribuciones analíticas de empleados. "
                    "Comprueba que los empleados del fichero tienen una etiqueta "
                    "analítica con distribución configurada y que sus códigos de "
                    "nómina coinciden con los del Excel."
                )
            )

        # Convert accumulated weighted amounts to percentages of the total cost.
        new_distribution = config_tag.analytic_distribution.copy()
        for account_id in new_distribution:
            pct = 0.0
            amount = weighted_amounts.get(account_id, 0.0)
            if amount and total_cost:
                pct = (amount * 100) / total_cost
            new_distribution[account_id] = round(pct, 2)

        _logger.debug(
            "New distribution before rounding adjustment: %s (total: %s)",
            new_distribution,
            sum(new_distribution.values()),
        )
        # Correct floating-point rounding so percentages sum exactly to 100.
        # Use sum() directly (not accumulated) to avoid compounding float errors.
        active = {k: v for k, v in new_distribution.items() if v > 0}
        if active:
            diff = round(100.0 - sum(new_distribution.values()), 2)
            if diff > 0:
                min_id = min(active, key=active.get)
                new_distribution[min_id] = round(new_distribution[min_id] + diff, 2)
            elif diff < 0:
                max_id = max(active, key=active.get)
                new_distribution[max_id] = round(new_distribution[max_id] + diff, 2)
        _logger.debug(
            "New distribution after rounding adjustment: %s (total: %s)",
            new_distribution,
            sum(new_distribution.values()),
        )
        config_tag.write({"analytic_distribution": new_distribution})
        return config_tag

    def update_distribution(self):
        """Validate, compute and apply the new analytic distribution.

        Returns an action that opens the updated analytic tag form.
        """
        self.check_data()
        excel_rows = self._parse_excel()
        config_tag = self._compute_distribution(excel_rows)
        return {
            "name": _("Account Analytic Tag"),
            "type": "ir.actions.act_window",
            "view_mode": "form",
            "res_model": "account.analytic.tag",
            "res_id": config_tag.id,
            "target": "current",
        }

    def update_distribution_detail(self):
        """Apply the distribution update and generate a detailed XLS breakdown report.

        Produces one row per employee per analytic account, showing the tag name,
        employee code, name, ID number, distributed amount and percentage.
        Returns a file download action for the generated report.
        """
        self.check_data()
        excel_rows = self._parse_excel()
        self._compute_distribution(excel_rows)

        employee_obj = self.env["hr.employee"]
        workbook = xlwt.Workbook({"in_memory": True})
        report_sheet = workbook.add_sheet(
            "Payroll Distribution Detail", cell_overwrite_ok=True
        )

        report_sheet.write_merge(0, 1, 0, 5, "Payroll Distribution Report")

        row_idx = 3
        report_sheet.write(row_idx, 0, "Tag")
        report_sheet.write(row_idx, 1, "Code")
        report_sheet.write(row_idx, 2, "Name")
        report_sheet.write(row_idx, 3, "ID Number")
        report_sheet.write(row_idx, 4, "Amount")
        report_sheet.write(row_idx, 5, "Percentage")
        row_idx += 1

        total_cost = 0.0
        for line in excel_rows:
            # Skip the totals row (empty code) and any blank line.
            code = line.get("code")
            if not code:
                continue
            employee = employee_obj.search([("payroll_code", "=", code)])
            row_idx = self.write_employee_row(row_idx, report_sheet, employee, line)
            total_cost += line.get("cost") or 0

        report_sheet.write_merge(
            row_idx + 2,
            row_idx + 2,
            3,
            5,
            "Total Company Cost: " + str(round(total_cost, 2)),
        )

        fp = io.BytesIO()
        workbook.save(fp)
        fp.seek(0)
        report_data = base64.b64encode(fp.read())
        fp.close()

        attachment = self.env["ir.attachment"].create(
            {"name": "distribution_summary.xls", "datas": report_data}
        )
        # Return a relative URL so the download stays on the same origin the user
        # is browsing. Prefixing web.base.url breaks the download when that
        # parameter does not match the host in use (common behind Docker/proxy).
        return {
            "type": "ir.actions.act_url",
            "url": "/web/content/%d?download=true" % attachment.id,
            "target": "self",
        }

    def write_employee_row(self, row_idx, report_sheet, employee, line):
        """Write XLS row(s) for a single employee's analytic distribution.

        Writes one row per analytic account for employees with multiple accounts,
        or a single consolidated row for single-account employees.
        Returns the updated row index.
        """
        employee_total_cost = round(line.get("cost") or 0, 3)
        distribution = employee.account_analytic_tag_id.analytic_distribution
        if not isinstance(distribution, dict):
            distribution = {}

        if len(distribution) > 1:
            for account_id_str, percentage in distribution.items():
                account = self.env["account.analytic.account"].browse(
                    int(account_id_str)
                )
                distributed_amount = round((employee_total_cost / 100) * percentage, 2)
                row_idx += 1
                report_sheet.write(row_idx, 0, account.display_name or "")
                report_sheet.write(row_idx, 1, line.get("code", ""))
                report_sheet.write(row_idx, 2, employee.name or "")
                report_sheet.write(row_idx, 3, employee.identification_id or "")
                report_sheet.write(row_idx, 4, distributed_amount)
                pct = (
                    round((distributed_amount * 100) / employee_total_cost, 3)
                    if employee_total_cost
                    else 0
                )
                report_sheet.write(row_idx, 5, pct)
        else:
            first_account_id = next(iter(distribution), None)
            display_name = ""
            if first_account_id:
                account = self.env["account.analytic.account"].browse(
                    int(first_account_id)
                )
                display_name = account.display_name or ""
            row_idx += 1
            report_sheet.write(row_idx, 0, display_name)
            report_sheet.write(row_idx, 1, line.get("code", ""))
            report_sheet.write(row_idx, 2, employee.name or "")
            report_sheet.write(row_idx, 3, employee.identification_id or "")
            report_sheet.write(row_idx, 4, employee_total_cost)
            # Single account: 100% of cost goes to this account.
            pct = (
                round((employee_total_cost * 100) / employee_total_cost, 3)
                if employee_total_cost
                else 0
            )
            report_sheet.write(row_idx, 5, pct)

        return row_idx
