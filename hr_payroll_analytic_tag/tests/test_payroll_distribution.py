# Copyright 2020 Netkia (https://netkia.es)
# License AGPL-3.0 or later (https://www.gnu.org/licenses/agpl).

import base64
import io

import xlwt
from psycopg2 import IntegrityError

from odoo.exceptions import ValidationError
from odoo.tests import tagged
from odoo.tests.common import TransactionCase


@tagged("post_install", "-at_install")
class TestHrPayrollAnalyticTag(TransactionCase):
    """Tests for the HR Payroll Analytic Tag module."""

    @classmethod
    def setUpClass(cls):
        super().setUpClass()
        cls.env = cls.env(
            context=dict(
                cls.env.context,
                mail_create_nolog=True,
                tracking_disable=True,
            )
        )
        plan = cls.env["account.analytic.plan"].create({"name": "Payroll Test Plan"})
        cls.account_1 = cls.env["account.analytic.account"].create(
            {"name": "Account 1", "plan_id": plan.id}
        )
        cls.account_2 = cls.env["account.analytic.account"].create(
            {"name": "Account 2", "plan_id": plan.id}
        )

        # Global distribution tag — its percentages are updated by the wizard.
        cls.config_tag = cls.env["account.analytic.tag"].create(
            {
                "name": "Global Payroll Tag",
                "active_analytic_distribution": True,
                "analytic_distribution": {
                    str(cls.account_1.id): 50.0,
                    str(cls.account_2.id): 50.0,
                },
            }
        )

        # Employee tag with two analytic accounts.
        cls.tag_multi = cls.env["account.analytic.tag"].create(
            {
                "name": "Employee Multi-Account Tag",
                "active_analytic_distribution": True,
                "analytic_distribution": {
                    str(cls.account_1.id): 60.0,
                    str(cls.account_2.id): 40.0,
                },
            }
        )

        # Employee tag with a single analytic account.
        cls.tag_single = cls.env["account.analytic.tag"].create(
            {
                "name": "Employee Single-Account Tag",
                "active_analytic_distribution": True,
                "analytic_distribution": {
                    str(cls.account_1.id): 100.0,
                },
            }
        )

        cls.employee_a = cls.env["hr.employee"].create(
            {
                "name": "Employee A",
                "payroll_code": "001",
                "account_analytic_tag_id": cls.tag_multi.id,
            }
        )
        cls.employee_b = cls.env["hr.employee"].create(
            {
                "name": "Employee B",
                "payroll_code": "002",
                "account_analytic_tag_id": cls.tag_multi.id,
            }
        )
        cls.employee_single = cls.env["hr.employee"].create(
            {
                "name": "Employee Single",
                "payroll_code": "003",
                "account_analytic_tag_id": cls.tag_single.id,
            }
        )

        cls.env["ir.config_parameter"].sudo().set_param(
            "hr_payroll_analytic_tag.employee_account_analytic_tag_id",
            str(cls.config_tag.id),
        )

    @staticmethod
    def _build_xls(data_rows, header=None):
        """Build a base64-encoded .xls with headers at row 3, data from row 4.

        An empty-code totals row is always appended so that:
        - _compute_distribution skips it (it is the last row).
        - check_data skips it (empty employee code).

        Args:
            data_rows: list of (code, name, cost) tuples.
            header: column names list; defaults to the expected format.
        """
        if header is None:
            header = ["", "EMPLOYEE NAME", "COMPANY COST"]
        wb = xlwt.Workbook()
        ws = wb.add_sheet("Sheet1")
        for col, val in enumerate(header):
            ws.write(3, col, val)
        for row_offset, (code, name, cost) in enumerate(data_rows):
            ws.write(4 + row_offset, 0, code)
            ws.write(4 + row_offset, 1, name)
            ws.write(4 + row_offset, 2, cost)
        totals_idx = len(data_rows)
        ws.write(4 + totals_idx, 0, "")
        ws.write(4 + totals_idx, 1, "Total")
        ws.write(4 + totals_idx, 2, sum(r[2] for r in data_rows))
        fp = io.BytesIO()
        wb.save(fp)
        fp.seek(0)
        return base64.b64encode(fp.read()).decode()

    def _create_wizard(self, xls_b64):
        """Return a wizard instance loaded with the given base64 XLS."""
        return self.env["payroll.distribution.process.wiz"].create(
            {"upload_file": xls_b64, "file_name": "test.xls"}
        )

    # ── hr.employee model ─────────────────────────────────────────────────────

    def test_employee_payroll_code_unique_constraint(self):
        """Duplicate payroll codes must be rejected by the DB constraint."""
        with self.assertRaises(IntegrityError), self.env.cr.savepoint():
            self.env["hr.employee"].create({"name": "Duplicate", "payroll_code": "001"})

    # ── res.config.settings ───────────────────────────────────────────────────

    def test_config_settings_get_values_no_param(self):
        """get_values returns a valid dict even when no param is stored."""
        self.env["ir.config_parameter"].sudo().set_param(
            "hr_payroll_analytic_tag.employee_account_analytic_tag_id", False
        )
        settings = self.env["res.config.settings"].create({})
        res = settings.get_values()
        self.assertNotIn("employee_account_analytic_tag_id", res)

    def test_config_settings_set_and_get_values(self):
        """set_values persists the tag; get_values reads it back correctly."""
        settings = self.env["res.config.settings"].create({})
        settings.employee_account_analytic_tag_id = self.config_tag
        settings.set_values()
        res = settings.get_values()
        self.assertEqual(res["employee_account_analytic_tag_id"], self.config_tag.id)

    def test_config_column_headers_set_and_get(self):
        """set_values persists column header names; get_values reads them back."""
        settings = self.env["res.config.settings"].create({})
        settings.payroll_col_code = "CODIGO"
        settings.payroll_col_name = "TRABAJADOR"
        settings.payroll_col_cost = "COSTE EMPRESA"
        settings.set_values()
        res = settings.get_values()
        self.assertEqual(res["payroll_col_code"], "CODIGO")
        self.assertEqual(res["payroll_col_name"], "TRABAJADOR")
        self.assertEqual(res["payroll_col_cost"], "COSTE EMPRESA")

    # ── _get_config_tag ───────────────────────────────────────────────────────

    def test_get_config_tag_no_param_returns_false(self):
        """_get_config_tag returns False when the config parameter is absent."""
        self.env["ir.config_parameter"].sudo().set_param(
            "hr_payroll_analytic_tag.employee_account_analytic_tag_id", False
        )
        wiz = self._create_wizard(self._build_xls([("001", "A", 100)]))
        self.assertFalse(wiz._get_config_tag())

    # ── check_data ────────────────────────────────────────────────────────────

    def test_check_data_no_file_raises(self):
        """check_data raises ValidationError when no file has been uploaded."""
        wiz = self.env["payroll.distribution.process.wiz"].new({})
        with self.assertRaises(ValidationError):
            wiz.check_data()

    def test_check_data_header_not_found_raises(self):
        """check_data raises ValidationError when a required header is not found.

        The Excel may have many columns; what matters is that the configured
        header names exist somewhere in row 4, regardless of their position.
        """
        xls = self._build_xls(
            [("001", "Employee A", 1000)],
            header=["COL_X", "COL_Y", "COL_Z"],
        )
        with self.assertRaises(ValidationError):
            self._create_wizard(xls).check_data()

    def test_check_data_columns_not_first_position(self):
        """check_data succeeds when required columns are not in the first 3 positions.

        Simulates a real A3 export where the file has many columns and the
        relevant ones appear after several informational columns.
        """
        # Build an XLS where cols 0-2 are irrelevant and the required headers
        # start at col 3.
        wb = xlwt.Workbook()
        ws = wb.add_sheet("Sheet1")
        for col, val in enumerate(
            ["EXTRA1", "EXTRA2", "EXTRA3", "CODE", "EMPLOYEE NAME", "COMPANY COST"]
        ):
            ws.write(3, col, val)
        ws.write(4, 0, "")
        ws.write(4, 1, "")
        ws.write(4, 2, "")
        ws.write(4, 3, "001")
        ws.write(4, 4, "Employee A")
        ws.write(4, 5, 1000.0)
        # Totals row (skipped automatically).
        ws.write(5, 3, "")
        ws.write(5, 4, "Total")
        ws.write(5, 5, 1000.0)
        fp = io.BytesIO()
        wb.save(fp)
        fp.seek(0)
        xls = base64.b64encode(fp.read()).decode()
        result = self._create_wizard(xls).check_data()
        self.assertEqual(result["params"]["type"], "success")

    def test_check_data_empty_string_code_header(self):
        """check_data succeeds when the code column header is an empty string.

        The default value for col_header_code is "" to match the A3 payroll
        export format, which uses an empty header for the employee code column.
        """
        # The default is already ""; no need to set the parameter explicitly.
        xls = self._build_xls(
            [("001", "Employee A", 1000)],
            header=["", "EMPLOYEE NAME", "COMPANY COST"],
        )
        result = self._create_wizard(xls).check_data()
        self.assertEqual(result["params"]["type"], "success")

    def test_check_data_no_config_tag_raises(self):
        """check_data raises when the global config tag is not configured."""
        self.env["ir.config_parameter"].sudo().set_param(
            "hr_payroll_analytic_tag.employee_account_analytic_tag_id", False
        )
        with self.assertRaises(ValidationError):
            self._create_wizard(
                self._build_xls([("001", "Employee A", 1000)])
            ).check_data()

    def test_check_data_inactive_distribution_raises(self):
        """check_data raises when the config tag has no active distribution."""
        self.config_tag.active_analytic_distribution = False
        with self.assertRaises(ValidationError):
            self._create_wizard(
                self._build_xls([("001", "Employee A", 1000)])
            ).check_data()

    def test_check_data_employee_not_found_raises(self):
        """check_data raises when an employee code has no matching record."""
        with self.assertRaises(ValidationError):
            self._create_wizard(
                self._build_xls([("999", "Unknown", 1000)])
            ).check_data()

    def test_check_data_employee_without_tag_raises(self):
        """check_data raises when an employee has no analytic tag assigned."""
        untagged = self.env["hr.employee"].create(
            {"name": "Untagged Employee", "payroll_code": "099"}
        )
        with self.assertRaises(ValidationError):
            self._create_wizard(
                self._build_xls([("099", "Untagged", 1000)])
            ).check_data()
        untagged.unlink()

    def test_check_data_accounts_mismatch_raises(self):
        """check_data raises when employee tag accounts are not in the config tag."""
        extra_plan = self.env["account.analytic.plan"].create({"name": "Extra Plan"})
        extra_account = self.env["account.analytic.account"].create(
            {"name": "Extra Account", "plan_id": extra_plan.id}
        )
        bad_tag = self.env["account.analytic.tag"].create(
            {
                "name": "Bad Tag",
                "active_analytic_distribution": True,
                "analytic_distribution": {str(extra_account.id): 100.0},
            }
        )
        bad_employee = self.env["hr.employee"].create(
            {
                "name": "Bad Employee",
                "payroll_code": "098",
                "account_analytic_tag_id": bad_tag.id,
            }
        )
        with self.assertRaises(ValidationError):
            self._create_wizard(
                self._build_xls([("098", "Bad Employee", 1000)])
            ).check_data()
        bad_employee.unlink()
        bad_tag.unlink()

    def test_check_data_success_returns_notification(self):
        """check_data returns a display_notification action on valid input."""
        xls = self._build_xls([("001", "Employee A", 1000)])
        result = self._create_wizard(xls).check_data()
        self.assertEqual(result["type"], "ir.actions.client")
        self.assertEqual(result["tag"], "display_notification")
        self.assertEqual(result["params"]["type"], "success")

    # ── update_distribution ───────────────────────────────────────────────────

    def test_update_distribution_computes_correctly(self):
        """update_distribution applies the weighted distribution and returns act_window.

        Both employees share tag_multi (A1=60%, A2=40%).
        Expected: A1=60%, A2=40% regardless of individual costs.
        """
        xls = self._build_xls(
            [("001", "Employee A", 1000), ("002", "Employee B", 2000)]
        )
        result = self._create_wizard(xls).update_distribution()
        self.assertEqual(result["type"], "ir.actions.act_window")
        self.assertEqual(result["res_model"], "account.analytic.tag")
        new_dist = self.config_tag.analytic_distribution
        self.assertAlmostEqual(new_dist[str(self.account_1.id)], 60.0, places=1)
        self.assertAlmostEqual(new_dist[str(self.account_2.id)], 40.0, places=1)

    def test_update_distribution_sum_always_100(self):
        """_compute_distribution always produces percentages summing to exactly 100.

        Uses a 3-account setup with non-terminating decimal thirds to exercise
        the rounding correction branch (percentage_sum < 100).
        """
        plan = self.env["account.analytic.plan"].create({"name": "Rounding Plan"})
        account_3 = self.env["account.analytic.account"].create(
            {"name": "Account 3", "plan_id": plan.id}
        )
        global_3 = self.env["account.analytic.tag"].create(
            {
                "name": "3-Account Global Tag",
                "active_analytic_distribution": True,
                "analytic_distribution": {
                    str(self.account_1.id): 33.33,
                    str(self.account_2.id): 33.33,
                    str(account_3.id): 33.34,
                },
            }
        )
        # Non-terminating thirds force a rounding gap after weighted aggregation.
        emp_3_tag = self.env["account.analytic.tag"].create(
            {
                "name": "3-Account Employee Tag",
                "active_analytic_distribution": True,
                "analytic_distribution": {
                    str(self.account_1.id): 33.333333,
                    str(self.account_2.id): 33.333333,
                    str(account_3.id): 33.333334,
                },
            }
        )
        emp_3a = self.env["hr.employee"].create(
            {
                "name": "Emp 3A",
                "payroll_code": "091",
                "account_analytic_tag_id": emp_3_tag.id,
            }
        )
        emp_3b = self.env["hr.employee"].create(
            {
                "name": "Emp 3B",
                "payroll_code": "092",
                "account_analytic_tag_id": emp_3_tag.id,
            }
        )
        self.env["ir.config_parameter"].sudo().set_param(
            "hr_payroll_analytic_tag.employee_account_analytic_tag_id",
            str(global_3.id),
        )
        xls = self._build_xls([("091", "Emp 3A", 1000), ("092", "Emp 3B", 2000)])
        self._create_wizard(xls).update_distribution()
        total_pct = sum(global_3.analytic_distribution.values())
        self.assertAlmostEqual(total_pct, 100.0, places=1)
        emp_3a.unlink()
        emp_3b.unlink()

    # ── update_distribution_detail ────────────────────────────────────────────

    def test_update_distribution_detail_multi_account(self):
        """update_distribution_detail generates a downloadable XLS for
        multi-account employees and returns a file download action."""
        xls = self._build_xls(
            [("001", "Employee A", 1000), ("002", "Employee B", 2000)]
        )
        result = self._create_wizard(xls).update_distribution_detail()
        self.assertEqual(result["type"], "ir.actions.act_url")
        self.assertIn("download=true", result["url"])

    def test_update_distribution_detail_single_account(self):
        """update_distribution_detail handles single-account employees correctly."""
        single_config = self.env["account.analytic.tag"].create(
            {
                "name": "Single Config Tag",
                "active_analytic_distribution": True,
                "analytic_distribution": {str(self.account_1.id): 100.0},
            }
        )
        self.env["ir.config_parameter"].sudo().set_param(
            "hr_payroll_analytic_tag.employee_account_analytic_tag_id",
            str(single_config.id),
        )
        xls = self._build_xls([("003", "Employee Single", 500)])
        result = self._create_wizard(xls).update_distribution_detail()
        self.assertEqual(result["type"], "ir.actions.act_url")

    def test_update_distribution_detail_zero_cost(self):
        """write_employee_row does not raise ZeroDivisionError when cost is 0.

        Covers the `else 0` branch in the percentage calculation for both
        single-account employees (tag_single) and — via the totals reset —
        the multi-account path when employee_total_cost is falsy.
        """
        single_config = self.env["account.analytic.tag"].create(
            {
                "name": "Zero Cost Config Tag",
                "active_analytic_distribution": True,
                "analytic_distribution": {str(self.account_1.id): 100.0},
            }
        )
        self.env["ir.config_parameter"].sudo().set_param(
            "hr_payroll_analytic_tag.employee_account_analytic_tag_id",
            str(single_config.id),
        )
        xls = self._build_xls([("003", "Employee Single", 0)])
        result = self._create_wizard(xls).update_distribution_detail()
        self.assertEqual(result["type"], "ir.actions.act_url")
