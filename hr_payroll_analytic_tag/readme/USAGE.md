Go to **Employees \> Payroll Distribution \> Payroll Distribution
Process**.

Upload an Excel file (**\`.xls\`** format) with the following structure:

- **Row 4 (index 3)**: header row. The wizard searches this row for the
  configured column names (order does not matter; other columns are ignored):
  - `CODE` (default): employee payroll code.
  - `EMPLOYEE NAME` (default): employee name (informational, not used for matching).
  - `COMPANY COST` (default): company cost for that employee in this payroll
    period.
- **Rows 5 onwards**: one data row per employee. The last row is treated
  as a totals row and is skipped automatically.

Actions available:

- **Check Data**: validates the file without making any changes. Shows a
  success notification if the file is consistent.
- **Update Distribution**: validates the file and recalculates the
  distribution percentages of the global payroll tag, then opens the
  updated tag record.
- **Update Distribution and Get Detail**: same as above, and
  additionally generates a downloadable `.xls` report with one row per
  employee per analytic account.
