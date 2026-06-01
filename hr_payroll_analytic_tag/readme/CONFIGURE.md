Go to **Configuration \> Settings \> Employees** and set the **Payroll
Distribution Tag** field to the analytic tag whose distribution
percentages should be updated by the wizard.

In the same Settings screen, you can also configure the expected column
headers for the payroll Excel file:

- **Payroll Column: Code** (default: `CODE`): header of the employee
  code column (col A).
- **Payroll Column: Employee Name** (default: `EMPLOYEE NAME`): header
  of the employee name column (col B).
- **Payroll Column: Company Cost** (default: `COMPANY COST`): header
  of the company cost column (col C).

These defaults match the English template. Change them to match your
payroll software's output language.

On each employee record, go to the **HR Settings** tab and fill in:

- **Payroll Code**: unique identifier matching the code column in the
  Excel file.
- **Account Analytic Tag**: the tag holding this employee's individual
  analytic distribution (analytic account percentages must be a subset
  of the global tag).
