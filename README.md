# Excel VBA Automation

Automated timesheet & payroll builder, operations dashboard, and dropdown automation macros for Excel — built for HR and operations teams that still do payroll by hand.

## The problem

Small HR and ops teams waste hours every month on the same manual loop: collect attendance, tally regular and overtime hours in a spreadsheet, apply hourly rates, look up tax slabs, then rebuild a summary table and a dashboard chart from scratch. It is slow, error-prone, and every formula edit risks breaking the whole sheet.

This repo replaces that loop with three importable VBA modules:

- **Timesheet & Payroll** — generates a monthly timesheet skeleton from your employee roster, then calculates regular/overtime hours, gross pay, slab-based tax, and net pay into a formatted payroll summary.
- **Operations Dashboard** — rebuilds a KPI dashboard from raw data rows: KPI tiles, a Top-5 table with colour-scale formatting, and auto-generated column + trend charts.
- **Dropdown Automation** — builds dependent (cascading) dropdowns from a simple Lists sheet via named ranges + data validation, with reset/refresh helpers.

## Features

- `GenerateMonthlyTimesheet` — one row per employee per weekday of any month, pre-filled at 8 hours; just correct the actuals.
- `CalculatePayroll` — aggregates hours per employee, splits regular (≤8/day) vs overtime (>8/day, 1.5x), computes gross pay, monthly tax (annualized slab schedule), net pay, and a totals row with borders and number formatting.
- `MonthlyTax` / `AnnualTax` — reusable tax-slab functions, also usable as worksheet formulas (e.g. `=MonthlyTax(50000)`).
- `RefreshDashboard` — KPI tiles (Total, Average, Peak, Records), Top-5 by value with colour-scale conditional formatting, category column chart and date trend chart — all rebuilt on every run.
- `BuildCascadingDropdowns` — master + dependent dropdowns from a plain Lists sheet; `ResetForm` clears the form, `RefreshLists` rebuilds named ranges after you edit the lists.
- Python companion: `python/build_demo_workbook.py` (openpyxl) generates `demo_timesheet.xlsx` with realistic sample data so the macros have something to run on immediately.
- `tests/test_workbook.py` — pytest suite covering workbook structure and the payroll math (zero hours, the 8-hour overtime threshold, tax slab boundaries).

## Quickstart

1. Open the demo workbook `demo_timesheet.xlsx` (or run `python python/build_demo_workbook.py` to regenerate it).
2. Press **Alt+F11** to open the VBA editor.
3. Right-click your workbook in the Project Explorer → **Insert → Module**, or simply **File → Import File** and select a `.bas` file from `vba/`.
4. Run a macro with **Alt+F8**:
   - `TimesheetPayroll.GenerateMonthlyTimesheet` → fill in hours → `TimesheetPayroll.CalculatePayroll`
   - `OperationsDashboard.RefreshDashboard`
   - `DropdownAutomation.BuildCascadingDropdowns`
5. Save the workbook as `.xlsm` (macro-enabled) to keep the modules.

## Worksheet layout

```
Employees   A: Emp ID   B: Name   C: Department   D: Role   E: Hourly Rate
Timesheet   A: Emp ID   B: Name   C: Date         D: Day    E: Hours Worked
Payroll     A: Emp ID   B: Name   C: Regular Hrs  D: OT Hrs E: Gross Pay
            F: Tax      G: Net Pay
Data        A: Date     B: Category  C: Metric    D: Value
Lists       Row 1 = master headers (Department | Role | Shift);
            each column below a header = allowed values for that category
Dashboard   (rebuilt by RefreshDashboard)
Form        B2 = master dropdown, B3 = dependent dropdown
```

## Tax slabs

The payroll module uses a simplified annualized slab schedule (mirrored in the Python tests):

| Annual income       | Rate |
|---------------------|------|
| Up to ₹4,00,000     | 0%   |
| ₹4,00,001–8,00,000  | 5%   |
| ₹8,00,001–12,00,000 | 10%  |
| ₹12,00,001–16,00,000| 15%  |
| ₹16,00,001–20,00,000| 20%  |
| ₹20,00,001–24,00,000| 25%  |
| Above ₹24,00,000    | 30%  |

Monthly tax = annual tax on (monthly gross × 12) ÷ 12. Adjust the slabs in `TimesheetPayroll.AnnualTax` to match your jurisdiction.

## Tech stack

- **VBA** (Excel object model only — no external dependencies, no Windows-only APIs)
- **Python 3** + **openpyxl** — demo workbook generator
- **pytest** — test suite

## Running the tests

```bash
pip install openpyxl pytest
python -m pytest tests/ -q
python python/build_demo_workbook.py   # regenerates demo_timesheet.xlsx
```

## Files

```
vba/TimesheetPayroll.bas      timesheet generation, hours & payroll calc, tax slabs
vba/OperationsDashboard.bas   KPI dashboard refresh, top-5 table, charts
vba/DropdownAutomation.bas    cascading dropdowns via named ranges + validation
python/build_demo_workbook.py demo workbook generator (openpyxl)
tests/test_workbook.py        pytest: structure + payroll math edge cases
demo_timesheet.xlsx           regenerate locally (see Quickstart) — not committed
```

## License

MIT — free to use and adapt.
