# Employee Attendance Tracker

![Screenshot](screenshot2.png)

A desktop application for managing employee attendance records stored in Excel, built with PySide6.

## Features

- **3-Column Layout**: Employees list on the left, monthly calendar grid in the middle, controls and totals on the right
- **Monthly Calendar View**: Weekday-aligned grid (Sun–Sat) with colored cells per leave type and today highlighted
- **Employee Management**: Search by name or G-number with tooltips for long names
- **Leave Types**: P (Present), SL (Sick Leave), AL (Annual Leave), AB (Absent), NG (New Guard), TR (Training), `-` (Unavailable)
- **Single Entry**: Click a day to set its status from the selected leave radio
- **Batch Entry**: Set a leave type across a date range (From / To days)
- **Auto Fill**: Fill empty day cells with "P" for the current month
- **Auto Organize**: Consolidates leave below a split point (default row 250) and inserts a yellow separator; NG rows stay in place
- **Days Selector**: Auto-detect days-in-month from the header, or override to 28/29/30/31
- **Split-at # Selector**: Choose which employee number is treated as the split boundary for Auto Organize
- **Total Attendance Panel**: Aggregate P/SL/AL/AB/NG/TR/`-` counts across all employees with total employees and entries
- **Mode Toggle**: One button runs either Auto Fill or Auto Organize based on the selected mode

## Running the App

### Option 1: Run .exe (Recommended)
1. Copy `dist/app_pyside.exe` to your desired location
2. Place `data.xlsx` in the same folder
3. Double-click `app_pyside.exe`

### Option 2: Run with Python
```bash
venv\Scripts\python.exe app_pyside.py
```

Requires Python 3.10+ and the packages listed in `requirements.txt` (PySide6, pandas, openpyxl).

## Data Format

The application reads from `data.xlsx` (`Sheet1`):
- Row 5: Month/year header (e.g., `FEB-2026`) and day-of-week row
- Column A: Row index / sequence number
- Column B: Employee ID (starts with `G`)
- Column C: Employee Name
- Columns F–AJ: Days 1–31 (column count adjusts to the days-in-month selector)
- Footer rows: COUNTIF totals per leave code

## Auto Organize Behavior

See [`auto organize.md`](auto%20organize.md) for the full spec. In short:

1. Count all leave codes in rows above the split point
2. Reset those rows to `P` (dashes and NG rows are preserved)
3. Insert two yellow separator rows after the split
4. Place collected leave below the separator in order: SL → AL → AB → TR
5. Update the conditional-formatting sqref ranges so colors still apply to shifted rows

## License

MIT
