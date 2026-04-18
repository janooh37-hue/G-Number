# CLAUDE.md

Guidance for Claude Code when working in this repository.

## Project

Employee Attendance Tracker — a PySide6 desktop app that reads/writes `data.xlsx` (Sheet1) using openpyxl. The canonical entry point is `app_pyside.py`; `app.py` is a legacy Tkinter version and should not be modified unless explicitly requested.

## Workbook layout

- `data.xlsx` / `Sheet1` is the live file; `before .xlsx` and `after .xlsx` are reference copies for testing
- Row 5 holds the month/year header (e.g., `FEB-2026`); day cells are 1-indexed across columns F–AJ
- Data rows start at row 6
- Column A = sequence number, B = G-number (employee ID), C = name, D = designation, E = days count
- Employee rows are those where column B starts with `G`
- Footer rows (below the last employee) hold `COUNTIF` formulas per leave code. The code labels (`D289` SL, `D290` AL, etc.) are referenced by the formulas, so trailing whitespace on those cells breaks the counts

## UI structure (`AttendanceApp.setup_ui`)

3-column `QHBoxLayout`:
- **Left (300px fixed)**: search, employee `QListWidget`, employee info card, stats label
- **Middle (flex)**: calendar title, weekday header row, `month_grid_widget` populated by `build_month_grid`, legend
- **Right (320px fixed)**: config group (Days + Split-at #), leave-type radios, date range inputs, Set Entry button, Auto Fill / Auto Organize button, mode selector, status label, `addStretch`, `totals_group`

The calendar grid is built per-month via `get_month_info` (parses the F5 header) and `build_month_grid` (positions day cells by weekday).

## Key methods to know

- `load_employees` — reads the sheet via pandas, populates `self.employees`
- `load_attendance_to_grid` — paints the monthly grid from the current employee's row
- `load_totals` — recounts leave codes across all employees for the right-side panel (call after any write)
- `auto_fill` — fills blank day cells with `P` for the current day range
- `auto_organize` — consolidates leave below the split row; preserves NG rows; updates conditional-formatting sqrefs after the shift
- `get_actual_days` — resolves the Days selector (Auto / 28 / 29 / 30 / 31) to a concrete column span
- `COLOR_MAP` — the source of truth for leave-type colors in both the UI and openpyxl fills

## Testing with a locked workbook

`data.xlsx` is often open in Excel during development. To test changes without closing Excel:
1. `cp "before .xlsx" test_data.xlsx`
2. Run with `app_pyside.FILE_PATH = 'test_data.xlsx'` monkey-patched before constructing `AttendanceApp`
3. Inspect the result with a separate openpyxl script

Set `PYTHONIOENCODING=utf-8` on Windows before running — the app prints characters like `●` that break cp1252.

## Conditional formatting after shifts

`ws.conditional_formatting._cf_rules` keys are `ConditionalFormatting` objects. Their cell ranges live at `cf_range.sqref.ranges` (each element a `CellRange` with `.bounds`). To update after a row shift, collect `(new_sqref_string, list(rules))` tuples, `clear()` the dict, and re-add via `ws.conditional_formatting.add(sqref_str, rule)` — do not try to mutate the existing keys in place (hash breaks) and do not swap the keys for plain `MultiCellRange` objects.

## Non-obvious invariants

- **Auto Organize must skip entire NG rows during the reset pass**, not just NG cells. Other codes (AB, SL, etc.) sitting inside an NG row are intentionally preserved.
- **Footer code labels must have no trailing whitespace**. `COUNTIF(range, D289)` silently returns 0 when `D289` is `"SL "` instead of `"SL"`.
- The conditional-formatting sqref for day-cell coloring must extend down to the new last employee row after Auto Organize inserts the yellow separator, or shifted rows render without color.
- Leave placement order under the yellow separator is fixed: SL → AL → AB → TR. TR last because it may exceed available `P` slots.

## Do not

- Add emojis to code or commits.
- Introduce a second source of truth for leave colors — use `COLOR_MAP`.
- Amend previous commits; create new ones.
- Touch `data.xlsx` in a test script — always copy to a `test_*.xlsx` first.
