# Auto Organize Feature - Implementation Plan

## Overview
Auto Organize reorganizes employee attendance data for company-specific workflow. It consolidates leave from rows 6-to the last row that have a number in column A  , and places new yellow background rows under 250 in columns while keeping NG at bottom unchanged.

## Data Range

- **Start**: Cell F6 (first employee data row)
- **End**: AJ284 (last column with data)
- **Dynamic Detection**:
  - Columns: Detect day columns based on month (F=day1 → AJ=day31)
  - Rows: Detect last row with employee number in Column A

## Leave Types

| Code | Description | Color |
|------|-------------|-------|
| P    | Present     | Default |
| AB   | Absent      | Red (#e74c3c) |
| SL   | Sick Leave  | Teal (#1abc9c) |
| TR   | Termination| Purple (#9b59b6) |
| AL   | Annual Leave| Blue (#3498db) |
| NG   | Resignation| Orange (#f39c12) |
| -    | Dash/Unavailable| Gray |

## UI Components

### 1. Days Selector (Top of Calendar)
- **Location**: Above or near search bar
- **Options**:
  - Auto (detect from month header)
  - Manual: 28 / 29 / 30 / 31
- **Auto-detect mapping**:
  - JAN, MAR, MAY, JUL, AUG, OCT, DEC → 31 days
  - APR, JUN, SEP, NOV → 30 days
  - FEB (non-leap) → 28 days
  - FEB (leap year) → 29 days

### 2. Auto Organize Button (Bottom Right)
- **Location**: /autofill bottom ( bottom will can be toggled to Change between this two functions )
- **Button**: "Auto Organize" with dropdown arrow ▾
- **Dropdown options**:
  - Auto Organize
  - Set Days: Auto (will show the how many days)
  - Set Days: 28
  - Set Days: 29
  - Set Days: 30
  - Set Days: 31

### 3. Yellow Separator Rows
- **Rows**: 2 blank rows between row 250 and 251
- **Background**: Yellow (#FFFF00)
- **Columns**: All columns (#, ID, Name, Designation, Days)

## Auto Organize Algorithm

### Step 1: Read Data
```
1. Find last row with employee number in Column A (last_employee_row)
2. Determine day_columns based on days selector (default: auto-detect)
3. Read attendance from F6 to (last_day_col)(last_employee_row)
```

### Step 2: Collect Leave from Rows 1-250
```
For each row in 1-250:
  - Count AB days 
  - Count SL days
  - Count AL days
  - Count TR days
  - Leave "-" unchanged
  - NG stays in original position (do not collect)
```

### Step 3: Replace Rows 1-250 to Present
```
For each row in 1-250:
  - Change any non-P, non-"-",  cell to "P"
```

### Step 4: Insert Yellow Separator
```
After row 250:
  - Insert 2 blank rows with yellow background
```

### Step 5: Place Leave Under 251 (SL → AB → AL → TR)
```
Priority order: SL → AB → AL → TR (TR last because it may exceed limit)

1. Find rows 251+ that have available P slots
2. Fill SL first:
   - Replace P with SL until collected SL days exhausted OR no P slots remain
3. Fill AB next:
   - Replace P with AB until collected AB days exhausted OR no P slots remain
4. Fill AL next:
   - Replace P with AL until collected AL days exhausted OR no P slots remain
5. Fill TR last:
   - Replace P with TR until collected TR days exhausted OR no P slots remain
   - Any remaining TR (exceeds available P) stays above 250 as original
```

### Step 6: Handle NG
```
- NG entries stay at their original position (untouched)
```

## Example Transformation

### Before (sample):
| Row | ID   | Day1 | Day2 | Day3 | ... |
|-----|------|------|------|------|-----|
| 13  | G3085| P    | P    | AB   | AB  |
| 17  | G3094| P    | TR   | TR   | TR  |
| 22  | G3102| AL   | AL   | AL   | AL  |
| 250 | G4660| P    | P    | P    | P   |
| 251 | G4662| P    | P    | SL   | SL  |
| ... | ... | ... | ... | ... | ... |

### After:
| Row | ID   | Day1 | Day2 | Day3 | ... |
|-----|------|------|------|------|-----|
| 1-250 | All P or "-" | P | P | P | P |
| 251   | YELLOW SEPARATOR | | | | |
| 252   | YELLOW SEPARATOR | | | | |
| 253+  | Leave consolidated: | SL... | SL... | SL... |
|       | Next row: AL... | | | | |
|       | Next row: AB... | | | | |
| NG rows | Stay at original | NG | NG | P | P |

## Data Limit
- **Rows**: Last employee number in Column A (dynamic)
- **Columns**: F to AJ (31 columns max)
- **Days selector**: Determines actual columns to process (28-31)

## Edge Cases

1. **No leave to organize**: If no leave found in rows 1-250, do nothing
2. **TR exceeds P limit**: Keep excess TR above 250
3. **Empty rows**: Skip empty rows in data processing
4. **NG handling**: NG always stays at original position
5. **- (dash)**: Leave as is, do not modify

## Implementation Notes
- Use openpyxl for Excel file manipulation
- Preserve original formatting where possible
- Auto-detect month from header (e.g., "FEB-2026")
- Dynamic row/column detection for flexibility
