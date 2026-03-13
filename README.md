# Employee Attendance Tracker

![Screenshot](screenshot2.png)

A desktop application for managing employee attendance records using Excel.

## Features

- **Employee Management**: Browse and search employees by ID or name
- **Attendance Tracking**: Mark daily attendance with leave types
- **Leave Types**: P (Present), SL (Sick Leave), AL (Annual Leave), AB (Absent), NG (New Guard), TR (Training), - (Resigned/Terminated)
- **Calendar View**: Visual 31-day calendar showing attendance status
- **Batch Entry**: Set attendance for multiple consecutive days
- **Auto Fill**: Fill all empty cells with "P" for the current month

## Running the App

### Option 1: Run .exe (Recommended)
1. Copy `dist/app_pyside.exe` to your desired location
2. Place `data.xlsx` in the same folder
3. Double-click `app_pyside.exe`

### Option 2: Run with Python
```bash
venv\Scripts\python.exe app_pyside.py
```

## Data Format

The application reads from `data.xlsx` (Sheet1):
- Column B: Employee ID (starts with "G")
- Column C: Employee Name
- Columns G-AH: Days 1-31

## License

MIT
