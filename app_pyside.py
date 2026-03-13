import sys
from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                                QHBoxLayout, QLabel, QListWidget, QLineEdit, QRadioButton, 
                                QPushButton, QFrame, QGroupBox, QScrollArea, QGridLayout, QMessageBox)
from PySide6.QtCore import Qt
import pandas as pd
from openpyxl import load_workbook
from datetime import datetime

FILE_PATH = 'data.xlsx'
SHEET_NAME = 'Sheet1'

COLOR_MAP = {
    "P": "#27ae60",
    "SL": "#e74c3c", 
    "AL": "#f39c12",
    "AB": "#9b59b6",
    "NG": "#3498db", 
    "TR": "#1abc9c",
    "-": "#7f8c8d"
}

class AttendanceApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Employee Attendance")
        self.setGeometry(100, 100, 1300, 700)
        self.current_employee = None
        self.employees = {}
        self.current_day = datetime.now().day
        
        self.load_employees()
        self.setup_ui()
        self.apply_styles()
    
    def load_employees(self):
        self.df = pd.read_excel(FILE_PATH, sheet_name=SHEET_NAME, header=None)
        for idx, row in self.df.iterrows():
            if idx >= 5:
                g_num = str(row[1]).strip() if pd.notna(row[1]) else ""
                name = str(row[2]).strip() if pd.notna(row[2]) else ""
                if g_num and g_num.startswith('G'):
                    self.employees[g_num] = {"name": name, "row": idx + 1}
    
    def setup_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        
        header = QLabel("Employee Attendance Entry")
        header.setObjectName("header")
        main_layout.addWidget(header)
        
        content = QHBoxLayout()
        main_layout.addLayout(content)
        
        left = QWidget()
        left.setFixedWidth(350)
        left_layout = QVBoxLayout(left)
        left_layout.setSpacing(6)
        
        emp_label = QLabel("Employees")
        emp_label.setObjectName("section_title")
        left_layout.addWidget(emp_label)
        
        self.search = QLineEdit()
        self.search.setPlaceholderText("Search...")
        self.search.textChanged.connect(self.on_search)
        left_layout.addWidget(self.search)
        
        self.emp_list = QListWidget()
        self.emp_list.setSpacing(2)
        self.emp_list.itemClicked.connect(self.on_select_employee)
        left_layout.addWidget(self.emp_list, 1)
        
        self.emp_info = QLabel("Select an employee")
        self.emp_info.setObjectName("emp_info")
        left_layout.addWidget(self.emp_info)
        
        self.stats_label = QLabel("")
        self.stats_label.setObjectName("stats_label")
        self.stats_label.setWordWrap(True)
        self.stats_label.setTextFormat(Qt.RichText)
        left_layout.addWidget(self.stats_label)
        
        set_label = QLabel("Set Attendance")
        set_label.setObjectName("section_title")
        left_layout.addWidget(set_label)
        
        type_group = QGroupBox("Leave Type")
        type_layout = QVBoxLayout()
        self.leave_type = "SL"
        
        for text, val in [("P - Present", "P"), ("SL - Sick Leave", "SL"), ("AL - Annual Leave", "AL"), 
                          ("AB - Absent", "AB"), ("NG - New Guard", "NG"), 
                          ("TR - Training", "TR"), ("- - Resigned/Terminated", "-")]:
            rb = QRadioButton(text)
            rb.toggled.connect(lambda checked, v=val: setattr(self, 'leave_type', v) if checked else None)
            if val == "P":
                rb.setChecked(True)
            type_layout.addWidget(rb)
        type_group.setLayout(type_layout)
        left_layout.addWidget(type_group)
        
        date_group = QGroupBox("Date Range")
        date_layout = QGridLayout()
        
        date_layout.addWidget(QLabel("Start Day:"), 0, 0)
        self.start_day = QLineEdit(str(datetime.now().day))
        self.start_day.setFixedWidth(60)
        date_layout.addWidget(self.start_day, 0, 1)
        
        date_layout.addWidget(QLabel("Days:"), 1, 0)
        self.num_days = QLineEdit("1")
        self.num_days.setFixedWidth(60)
        date_layout.addWidget(self.num_days, 1, 1)
        
        date_group.setLayout(date_layout)
        left_layout.addWidget(date_group)
        
        self.set_btn = QPushButton("Set Entry")
        self.set_btn.clicked.connect(self.set_entry)
        left_layout.addWidget(self.set_btn)
        
        self.status = QLabel("")
        self.status.setObjectName("status")
        left_layout.addWidget(self.status)
        
        content.addWidget(left)
        
        right = QWidget()
        right_layout = QVBoxLayout(right)
        
        cal_label = QLabel("Attendance Calendar (Days 1-31)")
        cal_label.setObjectName("section_title")
        right_layout.addWidget(cal_label)
        
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setObjectName("calendar_scroll")
        
        cal_widget = QWidget()
        cal_grid = QGridLayout(cal_widget)
        cal_grid.setSpacing(4)
        
        self.day_widgets = {}
        self.day_number_widgets = {}
        for day in range(1, 32):
            col_frame = QFrame()
            col_frame.setObjectName("day_cell")
            cell_layout = QVBoxLayout(col_frame)
            cell_layout.setContentsMargins(4, 2, 4, 2)
            cell_layout.setSpacing(2)
            cell_layout.setAlignment(Qt.AlignTop)
            
            day_label = QLabel(f"{day}")
            day_label.setObjectName("day_number")
            day_label.setFixedHeight(20)
            day_label.setAlignment(Qt.AlignCenter)
            cell_layout.addWidget(day_label)
            
            status_label = QLabel("P")
            status_label.setObjectName("day_status")
            status_label.setFixedHeight(28)
            status_label.setAlignment(Qt.AlignCenter)
            cell_layout.addWidget(status_label)
            
            self.day_widgets[day] = status_label
            self.day_number_widgets[day] = day_label
            cal_grid.addWidget(col_frame, 0, day-1)
        
        scroll.setWidget(cal_widget)
        right_layout.addWidget(scroll)
        
        content.addWidget(right, 1)
        
        self.populate_employee_list()
    
    def apply_styles(self):
        self.setStyleSheet("""
            QMainWindow {
                background-color: #1e1e2e;
            }
            #header {
                background-color: #313244;
                color: #cdd6f4;
                font-size: 18px;
                font-weight: bold;
                padding: 15px;
                qproperty-alignment: AlignCenter;
            }
            #section_title {
                color: #89b4fa;
                font-size: 14px;
                font-weight: bold;
                padding: 8px 0;
            }
            #emp_info {
                color: #a6adc8;
                font-size: 13px;
                font-weight: bold;
                padding: 10px;
                background-color: #313244;
                border-radius: 8px;
                qproperty-alignment: AlignCenter;
            }
            #stats_label {
                color: #cdd6f4;
                font-size: 11px;
                padding: 8px;
                background-color: #313244;
                border-radius: 8px;
                line-height: 1.4;
            }
            QLineEdit {
                background-color: #313244;
                color: #cdd6f4;
                border: 1px solid #45475a;
                border-radius: 6px;
                padding: 8px;
                font-size: 12px;
            }
            QLineEdit:focus {
                border: 1px solid #89b4fa;
            }
            QListWidget {
                background-color: #313244;
                color: #cdd6f4;
                border: 1px solid #45475a;
                border-radius: 8px;
                font-size: 13px;
            }
            QListWidget::item {
                padding: 6px 8px;
                border-radius: 4px;
                min-height: 24px;
            }
            QListWidget::item:selected {
                background-color: #89b4fa;
                color: #1e1e2e;
            }
            QListWidget::item:hover {
                background-color: #45475a;
            }
            QGroupBox {
                color: #89b4fa;
                font-weight: bold;
                border: 1px solid #45475a;
                border-radius: 8px;
                margin-top: 10px;
                padding-top: 10px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                subcontrol-position: top left;
                padding: 0 5px;
            }
            QRadioButton {
                color: #cdd6f4;
                padding: 4px;
            }
            QRadioButton::indicator {
                width: 16px;
                height: 16px;
            }
            QRadioButton::indicator:unchecked {
                border: 2px solid #45475a;
                border-radius: 8px;
                background-color: #313244;
            }
            QRadioButton::indicator:checked {
                border: 2px solid #89b4fa;
                border-radius: 8px;
                background-color: #89b4fa;
            }
            QPushButton {
                background-color: #89b4fa;
                color: #1e1e2e;
                border: none;
                border-radius: 8px;
                padding: 12px;
                font-size: 14px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #b4befe;
            }
            QPushButton:pressed {
                background-color: #74c7ec;
            }
            #status {
                color: #a6e3a1;
                font-size: 12px;
                qproperty-alignment: AlignCenter;
                padding: 5px;
            }
            #calendar_scroll {
                background-color: #181825;
                border: none;
            }
            #day_cell {
                background-color: #313244;
                border-radius: 6px;
                min-width: 45px;
                max-width: 55px;
            }
            #day_number {
                color: #6c7086;
                font-size: 10px;
                font-weight: bold;
                qproperty-alignment: AlignCenter;
                background-color: #45475a;
                border-radius: 3px;
                padding: 2px;
            }
            #day_status {
                color: black;
                font-size: 11px;
                font-weight: bold;
                qproperty-alignment: AlignCenter;
                background-color: #27ae60;
                border-radius: 4px;
            }
        """)
    
    def populate_employee_list(self, query=""):
        self.emp_list.clear()
        sorted_emps = sorted(self.employees.keys())
        
        for g_num in sorted_emps:
            name = self.employees[g_num]["name"]
            display = f"{g_num} - {name}"
            if not query or query.upper() in display.upper():
                self.emp_list.addItem(display)
    
    def on_search(self, text):
        self.populate_employee_list(text)
    
    def on_select_employee(self, item):
        g_num = item.text().split(" - ")[0]
        self.current_employee = g_num
        self.emp_info.setText(f"{self.employees[g_num]['name']}\n({g_num})")
        self.emp_info.setStyleSheet("color: #89b4fa; font-size: 13px; font-weight: bold; padding: 10px; background-color: #313244; border-radius: 8px; qproperty-alignment: AlignCenter;")
        self.load_attendance()
        self.load_stats(g_num)
    
    def get_day_column(self, day):
        return 6 + (day - 1)
    
    def load_attendance(self):
        if not self.current_employee:
            return
        
        g_num = self.current_employee
        row_num = self.employees[g_num]["row"]
        
        try:
            wb = load_workbook(FILE_PATH, data_only=True)
            ws = wb[SHEET_NAME]
            
            for day in range(1, 32):
                col_num = self.get_day_column(day)
                value = ws.cell(row=row_num, column=col_num).value
                display = str(value).strip() if value else "P"
                
                bg_color = COLOR_MAP.get(display, "#27ae60")
                fg_color = "white" if display != "P" else "black"
                
                if day < self.current_day:
                    day_border = "border: 2px solid #a6e3a1;"
                    day_bg = "#313244"
                    day_fg = "#a6e3a1"
                elif day == self.current_day:
                    day_border = "border: 2px solid #f9e2af;"
                    day_bg = "#313244"
                    day_fg = "#f9e2af"
                else:
                    day_border = "border: 2px solid #45475a;"
                    day_bg = "#45475a"
                    day_fg = "#6c7086"
                
                self.day_widgets[day].setText(display)
                self.day_widgets[day].setStyleSheet(f"""
                    color: {fg_color};
                    font-size: 11px;
                    font-weight: bold;
                    qproperty-alignment: AlignCenter;
                    background-color: {bg_color};
                    border-radius: 4px;
                """)
                
                self.day_number_widgets[day].setStyleSheet(f"""
                    color: {day_fg};
                    font-size: 10px;
                    font-weight: bold;
                    qproperty-alignment: AlignCenter;
                    background-color: {day_bg};
                    border-radius: 3px;
                    padding: 2px;
                    {day_border}
                """)
            
            wb.close()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to load: {str(e)}")
    
    def load_stats(self, g_num):
        row_num = self.employees[g_num]["row"]
        try:
            wb = load_workbook(FILE_PATH, data_only=True)
            ws = wb[SHEET_NAME]
            
            counts = {"P": 0, "SL": 0, "AL": 0, "AB": 0, "NG": 0, "TR": 0, "-": 0}
            
            for day in range(1, self.current_day + 1):
                col_num = self.get_day_column(day)
                value = ws.cell(row=row_num, column=col_num).value
                display = str(value).strip() if value else "P"
                if display in counts:
                    counts[display] += 1
            
            stats_text = f"""
<span style='color: #a6e3a1;'>● Present:</span> {counts['P']} &nbsp;
<span style='color: #e74c3c;'>● SL:</span> {counts['SL']} &nbsp;
<span style='color: #f39c12;'>● AL:</span> {counts['AL']}<br>
<span style='color: #9b59b6;'>● AB:</span> {counts['AB']} &nbsp;
<span style='color: #3498db;'>● NG:</span> {counts['NG']} &nbsp;
<span style='color: #1abc9c;'>● TR:</span> {counts['TR']}
"""
            self.stats_label.setText(stats_text)
            wb.close()
        except Exception as e:
            self.stats_label.setText("")
    
    def set_entry(self):
        if not self.current_employee:
            QMessageBox.warning(self, "Error", "Select an employee first!")
            return
        
        try:
            start_day = int(self.start_day.text())
            num_days = int(self.num_days.text())
            if start_day < 1 or start_day > 31 or num_days < 1:
                raise ValueError()
        except ValueError:
            QMessageBox.warning(self, "Error", "Invalid day or number!")
            return
        
        leave_type = self.leave_type
        g_num = self.current_employee
        row_num = self.employees[g_num]["row"]
        
        try:
            wb = load_workbook(FILE_PATH)
            ws = wb[SHEET_NAME]
            
            updated_days = []
            for i in range(num_days):
                day = start_day + i
                if day > 31:
                    break
                col_num = self.get_day_column(day)
                ws.cell(row=row_num, column=col_num, value=leave_type)
                updated_days.append(str(day))
            
            wb.save(FILE_PATH)
            wb.close()
            
            self.status.setText(f"✓ Set {leave_type} for days: {', '.join(updated_days)}")
            self.load_attendance()
            self.load_stats(g_num)
        except PermissionError:
            QMessageBox.critical(self, "Error", "Close data.xlsx in Excel first!")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed: {str(e)}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = AttendanceApp()
    window.show()
    sys.exit(app.exec())
