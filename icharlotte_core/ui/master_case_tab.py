import sys
import os
import json
import datetime
from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QTableWidget, QTableWidgetItem,
    QHeaderView, QSplitter, QFrame, QLabel, QLineEdit, QPushButton,
    QDateEdit, QListWidget, QListWidgetItem, QMenu, QMessageBox,
    QAbstractItemView, QDialog, QFormLayout, QDialogButtonBox,
    QComboBox, QProgressBar, QCheckBox, QToolTip, QTextEdit, QInputDialog,
    QCalendarWidget
)
from PyQt6.QtCore import Qt, QDate, QThread, pyqtSignal, QTimer
from PyQt6.QtGui import QColor, QAction, QFont, QShortcut, QKeySequence

from icharlotte_core.master_db import MasterCaseDatabase
from icharlotte_core.utils import log_event, get_case_path, BASE_PATH_WIN, parse_hearing_data
from icharlotte_core.sent_items_monitor import SentItemsMonitorWorker
from icharlotte_core.calendar import CalendarMonitorWorker

class CalendarDialog(QDialog):
    def __init__(self, parent=None, current_date=None):
        super().__init__(parent)
        self.setWindowTitle("Select Date")
        self.layout = QVBoxLayout(self)
        
        self.calendar = QCalendarWidget()
        if current_date:
            self.calendar.setSelectedDate(current_date)
        self.layout.addWidget(self.calendar)
        
        btn_layout = QHBoxLayout()
        ok_btn = QPushButton("OK")
        ok_btn.clicked.connect(self.accept)
        cancel_btn = QPushButton("Cancel")
        cancel_btn.clicked.connect(self.reject)
        
        # Clear/Reset button
        clear_btn = QPushButton("Clear")
        clear_btn.clicked.connect(self.clear_date)
        
        btn_layout.addWidget(clear_btn)
        btn_layout.addStretch()
        btn_layout.addWidget(cancel_btn)
        btn_layout.addWidget(ok_btn)
        self.layout.addLayout(btn_layout)
        
        self.selected_date = None
        self.cleared = False
        
    def accept(self):
        self.selected_date = self.calendar.selectedDate()
        super().accept()
        
    def clear_date(self):
        self.cleared = True
        super().accept()

class HearingCellWidget(QWidget):
    def __init__(self, hearings, parent=None):
        super().__init__(parent)
        layout = QHBoxLayout(self)
        layout.setContentsMargins(4, 0, 4, 0)
        layout.setSpacing(2)
        
        # Filter to only show future or today's hearings
        # hearings is a sorted list from parse_hearing_data
        today = datetime.datetime.now().date()
        
        self.future_hearings = []
        for h in hearings:
            # Check if date is today or future, or "no date" (9999)
            if h['date_obj'].date() >= today or h['date_obj'].year == 9999:
                self.future_hearings.append(h)
        
        # Display the first one (since list is sorted by date)
        self.display_hearing = self.future_hearings[0] if self.future_hearings else None
        
        text = self.display_hearing['display'] if self.display_hearing else ""
        self.label = QLabel(text)
        self.label.setStyleSheet("background: transparent; border: none;")
        layout.addWidget(self.label)
        
        # Show dropdown if there are multiple *future* hearings
        if len(self.future_hearings) > 1:
            self.btn = QPushButton("▼")
            self.btn.setFixedSize(16, 16)
            self.btn.setStyleSheet("""
                QPushButton {
                    border: none; 
                    font-size: 8px; 
                    color: gray;
                    background: transparent;
                }
                QPushButton:hover {
                    color: black;
                    background: #eee;
                    border-radius: 8px;
                }
            """)
            self.btn.setCursor(Qt.CursorShape.PointingHandCursor)
            self.btn.clicked.connect(lambda: self.show_menu(self.future_hearings))
            layout.addWidget(self.btn)
            
        layout.addStretch()
        
    def show_menu(self, hearings):
        menu = QMenu(self)
        for h in hearings:
            # Mark the one currently displayed with bold or something?
            action_text = h['display']
            if self.display_hearing and h == self.display_hearing:
                action_text = f"-> {action_text}"
            menu.addAction(action_text)
        menu.exec(self.btn.mapToGlobal(self.btn.rect().bottomLeft()))

class CaseScannerWorker(QThread):
    progress = pyqtSignal(str)
    finished = pyqtSignal(int, int) # Count of found cases, Count of removed cases

    def __init__(self, db):
        super().__init__()
        self.db = db
        self.running = True

    def run(self):
        try:
            import win32com.client
            import pythoncom
            import re

            pythoncom.CoInitialize()
            outlook = win32com.client.Dispatch("Outlook.Application")
            mapi = outlook.GetNamespace("MAPI")

            new_count = 0
            found_file_numbers = set()
            
            # Get existing cases to track removals
            all_existing = self.db.get_all_cases()
            existing_file_numbers = {c['file_number'] for c in all_existing}
            
            # Strategy: Look directly for "CASES" at the root of every store, similar to EmailSyncWorker.
            self.progress.emit("Connecting to Outlook...")

            for store in mapi.Stores:
                if not self.running: break
                try:
                    # Skip Archive stores
                    if "archive" in store.DisplayName.lower():
                        self.progress.emit(f"Skipping Archive Store: {store.DisplayName}")
                        continue

                    root = store.GetRootFolder()
                    cases_folder = None
                    
                    # Case-insensitive check for CASES folder
                    for f in root.Folders:
                        if f.Name.lower() == "cases":
                            cases_folder = f
                            break
                            
                    if cases_folder:
                        self.progress.emit(f"Scanning 'CASES' in {store.DisplayName}...")
                        
                        for sub in cases_folder.Folders:
                            if not self.running: break
                            
                            folder_name = sub.Name
                            # Expecting format like "3200-284 - Goulart"
                            # We need to extract File Number and Name.
                            
                            # Regex to find file number: ####.### or ####-###
                            match = re.search(r"(\d{4})[._-](\d{3})", folder_name)
                            if match:
                                # Normalize file number to ####.###
                                file_number = f"{match.group(1)}.{match.group(2)}"
                                
                                # Extract Name (everything after the number?)
                                name_part = folder_name.replace(match.group(0), "").strip()
                                # Clean up leading hyphens/underscores/spaces
                                name_part = name_part.lstrip(" -_").strip()
                                
                                if not name_part:
                                    name_part = folder_name # Fallback
                                
                                found_file_numbers.add(file_number)

                                # Try to resolve local path if possible, otherwise leave blank or try to find it
                                case_path = get_case_path(file_number) or ""

                                # Check DB if case exists
                                existing = self.db.get_case(file_number)
                                if not existing:
                                    self.db.upsert_case(file_number, name_part, "", "", case_path)
                                    new_count += 1
                                    self.progress.emit(f"Found New: {file_number} - {name_part}")
                                else:
                                    # Update name if changed (and not overridden), or path if missing
                                    # Also ensure we have the latest name from Outlook if it changed there
                                    
                                    # Determine if we should update the name
                                    should_update_name = False
                                    final_name = existing['plaintiff_last_name']
                                    
                                    if not existing.get('plaintiff_override', 0):
                                        if existing['plaintiff_last_name'] != name_part:
                                            should_update_name = True
                                            final_name = name_part
                                    
                                    if should_update_name or (case_path and not existing['case_path']):
                                         self.db.upsert_case(
                                             file_number, 
                                             final_name, 
                                             existing['next_hearing_date'], 
                                             existing['trial_date'], 
                                             case_path or existing['case_path'],
                                             plaintiff_override=existing.get('plaintiff_override', 0)
                                         )
                                         if should_update_name:
                                             self.progress.emit(f"Updated Name: {file_number}")

                except Exception as e:
                    pass
            
            # Remove cases that are in DB but not found in Outlook
            removed_count = 0
            # Safety check: Only proceed with deletion if we successfully found AT LEAST one folder in Outlook.
            # This prevents wiping the database if Outlook is offline or 'CASES' folders are temporarily inaccessible.
            if self.running and found_file_numbers:
                for old_file in existing_file_numbers:
                    if old_file not in found_file_numbers:
                        self.db.delete_case(old_file)
                        removed_count += 1
                        self.progress.emit(f"Removed: {old_file}")
            elif self.running and not found_file_numbers:
                self.progress.emit("No cases found in Outlook. Skipping deletion for safety.")

            pythoncom.CoUninitialize()
            self.finished.emit(new_count, removed_count)
            
        except Exception as e:
            self.progress.emit(f"Scan Error: {e}")
            self.finished.emit(0, 0) # Safety emit

    def stop(self):
        self.running = False


class DateTableWidgetItem(QTableWidgetItem):
    def __lt__(self, other):
        # Use UserRole+10 as the sort key (YYYY-MM-DD string)
        s1 = self.data(Qt.ItemDataRole.UserRole + 10)
        s2 = other.data(Qt.ItemDataRole.UserRole + 10)
        if s1 and s2:
            return s1 < s2
        return super().__lt__(other)

class TodoItemWidget(QWidget):
    statusChanged = pyqtSignal(bool) # Checked/Unchecked
    colorChanged = pyqtSignal(str)   # New color
    assignmentChanged = pyqtSignal(str, str) # New initials, New date string

    def __init__(self, text, status, color, created_date, assigned_to, assigned_date, case_assigned_attorney=""):
        super().__init__()
        self.case_assigned_attorney = case_assigned_attorney
        layout = QHBoxLayout(self)
        layout.setContentsMargins(5, 2, 5, 2)
        
        # Checkbox
        self.checkbox = QCheckBox()
        self.checkbox.setChecked(status == 'done')
        self.checkbox.stateChanged.connect(lambda: self.statusChanged.emit(self.checkbox.isChecked()))
        layout.addWidget(self.checkbox)
        
        # Color Button
        self.color_btn = QPushButton()
        self.color_btn.setFixedSize(16, 16)
        self.color_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.set_color_btn_style(color)
        self.color_btn.clicked.connect(self.cycle_color)
        self.current_color = color or 'yellow'
        layout.addWidget(self.color_btn)
        
        # Text
        self.label = QLabel(text)
        self.label.setStyleSheet("padding-left: 5px;")
        if status == 'done':
            f = self.label.font()
            f.setStrikeOut(True)
            self.label.setFont(f)
            self.label.setStyleSheet("color: gray; padding-left: 5px;")
        layout.addWidget(self.label)
        
        # Created Date [mm/dd/yy]
        if created_date:
            try:
                # Expecting YYYY-MM-DD, convert to MM/DD/YY
                dt = datetime.datetime.strptime(created_date, "%Y-%m-%d")
                fmt_date = dt.strftime("[%m/%d/%y]")
            except:
                fmt_date = f"[{created_date}]"
            
            date_lbl = QLabel(fmt_date)
            date_lbl.setStyleSheet("color: gray; font-size: 10px; margin-left: 5px;")
            layout.addWidget(date_lbl)

        # Assignment Box
        self.assign_btn = QPushButton(assigned_to or "  ") # Default to empty space if none
        self.assign_btn.setFixedSize(30, 20)
        self.assign_btn.setStyleSheet("border: 1px solid #999; border-radius: 3px; font-size: 10px; background-color: #eee;")
        self.assign_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.assign_btn.clicked.connect(self.cycle_assignment)
        self.current_assigned = assigned_to or ""
        layout.addWidget(self.assign_btn)

        # Assigned Date
        self.assign_date_lbl = QLabel(assigned_date or "")
        self.assign_date_lbl.setStyleSheet("color: gray; font-size: 10px; margin-left: 5px;")
        layout.addWidget(self.assign_date_lbl)
        
        layout.addStretch()

    def set_color_btn_style(self, color):
        c_map = {'red': '#f44336', 'yellow': '#ffeb3b', 'green': '#4caf50', 'blue': '#2196f3'}
        hex_c = c_map.get(color, '#ffeb3b')
        # Circle shape
        self.color_btn.setStyleSheet(
            f"background-color: {hex_c}; border-radius: 8px; border: 1px solid #ccc;"
        )
        self.color_btn.setToolTip(f"Priority: {color.capitalize()} (Click to change)")

    def cycle_color(self):
        order = ['yellow', 'green', 'red', 'blue']
        try:
            idx = order.index(self.current_color)
        except:
            idx = 0
        new_color = order[(idx + 1) % len(order)]
        self.current_color = new_color
        self.set_color_btn_style(new_color)
        self.colorChanged.emit(new_color)

    def cycle_assignment(self):
        # Cycle: Empty -> CE -> AS -> HP -> FM -> Empty
        options = ["", "CE", "AS", "HP", "FM"]
        
        # Filter based on case assigned attorney
        if self.case_assigned_attorney == "FM":
            if "HP" in options: options.remove("HP")
        elif self.case_assigned_attorney == "HP":
            if "FM" in options: options.remove("FM")
            
        try:
            idx = options.index(self.current_assigned)
        except:
            idx = 0
        
        new_assigned = options[(idx + 1) % len(options)]
        self.current_assigned = new_assigned
        
        # Update Button Text
        self.assign_btn.setText(new_assigned if new_assigned else "  ")
        
        # Update Date
        new_date_str = datetime.datetime.now().strftime("%m/%d/%y")
        self.assign_date_lbl.setText(new_date_str)
        
        self.assignmentChanged.emit(new_assigned, new_date_str)

class MasterCaseTab(QWidget):
    def __init__(self, main_window=None):
        super().__init__()
        self.main_window = main_window
        self.db = MasterCaseDatabase()

        self.save_timer = QTimer()
        self.save_timer.setSingleShot(True)
        self.save_timer.setInterval(500) # 500ms debounce
        self.save_timer.timeout.connect(self._perform_save_settings)

        # Email monitor worker (initialized when started)
        self.email_monitor_worker = None

        # Calendar monitor worker (initialized when started)
        self.calendar_monitor_worker = None

        self.setup_ui()
        self.refresh_data()

        # Auto-start email monitor
        QTimer.singleShot(1000, self._auto_start_email_monitor)

        # Auto-start calendar monitor (after email monitor)
        QTimer.singleShot(2000, self._auto_start_calendar_monitor)

    def setup_ui(self):
        main_layout = QVBoxLayout(self)
        
        # --- Top Bar ---
        top_bar = QHBoxLayout()
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Search cases (Name or Number)...")
        self.search_input.textChanged.connect(self.filter_table)
        top_bar.addWidget(self.search_input)
        
        self.refresh_btn = QPushButton("Refresh List")
        self.refresh_btn.clicked.connect(self.refresh_data)
        top_bar.addWidget(self.refresh_btn)
        
        self.scan_btn = QPushButton("Scan Outlook for Cases")
        self.scan_btn.clicked.connect(self.start_scan)
        top_bar.addWidget(self.scan_btn)

        self.clear_btn = QPushButton("Clear List")
        self.clear_btn.setStyleSheet("color: red;")
        self.clear_btn.clicked.connect(self.clear_list)
        top_bar.addWidget(self.clear_btn)

        # Email Monitor Toggle Button
        self.email_monitor_btn = QPushButton("Start Email Monitor")
        self.email_monitor_btn.setCheckable(True)
        self.email_monitor_btn.setStyleSheet("""
            QPushButton {
                background-color: #e8f5e9;
                border: 1px solid #4caf50;
                border-radius: 4px;
                padding: 5px 10px;
            }
            QPushButton:checked {
                background-color: #c8e6c9;
                border: 2px solid #2e7d32;
            }
        """)
        self.email_monitor_btn.clicked.connect(self.toggle_email_monitor)
        top_bar.addWidget(self.email_monitor_btn)

        # Email Monitor Status Label
        self.email_monitor_status = QLabel("")
        self.email_monitor_status.setStyleSheet("color: gray; font-size: 10px; margin-left: 5px;")
        top_bar.addWidget(self.email_monitor_status)

        # Calendar Monitor Toggle Button
        self.calendar_monitor_btn = QPushButton("Start Calendar")
        self.calendar_monitor_btn.setCheckable(True)
        self.calendar_monitor_btn.setStyleSheet("""
            QPushButton {
                background-color: #e3f2fd;
                border: 1px solid #2196f3;
                border-radius: 4px;
                padding: 5px 10px;
            }
            QPushButton:checked {
                background-color: #bbdefb;
                border: 2px solid #1565c0;
            }
        """)
        self.calendar_monitor_btn.clicked.connect(self.toggle_calendar_monitor)
        top_bar.addWidget(self.calendar_monitor_btn)

        # Calendar Monitor Status Label
        self.calendar_monitor_status = QLabel("")
        self.calendar_monitor_status.setStyleSheet("color: gray; font-size: 10px; margin-left: 5px;")
        top_bar.addWidget(self.calendar_monitor_status)

        main_layout.addLayout(top_bar)
        
        # Progress Bar for scan
        self.scan_progress = QProgressBar()
        self.scan_progress.setVisible(False)
        main_layout.addWidget(self.scan_progress)

        # --- Splitter ---
        splitter = QSplitter(Qt.Orientation.Horizontal)
        main_layout.addWidget(splitter)

        # --- Left: Case Table ---
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        left_layout.setContentsMargins(0,0,0,0)
        
        self.table = QTableWidget()
        self.table.setColumnCount(7) # File #, Plaintiff, Assigned, Hearings, Trial, Last Report, Tasks
        self.table.setHorizontalHeaderLabels(["File #", "Plaintiff Name", "Assigned", "Hearings", "Trial Date", "Last Report", "Tasks"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        self.table.setColumnWidth(0, 100) # File #
        self.table.setColumnWidth(1, 300) # Plaintiff Name
        self.table.setColumnWidth(2, 60) # Assigned
        self.table.setColumnWidth(3, 120) # Hearings
        self.table.setColumnWidth(4, 120) # Trial Date
        self.table.setColumnWidth(5, 120) # Last Report
        self.table.horizontalHeader().setSectionResizeMode(6, QHeaderView.ResizeMode.Stretch) # Tasks stretch
        
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.setSortingEnabled(True)
        self.table.itemSelectionChanged.connect(self.on_case_selected)
        self.table.itemDoubleClicked.connect(self.on_case_double_clicked)
        self.table.cellClicked.connect(self.on_cell_clicked)
        
        self.load_column_settings()
        self.table.horizontalHeader().sectionResized.connect(self.save_column_settings)

        # Enable Delete key for Table (Hearing/Trial)
        self.table_del_shortcut = QShortcut(QKeySequence(Qt.Key.Key_Delete), self.table)
        self.table_del_shortcut.setContext(Qt.ShortcutContext.WidgetShortcut)
        self.table_del_shortcut.activated.connect(self.delete_table_cell)
        
        left_layout.addWidget(self.table)
        splitter.addWidget(left_widget)



    # --- Right: Details Panel ---
        self.details_widget = QWidget()
        self.details_layout = QVBoxLayout(self.details_widget)
        self.details_widget.setEnabled(False)
        
        # Header
        self.case_title = QLabel("Select a Case")
        self.case_title.setStyleSheet("font-size: 16px; font-weight: bold; padding: 5px;")
        self.details_layout.addWidget(self.case_title)
        
        # Dates Editor
        dates_frame = QFrame()
        dates_frame.setFrameShape(QFrame.Shape.StyledPanel)
        dates_layout = QFormLayout(dates_frame)
        
        self.edit_hearing = QDateEdit()
        self.edit_hearing.setCalendarPopup(True)
        self.edit_hearing.setDisplayFormat("MM/dd/yyyy")
        self.edit_hearing.setSpecialValueText(" ") # Show blank for default/min date
        self.edit_hearing.dateChanged.connect(self.save_dates)
        
        self.edit_trial = QDateEdit()
        self.edit_trial.setCalendarPopup(True)
        self.edit_trial.setDisplayFormat("MM/dd/yyyy")
        self.edit_trial.setSpecialValueText(" ")
        self.edit_trial.dateChanged.connect(self.save_dates)
        
        dates_layout.addRow("Next Hearing:", self.edit_hearing)
        dates_layout.addRow("Trial Date:", self.edit_trial)
        self.details_layout.addWidget(dates_frame)
        
        # Case Summary
        self.details_layout.addWidget(QLabel("<b>Case Summary:</b>"))
        self.summary_edit = QTextEdit()
        self.summary_edit.setPlaceholderText("Case summary will appear here...")
        self.summary_edit.setFixedHeight(100) # Keep it small box as requested
        self.summary_edit.textChanged.connect(self.save_case_summary_debounced)
        self.details_layout.addWidget(self.summary_edit)

        # To-Dos
        self.details_layout.addWidget(QLabel("<b>To-Do Items:</b>"))
        self.todo_list = QListWidget()
        self.todo_list.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.todo_list.customContextMenuRequested.connect(self.show_todo_menu)
        self.details_layout.addWidget(self.todo_list)
        
        # Enable Delete key for To-Do List
        self.todo_del_shortcut = QShortcut(QKeySequence(Qt.Key.Key_Delete), self.todo_list)
        self.todo_del_shortcut.setContext(Qt.ShortcutContext.WidgetShortcut)
        self.todo_del_shortcut.activated.connect(self.delete_current_todo)
        
        todo_input_layout = QHBoxLayout()
        self.todo_input = QLineEdit()
        self.todo_input.setPlaceholderText("New To-Do Item... (Enter to add)")
        self.todo_input.setStyleSheet("""
            border: 2px solid #2196F3;
            border-radius: 4px;
            padding: 5px;
            background-color: #E3F2FD;
            font-size: 13px;
        """)
        self.todo_input.returnPressed.connect(self.add_todo)
        
        todo_input_layout.addWidget(self.todo_input)
        self.details_layout.addLayout(todo_input_layout)
        
        # History
        self.details_layout.addSpacing(10)
        self.details_layout.addWidget(QLabel("<b>History / Updates:</b>"))
        self.history_list = QListWidget() 
        self.history_list.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.history_list.customContextMenuRequested.connect(self.show_history_menu)
        self.history_list.itemDoubleClicked.connect(self.on_history_double_clicked)
        self.details_layout.addWidget(self.history_list)
        
        # Enable Delete key for History List
        self.hist_del_shortcut = QShortcut(QKeySequence(Qt.Key.Key_Delete), self.history_list)
        self.hist_del_shortcut.setContext(Qt.ShortcutContext.WidgetShortcut)
        self.hist_del_shortcut.activated.connect(self.delete_current_history)
        
        hist_input_layout = QHBoxLayout()
        self.hist_type = QComboBox()
        self.hist_type.addItems(["Status Update", "Note"])
        self.hist_type.setCurrentText("Note")
        self.hist_type.setStyleSheet("""
            padding: 5px;
            border: 1px solid #ccc;
            border-radius: 4px;
        """)
        self.hist_input = QLineEdit()
        self.hist_input.setPlaceholderText("Log entry... (Enter to add)")
        self.hist_input.setStyleSheet("""
            border: 2px solid #4CAF50;
            border-radius: 4px;
            padding: 5px;
            background-color: #E8F5E9;
            font-size: 13px;
        """)
        self.hist_input.returnPressed.connect(self.add_history)
        
        hist_input_layout.addWidget(self.hist_type)
        hist_input_layout.addWidget(self.hist_input)
        self.details_layout.addLayout(hist_input_layout)

        splitter.addWidget(self.details_widget)
        splitter.setSizes([600, 400])

    def delete_table_cell(self):
        # Handle deletion for Trial Date (4) or Hearing (3)
        row = self.table.currentRow()
        col = self.table.currentColumn()
        
        if row < 0 or col < 0: return
        
        file_number = self.table.item(row, 0).data(Qt.ItemDataRole.UserRole)
        
        if col == 3: # Hearings
            confirm = QMessageBox.question(self, "Clear Hearings", f"Clear all hearings for case {file_number}?", QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
            if confirm == QMessageBox.StandardButton.Yes:
                from icharlotte_core.config import GEMINI_DATA_DIR
                json_path = os.path.join(GEMINI_DATA_DIR, f"{file_number}.json")
                if os.path.exists(json_path):
                    try:
                        with open(json_path, 'r', encoding='utf-8') as f:
                            vars_data = json.load(f)
                        
                        # Clear JSON entry
                        keys_to_clear = ["other_hearings", "other_hearing", "next_hearing_date", "next_hearing"]
                        for k in keys_to_clear:
                            if isinstance(vars_data.get(k), dict):
                                vars_data[k]["value"] = ""
                            elif k in vars_data:
                                vars_data[k] = ""

                        with open(json_path, 'w', encoding='utf-8') as f:
                            json.dump(vars_data, f, indent=4)
                    except: pass
                
                # Clear DB
                self.db.update_hearing_date(file_number, "")
                self.refresh_data_row(file_number)
                
        elif col == 4: # Trial Date
            confirm = QMessageBox.question(self, "Clear Trial Date", f"Clear trial date for case {file_number}?", QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
            if confirm == QMessageBox.StandardButton.Yes:
                self.db.update_trial_date(file_number, "")
                self.refresh_data_row(file_number)

    def load_column_settings(self):
        settings_path = os.path.join(os.getcwd(), ".gemini", "config", "master_list_settings.json")
        if os.path.exists(settings_path):
            try:
                # Handle empty file case
                if os.path.getsize(settings_path) == 0:
                    return

                with open(settings_path, 'r') as f:
                    settings = json.load(f)
                    widths = settings.get("column_widths", [])
                    for i, width in enumerate(widths):
                        if i < self.table.columnCount() and isinstance(width, int) and width > 0:
                            self.table.setColumnWidth(i, width)
            except Exception as e:
                print(f"Error loading column settings: {e}")

    def save_column_settings(self, *args):
        # Debounce the save operation
        self.save_timer.start()

    def _perform_save_settings(self):
        settings_path = os.path.join(os.getcwd(), ".gemini", "config", "master_list_settings.json")
        widths = []
        for i in range(self.table.columnCount()):
            widths.append(self.table.columnWidth(i))
            
        settings = {"column_widths": widths}
        try:
            os.makedirs(os.path.dirname(settings_path), exist_ok=True)
            with open(settings_path, 'w') as f:
                json.dump(settings, f, indent=2)
        except Exception as e:
            print(f"Error saving column settings: {e}")

    def start_scan(self):
        self.scan_btn.setEnabled(False)
        self.scan_progress.setVisible(True)
        self.scan_progress.setRange(0, 0) # Indeterminate
        
        self.worker = CaseScannerWorker(self.db)
        self.worker.progress.connect(lambda s: self.case_title.setText(s)) # Show status in title area temp
        self.worker.finished.connect(self.on_scan_finished)
        self.worker.start()

    def on_scan_finished(self, new_count, removed_count):
        self.scan_progress.setVisible(False)
        self.scan_btn.setEnabled(True)
        self.case_title.setText("Select a Case")
        
        msg = []
        if new_count > 0:
            msg.append(f"Found {new_count} new cases.")
        if removed_count > 0:
            msg.append(f"Removed {removed_count} cases that no longer exist in Outlook.")
            
        if not msg:
            msg.append("Scan complete. No changes detected.")
            
        QMessageBox.information(self, "Scan Complete", "\n".join(msg))
        self.refresh_data()

    def clear_list(self):
        confirm = QMessageBox.question(
            self, 
            "Confirm Clear", 
            "Are you sure you want to clear the entire case list? This will remove all cases from this view (but not your files).",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        
        if confirm == QMessageBox.StandardButton.Yes:
            self.db.clear_all_cases()
            self.refresh_data()
            self.details_widget.setEnabled(False)
            self.case_title.setText("Select a Case")

    def refresh_data(self):
        # Visual Feedback
        if hasattr(self, 'refresh_btn'):
            self.refresh_btn.setText("Refreshed!")
            self.refresh_btn.setStyleSheet("color: green; font-weight: bold;")
            QTimer.singleShot(1000, lambda: self._reset_refresh_btn())

        self.table.setSortingEnabled(False)
        self.table.setRowCount(0)
        
        cases = self.db.get_all_cases()
        self.table.setRowCount(len(cases))
        
        today = QDate.currentDate()
        
        from icharlotte_core.config import GEMINI_DATA_DIR
        import json

        for i, case in enumerate(cases):
            file_number = case['file_number']
            
            # --- Auto-Populate Logic ---
            # If dates are missing in DB, try to load from case variables JSON
            hearing_date = case['next_hearing_date']
            trial_date = case['trial_date']
            
            updated = False
            
            if not hearing_date or not trial_date:
                json_path = os.path.join(GEMINI_DATA_DIR, f"{file_number}.json")
                if os.path.exists(json_path):
                    try:
                        with open(json_path, 'r', encoding='utf-8') as f:
                            vars_data = json.load(f)
                            
                        # Helper to extract value from potentially nested dict
                        def get_val(key):
                            v = vars_data.get(key)
                            if isinstance(v, dict) and "value" in v:
                                return v["value"]
                            return v or ""

                        if not hearing_date:
                            found_hearing = get_val("next_hearing_date") or get_val("next_hearing")
                            if found_hearing:
                                hearing_date = found_hearing
                                updated = True
                                
                        if not trial_date:
                            found_trial = get_val("trial_date")
                            if found_trial:
                                trial_date = found_trial
                                updated = True
                                
                        if updated:
                            self.db.upsert_case(
                                file_number, 
                                case['plaintiff_last_name'], 
                                hearing_date, 
                                trial_date, 
                                case['case_path'],
                                plaintiff_override=case.get('plaintiff_override', 0)
                            )
                    except:
                        pass
            
            # ---------------------------

            # File Number
            item_num = QTableWidgetItem(file_number)
            item_num.setData(Qt.ItemDataRole.UserRole, file_number)
            item_num.setForeground(QColor("black"))
            font = item_num.font()
            font.setUnderline(True)
            item_num.setFont(font)
            self.table.setItem(i, 0, item_num)
            
            # Name
            self.table.setItem(i, 1, QTableWidgetItem(case['plaintiff_last_name']))
            
            # Assigned Attorney
            assigned = case.get('assigned_attorney', '') or ""
            item_assigned = QTableWidgetItem(assigned)
            item_assigned.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.table.setItem(i, 2, item_assigned)

            # Hearings (Column 3) - Uses other_hearings variable
            json_path = os.path.join(GEMINI_DATA_DIR, f"{file_number}.json")
            other_hearings_str = ""
            if os.path.exists(json_path):
                try:
                    with open(json_path, 'r', encoding='utf-8') as f:
                        vars_data = json.load(f)
                        # Check for other_hearings or other_hearing
                        v = vars_data.get("other_hearings") or vars_data.get("other_hearing")
                        if isinstance(v, dict) and "value" in v:
                            other_hearings_str = v["value"]
                        elif isinstance(v, str):
                            other_hearings_str = v
                except: pass
            
            # If no other_hearings found in JSON, fall back to DB 'next_hearing_date'
            if not other_hearings_str and hearing_date:
                 other_hearings_str = hearing_date # Treat single date as hearing string
            
            parsed_hearings = parse_hearing_data(other_hearings_str)
            
            # Create Widget
            h_widget = HearingCellWidget(parsed_hearings)
            self.table.setCellWidget(i, 3, h_widget)
            
            # Set Item for Sorting (using the sort date of the displayed hearing)
            sort_val = "9999-99-99"
            if h_widget.display_hearing:
                sort_val = h_widget.display_hearing['date_sort']
                
            item_hearing = DateTableWidgetItem("") # Use DateTableWidgetItem to keep sort logic if needed, but we set UserRole+10
            item_hearing.setData(Qt.ItemDataRole.UserRole + 10, sort_val)
            self.table.setItem(i, 3, item_hearing)
            
            # Trial Date
            disp_trial = "N/A"
            sort_trial = trial_date or "9999-99-99"
            if trial_date and "-" in trial_date:
                parts = trial_date.split("-")
                if len(parts) == 3:
                    disp_trial = f"{parts[1]}/{parts[2]}/{parts[0]}"
            
            item_date = DateTableWidgetItem(disp_trial)
            item_date.setData(Qt.ItemDataRole.UserRole + 10, sort_trial)
            self.table.setItem(i, 4, item_date)
            
            # Color Coding for Trial Date
            if trial_date:
                try:
                    parts = trial_date.split('-')
                    if len(parts) == 3:
                        t_qdate = QDate(int(parts[0]), int(parts[1]), int(parts[2]))
                        days_to = today.daysTo(t_qdate)
                        
                        if days_to < 0:
                            item_date.setForeground(QColor("gray"))
                        elif days_to < 90: # < 3 months
                            item_date.setBackground(QColor("#ffcdd2"))
                        elif days_to < 180: # < 6 months
                            item_date.setBackground(QColor("#fff9c4"))
                except:
                    pass

            # --- Last Report (Column 5) ---
            # Prioritize manual override text if exists
            manual_report = case.get('last_report_text', '')
            last_report_date = self.db.get_last_status_update_date(file_number)
            
            disp_report = ""
            sort_report = ""
            
            if manual_report:
                disp_report = manual_report
                # If it looks like a date, sort as date, else sort as text (maybe 0000 prefix to keep at bottom or top?)
                # Let's try to see if it's YYYY-MM-DD
                if "-" in manual_report and len(manual_report) == 10:
                     sort_report = manual_report
                else:
                     sort_report = "0000-00-00" # Sort text entries as old/undefined
            elif last_report_date:
                 try:
                    # Expecting YYYY-MM-DD
                    if "-" in last_report_date:
                         parts = last_report_date.split("-")
                         if len(parts) == 3:
                             disp_report = f"{parts[1]}/{parts[2]}/{parts[0]}" # mm/dd/yyyy
                         sort_report = last_report_date
                 except:
                     pass
            else:
                 sort_report = "0000-00-00"

            item_report = DateTableWidgetItem(disp_report)
            item_report.setData(Qt.ItemDataRole.UserRole + 10, sort_report)
            self.table.setItem(i, 5, item_report)

            # --- Tasks Column (Index 6) ---
            todos = self.db.get_todos(file_number)
            pending = [t for t in todos if t['status'] != 'done']
            
            if pending:
                task_widget = QWidget()
                task_layout = QHBoxLayout(task_widget)
                task_layout.setContentsMargins(5, 5, 5, 5)
                task_layout.setSpacing(4)
                task_layout.setAlignment(Qt.AlignmentFlag.AlignLeft)
                
                c_map = {'red': '#f44336', 'yellow': '#ffeb3b', 'green': '#4caf50', 'blue': '#2196f3'}
                
                for t in pending:
                    color = t.get('color', 'yellow') or 'yellow'
                    hex_c = c_map.get(color, '#ffeb3b')
                    
                    lbl = QLabel()
                    lbl.setFixedSize(12, 12)
                    lbl.setStyleSheet(f"background-color: {hex_c}; border-radius: 6px;")
                    lbl.setToolTip(f"{t['item']}")
                    task_layout.addWidget(lbl)
                    
                self.table.setCellWidget(i, 6, task_widget)
            else:
                self.table.setItem(i, 6, QTableWidgetItem(""))

        self.table.setSortingEnabled(True)

    def _reset_refresh_btn(self):
        if hasattr(self, 'refresh_btn'):
            self.refresh_btn.setText("Refresh List")
            self.refresh_btn.setStyleSheet("")

    def filter_table(self, text):
        rows = self.table.rowCount()
        text = text.lower()
        for i in range(rows):
            match = False
            for j in range(3): # Check File, Name, Assigned
                item = self.table.item(i, j)
                if item and text in item.text().lower():
                    match = True
                    break
            self.table.setRowHidden(i, not match)

    def on_case_selected(self):
        rows = self.table.selectionModel().selectedRows()
        if not rows:
            self.details_widget.setEnabled(False)
            self.case_title.setText("Select a Case")
            return
            
        row = rows[0].row()
        file_number = self.table.item(row, 0).data(Qt.ItemDataRole.UserRole)
        name = self.table.item(row, 1).text()
        
        self.current_file_number = file_number
        self.case_title.setText(f"{file_number} - {name}")
        self.details_widget.setEnabled(True)
        
        # Load fields
        case = self.db.get_case(file_number)
        
        # Set Dates in QDateEdit (Expects YYYY-MM-DD in DB)
        self.edit_hearing.blockSignals(True)
        self.edit_trial.blockSignals(True)
        
        def set_qdate(widget, date_str):
            if date_str and "-" in date_str:
                parts = date_str.split("-")
                if len(parts) == 3:
                    widget.setDate(QDate(int(parts[0]), int(parts[1]), int(parts[2])))
                    return
            widget.setDate(QDate(2000, 1, 1)) # Default "empty" state for our logic
            
        set_qdate(self.edit_hearing, case.get('next_hearing_date', ''))
        set_qdate(self.edit_trial, case.get('trial_date', ''))
        
        self.edit_hearing.blockSignals(False)
        self.edit_trial.blockSignals(False)
        
        self.load_case_summary(file_number)
        
        self.refresh_details()

    def refresh_details(self):
        if not hasattr(self, 'current_file_number'): return
        
        # Get Case Info for Assigned Attorney
        case = self.db.get_case(self.current_file_number)
        case_assigned_attorney = case.get('assigned_attorney', '') if case else ''

        self.todo_list.blockSignals(True) 
        self.todo_list.clear()
        todos = self.db.get_todos(self.current_file_number)
        
        # Sort Logic: Red (0) > Yellow (1) > Green (2) > Blue (3) > Others
        # Then by ID desc (newest first)
        def sort_key(t):
            color = t.get('color', 'yellow')
            priority_map = {'red': 0, 'yellow': 1, 'green': 2, 'blue': 3}
            p = priority_map.get(color, 4)
            return (p, -t['id'])
            
        todos.sort(key=sort_key)
        
        for todo in todos:
            item = QListWidgetItem()
            
            created = todo.get('created_date', '')
            assigned = todo.get('assigned_to', '')
            assigned_date = todo.get('assigned_date', '')
            
            w = TodoItemWidget(
                todo['item'], 
                todo['status'], 
                todo.get('color', 'yellow'),
                created,
                assigned,
                assigned_date,
                case_assigned_attorney=case_assigned_attorney
            )
            
            w.statusChanged.connect(lambda s, tid=todo['id']: self.update_todo_status(tid, s))
            w.colorChanged.connect(lambda c, tid=todo['id']: self.update_todo_color(tid, c))
            w.assignmentChanged.connect(lambda a, d, tid=todo['id']: self.update_todo_assignment(tid, a, d))
            
            self.todo_list.addItem(item)
            item.setSizeHint(w.sizeHint())
            self.todo_list.setItemWidget(item, w)
            item.setData(Qt.ItemDataRole.UserRole, todo['id'])

        # History
        self.history_list.clear()
        hist = self.db.get_history(self.current_file_number)
        for h in hist:
            date_str = h['date']
            # Format to MM/DD/YY
            if date_str and "-" in date_str:
                try:
                    parts = date_str.split("-")
                    if len(parts) == 3:
                        date_str = f"{parts[1]}/{parts[2]}/{parts[0][2:]}"
                except: pass
                
            label = f"{date_str} [{h['type']}] {h['notes']}"
            item = QListWidgetItem(label)
            item.setData(Qt.ItemDataRole.UserRole, h['id']) # Store ID for editing
            item.setData(Qt.ItemDataRole.UserRole + 1, h['date']) # Store raw date
            self.history_list.addItem(item)
    
    def on_history_double_clicked(self, item):
        try:
            hist_id = item.data(Qt.ItemDataRole.UserRole)
            raw_date = item.data(Qt.ItemDataRole.UserRole + 1) # YYYY-MM-DD usually
            item_text = item.text() # Capture text before item might be deleted
            
            if not hist_id: return
            
            # Convert raw_date (YYYY-MM-DD) to MM/DD/YY for display in the dialog
            display_date = str(raw_date) if raw_date else ""
            
            if raw_date and "-" in raw_date:
                try:
                    parts = raw_date.split("-")
                    if len(parts) == 3:
                        # YYYY-MM-DD -> MM/DD/YY
                        display_date = f"{parts[1]}/{parts[2]}/{parts[0][2:]}"
                except: pass

            new_text, ok = QInputDialog.getText(self, "Edit Date", "Date (MM/DD/YY):", text=display_date)
            
            if ok and new_text:
                # Try to convert back to YYYY-MM-DD
                db_date = new_text.strip()
                try:
                    # Expecting MM/DD/YY or MM/DD/YYYY
                    if "/" in db_date:
                        parts = db_date.split("/")
                        if len(parts) == 3:
                            m, d, y = parts
                            if len(y) == 2: y = "20" + y
                            db_date = f"{y}-{m.zfill(2)}-{d.zfill(2)}"
                except:
                    pass
                
                self.db.update_history_date(hist_id, db_date)
                
                # Check text captured BEFORE refresh
                if "Status Update" in item_text:
                     self.refresh_data_row(self.current_file_number)
                
                # Refresh details last, as it clears the list and deletes 'item'
                self.refresh_details()
                
        except Exception as e:
            log_event(f"Error editing history date: {e}", "error")
            QMessageBox.warning(self, "Error", f"Could not edit date: {e}")

    def update_todo_status(self, tid, checked):
        status = 'done' if checked else 'pending'
        self.db.update_todo_status(tid, status)
        self.refresh_data()
        self.restore_selection()

    def update_todo_color(self, tid, color):
        self.db.update_todo_color(tid, color)
        # Refresh details to re-sort
        self.refresh_details()
        self.refresh_data()
        self.restore_selection()

    def update_todo_assignment(self, tid, assigned_to, assigned_date):
        self.db.update_todo_assignment(tid, assigned_to, assigned_date)

    def restore_selection(self):
        if hasattr(self, 'current_file_number'):
             for i in range(self.table.rowCount()):
                item = self.table.item(i, 0)
                if item and item.data(Qt.ItemDataRole.UserRole) == self.current_file_number:
                    self.table.selectRow(i)
                    break

    def save_case_summary_debounced(self):
        # We can reuse the existing save_timer or create a new one. 
        # The existing save_timer calls _perform_save_settings. 
        # Let's use a dedicated timer for content to avoid mixing settings/content.
        if not hasattr(self, 'summary_save_timer'):
            self.summary_save_timer = QTimer()
            self.summary_save_timer.setSingleShot(True)
            self.summary_save_timer.setInterval(1000) # 1s debounce
            self.summary_save_timer.timeout.connect(self._perform_save_summary)
        
        self.summary_save_timer.start()

    def _perform_save_summary(self):
        if hasattr(self, 'current_file_number'):
            text = self.summary_edit.toPlainText()
            self.db.update_case_summary(self.current_file_number, text)

    def load_case_summary(self, file_number):
        case = self.db.get_case(file_number)
        summary = case.get('case_summary')
        
        if not summary:
            # Try to load from variables JSON (factual_background)
            from icharlotte_core.config import GEMINI_DATA_DIR
            json_path = os.path.join(GEMINI_DATA_DIR, f"{file_number}.json")
            if os.path.exists(json_path):
                try:
                    with open(json_path, 'r', encoding='utf-8') as f:
                        vars_data = json.load(f)
                        
                    # Locate factual_background
                    fb = vars_data.get('factual_background')
                    if isinstance(fb, dict) and 'value' in fb:
                        summary = fb['value']
                    elif isinstance(fb, str):
                        summary = fb
                except:
                    pass
        
        self.summary_edit.blockSignals(True)
        self.summary_edit.setPlainText(summary or "")
        self.summary_edit.blockSignals(False)
        
        # If we loaded from file but DB was empty, save it to DB now to "init" it? 
        # Or wait for user edit? User said "autopopulated... persist". 
        # If we don't save now, next load will re-read file. 
        # If file changes, do we want to update? 
        # User said "changes made by the user... should not change procedural_history". 
        # Implied: This box is its own thing. 
        # Let's save to DB immediately if we pulled from factual_background so it becomes independent.
        if not case.get('case_summary') and summary:
            self.db.update_case_summary(file_number, summary)

    def save_dates(self):
        if not hasattr(self, 'current_file_number'): return
        
        # QDate to YYYY-MM-DD string
        h_qdate = self.edit_hearing.date()
        t_qdate = self.edit_trial.date()
        
        # Heuristic: If year is 2000, treat as empty
        hearing = h_qdate.toString("yyyy-MM-dd") if h_qdate.year() > 2000 else ""
        trial = t_qdate.toString("yyyy-MM-dd") if t_qdate.year() > 2000 else ""
        
        # Get existing case to preserve other fields
        case = self.db.get_case(self.current_file_number)
        if case:
            self.db.upsert_case(
                self.current_file_number, 
                case['plaintiff_last_name'], 
                hearing, 
                trial, 
                case['case_path'],
                plaintiff_override=case.get('plaintiff_override', 0)
            )
            # Full refresh to handle formatting and sorting properly
            self.refresh_data()
            
            # Reselect the row
            for i in range(self.table.rowCount()):
                if self.table.item(i, 0).data(Qt.ItemDataRole.UserRole) == self.current_file_number:
                    self.table.selectRow(i)
                    break

    def refresh_data_row(self, file_number):
        self.refresh_data()
        self.restore_selection()


    def add_todo(self):
        text = self.todo_input.text().strip()
        if text and hasattr(self, 'current_file_number'):
            self.db.add_todo(self.current_file_number, text, color='yellow')
            self.todo_input.clear()
            self.refresh_details()
            self.refresh_data()
            self.restore_selection()

    def add_history(self, file_number=None, type_val=None, notes=None):
        # Allow arguments or use UI inputs
        if file_number is None:
            # Called from UI returnPressed
            text = self.hist_input.text().strip()
            type_val = self.hist_type.currentText()
            if text and hasattr(self, 'current_file_number'):
                self.db.add_history(self.current_file_number, type_val, text)
                self.hist_input.clear()
                self.refresh_details()
                if type_val == "Status Update":
                    self.refresh_data()
                    self.restore_selection()
        else:
            # Programmatic add
            self.db.add_history(file_number, type_val, notes)
            
    def show_history_menu(self, pos):
        item = self.history_list.itemAt(pos)
        if not item: return

        menu = QMenu()
        hist_id = item.data(Qt.ItemDataRole.UserRole)
        
        change_cat_action = QAction("Change Category", self)
        change_cat_action.triggered.connect(lambda: self.change_history_category(hist_id))
        menu.addAction(change_cat_action)
        
        del_action = QAction("Delete", self)
        del_action.triggered.connect(lambda: self.delete_history(hist_id))
        menu.addAction(del_action)
        
        menu.exec(self.history_list.mapToGlobal(pos))

    def change_history_category(self, hist_id):
        categories = ["Status Update", "Note"]
        new_cat, ok = QInputDialog.getItem(self, "Change Category", "Select Category:", categories, 0, False)
        if ok and new_cat:
            self.db.update_history_type(hist_id, new_cat)
            self.refresh_details()
            # If it became a Status Update (or stopped being one), we might need to refresh the main table
            self.refresh_data_row(self.current_file_number)

    def delete_history(self, hist_id):
        confirm = QMessageBox.question(
            self, 
            "Confirm Delete", 
            "Are you sure you want to delete this history entry?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        if confirm == QMessageBox.StandardButton.Yes:
            self.db.delete_history(hist_id)
            self.refresh_details()
            self.refresh_data_row(self.current_file_number)

    def show_todo_menu(self, pos):
        item = self.todo_list.itemAt(pos)
        if not item: return
        
        menu = QMenu()
        tid = item.data(Qt.ItemDataRole.UserRole)
        # We can't toggle from context menu as easily with custom widget, 
        # but delete is fine.
        del_action = QAction("Delete", self)
        del_action.triggered.connect(lambda: self.delete_todo(tid))
        menu.addAction(del_action)
        
        menu.exec(self.todo_list.mapToGlobal(pos))
        
    def delete_current_todo(self):
        item = self.todo_list.currentItem()
        if item:
            tid = item.data(Qt.ItemDataRole.UserRole)
            self.delete_todo(tid)

    def delete_current_history(self):
        item = self.history_list.currentItem()
        if item:
            hist_id = item.data(Qt.ItemDataRole.UserRole)
            self.delete_history(hist_id)

    def delete_todo(self, tid):
        self.db.delete_todo(tid)
        self.refresh_details()
        self.refresh_data()
        self.restore_selection()

    def on_case_double_clicked(self, item):
        # Only load case and switch to Case View when double-clicking the Tasks column (index 6)
        if item.column() != 6:
            return

        row = item.row()
        file_number = self.table.item(row, 0).data(Qt.ItemDataRole.UserRole)

        if self.main_window:
            # Trigger switch in MainWindow
            if hasattr(self.main_window, 'load_case_by_number'):
                self.main_window.load_case_by_number(file_number)
            else:
                QMessageBox.information(self, "Info", f"Selected {file_number}. (Switching logic pending)")

    def on_cell_clicked(self, row, col):
        file_number = self.table.item(row, 0).data(Qt.ItemDataRole.UserRole)
        
        if col == 0: # File Number -> Open Folder
            case_path = get_case_path(file_number)
            if case_path and os.path.exists(case_path):
                try:
                    os.startfile(case_path)
                except Exception as e:
                    QMessageBox.warning(self, "Error", f"Could not open folder: {e}")
            else:
                QMessageBox.warning(self, "Error", f"Folder not found for {file_number}")
                
        elif col == 1: # Plaintiff Name -> Edit Text
            current_name = self.table.item(row, col).text()
            new_name, ok = QInputDialog.getText(self, "Edit Plaintiff", "Plaintiff Name:", text=current_name)
            if ok:
                self.db.update_plaintiff(file_number, new_name)
                self.table.item(row, col).setText(new_name)
                
        elif col == 2: # Assigned -> Edit Text
            current_val = self.table.item(row, col).text()
            # Offer defaults + editable
            items = ["HP", "FM", "JBW", "AS", "CE"]
            if current_val and current_val not in items:
                items.append(current_val)
                
            new_val, ok = QInputDialog.getItem(self, "Edit Assigned", "Attorney Initials:", items, 0, True)
            if ok:
                # User can type custom text in getItem if editable=True
                self.db.update_assigned_attorney(file_number, new_val)
                self.table.item(row, col).setText(new_val)
                
        elif col == 3: # Hearings -> Edit Text
            from icharlotte_core.config import GEMINI_DATA_DIR
            
            # Load current "other_hearings" from JSON
            json_path = os.path.join(GEMINI_DATA_DIR, f"{file_number}.json")
            current_val = ""
            vars_data = {}
            
            if os.path.exists(json_path):
                try:
                    with open(json_path, 'r', encoding='utf-8') as f:
                        vars_data = json.load(f)
                        # Check for other_hearings or other_hearing
                        v = vars_data.get("other_hearings") or vars_data.get("other_hearing")
                        if isinstance(v, dict) and "value" in v:
                            current_val = v["value"]
                        elif isinstance(v, str):
                            current_val = v
                except: pass
            
            # If no other_hearings found in JSON, fall back to DB 'next_hearing_date'
            if not current_val:
                case_info = self.db.get_case(file_number)
                if case_info:
                    current_val = case_info.get('next_hearing_date', "")

            # Show Dialog
            new_text, ok = QInputDialog.getMultiLineText(
                self, 
                "Edit Hearings", 
                "Hearings (One per line, e.g. 'CMC 2/5/26'):", 
                text=current_val
            )
            
            if ok:
                # Update DB 'next_hearing_date' for sorting/sync logic moved UP
                parsed = parse_hearing_data(new_text)
                today = datetime.datetime.now().date()
                future = [h for h in parsed if h['date_obj'].date() >= today or h['date_obj'].year == 9999]
                
                new_db_date = ""
                if future:
                    d = future[0]['date_sort']
                    if d != "9999-99-99":
                        new_db_date = d

                # Save to JSON
                if isinstance(vars_data.get("other_hearings"), dict):
                     vars_data["other_hearings"]["value"] = new_text
                else:
                     vars_data["other_hearings"] = new_text
                
                # Sync next_hearing_date to JSON as well
                if isinstance(vars_data.get("next_hearing_date"), dict):
                    vars_data["next_hearing_date"]["value"] = new_db_date
                else:
                    vars_data["next_hearing_date"] = new_db_date

                try:
                    os.makedirs(os.path.dirname(json_path), exist_ok=True)
                    with open(json_path, 'w', encoding='utf-8') as f:
                        json.dump(vars_data, f, indent=4)
                except Exception as e:
                    QMessageBox.warning(self, "Error", f"Could not save hearing data: {e}")
                
                self.db.update_hearing_date(file_number, new_db_date)
                self.refresh_data_row(file_number)

        elif col == 4: # Trial Date -> Calendar
            current_text = self.table.item(row, col).text()
            current_date = QDate.currentDate()
            if current_text and current_text != "N/A":
                try:
                    parts = current_text.split('/')
                    if len(parts) == 3:
                        current_date = QDate(int(parts[2]), int(parts[0]), int(parts[1]))
                except: pass
                
            dialog = CalendarDialog(self, current_date)
            if dialog.exec() == QDialog.DialogCode.Accepted:
                if dialog.cleared:
                    self.db.update_trial_date(file_number, "")
                else:
                    date_str = dialog.selected_date.toString("yyyy-MM-dd")
                    self.db.update_trial_date(file_number, date_str)
                self.refresh_data_row(file_number)
                
        elif col == 5: # Last Report -> Text/Date
            # Allow text or date. Just use InputDialog for maximum flexibility as requested ("date / text")
            current_val = self.table.item(row, col).text()
            new_val, ok = QInputDialog.getText(self, "Edit Last Report", "Date (YYYY-MM-DD) or Text:", text=current_val)
            if ok:
                self.db.update_last_report_text(file_number, new_val)
                self.refresh_data_row(file_number)

    # --- Email Monitor Methods ---

    def _auto_start_email_monitor(self):
        """Auto-start the email monitor on app launch."""
        self.email_monitor_btn.setChecked(True)
        self.start_email_monitor()

    def toggle_email_monitor(self, checked):
        """Toggle the email monitor on/off."""
        if checked:
            self.start_email_monitor()
        else:
            self.stop_email_monitor()

    def start_email_monitor(self):
        """Start the sent items monitor worker."""
        if self.email_monitor_worker is not None and self.email_monitor_worker.isRunning():
            return  # Already running

        self.email_monitor_worker = SentItemsMonitorWorker(self.db)
        self.email_monitor_worker.todo_created.connect(self.on_email_todo_created)
        self.email_monitor_worker.error.connect(self.on_email_monitor_error)
        self.email_monitor_worker.status.connect(self.on_email_monitor_status)
        self.email_monitor_worker.finished.connect(self.on_email_monitor_finished)
        self.email_monitor_worker.start()

        self.email_monitor_btn.setText("Stop Email Monitor")
        self.email_monitor_status.setText("Running...")
        self.email_monitor_status.setStyleSheet("color: green; font-size: 10px; margin-left: 5px;")

    def stop_email_monitor(self):
        """Stop the sent items monitor worker."""
        if self.email_monitor_worker is None:
            return

        self.email_monitor_worker.request_stop()
        self.email_monitor_status.setText("Stopping...")
        self.email_monitor_status.setStyleSheet("color: orange; font-size: 10px; margin-left: 5px;")

    def on_email_todo_created(self, file_number, todo_text):
        """Handle todo created signal from email monitor."""
        # Refresh the UI
        self.refresh_data()

        # If currently viewing this case, refresh details
        if hasattr(self, 'current_file_number') and self.current_file_number == file_number:
            self.refresh_details()
            self.restore_selection()

        # Update status
        self.email_monitor_status.setText(f"Todo added: {file_number}")

    def on_email_monitor_error(self, message):
        """Handle error signal from email monitor."""
        self.email_monitor_status.setText(f"Error: {message[:30]}")
        self.email_monitor_status.setStyleSheet("color: red; font-size: 10px; margin-left: 5px;")
        log_event(f"Email monitor error: {message}", "error")

    def on_email_monitor_status(self, message):
        """Handle status signal from email monitor."""
        # Show truncated status
        display_msg = message[:40] + "..." if len(message) > 40 else message
        self.email_monitor_status.setText(display_msg)

    def on_email_monitor_finished(self):
        """Handle worker finished signal."""
        self.email_monitor_btn.setChecked(False)
        self.email_monitor_btn.setText("Start Email Monitor")
        self.email_monitor_status.setText("Stopped")
        self.email_monitor_status.setStyleSheet("color: gray; font-size: 10px; margin-left: 5px;")
        self.email_monitor_worker = None

    # --- Calendar Monitor Methods ---

    def _auto_start_calendar_monitor(self):
        """Auto-start the calendar monitor on app launch."""
        self.calendar_monitor_btn.setChecked(True)
        self.start_calendar_monitor()

    def toggle_calendar_monitor(self, checked):
        """Toggle the calendar monitor on/off."""
        if checked:
            self.start_calendar_monitor()
        else:
            self.stop_calendar_monitor()

    def start_calendar_monitor(self):
        """Start the calendar monitor worker."""
        if self.calendar_monitor_worker is not None and self.calendar_monitor_worker.isRunning():
            return  # Already running

        self.calendar_monitor_worker = CalendarMonitorWorker()
        self.calendar_monitor_worker.calendar_event_created.connect(self.on_calendar_event_created)
        self.calendar_monitor_worker.error.connect(self.on_calendar_monitor_error)
        self.calendar_monitor_worker.status.connect(self.on_calendar_monitor_status)
        self.calendar_monitor_worker.finished.connect(self.on_calendar_monitor_finished)
        self.calendar_monitor_worker.start()

        self.calendar_monitor_btn.setText("Stop Calendar")
        self.calendar_monitor_status.setText("Starting...")
        self.calendar_monitor_status.setStyleSheet("color: blue; font-size: 10px; margin-left: 5px;")

    def stop_calendar_monitor(self):
        """Stop the calendar monitor worker."""
        if self.calendar_monitor_worker is None:
            return

        self.calendar_monitor_worker.request_stop()
        self.calendar_monitor_status.setText("Stopping...")
        self.calendar_monitor_status.setStyleSheet("color: orange; font-size: 10px; margin-left: 5px;")

    def on_calendar_event_created(self, file_number, event_title):
        """Handle calendar event created signal."""
        self.calendar_monitor_status.setText(f"Event: {file_number}")
        self.calendar_monitor_status.setStyleSheet("color: green; font-size: 10px; margin-left: 5px;")

    def on_calendar_monitor_error(self, message):
        """Handle error signal from calendar monitor."""
        self.calendar_monitor_status.setText(f"Error: {message[:30]}")
        self.calendar_monitor_status.setStyleSheet("color: red; font-size: 10px; margin-left: 5px;")
        log_event(f"Calendar monitor error: {message}", "error")

    def on_calendar_monitor_status(self, message):
        """Handle status signal from calendar monitor."""
        display_msg = message[:40] + "..." if len(message) > 40 else message
        self.calendar_monitor_status.setText(display_msg)
        if "Running" in message or "authenticated" in message.lower():
            self.calendar_monitor_status.setStyleSheet("color: green; font-size: 10px; margin-left: 5px;")

    def on_calendar_monitor_finished(self):
        """Handle worker finished signal."""
        self.calendar_monitor_btn.setChecked(False)
        self.calendar_monitor_btn.setText("Start Calendar")
        self.calendar_monitor_status.setText("Stopped")
        self.calendar_monitor_status.setStyleSheet("color: gray; font-size: 10px; margin-left: 5px;")
        self.calendar_monitor_worker = None