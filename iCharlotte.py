import sys
import os

# Load environment variables from .env file
from dotenv import load_dotenv
load_dotenv()

# MUST register schemes before QApplication is created
from icharlotte_core.bridge import register_custom_schemes
register_custom_schemes()

import re
import glob
import json
import uuid
import subprocess
import datetime
from functools import partial

# --- Imports ---
try:
    from PyQt6.QtWidgets import (
        QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
        QPushButton, QTreeWidgetItem, QHeaderView, QMessageBox, QLabel, 
        QFrame, QSplitter, QAbstractItemView, QLineEdit, 
        QTreeWidgetItemIterator, QTabWidget, QScrollArea, QMenu, QDialog,
        QFileIconProvider, QToolButton
    )
    from PyQt6.QtCore import Qt, QThread, pyqtSignal, QFileInfo
    from PyQt6.QtGui import QAction
    from PyQt6.QtWebEngineCore import QWebEngineUrlScheme
except ImportError:
    print("Error: PyQt6 or its components are not installed. Please run: pip install PyQt6 PyQt6-WebEngine")
    sys.exit(1)

# --- Core Modules ---
from icharlotte_core.config import SCRIPTS_DIR, GEMINI_DATA_DIR, BASE_PATH_WIN
from icharlotte_core.utils import (
    log_event, get_case_path, sanitize_filename, format_date_to_mm_dd_yyyy
)
from icharlotte_core.ui.widgets import (
    StatusWidget, AgentRunner, FileTreeWidget
)
from icharlotte_core.ui.dialogs import FileNumberDialog, VariablesDialog, PromptsDialog
from icharlotte_core.ui.tabs import ChatTab, IndexTab
from icharlotte_core.ui.browser import NoteTakerTab
from icharlotte_core.ui.email_tab import EmailTab
from icharlotte_core.ui.email_update_tab import EmailUpdateTab
from icharlotte_core.ui.report_tab import ReportTab
from icharlotte_core.ui.logs_tab import LogsTab
from icharlotte_core.ui.liability_tab import LiabilityExposureTab
from icharlotte_core.ui.master_case_tab import MasterCaseTab

class DirectoryTreeWorker(QThread):
    data_ready = pyqtSignal(list) # Emits (root, dirs, files) tuples
    finished = pyqtSignal()
    
    def __init__(self, root_path):
        super().__init__()
        self.root_path = root_path
        self.running = True
        
    def run(self):
        batch = []
        for root, dirs, files in os.walk(self.root_path):
            if not self.running:
                break
            # Skip hidden files/dirs
            dirs[:] = [d for d in dirs if not d.startswith('.') and not d.startswith('~$')]
            
            file_data = []
            for f in files:
                if f.startswith('.') or f.startswith('~$'):
                    continue
                path = os.path.join(root, f)
                try:
                    stat = os.stat(path)
                    size = stat.st_size
                    mtime = stat.st_mtime
                    file_data.append((f, size, mtime))
                except:
                    file_data.append((f, 0, 0))
            
            batch.append((root, dirs, file_data))
            if len(batch) >= 10:
                self.data_ready.emit(batch)
                batch = []
                self.msleep(10) # Yield to UI
                
        if batch:
            self.data_ready.emit(batch)
        self.finished.emit()

    def stop(self):
        self.running = False

class MainWindow(QMainWindow):
    def __init__(self, file_number, case_path):
        super().__init__()
        self.file_number = file_number
        self.case_path = case_path
        self.setWindowTitle(f"iCharlotte - {self.file_number} - {os.path.basename(self.case_path)}")
        self.resize(1200, 800)
        
        self.agent_runners = [] # Keep references to prevent GC
        self.cached_models = {} # Cache for models: {provider: [list]}
        self.fetcher = None
        
        log_event(f"Initializing MainWindow for {file_number} at {case_path}")
        self.icon_provider = QFileIconProvider()
        self.setup_ui()
        self.populate_tree()
        self.load_status_history()
        self.check_docket_expiry(file_number)

    def setup_ui(self):
        self.tabs = QTabWidget()
        self.setCentralWidget(self.tabs)
        
        # --- Tab 0: Master List ---
        self.master_tab = MasterCaseTab(self)
        self.tabs.addTab(self.master_tab, "Master List")

        # --- Tab 1: Case View ---
        case_view_widget = QWidget()
        self.tabs.addTab(case_view_widget, "Case View")
        
        main_layout = QVBoxLayout(case_view_widget)
        
        # Top Toolbar
        toolbar_layout = QHBoxLayout()
        
        btn_view_docket = QPushButton("ViewDocket")
        btn_view_docket.clicked.connect(self.view_docket)
        toolbar_layout.addWidget(btn_view_docket)
        
        btn_vars = QPushButton("Variables")
        btn_vars.clicked.connect(self.manage_variables)
        toolbar_layout.addWidget(btn_vars)
        
        btn_prompts = QPushButton("Prompts")
        btn_prompts.clicked.connect(self.manage_prompts)
        toolbar_layout.addWidget(btn_prompts)
        
        toolbar_layout.addStretch()
        
        # Wrapper for vertical layout of Case View
        wrapper_layout = QVBoxLayout()
        wrapper_layout.addLayout(toolbar_layout)
        
        splitter = QSplitter(Qt.Orientation.Horizontal)
        wrapper_layout.addWidget(splitter)
        
        main_layout.addLayout(wrapper_layout)
        
        # Left Panel (Case Agents)
        left_panel = QFrame()
        left_panel.setFrameShape(QFrame.Shape.StyledPanel)
        left_panel.setFixedWidth(170)
        left_layout = QVBoxLayout(left_panel)
        left_layout.setAlignment(Qt.AlignmentFlag.AlignTop)
        left_layout.setContentsMargins(2, 5, 2, 5)
        
        title_label = QLabel("Case Agents")
        title_label.setStyleSheet("font-weight: bold; font-size: 14px; margin-bottom: 10px;")
        left_layout.addWidget(title_label)
        
        # Agent Buttons
        self.create_agent_button("Docket Agent", "docket.py", left_layout, arg_type="file_number")
        self.create_agent_button("Complaint Agent", "complaint.py", left_layout, arg_type="file_number")
        self.create_agent_button("Report Agent", "report.py", left_layout, arg_type="file_number")
        self.create_agent_button("Discovery Generate Agent", "discovery_requests.py", left_layout, arg_type="file_number", extra_flags=["--interactive"])
        self.create_agent_button("Subpoena Tracker Agent", "subpoena_tracker.py", left_layout, arg_type="file_number")
        
        # Directory-based Case Agents
        left_layout.addSpacing(10)
        self.create_agent_button("Liability Agent", "liability.py", left_layout, arg_type="case_path")
        self.create_agent_button("Exposure Agent", "exposure.py", left_layout, arg_type="case_path")
        
        left_layout.addStretch()
        splitter.addWidget(left_panel)
        
        # Right Panel (File Tree)
        right_panel = QFrame()
        right_layout = QVBoxLayout(right_panel)
        
        # Header Layout (Status Label + Expand/Collapse Button)
        header_layout = QHBoxLayout()
        
        self.refresh_btn = QPushButton()
        self.refresh_btn.setIcon(self.style().standardIcon(self.style().StandardPixmap.SP_BrowserReload))
        self.refresh_btn.setToolTip("Refresh Tree")
        self.refresh_btn.clicked.connect(self.populate_tree)
        header_layout.addWidget(self.refresh_btn)

        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Filter files...")
        self.search_input.textChanged.connect(self.filter_tree)
        header_layout.addWidget(self.search_input)
        
        self.smart_select_btn = QPushButton("Select...")
        self.smart_select_menu = QMenu(self)
        
        act_pdfs = QAction("All PDFs -> OCR", self)
        act_pdfs.triggered.connect(lambda: self.smart_select("pdf_ocr"))
        self.smart_select_menu.addAction(act_pdfs)
        
        self.smart_select_btn.setMenu(self.smart_select_menu)
        header_layout.addWidget(self.smart_select_btn)
        
        self.expand_btn = QPushButton("Expand All")
        self.expand_btn.setCheckable(True)
        self.expand_btn.clicked.connect(self.toggle_expand)
        header_layout.addWidget(self.expand_btn)

        self.clear_all_btn = QPushButton("Clear All")
        self.clear_all_btn.clicked.connect(self.clear_all_checkboxes)
        header_layout.addWidget(self.clear_all_btn)
        
        right_layout.addLayout(header_layout)
        
        self.status_label = QLabel("Ready")
        right_layout.addWidget(self.status_label)
        
        # AGENTS definition for the Task Queue
        self.AGENTS = [
            {"id": "separate", "name": "Separate", "script": "separate.py", "color": "#e91e63", "short": "SEP"},
            {"id": "summarize", "name": "Summarize", "script": "summarize.py", "color": "#2196f3", "short": "SUM"},
            {"id": "sum_disc", "name": "Sum. Disc.", "script": "summarize_discovery.py", "color": "#4caf50", "short": "DISC"},
            {"id": "sum_depo", "name": "Sum. Depo.", "script": "summarize_deposition.py", "color": "#ff9800", "short": "DEPO"},
            {"id": "med_rec", "name": "Med Rec", "script": "med_record.py", "color": "#9c27b0", "short": "MED"},
            {"id": "med_chron", "name": "Med Chron", "script": "med_chron.py", "color": "#00bcd4", "short": "CHRON"},
            {"id": "ocr", "name": "OCR", "script": "ocr.py", "color": "#795548", "short": "OCR"},
            {"id": "organize", "name": "Organize", "script": "organizer.py", "color": "#607d8b", "short": "ORG"},
        ]

        self.tree = FileTreeWidget()
        self.tree.item_moved.connect(lambda: self.populate_tree())
        self.tree.setHeaderLabels([
            "Category / File", 
            "Size",
            "Date Modified",
            "Queued Tasks (Click to Add ➕)"
        ])
        self.tree.setSortingEnabled(True)
        self.tree.setColumnWidth(1, 80)
        self.tree.setColumnWidth(2, 120)
        self.tree.setColumnWidth(3, 250)
        self.tree.header().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.tree.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.tree.setAlternatingRowColors(True)
        self.tree.itemSelectionChanged.connect(self.on_tree_selection_changed)
        self.tree.itemDoubleClicked.connect(self.on_tree_double_click) 
        self.tree.itemClicked.connect(self.on_tree_item_clicked)
        right_layout.addWidget(self.tree)
        
        btn_layout = QHBoxLayout()
        self.process_btn = QPushButton("Process All Queued Tasks")
        self.process_btn.setStyleSheet("background-color: #4CAF50; color: white; font-weight: bold; padding: 10px;")
        self.process_btn.clicked.connect(self.process_checked_items)
        btn_layout.addWidget(self.process_btn)

        self.organize_btn = QPushButton("Quick Organize (AI)")
        self.organize_btn.setStyleSheet("background-color: #2196F3; color: white; font-weight: bold; padding: 10px;")
        self.organize_btn.clicked.connect(self.organize_checked_items)
        btn_layout.addWidget(self.organize_btn)
        
        right_layout.addLayout(btn_layout)
        
        splitter.addWidget(right_panel)
        splitter.setSizes([170, 1030])
        splitter.setStretchFactor(0, 0)
        splitter.setStretchFactor(1, 1)
        
        # --- Tab 2: Status ---
        self.status_tab = QWidget()
        self.tabs.addTab(self.status_tab, "Status")
        status_layout = QVBoxLayout(self.status_tab)
        
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        self.status_container = QWidget()
        self.status_list_layout = QVBoxLayout(self.status_container)
        self.status_list_layout.setAlignment(Qt.AlignmentFlag.AlignTop)
        
        scroll.setWidget(self.status_container)
        status_layout.addWidget(scroll)
        
        clear_btn = QPushButton("Clear Completed")
        clear_btn.clicked.connect(self.clear_completed_status)
        status_layout.addWidget(clear_btn)

        # --- Tab 3: Index ---
        self.index_tab = IndexTab(self)
        self.tabs.addTab(self.index_tab, "Index")
        if self.file_number:
            self.index_tab.load_data(self.file_number)

        # --- Tab 4: Note Taker ---
        self.note_taker_tab = NoteTakerTab(self)
        self.tabs.addTab(self.note_taker_tab, "Note Taker")

        # --- Tab 5: Chat ---
        self.chat_tab = ChatTab()
        self.tabs.addTab(self.chat_tab, "Chat")

        # --- Tab 6: Email ---
        self.email_tab = EmailTab()
        self.tabs.addTab(self.email_tab, "Email")
        if self.file_number:
            # Force initialization now that it's part of the window hierarchy
            # (check_db_init relies on self.window().file_number which works now)
            self.email_tab.check_db_init() 
            self.email_tab.perform_search()

        # --- Tab: Email Update ---
        self.email_update_tab = EmailUpdateTab()
        self.tabs.addTab(self.email_update_tab, "Email Update")

        # --- Tab 7: Report ---
        self.report_tab = ReportTab(main_window=self)
        self.tabs.addTab(self.report_tab, "Report")

        # --- Tab: Liability & Exposure ---
        self.liability_tab = LiabilityExposureTab()
        self.tabs.addTab(self.liability_tab, "Liability & Exposure")

        # --- Tab 8: Logs ---
        self.logs_tab = LogsTab(self)
        self.tabs.addTab(self.logs_tab, "Logs")

        # Add Restart and View Buttons next to the tabs
        self.corner_widget = QWidget()
        self.corner_layout = QHBoxLayout(self.corner_widget)
        self.corner_layout.setContentsMargins(0, 0, 0, 0)
        self.corner_layout.setSpacing(5)

        self.btn_open_root = QPushButton("Open File")
        self.btn_open_root.clicked.connect(self.open_root_folder)
        self.corner_layout.addWidget(self.btn_open_root)

        self.btn_change_file = QPushButton("Change File")
        self.btn_change_file.clicked.connect(self.change_file)
        self.corner_layout.addWidget(self.btn_change_file)

        self.setup_view_menu()
        self.corner_layout.addWidget(self.view_btn)

        self.restart_btn = QPushButton("Restart")
        self.restart_btn.setStyleSheet("background-color: #f44336; color: white; font-weight: bold;")
        self.restart_btn.clicked.connect(self.restart_app)
        self.corner_layout.addWidget(self.restart_btn)

        self.tabs.setCornerWidget(self.corner_widget, Qt.Corner.TopRightCorner)

    def setup_view_menu(self):
        self.view_btn = QToolButton()
        self.view_btn.setText("View ▾") 
        self.view_btn.setPopupMode(QToolButton.ToolButtonPopupMode.InstantPopup)
        self.view_btn.setStyleSheet("font-weight: bold; padding: 5px;")
        
        self.view_menu = QMenu(self)
        self.view_btn.setMenu(self.view_menu)
        
        # Load and apply settings
        settings = self.load_tab_settings()
        
        # Populate menu based on current tabs
        for i in range(self.tabs.count()):
            tab_text = self.tabs.tabText(i)
            
            # Apply saved setting if available
            if tab_text in settings:
                self.tabs.setTabVisible(i, settings[tab_text])
            
            action = QAction(tab_text, self)
            action.setCheckable(True)
            action.setChecked(self.tabs.isTabVisible(i))
            # Use partial to capture the current loop variable 'i'
            action.toggled.connect(partial(self.toggle_tab_visibility, i))
            self.view_menu.addAction(action)

    def toggle_tab_visibility(self, index, visible):
        self.tabs.setTabVisible(index, visible)
        self.save_tab_settings()

    def load_tab_settings(self):
        config_dir = os.path.join(GEMINI_DATA_DIR, "..", "config")
        settings_path = os.path.join(config_dir, "view_settings.json")
        if os.path.exists(settings_path):
            try:
                with open(settings_path, 'r') as f:
                    return json.load(f)
            except Exception as e:
                log_event(f"Error loading view settings: {e}", "error")
        return {}

    def save_tab_settings(self):
        config_dir = os.path.join(GEMINI_DATA_DIR, "..", "config")
        if not os.path.exists(config_dir):
            try:
                os.makedirs(config_dir)
            except:
                pass
            
        settings_path = os.path.join(config_dir, "view_settings.json")
        settings = {}
        for i in range(self.tabs.count()):
            settings[self.tabs.tabText(i)] = self.tabs.isTabVisible(i)
            
        try:
            with open(settings_path, 'w') as f:
                json.dump(settings, f)
        except Exception as e:
            log_event(f"Error saving view settings: {e}", "error")

    def restart_app(self):
        log_event("User requested manual restart. Spawning new process...")
        # Close all agent runners if any are running
        for runner in self.agent_runners:
            try:
                runner.terminate()
            except:
                pass
        
        # Get the current arguments
        args = sys.argv[:]
        
        # Spawn new process
        subprocess.Popen([sys.executable] + args)
        
        # Exit current process
        QApplication.quit()

    def clear_all_checkboxes(self):
        self.tree.blockSignals(True)
        iterator = QTreeWidgetItemIterator(self.tree)
        while iterator.value():
            item = iterator.value()
            item.setData(0, Qt.ItemDataRole.UserRole + 2, [])
            self.update_item_tasks_ui(item)
            iterator += 1
        self.tree.blockSignals(False)

    def on_tree_item_clicked(self, item, column):
        if column == 3:
            # Get click position relative to the tree's viewport
            pos = self.tree.visualItemRect(item).bottomLeft()
            pos.setX(self.tree.columnViewportPosition(3))
            global_pos = self.tree.viewport().mapToGlobal(pos)
            self.show_agent_menu(item, global_pos)

    def show_agent_menu(self, item, global_pos):
        menu = QMenu(self)
        current_tasks = item.data(0, Qt.ItemDataRole.UserRole + 2) or []
        
        for agent in self.AGENTS:
            action = QAction(agent["name"], menu)
            action.setCheckable(True)
            action.setChecked(agent["id"] in current_tasks)
            action.triggered.connect(partial(self.toggle_agent_task, item, agent["id"]))
            menu.addAction(action)
            
        menu.addSeparator()
        clear_act = QAction("Clear All Tasks", menu)
        clear_act.triggered.connect(lambda: self.set_item_tasks(item, []))
        menu.addAction(clear_act)
        
        menu.exec(global_pos)

    def toggle_agent_task(self, item, agent_id):
        current_tasks = item.data(0, Qt.ItemDataRole.UserRole + 2) or []
        if agent_id in current_tasks:
            current_tasks.remove(agent_id)
        else:
            current_tasks.append(agent_id)
        self.set_item_tasks(item, current_tasks)

    def set_item_tasks(self, item, task_ids, recursive=True):
        self.tree.blockSignals(True)
        item.setData(0, Qt.ItemDataRole.UserRole + 2, task_ids)
        self.update_item_tasks_ui(item)
        
        if recursive:
            for i in range(item.childCount()):
                self._recursive_set_tasks(item.child(i), task_ids)
        self.tree.blockSignals(False)

    def _recursive_set_tasks(self, item, task_ids):
        item.setData(0, Qt.ItemDataRole.UserRole + 2, task_ids)
        self.update_item_tasks_ui(item)
        for i in range(item.childCount()):
            self._recursive_set_tasks(item.child(i), task_ids)

    def update_item_tasks_ui(self, item):
        task_ids = item.data(0, Qt.ItemDataRole.UserRole + 2) or []
        if not task_ids:
            item.setText(3, " [ + Add Tasks ]")
            item.setForeground(3, Qt.GlobalColor.gray)
            return
            
        # Create a visual string of short tags
        tags = []
        for tid in task_ids:
            agent = next((a for a in self.AGENTS if a["id"] == tid), None)
            if agent:
                tags.append(f"[{agent['short']}]")
        
        item.setText(3, " ".join(tags))
        # Optional: set color for column 3 text
        item.setForeground(3, Qt.GlobalColor.blue)

    def filter_tree(self, text):
        search_text = text.lower()
        iterator = QTreeWidgetItemIterator(self.tree)
        
        # If empty, show all
        if not search_text:
            while iterator.value():
                item = iterator.value()
                item.setHidden(False)
                iterator += 1
            return

        # First pass: Hide all
        items = []
        while iterator.value():
            items.append(iterator.value())
            iterator += 1
            
        for item in items:
            item.setHidden(True)
            
        # Second pass: Show matches and parents
        for item in items:
            if search_text in item.text(0).lower():
                item.setHidden(False)
                parent = item.parent()
                while parent:
                    parent.setHidden(False)
                    parent.setExpanded(True) # Expand path to match
                    parent = parent.parent()

    def smart_select(self, mode):
        self.tree.blockSignals(True)
        iterator = QTreeWidgetItemIterator(self.tree)
        while iterator.value():
            item = iterator.value()
            path = item.data(0, Qt.ItemDataRole.UserRole)
            is_file = item.data(0, Qt.ItemDataRole.UserRole + 1) == "file"
            
            if is_file and path:
                if mode == "pdf_ocr":
                    if path.lower().endswith(".pdf"):
                        current_tasks = item.data(0, Qt.ItemDataRole.UserRole + 2) or []
                        if "ocr" not in current_tasks:
                            current_tasks.append("ocr")
                            item.setData(0, Qt.ItemDataRole.UserRole + 2, current_tasks)
                            self.update_item_tasks_ui(item)
            
            iterator += 1
        self.tree.blockSignals(False)

    def toggle_expand(self):
        if self.expand_btn.isChecked():
            self.tree.expandAll()
            self.expand_btn.setText("Collapse to Root")
        else:
            self.tree.collapseAll()
            # Expand only the top-level root item
            if self.tree.topLevelItemCount() > 0:
                self.tree.topLevelItem(0).setExpanded(True)
            self.expand_btn.setText("Expand All")

    def change_file(self):
        dialog = FileNumberDialog(self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            new_file_num = dialog.get_file_number()
            new_path = get_case_path(new_file_num)
            
            if new_path:
                self.save_status_history()
                self.file_number = new_file_num
                self.case_path = new_path
                self.setWindowTitle(f"iCharlotte - {self.file_number} - {os.path.basename(self.case_path)}")
                self.populate_tree()
                # Clear status list to reset state
                self.clear_all_status()
                self.load_status_history()
                
                # Reset Tabs for new case isolation
                if hasattr(self, 'index_tab'):
                    self.index_tab.load_data(self.file_number)
                if hasattr(self, 'chat_tab'):
                    self.chat_tab.reset_state()
                if hasattr(self, 'note_taker_tab'):
                    self.note_taker_tab.reset_state()
                if hasattr(self, 'report_tab'):
                    self.report_tab.reset_state()
                    self.report_tab.refresh_ai_outputs()
                if hasattr(self, 'liability_tab'):
                    self.liability_tab.reset_state()
                if hasattr(self, 'email_tab'):
                    self.email_tab.search_bar.clear()
                    self.email_tab.check_db_init()
                    self.email_tab.perform_search()
                    
                log_event(f"Switched to case {new_file_num}")
                self.check_docket_expiry(new_file_num)
            else:
                QMessageBox.critical(self, "Error", f"Could not find case directory for {new_file_num}")

    def open_root_folder(self):
        if os.path.exists(self.case_path):
            try:
                os.startfile(self.case_path)
            except Exception as e:
                QMessageBox.warning(self, "Error", f"Could not open directory: {e}")

    def load_case_by_number(self, file_number):
        new_path = get_case_path(file_number)
        
        if new_path:
            self.save_status_history()
            self.file_number = file_number
            self.case_path = new_path
            self.setWindowTitle(f"iCharlotte - {self.file_number} - {os.path.basename(self.case_path)}")
            self.populate_tree()
            self.clear_all_status()
            self.load_status_history()
            
            # Switch to Case View tab (Index 1)
            self.tabs.setCurrentIndex(1)
            
            # Reset Tabs
            if hasattr(self, 'index_tab'):
                self.index_tab.load_data(self.file_number)
            if hasattr(self, 'chat_tab'):
                self.chat_tab.reset_state()
            if hasattr(self, 'note_taker_tab'):
                self.note_taker_tab.reset_state()
            if hasattr(self, 'report_tab'):
                self.report_tab.reset_state()
                self.report_tab.refresh_ai_outputs()
            if hasattr(self, 'liability_tab'):
                self.liability_tab.reset_state()
            if hasattr(self, 'email_tab'):
                self.email_tab.search_bar.clear()
                self.email_tab.check_db_init()
                self.email_tab.perform_search()
                
            log_event(f"Switched to case {self.file_number}")
            self.check_docket_expiry(file_number)
        else:
            QMessageBox.critical(self, "Error", f"Could not find case directory for {file_number}")

    def check_docket_expiry(self, file_number):
        """Checks if the last docket download was more than 30 days ago and runs the agent if so."""
        try:
            from icharlotte_core.master_db import MasterCaseDatabase
            db = MasterCaseDatabase()
            case = db.get_case(file_number)
            
            if not case:
                return

            last_download = case.get('last_docket_download')
            run_agent = False
            
            if not last_download:
                # Try to find existing docket in AI OUTPUT
                ai_output_dir = os.path.join(self.case_path, "NOTES", "AI OUTPUT")
                dockets = glob.glob(os.path.join(ai_output_dir, "Docket_*.pdf"))
                if dockets:
                    dockets.sort(key=os.path.getmtime, reverse=True)
                    latest_docket = dockets[0]
                    # Try to extract date from filename (Docket_YYYY.MM.DD.pdf)
                    match = re.search(r"Docket_(\d{4}\.\d{2}\.\d{2})", os.path.basename(latest_docket))
                    if match:
                        file_date_str = match.group(1).replace(".", "-")
                        try:
                            file_date = datetime.datetime.strptime(file_date_str, "%Y-%m-%d")
                            db.update_last_docket_download(file_number, file_date_str)
                            last_download = file_date_str
                            log_event(f"Detected existing docket from {file_date_str} for {file_number}. Updated database.")
                        except ValueError:
                            pass
                    
                    if not last_download:
                        # Fallback to file mtime if filename date fails
                        mtime = os.path.getmtime(latest_docket)
                        mtime_str = datetime.datetime.fromtimestamp(mtime).strftime("%Y-%m-%d")
                        db.update_last_docket_download(file_number, mtime_str)
                        last_download = mtime_str
                        log_event(f"Detected existing docket (mtime: {mtime_str}) for {file_number}. Updated database.")

            if not last_download:
                log_event(f"No previous docket download recorded for {file_number}. Triggering Docket Agent.")
                run_agent = True
            else:
                try:
                    last_date = datetime.datetime.strptime(last_download, "%Y-%m-%d")
                    days_elapsed = (datetime.datetime.now() - last_date).days
                    if days_elapsed > 30:
                        log_event(f"Last docket download for {file_number} was {days_elapsed} days ago (> 30). Triggering Docket Agent.")
                        run_agent = True
                except ValueError:
                    log_event(f"Invalid docket download date format for {file_number}: {last_download}. Triggering Docket Agent.")
                    run_agent = True
            
            if run_agent:
                # Run docket.py using run_agent
                self.run_agent("Docket Agent (Auto)", "docket.py", "file_number", None)
                
        except Exception as e:
            log_event(f"Error checking docket expiry: {e}", "error")

    def view_docket(self):
        # Look in NOTES/AI OUTPUT for Docket_*.pdf
        ai_output_dir = os.path.join(self.case_path, "NOTES", "AI OUTPUT")
        if not os.path.exists(ai_output_dir):
            QMessageBox.information(self, "Info", "AI OUTPUT directory not found.")
            return
            
        dockets = glob.glob(os.path.join(ai_output_dir, "Docket_*.pdf"))
        if dockets:
            # Sort by modification time, newest first
            dockets.sort(key=os.path.getmtime, reverse=True)
            latest = dockets[0]
            try:
                os.startfile(latest)
            except Exception as e:
                QMessageBox.warning(self, "Error", f"Could not open docket: {e}")
        else:
            QMessageBox.information(self, "Info", "No Docket PDF found in AI OUTPUT.")

    def manage_variables(self):
        dialog = VariablesDialog(self.file_number, self)
        dialog.exec()

    def manage_prompts(self):
        dialog = PromptsDialog(self)
        dialog.exec()

    def on_tree_double_click(self, item, column):
        path = item.data(0, Qt.ItemDataRole.UserRole)
        if path and os.path.exists(path) and os.path.isfile(path):
            try:
                os.startfile(path)
            except Exception as e:
                QMessageBox.warning(self, "Error", f"Could not open file: {e}")

    def on_tree_selection_changed(self):
        pass

    def run_separator_path(self, path):
        # Switch to Status Tab to show progress
        self.tabs.setCurrentIndex(1)
        
        status_widget = StatusWidget("Separator Agent", f"Analyzing {os.path.basename(path)}")
        self.status_list_layout.insertWidget(0, status_widget)
        
        script_path = os.path.join(SCRIPTS_DIR, "separate.py")
        args = [script_path, "--headless", path]
        
        runner = AgentRunner(sys.executable, args, status_widget)
        self.agent_runners.append(runner)
        
        # Capture output for JSON extraction
        separator_output_container = {"text": ""}
        def collect_stdout(text):
            separator_output_container["text"] += text
            
        runner.log_update.connect(collect_stdout)
        
        def on_finished(success):
            if success:
                # Find JSON path
                match = re.search(r"JSON_MAP: (.+)", separator_output_container["text"])
                if match:
                    json_path = match.group(1).strip()
                    try:
                        with open(json_path, 'r') as f:
                            docs = json.load(f)
                        
                        # Add to Index Tab
                        self.index_tab.add_pdf(path, docs)
                        
                        # Cleanup temp file
                        try:
                            os.remove(json_path)
                        except:
                            pass
                        
                    except Exception as e:
                        QMessageBox.critical(self, "Error", f"Failed to load separator result: {e}")
                else:
                     log_event(f"Warning: Could not find JSON output from separator for {path}", "warning")
            
            self.cleanup_runner(runner)
            
        runner.finished.connect(on_finished)
        runner.start()

    def create_agent_button(self, name, script, layout, arg_type="file_number", extra_flags=None):
        btn = QPushButton(name)
        btn.setFixedHeight(35)
        btn.setStyleSheet("font-size: 11px; padding: 2px;")
        btn.clicked.connect(partial(self.run_agent, name, script, arg_type, extra_flags))
        layout.addWidget(btn)

    def add_status_task(self, name, details, command, args):
        status_widget = StatusWidget(name, details)
        self.status_list_layout.insertWidget(0, status_widget) # Add to top
        
        runner = AgentRunner(command, args, status_widget, task_id=status_widget.task_id, file_number=self.file_number)
        self.agent_runners.append(runner) # Keep alive
        runner.finished.connect(lambda: self.cleanup_runner(runner))
        runner.start()
        return runner

    def cleanup_runner(self, runner):
        # Only cleanup if it matches the current file number (meaning the user saw it finish)
        if getattr(runner, 'file_number', None) == self.file_number:
            if runner in self.agent_runners:
                self.agent_runners.remove(runner)

    def clear_completed_status(self):
        for i in range(self.status_list_layout.count() - 1, -1, -1):
            item = self.status_list_layout.itemAt(i)
            widget = item.widget()
            if widget:
                if widget.is_finished:
                    widget.deleteLater()
                    
    def clear_all_status(self):
        for i in range(self.status_list_layout.count() - 1, -1, -1):
            item = self.status_list_layout.itemAt(i)
            widget = item.widget()
            if widget:
                widget.deleteLater()

    def save_status_history(self):
        history = []
        # Loop from 0 to count-1 (top to bottom)
        for i in range(self.status_list_layout.count()):
            item = self.status_list_layout.itemAt(i)
            widget = item.widget()
            if isinstance(widget, StatusWidget):
                history.append(widget.to_dict())
        
        if not os.path.exists(GEMINI_DATA_DIR):
            os.makedirs(GEMINI_DATA_DIR, exist_ok=True)
            
        save_path = os.path.join(GEMINI_DATA_DIR, f"{self.file_number}_status_history.json")
        try:
            with open(save_path, 'w') as f:
                json.dump(history, f, indent=2)
            log_event(f"Saved status history to {save_path}")
        except Exception as e:
            log_event(f"Error saving status history: {e}", "error")

    def load_status_history(self):
        save_path = os.path.join(GEMINI_DATA_DIR, f"{self.file_number}_status_history.json")
        if not os.path.exists(save_path):
            return

        try:
            with open(save_path, 'r') as f:
                history = json.load(f)
            
            for item_data in history:
                widget = StatusWidget.from_dict(item_data)
                
                # Try to reconnect to running agent
                reconnected = False
                if not widget.is_finished:
                    task_id = getattr(widget, 'task_id', None)
                    if task_id:
                        for runner in self.agent_runners:
                            if getattr(runner, 'task_id', None) == task_id:
                                runner.reconnect_widget(widget)
                                reconnected = True
                                # If runner is already finished, we can remove it now that we synced state
                                if runner.success is not None:
                                     self.agent_runners.remove(runner)
                                break
                    
                    if not reconnected:
                        # Mark as interrupted if we couldn't find the runner
                        widget.status_text_label.setText(widget.status_text_label.text() + " (Interrupted)")
                        widget.status_text_label.setStyleSheet("color: orange; font-weight: bold;")
                        widget.is_finished = True
                
                self.status_list_layout.addWidget(widget)
                
            log_event(f"Loaded {len(history)} status items from history")
        except Exception as e:
            log_event(f"Error loading status history: {e}", "error")

    def closeEvent(self, event):
        self.save_status_history()
        super().closeEvent(event)

    def run_agent(self, name, script, arg_type, extra_flags):
        log_event(f"Button clicked: {name}")
        script_path = os.path.join(SCRIPTS_DIR, script)
        if not os.path.exists(script_path):
            QMessageBox.critical(self, "Error", f"Script not found: {script_path}")
            return

        args = [script_path] # First arg for python is script path
        
        details = ""
        if arg_type == "file_number":
            args.append(self.file_number)
            details = f"File: {self.file_number}"
        elif arg_type == "case_path":
            args.append(self.case_path)
            details = "Scanning Case Directory"
            
        if extra_flags:
            args.extend(extra_flags)
            
        is_interactive = extra_flags and "--interactive" in extra_flags
        
        if is_interactive:
            try:
                creation_flags = 0x00000010 if os.name == 'nt' else 0
                subprocess.Popen([sys.executable] + args, creationflags=creation_flags)
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to launch: {e}")
        else:
            if script in ["docket.py", "complaint.py"]:
                args.append("--headless")
                
            self.add_status_task(name, details, sys.executable, args)

    def organize_checked_items(self):
        log_event("Organizing checked items...")
        paths = []
        
        iterator = QTreeWidgetItemIterator(self.tree)
        while iterator.value():
            item = iterator.value()
            path = item.data(0, Qt.ItemDataRole.UserRole)
            item_type = item.data(0, Qt.ItemDataRole.UserRole + 1)
            task_ids = item.data(0, Qt.ItemDataRole.UserRole + 2) or []
            
            if path and item_type == "file" and "organize" in task_ids:
                paths.append(path)
            
            iterator += 1
            
        if not paths:
            QMessageBox.warning(self, "No Selection", "No files selected for organization.")
            return

        script_path = os.path.join(SCRIPTS_DIR, "organizer.py")
        args = ["-u", script_path, "--dry-run"] + paths
        
        self.tabs.setCurrentIndex(1) # Switch to status
        status_widget = StatusWidget("Organizer (Analyze)", f"Analyzing {len(paths)} files...")
        self.status_list_layout.insertWidget(0, status_widget)
        
        runner = AgentRunner(sys.executable, args, status_widget)
        self.agent_runners.append(runner)
        
        self.organize_output = ""
        def collect_stdout(text):
            self.organize_output += text
            
        runner.log_update.connect(collect_stdout)
        
        def on_finished(success):
            if success:
                try:
                    output = self.organize_output.strip()
                    start_marker = "<<<JSON_START>>>"
                    end_marker = "<<<JSON_END>>>"
                    
                    m_start = output.find(start_marker)
                    m_end = output.find(end_marker)
                    
                    if m_start != -1 and m_end != -1:
                        json_str = output[m_start + len(start_marker):m_end].strip()
                        suggestions = json.loads(json_str)
                        self.index_tab.add_organization_suggestions(suggestions)
                        self.tabs.setCurrentWidget(self.index_tab)
                    else:
                        # Fallback: Try to find the JSON array directly (prone to errors with logs)
                        start = output.find("[")
                        end = output.rfind("]") + 1
                        if start != -1 and end != -1:
                            json_str = output[start:end]
                            suggestions = json.loads(json_str)
                            self.index_tab.add_organization_suggestions(suggestions)
                            self.tabs.setCurrentWidget(self.index_tab)
                        else:
                            QMessageBox.critical(self, "Error", "Could not parse organizer output (No JSON found).")
                except Exception as e:
                    QMessageBox.critical(self, "Error", f"Failed to parse suggestions: {e}\n\nOutput: {self.organize_output[:500]}")
            else:
                QMessageBox.critical(self, "Error", "Organization analysis failed.")
            self.cleanup_runner(runner)

        runner.finished.connect(on_finished)
        runner.start()
        self.clear_all_checkboxes()

    def apply_organization(self, approved):
        script_path = os.path.join(SCRIPTS_DIR, "organizer.py")
        json_data = json.dumps(approved)
        
        status_widget = StatusWidget("Organizer (Apply)", f"Moving {len(approved)} files...")
        self.status_list_layout.insertWidget(0, status_widget)
        
        args = ["-u", script_path, "--apply", json_data]
        
        runner = AgentRunner(sys.executable, args, status_widget)
        self.agent_runners.append(runner)
        
        def on_finished(success):
            if success:
                self.populate_tree()
                QMessageBox.information(self, "Success", "Files organized successfully.")
            else:
                QMessageBox.critical(self, "Error", "Failed to apply organization changes.")
            self.cleanup_runner(runner)
            
        runner.finished.connect(on_finished)
        runner.start()

    def process_checked_items(self):
        log_event("Processing checked items...")
        
        count = 0
        iterator = QTreeWidgetItemIterator(self.tree)
        while iterator.value():
            item = iterator.value()
            path = item.data(0, Qt.ItemDataRole.UserRole)
            item_type = item.data(0, Qt.ItemDataRole.UserRole + 1)
            task_ids = item.data(0, Qt.ItemDataRole.UserRole + 2) or []
            
            if not path or item_type != "file" or not task_ids:
                iterator += 1
                continue

            for tid in task_ids:
                if tid == "organize":
                    continue # Handled by the Organize button specifically
                    
                agent = next((a for a in self.AGENTS if a["id"] == tid), None)
                if not agent: continue
                
                if tid == "separate":
                    self.run_separator_path(path)
                    count += 1
                else:
                    filename = os.path.basename(path)
                    display_name = agent["name"]
                    details = f"{filename}"
                    
                    args = [os.path.join(SCRIPTS_DIR, agent["script"]), path]
                    self.add_status_task(display_name, details, sys.executable, args)
                    count += 1
            
            iterator += 1

        self.clear_all_checkboxes() # Clears the task data and UI

        if count > 0:
            self.tabs.setCurrentIndex(1)
        else:
            QMessageBox.warning(self, "No Selection", "No files selected for processing. (Did you queue any tasks?)")

    def load_cache(self):
        cache_path = os.path.join(GEMINI_DATA_DIR, f"{self.file_number}_tree.json")
        if not os.path.exists(cache_path):
            return False
            
        try:
            with open(cache_path, 'r') as f:
                data = json.load(f)
            
            entries = sorted(data, key=lambda x: len(x['path']))
            
            for entry in entries:
                path = entry['path']
                if path == self.case_path: continue
                
                parent_path = os.path.dirname(path)
                parent_item = self.tree_item_map.get(parent_path)
                
                if parent_item:
                    item = QTreeWidgetItem(parent_item)
                    item.setText(0, os.path.basename(path))
                    item.setData(0, Qt.ItemDataRole.UserRole, path)
                    item.setData(0, Qt.ItemDataRole.UserRole + 1, entry['type'])
                    
                    if entry['type'] == 'dir':
                        item.setIcon(0, self.icon_provider.icon(QFileInfo(path)))
                        item.setExpanded(False)
                    else:
                        item.setIcon(0, self.icon_provider.icon(QFileInfo(path)))
                        item.setText(1, entry.get('size_str', ''))
                        
                        # Convert cached date to standard format
                        date_str = format_date_to_mm_dd_yyyy(entry.get('date_str', ''))
                        item.setText(2, date_str)
                    
                    # Restore tasks
                    task_ids = entry.get('task_ids', [])
                    item.setData(0, Qt.ItemDataRole.UserRole + 2, task_ids)
                    self.update_item_tasks_ui(item)
                        
                    self.tree_item_map[path] = item
            return True
        except Exception as e:
            log_event(f"Error loading cache: {e}", "error")
            return False

    def save_cache(self):
        data = []
        for path, item in self.tree_item_map.items():
            if path == self.case_path: continue
            
            entry = {
                'path': path,
                'type': item.data(0, Qt.ItemDataRole.UserRole + 1),
                'size_str': item.text(1),
                'date_str': item.text(2),
                'task_ids': item.data(0, Qt.ItemDataRole.UserRole + 2)
            }
            data.append(entry)
            
        if not os.path.exists(GEMINI_DATA_DIR):
            os.makedirs(GEMINI_DATA_DIR)
            
        cache_path = os.path.join(GEMINI_DATA_DIR, f"{self.file_number}_tree.json")
        try:
            with open(cache_path, 'w') as f:
                json.dump(data, f)
        except Exception as e:
            log_event(f"Error saving cache: {e}", "error")

    def populate_tree(self):
        self.tree.clear()
        self.status_label.setText("Scanning directory structure... Please wait.")
        self.tree.setEnabled(False)
        self.process_btn.setEnabled(False)
        self.tree.setSortingEnabled(False)
        
        self.tree_item_map = {}
        self.visited_paths = set()
        self.visited_paths.add(self.case_path)
        
        root_text = f"{os.path.basename(self.case_path)} ({self.file_number})"
        root_item = QTreeWidgetItem(self.tree)
        root_item.setText(0, root_text)
        root_item.setIcon(0, self.icon_provider.icon(QFileInfo(self.case_path)))
        root_item.setData(0, Qt.ItemDataRole.UserRole, self.case_path)
        root_item.setExpanded(True)
        self.tree_item_map[self.case_path] = root_item

        if self.load_cache():
             self.tree.setEnabled(True)
             self.process_btn.setEnabled(True)
             self.status_label.setText(f"Loaded from cache. Verifying...")

        if hasattr(self, 'worker') and self.worker is not None:
            self.worker.stop()
            try:
                self.worker.disconnect()
            except:
                pass

        self.worker = DirectoryTreeWorker(self.case_path)
        self.worker.data_ready.connect(self.add_tree_batch)
        self.worker.finished.connect(self.on_scan_complete)
        self.worker.start()

    def add_tree_batch(self, batch):
        for root, dirs, files in batch:
            self.visited_paths.add(root)
            parent_item = self.tree_item_map.get(root)
            if not parent_item:
                continue
                
            dirs.sort(key=str.lower)
            files.sort(key=lambda x: x[0].lower())
            
            for d in dirs:
                dir_path = os.path.join(root, d)
                if dir_path not in self.tree_item_map:
                    d_item = QTreeWidgetItem(parent_item)
                    d_item.setText(0, d)
                    d_item.setIcon(0, self.icon_provider.icon(QFileInfo(dir_path)))
                    d_item.setData(0, Qt.ItemDataRole.UserRole, dir_path)
                    d_item.setData(0, Qt.ItemDataRole.UserRole + 1, "dir")
                    d_item.setExpanded(False)
                    self.tree_item_map[dir_path] = d_item
                
            for f, size, mtime in files:
                file_path = os.path.join(root, f)
                self.visited_paths.add(file_path)
                
                size_str = f"{size / 1024:.1f} KB" if size < 1024 * 1024 else f"{size / (1024 * 1024):.1f} MB"
                date_str = format_date_to_mm_dd_yyyy(mtime)

                if file_path not in self.tree_item_map:
                    f_item = QTreeWidgetItem(parent_item)
                    f_item.setText(0, f)
                    
                    # Use QFileIconProvider for typical PDF/Word icons
                    icon = self.icon_provider.icon(QFileInfo(file_path))
                    f_item.setIcon(0, icon)
                    
                    f_item.setData(0, Qt.ItemDataRole.UserRole, file_path)
                    f_item.setData(0, Qt.ItemDataRole.UserRole + 1, "file")
                    f_item.setText(1, size_str)
                    f_item.setText(2, date_str)
                    
                    self.tree_item_map[file_path] = f_item
                else:
                    f_item = self.tree_item_map[file_path]
                    try:
                        f_item.setText(1, size_str)
                        f_item.setText(2, date_str)
                    except RuntimeError:
                        continue
        
        QApplication.processEvents()

    def on_scan_complete(self):
        # Prune items not visited (deleted files/folders)
        to_remove = []
        for path, item in self.tree_item_map.items():
            if path != self.case_path and path not in self.visited_paths:
                to_remove.append(path)
                
        for path in to_remove:
            item = self.tree_item_map.pop(path)
            try:
                parent = item.parent()
                if parent:
                    parent.removeChild(item)
            except RuntimeError:
                pass
                
        self.save_cache()
        self.tree.setEnabled(True)
        self.process_btn.setEnabled(True)
        self.status_label.setText(f"Scan Complete. Case: {self.file_number}")
        self.tree.setSortingEnabled(True)

def exception_hook(exctype, value, traceback):
    print("------------------------------------------------------------------------------")
    print("CRITICAL ERROR CAUGHT BY EXCEPTION HOOK")
    print("------------------------------------------------------------------------------")
    import traceback as tb
    tb.print_exception(exctype, value, traceback)
    print("------------------------------------------------------------------------------")
    sys.__excepthook__(exctype, value, traceback)

sys.excepthook = exception_hook

if __name__ == "__main__":
    try:
        # Disable Chromium sandbox and security for local file editing
        os.environ["QTWEBENGINE_CHROMIUM_FLAGS"] = "--no-sandbox --disable-web-security"

        app = QApplication(sys.argv)
        
        # 1. Ask for File Number
        dialog = FileNumberDialog()
        if dialog.exec() == QDialog.DialogCode.Accepted:
            file_num = dialog.get_file_number()
            
            # 2. Resolve Path
            case_path = get_case_path(file_num)
            
            if case_path:
                # 3. Launch Main Window
                window = MainWindow(file_num, case_path)
                window.show()
                sys.exit(app.exec())
            else:
                QMessageBox.critical(None, "Error", f"Could not find case directory for {file_num}")
                sys.exit(1)
        else:
            sys.exit(0)
    except Exception as e:
        print(f"CRITICAL MAIN ERROR: {e}")
        import traceback
        traceback.print_exc()
        # Keep window open if possible or wait for input so user can see error
        input("Press Enter to exit...")
