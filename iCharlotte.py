import sys
import os
import argparse

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
    from PySide6.QtWidgets import (
        QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
        QPushButton, QTreeWidgetItem, QHeaderView, QMessageBox, QLabel,
        QFrame, QSplitter, QAbstractItemView, QLineEdit,
        QTreeWidgetItemIterator, QTabWidget, QScrollArea, QMenu, QDialog,
        QFileIconProvider, QToolButton, QGroupBox, QCheckBox, QComboBox,
        QInputDialog
    )
    from PySide6.QtCore import Qt, QThread, Signal, QFileInfo
    from PySide6.QtGui import QAction
    from PySide6.QtWebEngineCore import QWebEngineUrlScheme
except ImportError:
    print("Error: PySide6 or its components are not installed. Please run: pip install PySide6 PySide6-WebEngine")
    sys.exit(1)

# --- Core Modules ---
from icharlotte_core.config import SCRIPTS_DIR, GEMINI_DATA_DIR, BASE_PATH_WIN
from icharlotte_core.utils import (
    log_event, get_case_path, sanitize_filename, format_date_to_mm_dd_yyyy
)
from icharlotte_core.ui.widgets import (
    StatusWidget, AgentRunner, FileTreeWidget
)
from icharlotte_core.ui.case_view_enhanced import (
    EnhancedAgentButton, AgentSettingsDB, AgentSettingsDialog, ContradictionSettingsDialog,
    AdvancedFilterWidget, FilePreviewWidget, OutputBrowserWidget,
    ProcessingLogWidget, ProcessingLogDB, FileTagsDB, EnhancedFileTreeWidget
)
from icharlotte_core.ui.dialogs import FileNumberDialog, VariablesDialog, PromptsDialog, LLMSettingsDialog
from icharlotte_core.ui.tabs import ChatTab, IndexTab
from icharlotte_core.ui.email_tab import EmailTab
from icharlotte_core.ui.email_update_tab import EmailUpdateTab
from icharlotte_core.ui.report_tab import ReportTab
from icharlotte_core.ui.logs_tab import LogsTab
from icharlotte_core.ui.liability_tab import LiabilityExposureTab
from icharlotte_core.ui.master_case_tab import MasterCaseTab
from icharlotte_core.ui.templates_resources_tab import TemplatesResourcesTab

class DirectoryTreeWorker(QThread):
    data_ready = Signal(list) # Emits (root, dirs, files) tuples
    finished = Signal()
    
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
    def __init__(self, file_number=None, case_path=None, initial_tab=None):
        super().__init__()
        self.file_number = file_number
        self.case_path = case_path
        self._update_window_title()
        self.resize(1200, 800)

        self.agent_runners = [] # Keep references to prevent GC
        self.cached_models = {} # Cache for models: {provider: [list]}
        self.fetcher = None

        log_event(f"Initializing MainWindow for {file_number} at {case_path}")
        self.icon_provider = QFileIconProvider()
        # Cache icons to avoid slow network file access
        self._icon_cache = {}
        self._folder_icon = self.icon_provider.icon(QFileIconProvider.IconType.Folder)
        self._file_icon = self.icon_provider.icon(QFileIconProvider.IconType.File)
        self.setup_ui()

        # Restore tab if specified
        if initial_tab is not None and 0 <= initial_tab < self.tabs.count():
            self.tabs.setCurrentIndex(initial_tab)

        # Only populate tree and check docket if a case is loaded
        if self.case_path:
            self.populate_tree()
            self.load_status_history()
            self.check_docket_expiry(file_number)

    def _update_window_title(self):
        """Update window title based on current file_number and case_path."""
        if self.file_number and self.case_path:
            self.setWindowTitle(f"iCharlotte - {self.file_number} - {os.path.basename(self.case_path)}")
        else:
            self.setWindowTitle("iCharlotte")

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

        # Output Browser Button
        btn_outputs = QPushButton("Output Browser")
        btn_outputs.clicked.connect(self.open_output_browser)
        toolbar_layout.addWidget(btn_outputs)

        # Processing Log Button
        btn_proc_log = QPushButton("Processing Log")
        btn_proc_log.clicked.connect(self.open_processing_log)
        toolbar_layout.addWidget(btn_proc_log)

        toolbar_layout.addStretch()

        # Wrapper for vertical layout of Case View
        wrapper_layout = QVBoxLayout()
        wrapper_layout.addLayout(toolbar_layout)

        # Main horizontal splitter (agents | tree | preview)
        splitter = QSplitter(Qt.Orientation.Horizontal)
        wrapper_layout.addWidget(splitter)

        main_layout.addLayout(wrapper_layout)

        # Initialize agent settings database
        self.agent_settings_db = AgentSettingsDB()
        self.agent_buttons = {}  # Track enhanced agent buttons
        self.running_agents = {}  # Track which agents are running {script: file_number}

        # Left Panel (Case Agents)
        left_panel = QFrame()
        left_panel.setFrameShape(QFrame.Shape.StyledPanel)
        left_panel.setFixedWidth(180)
        left_layout = QVBoxLayout(left_panel)
        left_layout.setAlignment(Qt.AlignmentFlag.AlignTop)
        left_layout.setContentsMargins(4, 5, 4, 5)
        left_layout.setSpacing(4)

        title_label = QLabel("Case Agents")
        title_label.setStyleSheet("font-weight: bold; font-size: 14px; margin-bottom: 5px;")
        left_layout.addWidget(title_label)

        # Enhanced Agent Buttons with running indicator, status, and settings
        self.create_enhanced_agent_button("Docket Agent", "docket.py", left_layout, arg_type="file_number")
        self.create_enhanced_agent_button("Complaint Agent", "complaint.py", left_layout, arg_type="file_number")
        self.create_enhanced_agent_button("Report Agent", "report.py", left_layout, arg_type="file_number")
        self.create_enhanced_agent_button("Discovery Generate", "discovery_requests.py", left_layout, arg_type="file_number", extra_flags=["--interactive"])
        self.create_enhanced_agent_button("Subpoena Tracker", "subpoena_tracker.py", left_layout, arg_type="file_number")

        # Directory-based Case Agents
        left_layout.addSpacing(8)
        dir_label = QLabel("Analysis Agents")
        dir_label.setStyleSheet("font-weight: bold; font-size: 12px; color: #666;")
        left_layout.addWidget(dir_label)

        self.create_enhanced_agent_button("Liability Agent", "liability.py", left_layout, arg_type="case_path")
        self.create_enhanced_agent_button("Exposure Agent", "exposure.py", left_layout, arg_type="case_path")

        # New Agents: Timeline and Contradiction Detector
        left_layout.addSpacing(8)
        new_label = QLabel("Document Agents")
        new_label.setStyleSheet("font-weight: bold; font-size: 12px; color: #666;")
        left_layout.addWidget(new_label)

        self.create_enhanced_agent_button("Timeline Agent", "extract_timeline.py", left_layout, arg_type="file_picker")
        self.create_enhanced_agent_button("Contradiction Detector", "detect_contradictions.py", left_layout, arg_type="file_number")

        left_layout.addStretch()
        splitter.addWidget(left_panel)
        
        # Center Panel (File Tree with Enhanced Features)
        center_panel = QFrame()
        center_layout = QVBoxLayout(center_panel)

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

        act_unprocessed = QAction("Unprocessed Files -> OCR", self)
        act_unprocessed.triggered.connect(lambda: self.smart_select("unprocessed_ocr"))
        self.smart_select_menu.addAction(act_unprocessed)

        self.smart_select_btn.setMenu(self.smart_select_menu)
        header_layout.addWidget(self.smart_select_btn)

        # Advanced Filter Toggle
        self.filter_toggle_btn = QPushButton("▼ Filters")
        self.filter_toggle_btn.setCheckable(True)
        self.filter_toggle_btn.clicked.connect(self.toggle_advanced_filters)
        header_layout.addWidget(self.filter_toggle_btn)

        self.expand_btn = QPushButton("Expand All")
        self.expand_btn.setCheckable(True)
        self.expand_btn.clicked.connect(self.toggle_expand)
        header_layout.addWidget(self.expand_btn)

        # Preview Toggle
        self.preview_toggle_btn = QPushButton("Preview")
        self.preview_toggle_btn.setCheckable(True)
        self.preview_toggle_btn.clicked.connect(self.toggle_preview_pane)
        header_layout.addWidget(self.preview_toggle_btn)

        self.clear_all_btn = QPushButton("Clear All")
        self.clear_all_btn.clicked.connect(self.clear_all_checkboxes)
        header_layout.addWidget(self.clear_all_btn)

        center_layout.addLayout(header_layout)

        # Advanced Filter Widget (Hidden by default)
        self.advanced_filter = AdvancedFilterWidget()
        self.advanced_filter.filter_changed.connect(self.apply_advanced_filters)
        self.advanced_filter.hide()
        center_layout.addWidget(self.advanced_filter)

        self.status_label = QLabel("Ready")
        center_layout.addWidget(self.status_label)

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
            {"id": "timeline", "name": "Timeline", "script": "extract_timeline.py", "color": "#3f51b5", "short": "TIME"},
            {"id": "contradict", "name": "Conflicts", "script": "detect_contradictions.py", "color": "#f44336", "short": "CONF"},
        ]

        # Enhanced File Tree with additional columns
        self.tree = EnhancedFileTreeWidget()
        self.tree.item_moved.connect(lambda: self.populate_tree())
        self.tree.folder_created.connect(lambda p: self.populate_tree())
        self.tree.setHeaderLabels([
            "Category / File",
            "Size",
            "Date Modified",
            "Pages",
            "Status",
            "Tags",
            "Queued Tasks (Click to Add ➕)"
        ])
        self.tree.setSortingEnabled(True)
        self.tree.setColumnWidth(0, 300)
        self.tree.setColumnWidth(1, 70)
        self.tree.setColumnWidth(2, 100)
        self.tree.setColumnWidth(3, 50)
        self.tree.setColumnWidth(4, 80)
        self.tree.setColumnWidth(5, 100)
        self.tree.setColumnWidth(6, 200)
        self.tree.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.tree.setAlternatingRowColors(True)
        self.tree.itemSelectionChanged.connect(self.on_tree_selection_changed)
        self.tree.itemDoubleClicked.connect(self.on_tree_double_click)
        self.tree.itemClicked.connect(self.on_tree_item_clicked)
        center_layout.addWidget(self.tree)

        btn_layout = QHBoxLayout()
        self.process_btn = QPushButton("Process All Queued Tasks")
        self.process_btn.setStyleSheet("background-color: #4CAF50; color: white; font-weight: bold; padding: 10px;")
        self.process_btn.clicked.connect(self.process_checked_items)
        btn_layout.addWidget(self.process_btn)

        self.organize_btn = QPushButton("Quick Organize (AI)")
        self.organize_btn.setStyleSheet("background-color: #2196F3; color: white; font-weight: bold; padding: 10px;")
        self.organize_btn.clicked.connect(self.organize_checked_items)
        btn_layout.addWidget(self.organize_btn)

        center_layout.addLayout(btn_layout)

        splitter.addWidget(center_panel)

        # Right Panel (File Preview - Hidden by default)
        self.preview_pane = FilePreviewWidget()
        self.preview_pane.hide()
        splitter.addWidget(self.preview_pane)

        splitter.setSizes([180, 800, 0])
        splitter.setStretchFactor(0, 0)
        splitter.setStretchFactor(1, 1)
        splitter.setStretchFactor(2, 0)
        
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

        # --- Tab 4: Chat ---
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

        # --- Tab: Templates / Resources ---
        self.templates_tab = TemplatesResourcesTab(main_window=self)
        self.tabs.addTab(self.templates_tab, "Templates / Resources")

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

        self.settings_btn = QPushButton("Settings")
        self.settings_btn.clicked.connect(self.open_settings_dialog)
        self.corner_layout.addWidget(self.settings_btn)

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

    def open_settings_dialog(self):
        """Open the LLM settings dialog."""
        dialog = LLMSettingsDialog(self)
        dialog.exec()

    def restart_app(self):
        log_event("User requested manual restart. Spawning new process...")
        # Close all agent runners if any are running
        for runner in self.agent_runners:
            try:
                runner.terminate()
            except:
                pass

        # Get the absolute path to this script
        script_path = os.path.abspath(__file__)

        # Build restart arguments with current state
        args = [script_path]

        # Add current file number if loaded
        if self.file_number:
            args.extend(['--file-number', str(self.file_number)])

        # Add current case path if loaded
        if self.case_path:
            args.extend(['--case-path', str(self.case_path)])

        # Add current tab index
        current_tab = self.tabs.currentIndex()
        args.extend(['--tab', str(current_tab)])

        log_event(f"Restarting with: python={sys.executable}, args={args}")
        log_event(f"Current state: file_number={self.file_number}, case_path={self.case_path}, tab={current_tab}")

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
        if column == 6:  # Queued Tasks column (index 6)
            # Get click position relative to the tree's viewport
            pos = self.tree.visualItemRect(item).bottomLeft()
            pos.setX(self.tree.columnViewportPosition(6))
            global_pos = self.tree.viewport().mapToGlobal(pos)

            # Get all selected items (for multi-select support)
            selected_items = self.tree.selectedItems()

            # If clicked item is not in selection, use just the clicked item
            if item not in selected_items:
                selected_items = [item]

            self.show_agent_menu(selected_items, global_pos)

    def show_agent_menu(self, items, global_pos):
        """Show agent menu for one or more selected items."""
        menu = QMenu(self)

        # For multiple items, show count in header
        if len(items) > 1:
            header_action = QAction(f"Apply to {len(items)} selected items:", menu)
            header_action.setEnabled(False)
            menu.addAction(header_action)
            menu.addSeparator()

        # Determine which tasks are common across all items
        # A task is "checked" if ALL items have it, "partial" if some have it
        task_states = {}
        for agent in self.AGENTS:
            agent_id = agent["id"]
            has_count = sum(1 for item in items if agent_id in (item.data(0, Qt.ItemDataRole.UserRole + 2) or []))
            if has_count == len(items):
                task_states[agent_id] = "all"
            elif has_count > 0:
                task_states[agent_id] = "partial"
            else:
                task_states[agent_id] = "none"

        for agent in self.AGENTS:
            state = task_states[agent["id"]]
            action = QAction(agent["name"], menu)
            action.setCheckable(True)
            action.setChecked(state in ["all", "partial"])

            # Visual indication for partial selection
            if state == "partial":
                action.setText(f"{agent['name']} (partial)")

            action.triggered.connect(partial(self.toggle_agent_task_multi, items, agent["id"], state))
            menu.addAction(action)

        menu.addSeparator()
        clear_act = QAction("Clear All Tasks", menu)
        clear_act.triggered.connect(lambda: self.clear_tasks_multi(items))
        menu.addAction(clear_act)

        menu.exec(global_pos)

    def toggle_agent_task_multi(self, items, agent_id, current_state):
        """Toggle a task for multiple items. If partial/none, add to all. If all, remove from all."""
        self.tree.blockSignals(True)

        # If all items have it, remove from all. Otherwise, add to all.
        should_add = current_state != "all"

        for item in items:
            current_tasks = list(item.data(0, Qt.ItemDataRole.UserRole + 2) or [])
            if should_add:
                if agent_id not in current_tasks:
                    current_tasks.append(agent_id)
            else:
                if agent_id in current_tasks:
                    current_tasks.remove(agent_id)

            item.setData(0, Qt.ItemDataRole.UserRole + 2, current_tasks)
            self.update_item_tasks_ui(item)

        self.tree.blockSignals(False)

    def clear_tasks_multi(self, items):
        """Clear all tasks from multiple items."""
        self.tree.blockSignals(True)
        for item in items:
            item.setData(0, Qt.ItemDataRole.UserRole + 2, [])
            self.update_item_tasks_ui(item)
        self.tree.blockSignals(False)

    def toggle_agent_task(self, item, agent_id):
        """Toggle a task for a single item (legacy support)."""
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
            item.setText(6, " [ + Add Tasks ]")
            item.setForeground(6, Qt.GlobalColor.gray)
            return

        # Create a visual string of short tags
        tags = []
        for tid in task_ids:
            agent = next((a for a in self.AGENTS if a["id"] == tid), None)
            if agent:
                tags.append(f"[{agent['short']}]")

        item.setText(6, " ".join(tags))
        # Optional: set color for column 6 text
        item.setForeground(6, Qt.GlobalColor.blue)

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

                elif mode == "unprocessed_ocr":
                    # Only select unprocessed PDF files for OCR
                    if path.lower().endswith(".pdf"):
                        # Check if already processed
                        proc_status = item.text(4)  # Status column
                        if not proc_status or "OCR" not in proc_status:
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
                    self.chat_tab.load_case(self.file_number)
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
            self._update_window_title()
            self.populate_tree()
            self.clear_all_status()

            # Reset all agent button running states when switching cases
            for btn in self.agent_buttons.values():
                btn.set_running(False)

            self.load_status_history()

            # Switch to Case View tab (Index 1)
            self.tabs.setCurrentIndex(1)

            # Reset Tabs
            if hasattr(self, 'index_tab'):
                self.index_tab.load_data(self.file_number)
            if hasattr(self, 'chat_tab'):
                self.chat_tab.load_case(self.file_number)
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
        selected = self.tree.selectedItems()
        count = len(selected)
        if count > 1:
            # Count only files (not directories)
            file_count = sum(1 for item in selected if item.data(0, Qt.ItemDataRole.UserRole + 1) == "file")
            if file_count > 0:
                self.status_label.setText(f"{file_count} files selected - Click 'Queued Tasks' column to apply tasks to all")
            # Clear preview for multi-selection
            if hasattr(self, 'preview_pane') and self.preview_pane.isVisible():
                self.preview_pane.clear()
        elif count == 1:
            item = selected[0]
            path = item.data(0, Qt.ItemDataRole.UserRole)
            item_type = item.data(0, Qt.ItemDataRole.UserRole + 1)
            if path:
                self.status_label.setText(f"Selected: {os.path.basename(path)}")
                # Update preview pane if visible and it's a file
                if hasattr(self, 'preview_pane') and self.preview_pane.isVisible() and item_type == "file":
                    self.preview_pane.show_file(path)
        else:
            if hasattr(self, 'file_number') and self.file_number:
                self.status_label.setText(f"Case: {self.file_number}")
            if hasattr(self, 'preview_pane') and self.preview_pane.isVisible():
                self.preview_pane.clear()

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

    def create_enhanced_agent_button(self, name, script, layout, arg_type="file_number", extra_flags=None):
        """Create an enhanced agent button with running indicator, status, and settings."""
        enhanced_btn = EnhancedAgentButton(name, script)
        enhanced_btn.clicked.connect(partial(self.run_enhanced_agent, name, script, arg_type, extra_flags, enhanced_btn))
        enhanced_btn.settings_clicked.connect(partial(self.open_agent_settings, script))

        # Store reference for updating status
        self.agent_buttons[script] = enhanced_btn

        # Update last docket download date for docket agent
        if script == "docket.py":
            self.update_docket_agent_status()

        layout.addWidget(enhanced_btn)

    def run_enhanced_agent(self, name, script, arg_type, extra_flags, btn_widget):
        """Run agent with enhanced status tracking."""
        # Set button to running state
        btn_widget.set_running(True)
        # Store the file_number this agent was started for
        started_for_case = self.file_number
        self.running_agents[script] = started_for_case

        # Run the agent
        runner = self.run_agent(name, script, arg_type, extra_flags)

        # Connect finished signal to update button
        if runner:
            runner.finished.connect(partial(self.on_agent_finished, script, btn_widget, started_for_case))
        else:
            # User cancelled (e.g., file picker dialog) - reset button state
            btn_widget.set_running(False)
            if script in self.running_agents:
                del self.running_agents[script]

    def on_agent_finished(self, script, btn_widget, started_for_case, success):
        """Handle agent completion and update button status."""
        # Clear the running state
        if script in self.running_agents:
            del self.running_agents[script]

        # Only update button UI if we're still on the same case the agent was started for
        if self.file_number == started_for_case:
            btn_widget.set_running(False)
            # Update status message
            if success:
                btn_widget.set_status("Last: Just now")
            else:
                btn_widget.set_status("Last: Failed")

        # Update docket download date for the ORIGINAL case (not current case)
        if script == "docket.py" and success and started_for_case:
            from datetime import datetime
            today = datetime.now().strftime("%Y-%m-%d")
            self.master_db.update_last_docket_download(started_for_case, today)
            # Only update button if still on the same case
            if self.file_number == started_for_case:
                btn_widget.set_last_run(today)

    def update_docket_agent_status(self):
        """Update the docket agent button with last download date."""
        if not self.file_number or "docket.py" not in self.agent_buttons:
            return

        case_data = self.master_tab.db.get_case(self.file_number)
        if case_data:
            last_download = case_data.get("last_docket_download")
            if last_download:
                self.agent_buttons["docket.py"].set_last_run(last_download)
            else:
                self.agent_buttons["docket.py"].set_status("Never downloaded")

    def open_agent_settings(self, script):
        """Open the settings dialog for an agent."""
        if script == "detect_contradictions.py":
            # Use custom dialog for Contradiction Detector
            dialog = ContradictionSettingsDialog(self.file_number, self.agent_settings_db, self)
        else:
            dialog = AgentSettingsDialog(script, self.agent_settings_db, self)
        dialog.exec()

    def open_output_browser(self):
        """Open the output browser dialog."""
        if not self.case_path or not self.file_number:
            QMessageBox.warning(self, "No Case", "Please load a case first.")
            return

        dialog = OutputBrowserWidget(self.case_path, self.file_number, self)
        dialog.exec()

    def open_processing_log(self):
        """Open the processing log dialog."""
        if not self.file_number:
            QMessageBox.warning(self, "No Case", "Please load a case first.")
            return

        dialog = ProcessingLogWidget(self.file_number, self)
        dialog.exec()

    def toggle_advanced_filters(self, checked):
        """Toggle visibility of advanced filter panel."""
        if checked:
            self.advanced_filter.show()
            self.filter_toggle_btn.setText("▲ Filters")
            # Update available tags
            if self.file_number:
                tags_db = FileTagsDB(self.file_number)
                self.advanced_filter.set_available_tags(tags_db.get_all_tags())
        else:
            self.advanced_filter.hide()
            self.filter_toggle_btn.setText("▼ Filters")

    def toggle_preview_pane(self, checked):
        """Toggle visibility of file preview pane."""
        if checked:
            self.preview_pane.show()
            # Get current splitter parent
            splitter = self.preview_pane.parentWidget()
            if splitter and hasattr(splitter, 'setSizes'):
                current = splitter.sizes()
                if len(current) >= 3:
                    # Allocate space for preview
                    total = sum(current)
                    splitter.setSizes([180, int(total * 0.6), int(total * 0.25)])
        else:
            self.preview_pane.hide()
            self.preview_pane.clear()

    def apply_advanced_filters(self, filters):
        """Apply advanced filters to the file tree."""
        iterator = QTreeWidgetItemIterator(self.tree)

        while iterator.value():
            item = iterator.value()
            file_path = item.data(0, Qt.ItemDataRole.UserRole)
            item_type = item.data(0, Qt.ItemDataRole.UserRole + 1)

            if item_type == "file" and file_path:
                should_show = self._file_matches_filters(file_path, filters)
                item.setHidden(not should_show)
            elif item_type == "dir":
                # Directories stay visible if any child is visible
                pass

            iterator += 1

        # Update directory visibility based on children
        self._update_directory_visibility()

    def _file_matches_filters(self, file_path, filters):
        """Check if a file matches the given filters."""
        from datetime import datetime, timedelta

        # File type filter
        ext = os.path.splitext(file_path)[1].lower().lstrip('.')
        file_types = filters.get("file_types", [])

        if file_types:
            type_match = False
            if ext in ["pdf"] and "pdf" in file_types:
                type_match = True
            elif ext in ["doc", "docx"] and any(t in file_types for t in ["doc", "docx"]):
                type_match = True
            elif ext in ["jpg", "jpeg", "png", "gif", "bmp", "tiff"] and any(t in file_types for t in ["jpg", "jpeg", "png", "gif", "bmp", "tiff"]):
                type_match = True
            elif "other" in file_types and ext not in ["pdf", "doc", "docx", "jpg", "jpeg", "png", "gif", "bmp", "tiff"]:
                type_match = True

            if not type_match:
                return False

        # Processing status filter
        status_filter = filters.get("processing_status", [])
        if status_filter and "all" not in status_filter:
            if self.file_number:
                proc_log = ProcessingLogDB(self.file_number)
                file_status = proc_log.get_file_processing_status(file_path)

                has_ocr = any("ocr" in log.get("task_type", "").lower() and log.get("status") == "success" for log in file_status)
                has_summary = any("summar" in log.get("task_type", "").lower() and log.get("status") == "success" for log in file_status)
                is_unprocessed = not file_status or not any(log.get("status") == "success" for log in file_status)

                status_match = False
                if "unprocessed" in status_filter and is_unprocessed:
                    status_match = True
                if "ocr" in status_filter and has_ocr:
                    status_match = True
                if "summarized" in status_filter and has_summary:
                    status_match = True

                if not status_match:
                    return False

        # Date filter
        date_filter = filters.get("date_filter", "Any time")
        if date_filter != "Any time":
            try:
                mtime = os.path.getmtime(file_path)
                file_date = datetime.fromtimestamp(mtime)
                now = datetime.now()

                if date_filter == "Today":
                    if file_date.date() != now.date():
                        return False
                elif date_filter == "Last 7 days":
                    if (now - file_date).days > 7:
                        return False
                elif date_filter == "Last 30 days":
                    if (now - file_date).days > 30:
                        return False
                elif date_filter == "Last 90 days":
                    if (now - file_date).days > 90:
                        return False
                elif date_filter == "Custom range...":
                    date_from = filters.get("date_from")
                    date_to = filters.get("date_to")
                    if date_from:
                        from_dt = datetime.strptime(date_from, "%Y-%m-%d")
                        if file_date < from_dt:
                            return False
                    if date_to:
                        to_dt = datetime.strptime(date_to, "%Y-%m-%d")
                        if file_date > to_dt:
                            return False
            except:
                pass

        # Tag filter
        tag_filter = filters.get("tag")
        if tag_filter and self.file_number:
            tags_db = FileTagsDB(self.file_number)
            file_tags = tags_db.get_tags(file_path)
            if tag_filter not in file_tags:
                return False

        return True

    def _update_directory_visibility(self):
        """Update directory visibility based on visible children."""
        def check_item(item):
            if item.data(0, Qt.ItemDataRole.UserRole + 1) == "dir":
                has_visible_child = False
                for i in range(item.childCount()):
                    child = item.child(i)
                    if not child.isHidden():
                        has_visible_child = True
                        break
                    # Recursively check nested directories
                    if child.data(0, Qt.ItemDataRole.UserRole + 1) == "dir":
                        check_item(child)
                        if not child.isHidden():
                            has_visible_child = True

                item.setHidden(not has_visible_child)

        for i in range(self.tree.topLevelItemCount()):
            check_item(self.tree.topLevelItem(i))

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
        # Disconnect all runners from their widgets before deleting
        # to prevent signals being sent to deleted widgets
        for runner in self.agent_runners:
            runner.disconnect_widget()

        for i in range(self.status_list_layout.count() - 1, -1, -1):
            item = self.status_list_layout.itemAt(i)
            widget = item.widget()
            if widget:
                widget.deleteLater()

    def save_status_history(self):
        if not self.file_number:
            return

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
        if not self.file_number:
            return

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

            # Special handling for Contradiction Detector - pass selected summaries
            if script == "detect_contradictions.py":
                settings = self.agent_settings_db.get_settings(script)
                selected_summaries = settings.get("selected_summaries")

                if selected_summaries:
                    args.extend(["--summaries", ",".join(selected_summaries)])
                    details = f"{len(selected_summaries)} summaries"

        elif arg_type == "case_path":
            args.append(self.case_path)
            details = "Scanning Case Directory"
        elif arg_type == "file_picker":
            # Show file picker dialog, starting in the case directory
            start_dir = self.case_path if hasattr(self, 'case_path') and self.case_path else ""
            files, _ = QFileDialog.getOpenFileNames(
                self,
                f"Select files for {name}",
                start_dir,
                "Documents (*.pdf *.docx *.txt);;All Files (*.*)"
            )
            if not files:
                return  # User cancelled
            args.extend(files)
            details = f"{len(files)} file(s) selected"

        if extra_flags:
            args.extend(extra_flags)
            
        is_interactive = extra_flags and "--interactive" in extra_flags
        
        if is_interactive:
            try:
                creation_flags = 0x00000010 if os.name == 'nt' else 0
                subprocess.Popen([sys.executable] + args, creationflags=creation_flags)
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to launch: {e}")
            return None
        else:
            if script in ["docket.py", "complaint.py"]:
                args.append("--headless")

            return self.add_status_task(name, details, sys.executable, args)

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
                        item.setIcon(0, self._get_cached_icon(path, is_dir=True))
                        item.setExpanded(False)
                    else:
                        item.setIcon(0, self._get_cached_icon(path))
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

    def _get_cached_icon(self, file_path, is_dir=False):
        """Get icon from cache based on file extension to avoid network access."""
        if is_dir:
            return self._folder_icon

        # Get extension and check cache
        ext = os.path.splitext(file_path)[1].lower()
        if ext in self._icon_cache:
            return self._icon_cache[ext]

        # For first occurrence of each extension, get icon from a local temp file
        # This avoids accessing network files while still getting proper icons
        if ext in ['.pdf', '.doc', '.docx', '.xls', '.xlsx', '.ppt', '.pptx',
                   '.txt', '.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tiff',
                   '.mp3', '.mp4', '.avi', '.mov', '.zip', '.rar', '.7z']:
            # Use the icon provider with just the extension info
            icon = self.icon_provider.icon(QFileInfo(f"dummy{ext}"))
            self._icon_cache[ext] = icon
            return icon

        # Default file icon for unknown extensions
        self._icon_cache[ext] = self._file_icon
        return self._file_icon

    def _get_pdf_page_count(self, file_path):
        """Get page count for a PDF file."""
        # DISABLED: Reading PDF files on network drives causes UI freeze
        # TODO: Move this to background thread or cache results
        return ""

    def _get_file_processing_status(self, file_path):
        """Get processing status for a file."""
        if not hasattr(self, 'file_number') or not self.file_number:
            return ""

        try:
            # Use cached processing log instead of creating new instance per file
            if not hasattr(self, '_cached_proc_log') or self._cached_proc_log is None:
                return ""

            logs = self._cached_proc_log.get_file_processing_status(file_path)
            if not logs:
                return ""

            # Get unique successful task types
            successful = set()
            for log in logs:
                if log.get("status") == "success":
                    task = log.get("task_type", "").lower()
                    if "ocr" in task:
                        successful.add("OCR")
                    elif "summar" in task:
                        successful.add("SUM")
                    elif "timeline" in task:
                        successful.add("TL")
                    elif "chron" in task:
                        successful.add("CHR")

            return ", ".join(sorted(successful)) if successful else ""
        except:
            return ""

    def _get_file_tags(self, file_path):
        """Get tags for a file."""
        if not hasattr(self, 'file_number') or not self.file_number:
            return ""

        try:
            # Use cached tags db instead of creating new instance per file
            if not hasattr(self, '_cached_tags_db') or self._cached_tags_db is None:
                return ""

            tags = self._cached_tags_db.get_tags(file_path)
            return ", ".join(tags) if tags else ""
        except:
            return ""

    def populate_tree(self):
        self.tree.clear()
        self.status_label.setText("Scanning directory structure... Please wait.")
        self.tree.setEnabled(False)
        self.process_btn.setEnabled(False)
        self.tree.setSortingEnabled(False)

        # Set up enhanced tree databases for the current case
        # Cache these once to avoid re-reading JSON for every file
        if hasattr(self, 'file_number') and self.file_number:
            self.tree.set_databases(self.file_number)
            self._cached_proc_log = ProcessingLogDB(self.file_number)
            self._cached_tags_db = FileTagsDB(self.file_number)
            # Also update docket agent status
            self.update_docket_agent_status()
        else:
            self._cached_proc_log = None
            self._cached_tags_db = None

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
                    d_item.setIcon(0, self._get_cached_icon(dir_path, is_dir=True))
                    d_item.setData(0, Qt.ItemDataRole.UserRole, dir_path)
                    d_item.setData(0, Qt.ItemDataRole.UserRole + 1, "dir")
                    d_item.setExpanded(False)
                    self.tree_item_map[dir_path] = d_item

            for f, size, mtime in files:
                file_path = os.path.join(root, f)
                self.visited_paths.add(file_path)

                size_str = f"{size / 1024:.1f} KB" if size < 1024 * 1024 else f"{size / (1024 * 1024):.1f} MB"
                date_str = format_date_to_mm_dd_yyyy(mtime)

                # Page count disabled - too slow on network drives
                page_count = ""

                # Get processing status (uses cached DB)
                proc_status = self._get_file_processing_status(file_path) if self.file_number else ""

                # Get tags (uses cached DB)
                tags_str = self._get_file_tags(file_path) if self.file_number else ""

                if file_path not in self.tree_item_map:
                    f_item = QTreeWidgetItem(parent_item)
                    f_item.setText(0, f)

                    # Use cached extension-based icons to avoid network access
                    f_item.setIcon(0, self._get_cached_icon(file_path))

                    f_item.setData(0, Qt.ItemDataRole.UserRole, file_path)
                    f_item.setData(0, Qt.ItemDataRole.UserRole + 1, "file")
                    f_item.setText(1, size_str)
                    f_item.setText(2, date_str)
                    f_item.setText(3, page_count)
                    f_item.setText(4, proc_status)
                    f_item.setText(5, tags_str)

                    self.tree_item_map[file_path] = f_item
                else:
                    f_item = self.tree_item_map[file_path]
                    try:
                        f_item.setText(1, size_str)
                        f_item.setText(2, date_str)
                        f_item.setText(3, page_count)
                        f_item.setText(4, proc_status)
                        f_item.setText(5, tags_str)
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
        # Parse command-line arguments for restart state restoration
        parser = argparse.ArgumentParser(description='iCharlotte Legal Document Management Suite')
        parser.add_argument('--file-number', type=str, help='File number to load on startup')
        parser.add_argument('--case-path', type=str, help='Case path to load on startup')
        parser.add_argument('--tab', type=int, help='Tab index to open on startup')
        args, remaining = parser.parse_known_args()

        # Debug: Log received arguments
        print(f"DEBUG: sys.argv = {sys.argv}")
        print(f"DEBUG: Parsed args = file_number={args.file_number}, case_path={args.case_path}, tab={args.tab}")

        # Disable Chromium sandbox and security for local file editing
        os.environ["QTWEBENGINE_CHROMIUM_FLAGS"] = "--no-sandbox --disable-web-security"

        app = QApplication(sys.argv)

        # Log startup parameters
        log_event(f"Starting with args: file_number={args.file_number}, case_path={args.case_path}, tab={args.tab}")

        # Launch Main Window with restored state if provided
        window = MainWindow(
            file_number=args.file_number,
            case_path=args.case_path,
            initial_tab=args.tab
        )
        window.show()
        sys.exit(app.exec())
    except Exception as e:
        print(f"CRITICAL MAIN ERROR: {e}")
        import traceback
        traceback.print_exc()
        # Keep window open if possible or wait for input so user can see error
        input("Press Enter to exit...")
