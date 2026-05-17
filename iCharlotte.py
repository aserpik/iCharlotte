import sys
import os
import argparse
import ctypes

# Workaround: WMI queries hang indefinitely on some Windows machines.
# platform.system(), platform.platform(), etc. all call _wmi_query internally.
# Pre-populate the caches and disable _wmi_query before any library imports.
import platform as _platform_mod
_platform_mod._uname_cache = _platform_mod.uname_result(
    system='Windows',
    node=os.environ.get('COMPUTERNAME', ''),
    release='11',
    version='10.0.26200',
    machine=os.environ.get('PROCESSOR_ARCHITECTURE', 'AMD64'),
)
_platform_mod.win32_ver = lambda *a, **kw: ('11', '10.0.26200', '', 'Multiprocessor Free')
_orig_wmi_query = getattr(_platform_mod, '_wmi_query', None)
_platform_mod._wmi_query = lambda *a, **kw: (_ for _ in ()).throw(OSError('WMI disabled'))

# Set Windows App User Model ID BEFORE any Qt imports for proper taskbar icon
try:
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID('iCharlotte.LegalSuite.1')
except Exception:
    pass

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
import time
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
    from PySide6.QtCore import Qt, QThread, Signal, QFileInfo, QMetaObject, Q_ARG, QSettings, QTimer
    from PySide6.QtGui import QAction, QShortcut, QKeySequence, QIcon, QCursor
    from PySide6.QtWebEngineCore import QWebEngineUrlScheme
except ImportError:
    print("Error: PySide6 or its components are not installed. Please run: pip install PySide6 PySide6-WebEngine")
    sys.exit(1)

# Global hotkey support
try:
    import keyboard
    KEYBOARD_AVAILABLE = True
except (ImportError, OSError):
    KEYBOARD_AVAILABLE = False
    print("Warning: 'keyboard' library not available. Global hotkeys (Win+F, Win+C) disabled.")

# --- Core Modules ---
from icharlotte_core.config import SCRIPTS_DIR, GEMINI_DATA_DIR, BASE_PATH_WIN
from icharlotte_core.utils import (
    log_event, get_case_path, sanitize_filename, format_date_to_mm_dd_yyyy
)
from icharlotte_core.ui.widgets import (
    StatusWidget, AgentRunner, FileTreeWidget
)
from icharlotte_core.ui.case_view_enhanced import (
    EnhancedAgentButton, AgentSettingsDB, AgentSettingsDialog,
    AdvancedFilterWidget, FilePreviewWidget, OutputBrowserWidget,
    ProcessingLogWidget, ProcessingLogDB, FileTagsDB, EnhancedFileTreeWidget
)
from icharlotte_core.ui.dialogs import FileNumberDialog, VariablesDialog, PromptsDialog, LLMSettingsDialog
from icharlotte_core.ui.report_generator_dialog import ReportGeneratorDialog, ReportPipelineWorker
from icharlotte_core.subpoena_tracker import SubpoenaTrackerWorker
from icharlotte_core.ui.tabs import ChatTab, IndexTab
from icharlotte_core.ui.email_tab import EmailTab
from icharlotte_core.ui.email_update_tab import EmailUpdateTab
from icharlotte_core.ui.logs_tab import LogsTab
from icharlotte_core.ui.liability_tab import LiabilityExposureTab
from icharlotte_core.ui.master_case_tab import MasterCaseTab
from icharlotte_core.master_db import MasterCaseDatabase
from icharlotte_core.ui.templates_resources_tab import TemplatesResourcesTab
from icharlotte_core.ui.deposition_tab import DepositionTab
from icharlotte_core.ui.discovery_tab import DiscoveryTab
from icharlotte_core.word_hotkey import init_word_hotkey, stop_word_hotkey
from icharlotte_core.ui.zoom_handler import ZoomEventFilter
from icharlotte_core.app_crash_handler import (
    install_crash_handler, checkpoint, add_context,
    log_info, log_warning, log_error, log_debug, safe_slot
)

class QuickOpenDialog(QDialog):
    """Lightweight popup dialog for quick file number or plaintiff name entry via double-Ctrl tap."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Quick Open Case Folder")
        self.setFixedSize(400, 100)
        self.setWindowFlags(Qt.WindowType.Dialog | Qt.WindowType.FramelessWindowHint | Qt.WindowType.WindowStaysOnTopHint)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground, False)

        # Load recent cases for autocomplete
        self.recent_file = os.path.join(GEMINI_DATA_DIR, "recent_cases.json")
        self.recent_cases = self._load_recent_cases()

        # Load all cases from database for plaintiff name lookup
        self.db = MasterCaseDatabase()
        self.all_cases = self.db.get_all_cases()
        # Build a mapping of plaintiff names to file numbers (lowercase for matching)
        self.plaintiff_to_file = {}
        for case in self.all_cases:
            name = case.get('plaintiff_last_name', '').strip()
            file_num = case.get('file_number', '')
            if name and file_num:
                self.plaintiff_to_file[name.lower()] = file_num

        layout = QVBoxLayout(self)
        layout.setContentsMargins(15, 15, 15, 15)

        # Style the dialog (light/white theme)
        self.setStyleSheet("""
            QDialog {
                background-color: #ffffff;
                border: 2px solid #1565C0;
                border-radius: 10px;
            }
            QComboBox {
                background-color: #f5f5f5;
                color: #333;
                border: 1px solid #ccc;
                border-radius: 5px;
                padding: 8px 12px;
                font-size: 16px;
                min-height: 30px;
            }
            QComboBox:focus {
                border: 2px solid #1565C0;
                background-color: #fff;
            }
            QComboBox::drop-down {
                border: none;
                width: 20px;
            }
            QComboBox QAbstractItemView {
                background-color: #fff;
                color: #333;
                selection-background-color: #1565C0;
                selection-color: #fff;
            }
            QLabel {
                color: #666;
                font-size: 11px;
            }
        """)

        # File number input with autocomplete (includes both file numbers and plaintiff names)
        self.file_input = QComboBox()
        self.file_input.setEditable(True)
        self.file_input.lineEdit().setPlaceholderText("Enter file number or plaintiff name")

        # Build autocomplete list: recent cases first, then all cases with names
        autocomplete_items = list(self.recent_cases)  # Recent file numbers
        for case in self.all_cases:
            file_num = case.get('file_number', '')
            name = case.get('plaintiff_last_name', '').strip()
            if name and file_num:
                # Add "Name (####.###)" format for easy selection
                display = f"{name} ({file_num})"
                if display not in autocomplete_items and file_num not in autocomplete_items:
                    autocomplete_items.append(display)

        self.file_input.addItems(autocomplete_items)
        self.file_input.setCurrentIndex(-1)  # Start empty
        self.file_input.lineEdit().returnPressed.connect(self._on_enter)
        layout.addWidget(self.file_input)

        # Help text
        help_label = QLabel("Press Enter to open folder • Esc to close")
        help_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(help_label)

        self.result_file_number = None

    def _load_recent_cases(self):
        if os.path.exists(self.recent_file):
            try:
                with open(self.recent_file, 'r') as f:
                    return json.load(f)
            except:
                return []
        return []

    def _save_recent_case(self, file_num):
        if not file_num:
            return
        if file_num in self.recent_cases:
            self.recent_cases.remove(file_num)
        self.recent_cases.insert(0, file_num)
        self.recent_cases = self.recent_cases[:10]

        if not os.path.exists(GEMINI_DATA_DIR):
            os.makedirs(GEMINI_DATA_DIR)

        try:
            with open(self.recent_file, 'w') as f:
                json.dump(self.recent_cases, f)
        except:
            pass

    def _resolve_to_file_number(self, text):
        """Resolve user input to a file number.

        Handles:
        - Direct file number (####.###)
        - "Name (####.###)" format from autocomplete
        - Plaintiff name lookup
        """
        text = text.strip()
        if not text:
            return None

        # Check if it's the "Name (####.###)" format from autocomplete
        if '(' in text and text.endswith(')'):
            # Extract file number from parentheses
            start = text.rfind('(')
            file_num = text[start + 1:-1].strip()
            if re.match(r'^\d{4}[.\-]\d{3}$', file_num):
                return file_num.replace('-', '.')

        # Check if it's a file number format (####.### or ####-###)
        if re.match(r'^\d{4}[.\-]\d{3}$', text):
            return text.replace('-', '.')

        # Otherwise, try to find by plaintiff name
        text_lower = text.lower()

        # Exact match first
        if text_lower in self.plaintiff_to_file:
            return self.plaintiff_to_file[text_lower]

        # Partial match (find first case where name contains the search text)
        for name, file_num in self.plaintiff_to_file.items():
            if text_lower in name:
                return file_num

        # No match found - return input as-is (will fail gracefully in caller)
        return text

    def _on_enter(self):
        user_input = self.file_input.currentText().strip()
        if user_input:
            file_num = self._resolve_to_file_number(user_input)
            self.result_file_number = file_num
            self._save_recent_case(file_num)
            self.accept()

    def keyPressEvent(self, event):
        if event.key() == Qt.Key.Key_Escape:
            self.reject()
        else:
            super().keyPressEvent(event)

    def showEvent(self, event):
        super().showEvent(event)
        # Center on the screen where the mouse cursor is (for multi-monitor support)
        cursor_pos = QCursor.pos()
        screen = QApplication.screenAt(cursor_pos)
        if screen is None:
            screen = QApplication.primaryScreen()
        screen_geom = screen.geometry()
        self.move(
            screen_geom.x() + (screen_geom.width() - self.width()) // 2,
            screen_geom.y() + screen_geom.height() // 3
        )
        # Use a timer to ensure focus is set after window is fully rendered
        # This is necessary for multi-monitor setups where focus can be unreliable
        # 100ms delay needed on Windows for reliable focus stealing
        QTimer.singleShot(100, self._ensure_focus)

    def _ensure_focus(self):
        """Ensure the input field has focus after the window is shown."""
        # On Windows, we need to use native APIs to force foreground focus
        if sys.platform == 'win32':
            try:
                import ctypes
                hwnd = int(self.winId())
                # Attach to the foreground window's thread to get permission to set foreground
                user32 = ctypes.windll.user32
                foreground_hwnd = user32.GetForegroundWindow()
                foreground_thread = user32.GetWindowThreadProcessId(foreground_hwnd, None)
                current_thread = ctypes.windll.kernel32.GetCurrentThreadId()
                # Attach threads to allow focus stealing
                user32.AttachThreadInput(foreground_thread, current_thread, True)
                user32.SetForegroundWindow(hwnd)
                user32.AttachThreadInput(foreground_thread, current_thread, False)
            except Exception:
                pass  # Fall back to Qt methods if native approach fails

        self.activateWindow()
        self.raise_()
        self.file_input.setFocus()
        self.file_input.lineEdit().setFocus()
        self.file_input.lineEdit().selectAll()


class DirectoryTreeWorker(QThread):
    data_ready = Signal(list) # Emits (root, dirs, files) tuples
    finished = Signal()

    def __init__(self, root_path):
        super().__init__()
        self.root_path = root_path
        self.running = True
        
    def run(self):
        import time as _time
        _start = _time.monotonic()
        _batch_count = 0
        _total_dirs = 0
        _total_files = 0
        log_debug(f"DirectoryTreeWorker: scanning {self.root_path}")
        try:
            batch = []
            for root, dirs, files in os.walk(self.root_path):
                if not self.running:
                    log_debug(f"DirectoryTreeWorker: stopped early after {_time.monotonic()-_start:.1f}s")
                    break
                # Skip hidden files/dirs
                dirs[:] = [d for d in dirs if not d.startswith('.') and not d.startswith('~$')]

                file_data = []
                for f in files:
                    if not self.running:
                        break
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

                _total_dirs += len(dirs)
                _total_files += len(file_data)
                batch.append((root, dirs, file_data))
                if len(batch) >= 10:
                    _batch_count += 1
                    self.data_ready.emit(batch)
                    batch = []
                    self.msleep(10) # Yield to UI

            if batch:
                _batch_count += 1
                self.data_ready.emit(batch)
            log_debug(f"DirectoryTreeWorker: done in {_time.monotonic()-_start:.1f}s, {_batch_count} batches, {_total_dirs} dirs, {_total_files} files")
            self.finished.emit()
        except Exception as e:
            log_error(f"DirectoryTreeWorker CRASHED after {_time.monotonic()-_start:.1f}s: {e}", exc_info=True)
            self.finished.emit()

    def stop(self):
        self.running = False

class MainWindow(QMainWindow):
    # Signals for thread-safe hotkey callbacks
    open_file_signal = Signal()
    change_file_signal = Signal()
    quick_open_signal = Signal()  # For double-Ctrl quick open
    ctrl_press_signal = Signal()  # For thread-safe Ctrl press handling
    ctrl_release_signal = Signal()  # For thread-safe Ctrl release handling

    def __init__(self, file_number=None, case_path=None, initial_tab=None):
        super().__init__()
        checkpoint("MainWindow.__init__ starting", file_number=file_number)

        # Connect signals for global hotkeys (thread-safe)
        self.open_file_signal.connect(self._on_open_file_hotkey)
        self.change_file_signal.connect(self._on_change_file_hotkey)
        self.quick_open_signal.connect(self._on_quick_open_hotkey)
        self.ctrl_press_signal.connect(self._handle_ctrl_press)
        self.ctrl_release_signal.connect(self._handle_ctrl_release)

        # Double-Ctrl tap detection state
        self._last_ctrl_tap_time = None  # Time of last valid tap (quick press+release), None = no recent tap
        self._ctrl_press_time = 0  # When Ctrl was pressed
        self._ctrl_tap_threshold = 0.5  # Max seconds between taps for double-tap
        self._ctrl_hold_threshold = 0.2  # Max seconds Ctrl can be held to count as tap
        self.file_number = file_number
        self.case_path = case_path
        add_context('current_file_number', file_number)
        add_context('current_case_path', case_path)
        self._update_window_title()
        self.resize(1200, 800)

        # Set application icon
        icon_path = os.path.join(os.path.dirname(__file__), 'icharlotte.ico')
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))

        # Set AppUserModelID on window handle for proper taskbar grouping
        try:
            from ctypes import wintypes
            hwnd = int(self.winId())
            shell32 = ctypes.windll.shell32
            propvariant = ctypes.create_unicode_buffer('iCharlotte.LegalSuite.1')
            shell32.SHGetPropertyStoreForWindow.argtypes = [wintypes.HWND, ctypes.POINTER(ctypes.c_void_p), ctypes.POINTER(ctypes.c_void_p)]
            shell32.SHGetPropertyStoreForWindow.restype = ctypes.HRESULT
        except Exception:
            pass

        self.agent_runners = [] # Keep references to prevent GC
        self._agent_queue = []  # Queued agents waiting to run (when concurrency limit reached)
        self.MAX_CONCURRENT_AGENTS = 4  # Max subprocess agents running simultaneously
        self._tree_generation = 0  # Incremented on each populate_tree to ignore stale worker callbacks
        self._populating_tree = False  # Re-entrancy guard for populate_tree
        self._tree_refresh_timer = QTimer()
        self._tree_refresh_timer.setSingleShot(True)
        self._tree_refresh_timer.timeout.connect(self.populate_tree)
        self.cached_models = {} # Cache for models: {provider: [list]}
        self.fetcher = None

        log_info(f"Initializing MainWindow for {file_number} at {case_path}")
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

        # Register global hotkeys (Win+F for Open File, Win+C for Change File)
        self._setup_global_hotkeys()

    def _setup_global_hotkeys(self):
        """Register global hotkeys for Open File (Win+F), Change File (Win+C), double-Ctrl quick open, and Win+V for Word AI."""
        if not KEYBOARD_AVAILABLE:
            return

        try:
            # Register Win+F for Open File
            keyboard.add_hotkey('win+f', lambda: self.open_file_signal.emit(), suppress=True)
            # Register Win+C for Change File
            keyboard.add_hotkey('win+c', lambda: self.change_file_signal.emit(), suppress=True)
            # Register Ctrl key press and release for double-tap detection
            keyboard.on_press_key('ctrl', self._on_ctrl_press)
            keyboard.on_release_key('ctrl', self._on_ctrl_release)
            log_event("Global hotkeys registered: Win+F (Open File), Win+C (Change File), Double-Ctrl (Quick Open)")
        except Exception as e:
            log_event(f"Failed to register global hotkeys: {e}", "error")

        # Register Word AI hotkey (Win+V)
        try:
            if init_word_hotkey(self):
                log_event("Word AI hotkey registered: Win+V")
        except Exception as e:
            log_event(f"Failed to register Word AI hotkey: {e}", "error")

    def _on_ctrl_press(self, event):
        """Called from keyboard library thread - emit signal for thread safety."""
        self.ctrl_press_signal.emit()

    def _on_ctrl_release(self, event):
        """Called from keyboard library thread - emit signal for thread safety."""
        self.ctrl_release_signal.emit()

    def _handle_ctrl_press(self):
        """Handle Ctrl press on main thread - record time for tap duration calculation."""
        self._ctrl_press_time = time.time()

    def _handle_ctrl_release(self):
        """Handle Ctrl release on main thread - detect double-tap for quick open dialog.

        Only triggers if:
        1. Ctrl was held briefly (< 200ms) - this is a "tap", not a held key
        2. Two taps occurred within 500ms of each other
        """
        current_time = time.time()
        hold_duration = current_time - self._ctrl_press_time

        # Only count as a tap if Ctrl was held briefly (not held for shortcuts)
        if hold_duration > self._ctrl_hold_threshold:
            # Ctrl was held too long - this was probably a shortcut, not a tap
            self._last_ctrl_tap_time = None
            return

        # This is a valid tap - check if it's a double-tap
        if self._last_ctrl_tap_time is not None:
            time_since_last_tap = current_time - self._last_ctrl_tap_time
            if time_since_last_tap < self._ctrl_tap_threshold:
                # Double-tap detected - emit signal (thread-safe)
                self.quick_open_signal.emit()
                self._last_ctrl_tap_time = None  # Reset to prevent triple-tap triggering
                return

        # Record this tap for potential double-tap detection
        self._last_ctrl_tap_time = current_time

    def _on_open_file_hotkey(self):
        """Handle Win+F hotkey - bring window to front and open file."""
        self.activateWindow()
        self.raise_()
        self.showNormal()
        self.open_root_folder()

    def _on_change_file_hotkey(self):
        """Handle Win+C hotkey - bring window to front and change file."""
        self.activateWindow()
        self.raise_()
        self.showNormal()
        self.change_file()

    def _on_quick_open_hotkey(self):
        """Handle double-Ctrl hotkey - show quick open dialog to open case folder."""
        dialog = QuickOpenDialog(None)  # No parent for independent focus
        dialog.setWindowFlag(Qt.WindowType.WindowStaysOnTopHint, True)
        if dialog.exec() == QDialog.DialogCode.Accepted and dialog.result_file_number:
            file_num = dialog.result_file_number
            case_path = get_case_path(file_num)
            if case_path and os.path.exists(case_path):
                try:
                    os.startfile(case_path)
                    log_event(f"Quick open: Opened folder for {file_num}")
                except Exception as e:
                    QMessageBox.warning(self, "Error", f"Could not open folder: {e}")
            else:
                QMessageBox.warning(self, "Error", f"Folder not found for {file_num}")

    def _update_window_title(self):
        """Update window title based on current file_number and case_path."""
        if self.file_number and self.case_path:
            self.setWindowTitle(f"iCharlotte - {self.file_number} - {os.path.basename(self.case_path)}")
        else:
            self.setWindowTitle("iCharlotte")

    def setup_ui(self):
        checkpoint("setup_ui starting - creating tabs")
        self.tabs = QTabWidget()
        self.setCentralWidget(self.tabs)

        # Make tabs expand to fill available width (no scroll arrows needed)
        self.tabs.tabBar().setExpanding(True)
        self.tabs.setUsesScrollButtons(False)

        # Style the tab bar for larger, more visible tabs
        self.tabs.setStyleSheet("""
            QTabWidget::pane {
                border: 1px solid #ccc;
                border-top: none;
            }
            QTabBar::tab {
                background-color: #e0e0e0;
                color: #333;
                font-size: 13px;
                font-weight: 500;
                padding: 10px 6px;
                margin-right: 1px;
                border: 1px solid #ccc;
                border-bottom: none;
                border-top-left-radius: 6px;
                border-top-right-radius: 6px;
            }
            QTabBar::tab:selected {
                background-color: #fff;
                color: #1565C0;
                font-weight: bold;
                border-bottom: 2px solid #1565C0;
            }
            QTabBar::tab:hover:!selected {
                background-color: #f0f0f0;
                color: #1976D2;
            }
        """)

        # --- Tab 0: Master List ---
        checkpoint("Creating MasterCaseTab")
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

        btn_notes = QPushButton("Notes")
        btn_notes.clicked.connect(self.open_notes)
        toolbar_layout.addWidget(btn_notes)

        btn_vars = QPushButton("Variables")
        btn_vars.clicked.connect(self.manage_variables)
        toolbar_layout.addWidget(btn_vars)

        # Output Browser Button
        btn_outputs = QPushButton("Output Browser")
        btn_outputs.clicked.connect(self.open_output_browser)
        toolbar_layout.addWidget(btn_outputs)

        # Processing Log Button
        btn_proc_log = QPushButton("Processing Log")
        btn_proc_log.clicked.connect(self.open_processing_log)
        toolbar_layout.addWidget(btn_proc_log)

        btn_generate_report = QPushButton("Generate Report")
        btn_generate_report.setStyleSheet("background-color: #2196F3; color: white; font-weight: bold;")
        btn_generate_report.clicked.connect(self.open_report_generator)
        toolbar_layout.addWidget(btn_generate_report)

        toolbar_layout.addStretch()

        # Wrapper for vertical layout of Case View
        wrapper_layout = QVBoxLayout()
        wrapper_layout.addLayout(toolbar_layout)

        # Main horizontal splitter (agents | tree | preview)
        self.main_splitter = QSplitter(Qt.Orientation.Horizontal)
        wrapper_layout.addWidget(self.main_splitter)
        self.main_splitter.splitterMoved.connect(self._on_splitter_moved)

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
        # Subpoena Tracker — in-process QThread worker (not a subprocess agent)
        self._subpoena_btn = EnhancedAgentButton("Subpoena Tracker", "subpoena_tracker")
        self._subpoena_btn.clicked.connect(self._run_subpoena_tracker)
        self.agent_buttons["subpoena_tracker"] = self._subpoena_btn
        left_layout.addWidget(self._subpoena_btn)
        # Med Record Extractor — in-process QThread worker
        self._med_extract_btn = EnhancedAgentButton("Med Record Extractor", "med_record_extractor")
        self._med_extract_btn.clicked.connect(self._run_med_record_extractor)
        self.agent_buttons["med_record_extractor"] = self._med_extract_btn
        left_layout.addWidget(self._med_extract_btn)

        # Document Agents
        left_layout.addSpacing(8)
        new_label = QLabel("Document Agents")
        new_label.setStyleSheet("font-weight: bold; font-size: 12px; color: #666;")
        left_layout.addWidget(new_label)

        left_layout.addStretch()
        self.main_splitter.addWidget(left_panel)
        
        # Center Panel (File Tree with Enhanced Features)
        center_panel = QFrame()
        center_layout = QVBoxLayout(center_panel)

        # Header Layout (Status Label + Expand/Collapse Button)
        header_layout = QHBoxLayout()

        self.refresh_btn = QPushButton()
        self.refresh_btn.setIcon(self.style().standardIcon(self.style().StandardPixmap.SP_BrowserReload))
        self.refresh_btn.setToolTip("Refresh Tree")
        self.refresh_btn.clicked.connect(self._schedule_tree_refresh)
        header_layout.addWidget(self.refresh_btn)

        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Filter files...")
        self.search_input.textChanged.connect(self.filter_tree)
        header_layout.addWidget(self.search_input)

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
        ]

        # Enhanced File Tree with additional columns
        self.tree = EnhancedFileTreeWidget()
        self.tree.item_moved.connect(lambda: self._schedule_tree_refresh())
        self.tree.folder_created.connect(lambda p: self._schedule_tree_refresh())
        self.tree.setHeaderLabels([
            "Category / File",
            "Queued Tasks (Click to Add ➕)",
            "Size",
            "Date Modified",
            "Status"
        ])
        self.tree.setSortingEnabled(True)
        self.tree.setColumnWidth(0, 300)
        self.tree.setColumnWidth(1, 200)
        self.tree.setColumnWidth(2, 70)
        self.tree.setColumnWidth(3, 100)
        self.tree.setColumnWidth(4, 80)
        self.tree.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.tree.setAlternatingRowColors(True)
        self.tree.itemSelectionChanged.connect(self.on_tree_selection_changed)
        self.tree.itemDoubleClicked.connect(self.on_tree_double_click)
        self.tree.itemClicked.connect(self.on_tree_item_clicked)
        self.tree.tasks_column_clicked.connect(self.on_tree_item_clicked)
        center_layout.addWidget(self.tree)

        btn_layout = QHBoxLayout()
        self.process_btn = QPushButton("Process All Queued Tasks")
        self.process_btn.setStyleSheet("background-color: #4CAF50; color: white; font-weight: bold; padding: 10px;")
        self.process_btn.clicked.connect(self.process_checked_items)
        btn_layout.addWidget(self.process_btn)

        center_layout.addLayout(btn_layout)

        self.main_splitter.addWidget(center_panel)

        # Right Panel (File Preview - Hidden by default)
        self.preview_pane = FilePreviewWidget()
        self.preview_pane.hide()
        self.main_splitter.addWidget(self.preview_pane)

        self.main_splitter.setSizes([180, 800, 0])
        self.main_splitter.setStretchFactor(0, 0)
        self.main_splitter.setStretchFactor(1, 1)
        self.main_splitter.setStretchFactor(2, 0)
        
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
        checkpoint("Creating ChatTab")
        self.chat_tab = ChatTab()
        self.tabs.addTab(self.chat_tab, "Chat")
        if self.file_number:
            self.chat_tab.load_case(self.file_number)

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
        if self.file_number:
            self.email_update_tab.on_case_changed(self.file_number)

        # --- Tab: Depositions ---
        self.deposition_tab = DepositionTab()
        self.tabs.addTab(self.deposition_tab, "Depositions")
        if self.file_number:
            self.deposition_tab.load_case(self.file_number)

        # --- Tab: Discovery ---
        self.discovery_tab = DiscoveryTab()
        self.tabs.addTab(self.discovery_tab, "Discovery")
        if self.file_number:
            self.discovery_tab.load_case(self.file_number)

        # --- Tab: Liability & Exposure ---
        self.liability_tab = LiabilityExposureTab()
        self.tabs.addTab(self.liability_tab, "Liability & Exposure")

        # --- Tab: Templates / Resources ---
        self.templates_tab = TemplatesResourcesTab(main_window=self)
        self.tabs.addTab(self.templates_tab, "Templates / Resources")

        # --- Tab 8: Logs ---
        self.logs_tab = LogsTab(self)
        self.tabs.addTab(self.logs_tab, "Logs")

        # Add corner buttons next to the tabs
        self.corner_widget = QWidget()
        self.corner_layout = QHBoxLayout(self.corner_widget)
        self.corner_layout.setContentsMargins(5, 5, 10, 5)
        self.corner_layout.setSpacing(8)

        # Common button style for primary actions (colored buttons)
        primary_btn_style = """
            QPushButton {{
                background-color: {bg};
                color: white;
                font-weight: bold;
                font-size: 13px;
                padding: 0px 16px;
                border: none;
                border-radius: 4px;
                height: 34px;
            }}
            QPushButton:hover {{
                background-color: {hover};
            }}
            QPushButton:pressed {{
                background-color: {pressed};
            }}
        """

        # Common button style for secondary actions (neutral buttons)
        secondary_btn_style = """
            QPushButton {
                background-color: #f5f5f5;
                color: #333;
                font-weight: 500;
                font-size: 13px;
                padding: 0px 14px;
                border: 1px solid #ccc;
                border-radius: 4px;
                height: 34px;
            }
            QPushButton:hover {
                background-color: #e8e8e8;
                border-color: #999;
            }
            QPushButton:pressed {
                background-color: #ddd;
            }
        """

        # Open File button (blue - primary)
        self.btn_open_root = QPushButton("Open File")
        self.btn_open_root.setStyleSheet(primary_btn_style.format(
            bg="#1976D2", hover="#1565C0", pressed="#0D47A1"
        ))
        self.btn_open_root.setToolTip("Open case folder in Explorer (Win+F)")
        self.btn_open_root.clicked.connect(self.open_root_folder)
        self.corner_layout.addWidget(self.btn_open_root)

        # Change File button (green - primary)
        self.btn_change_file = QPushButton("Change File")
        self.btn_change_file.setStyleSheet(primary_btn_style.format(
            bg="#388E3C", hover="#2E7D32", pressed="#1B5E20"
        ))
        self.btn_change_file.setToolTip("Switch to different case (Win+C)")
        self.btn_change_file.clicked.connect(self.change_file)
        self.corner_layout.addWidget(self.btn_change_file)

        # View menu button
        self.setup_view_menu()
        self.corner_layout.addWidget(self.view_btn)

        # Prompts button (secondary)
        self.prompts_btn = QPushButton("Prompts")
        self.prompts_btn.setStyleSheet(secondary_btn_style)
        self.prompts_btn.setToolTip("Open Prompt Engineering Workbench")
        self.prompts_btn.clicked.connect(self.manage_prompts)
        self.corner_layout.addWidget(self.prompts_btn)

        # Settings button with dropdown menu
        self.settings_btn = QToolButton()
        self.settings_btn.setText("Settings ▾")
        self.settings_btn.setPopupMode(QToolButton.ToolButtonPopupMode.InstantPopup)
        self.settings_btn.setStyleSheet(secondary_btn_style.replace("QPushButton", "QToolButton") + """
            QToolButton::menu-indicator { image: none; }
        """)
        self.settings_menu = QMenu(self)
        self.settings_menu.addAction("LLM Settings", self.open_settings_dialog)
        self.settings_menu.addSeparator()
        self.email_monitor_action = self.settings_menu.addAction("Email Monitor")
        self.email_monitor_action.setCheckable(True)
        self.docket_refresh_action = self.settings_menu.addAction("Auto Docket Refresh")
        self.docket_refresh_action.setCheckable(True)
        self.settings_btn.setMenu(self.settings_menu)
        self.corner_layout.addWidget(self.settings_btn)

        # Wire settings actions to master_case_tab's toggle methods
        self.master_tab.email_monitor_action = self.email_monitor_action
        self.master_tab.docket_refresh_action = self.docket_refresh_action
        self.email_monitor_action.toggled.connect(self.master_tab.toggle_email_monitor)
        self.docket_refresh_action.toggled.connect(self.master_tab.toggle_docket_refresh)

        # Restart button (red - danger)
        self.restart_btn = QPushButton("Restart")
        self.restart_btn.setStyleSheet(primary_btn_style.format(
            bg="#d32f2f", hover="#c62828", pressed="#b71c1c"
        ))
        self.restart_btn.clicked.connect(self.restart_app)
        self.corner_layout.addWidget(self.restart_btn)

        self.tabs.setCornerWidget(self.corner_widget, Qt.Corner.TopRightCorner)

        # Ctrl+Scroll zoom for all tabs (skips PDF viewers)
        self._zoom_filter = ZoomEventFilter(self.tabs, self)
        QApplication.instance().installEventFilter(self._zoom_filter)

        checkpoint("setup_ui complete - all tabs created")

    def setup_view_menu(self):
        self.view_btn = QToolButton()
        self.view_btn.setText("View ▾")
        self.view_btn.setPopupMode(QToolButton.ToolButtonPopupMode.InstantPopup)
        self.view_btn.setStyleSheet("""
            QToolButton {
                background-color: #f5f5f5;
                color: #333;
                font-weight: 500;
                font-size: 13px;
                padding: 0px 14px;
                border: 1px solid #ccc;
                border-radius: 4px;
                height: 34px;
            }
            QToolButton:hover {
                background-color: #e8e8e8;
                border-color: #999;
            }
            QToolButton:pressed {
                background-color: #ddd;
            }
            QToolButton::menu-indicator {
                image: none;
            }
        """)
        
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
        if column == 1:  # Queued Tasks column (index 1)
            # Get click position relative to the tree's viewport
            pos = self.tree.visualItemRect(item).bottomLeft()
            pos.setX(self.tree.columnViewportPosition(1))
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
            try:
                current_tasks = list(item.data(0, Qt.ItemDataRole.UserRole + 2) or [])
            except RuntimeError:
                continue
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
            try:
                item.setData(0, Qt.ItemDataRole.UserRole + 2, [])
                self.update_item_tasks_ui(item)
            except RuntimeError:
                continue
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
            item.setText(1, " [ + Add Tasks ]")
            item.setForeground(1, Qt.GlobalColor.gray)
            return

        # Create a visual string of short tags
        tags = []
        for tid in task_ids:
            agent = next((a for a in self.AGENTS if a["id"] == tid), None)
            if agent:
                tags.append(f"[{agent['short']}]")

        item.setText(1, " ".join(tags))
        item.setForeground(1, Qt.GlobalColor.blue)

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
                # Save chat conversation before switching cases
                if hasattr(self, 'chat_tab') and self.chat_tab:
                    self.chat_tab.save_current_state()
                self.file_number = new_file_num
                self.case_path = new_path
                self.setWindowTitle(f"iCharlotte - {self.file_number} - {os.path.basename(self.case_path)}")
                self.populate_tree()
                # Clear status list to reset state
                self.clear_all_status()
                self.load_status_history()

                # Reset agent buttons, then restore running state for agents on this case
                for btn in self.agent_buttons.values():
                    btn.set_running(False)
                for script, case_num in self.running_agents.items():
                    if case_num == new_file_num and script in self.agent_buttons:
                        self.agent_buttons[script].set_running(True)

                # Reset Tabs for new case isolation
                if hasattr(self, 'index_tab'):
                    self.index_tab.load_data(self.file_number)
                if hasattr(self, 'chat_tab'):
                    self.chat_tab.load_case(self.file_number)
                if hasattr(self, 'liability_tab'):
                    self.liability_tab.reset_state()
                if hasattr(self, 'email_tab'):
                    self.email_tab.search_bar.clear()
                    self.email_tab.check_db_init()
                    self.email_tab.perform_search()
                if hasattr(self, 'email_update_tab'):
                    self.email_update_tab.on_case_changed(new_file_num)
                if hasattr(self, 'discovery_tab'):
                    self.discovery_tab.load_case(new_file_num)

                log_event(f"Switched to case {new_file_num}")
            else:
                QMessageBox.critical(self, "Error", f"Could not find case directory for {new_file_num}")

    def open_root_folder(self):
        if not self.case_path:
            QMessageBox.information(self, "Info", "No case loaded.")
            return
        if os.path.exists(self.case_path):
            try:
                os.startfile(self.case_path)
            except Exception as e:
                QMessageBox.warning(self, "Error", f"Could not open directory: {e}")

    def load_case_by_number(self, file_number):
        log_debug(f"load_case_by_number: switching to {file_number}")
        new_path = get_case_path(file_number)

        if new_path:
            log_debug(f"load_case_by_number: path={new_path}")
            self.save_status_history()
            # Save chat conversation before switching cases
            if hasattr(self, 'chat_tab') and self.chat_tab:
                self.chat_tab.save_current_state()
            self.file_number = file_number
            self.case_path = new_path
            self._update_window_title()
            self.populate_tree()
            self.clear_all_status()

            # Reset agent buttons, then restore running state for agents on this case
            for btn in self.agent_buttons.values():
                btn.set_running(False)
            for script, case_num in self.running_agents.items():
                if case_num == file_number and script in self.agent_buttons:
                    self.agent_buttons[script].set_running(True)

            self.load_status_history()

            # Switch to Case View tab (Index 1)
            self.tabs.setCurrentIndex(1)

            # Reset Tabs
            if hasattr(self, 'index_tab'):
                self.index_tab.load_data(self.file_number)
            if hasattr(self, 'chat_tab'):
                self.chat_tab.load_case(self.file_number)
            if hasattr(self, 'liability_tab'):
                self.liability_tab.reset_state()
            if hasattr(self, 'email_tab'):
                self.email_tab.search_bar.clear()
                self.email_tab.check_db_init()
                self.email_tab.perform_search()
            if hasattr(self, 'email_update_tab'):
                self.email_update_tab.on_case_changed(file_number)
            if hasattr(self, 'deposition_tab'):
                self.deposition_tab.load_case(file_number)
            if hasattr(self, 'discovery_tab'):
                self.discovery_tab.load_case(file_number)

            log_event(f"Switched to case {self.file_number}")
        else:
            QMessageBox.critical(self, "Error", f"Could not find case directory for {file_number}")

    def view_docket(self):
        if not self.case_path:
            QMessageBox.information(self, "Info", "No case loaded.")
            return
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

    def open_notes(self):
        """Open or create the AS NOTES document for the current case."""
        if not hasattr(self, 'case_path') or not self.case_path:
            QMessageBox.information(self, "Info", "No case loaded.")
            return
        if not hasattr(self, 'file_number') or not self.file_number:
            QMessageBox.information(self, "Info", "No case loaded.")
            return

        notes_dir = os.path.join(self.case_path, "NOTES")
        notes_filename = f"AS NOTES - {self.file_number}.docx"
        notes_path = os.path.join(notes_dir, notes_filename)

        # Create NOTES directory if it doesn't exist
        if not os.path.exists(notes_dir):
            try:
                os.makedirs(notes_dir)
            except Exception as e:
                QMessageBox.warning(self, "Error", f"Could not create NOTES directory: {e}")
                return

        # Create the document if it doesn't exist
        if not os.path.exists(notes_path):
            try:
                from docx import Document
                doc = Document()
                doc.add_heading(f"AS Notes - {self.file_number}", level=1)
                doc.add_paragraph("")
                doc.save(notes_path)
            except Exception as e:
                QMessageBox.warning(self, "Error", f"Could not create notes document: {e}")
                return

        # Open the document
        try:
            os.startfile(notes_path)
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Could not open notes document: {e}")

    def manage_variables(self):
        dialog = VariablesDialog(self.file_number, self)
        dialog.exec()

    def manage_prompts(self):
        # Get current tab index to auto-select relevant agent
        current_tab_index = self.tabs.currentIndex()
        current_tab_name = self.tabs.tabText(current_tab_index)
        dialog = PromptsDialog(self, current_tab=current_tab_name)
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
                self.status_label.setText(f"{file_count} files selected - Click 'Queued Tasks' column to apply tasks to all selected")
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

    def run_separator_path(self, path, sensitivity=2):
        # Switch to Status Tab to show progress
        self.tabs.setCurrentIndex(1)
        
        status_widget = StatusWidget("Separator Agent", f"Analyzing {os.path.basename(path)}")
        self.status_list_layout.insertWidget(0, status_widget)
        
        script_path = os.path.join(SCRIPTS_DIR, "separate.py")
        args = [script_path, "--headless", "--sensitivity", str(sensitivity), path]
        
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

            # Re-enable sensitivity controls even on failure
            if hasattr(self, 'index_tab') and hasattr(self.index_tab, 'reanalyze_btn'):
                self.index_tab.reanalyze_btn.setEnabled(True)
                self.index_tab.sensitivity_slider.setEnabled(True)

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
        try:
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
        except Exception as e:
            log_event(f"Error in on_agent_finished: {e}", "error")

    def _run_subpoena_tracker(self):
        """Launch the in-process subpoena tracker worker."""
        if "subpoena_tracker" in self.running_agents:
            return  # Already running

        if not self.case_path:
            from PySide6.QtWidgets import QMessageBox
            QMessageBox.warning(self, "No Case", "No case is currently loaded.")
            return

        btn = self._subpoena_btn
        btn.set_running(True)
        started_for_case = self.file_number
        self.running_agents["subpoena_tracker"] = started_for_case

        worker = SubpoenaTrackerWorker(self.case_path, file_number=self.file_number)
        worker.progress.connect(lambda msg: log_event(f"Subpoena Tracker: {msg}"))
        worker.warning.connect(lambda msg: log_event(f"Subpoena Tracker warning: {msg}", "warning"))
        worker.finished_result.connect(
            lambda success, result: self._on_subpoena_tracker_finished(
                worker, btn, started_for_case, success, result
            )
        )
        self.agent_runners.append(worker)
        worker.start()

    def _on_subpoena_tracker_finished(self, worker, btn_widget, started_for_case, success, result):
        """Handle subpoena tracker completion."""
        try:
            if "subpoena_tracker" in self.running_agents:
                del self.running_agents["subpoena_tracker"]

            # Clean up the worker reference
            if worker in self.agent_runners:
                self.agent_runners.remove(worker)

            if self.file_number == started_for_case:
                btn_widget.set_running(False)
                if success:
                    btn_widget.set_status("Last: Just now")
                    log_event(f"Subpoena Tracker complete: {result}")
                else:
                    btn_widget.set_status("Last: Failed")
                    log_event(f"Subpoena Tracker failed: {result}", "error")
        except Exception as e:
            log_event(f"Error in _on_subpoena_tracker_finished: {e}", "error")

    def _run_med_record_extractor(self):
        """Open dialog and launch med record extraction worker."""
        if "med_record_extractor" in self.running_agents:
            return
        if not self.case_path:
            from PySide6.QtWidgets import QMessageBox
            QMessageBox.warning(self, "No Case", "No case is currently loaded.")
            return

        from icharlotte_core.ui.med_record_extractor_dialog import MedRecordExtractorDialog
        dialog = MedRecordExtractorDialog(parent=self)
        if not dialog.exec():
            return

        user_text = dialog.get_text()
        if not user_text:
            return

        btn = self._med_extract_btn
        btn.set_running(True)
        started_for_case = self.file_number
        self.running_agents["med_record_extractor"] = started_for_case

        # Add status widget to Status tab
        from icharlotte_core.ui.widgets import StatusWidget
        status_widget = StatusWidget("Med Record Extractor", f"Case {self.file_number}")
        self.status_list_layout.insertWidget(0, status_widget)

        from icharlotte_core.med_record_extractor import MedRecordExtractorWorker
        worker = MedRecordExtractorWorker(self.case_path, self.file_number, user_text)
        worker.progress.connect(lambda msg: (
            log_event(f"Med Record Extractor: {msg}"),
            status_widget.append_log(msg + "\n"),
            status_widget.status_text_label.setText(msg),
        ))
        worker.warning.connect(lambda msg: (
            log_event(f"Med Record Extractor warning: {msg}", "warning"),
            status_widget.append_log(f"WARNING: {msg}\n"),
        ))
        worker.finished_result.connect(
            lambda success, result: self._on_med_extract_finished(
                worker, btn, started_for_case, success, result, status_widget
            )
        )
        self.agent_runners.append(worker)
        worker.start()

    def _on_med_extract_finished(self, worker, btn_widget, started_for_case, success, result, status_widget=None):
        """Handle med record extractor completion."""
        try:
            if "med_record_extractor" in self.running_agents:
                del self.running_agents["med_record_extractor"]
            if worker in self.agent_runners:
                self.agent_runners.remove(worker)
            if self.file_number == started_for_case:
                btn_widget.set_running(False)
                if success:
                    btn_widget.set_status("Last: Just now")
                    log_event(f"Med Record Extractor complete: {result}")
                    if status_widget:
                        status_widget.append_log(f"\n{result}\n")
                        status_widget.set_finished(True)
                        # Set output to the extracts folder
                        import os
                        out_dir = os.path.join(self.case_path, "NOTES", "AI OUTPUT", "Med Record Extracts")
                        if os.path.isdir(out_dir):
                            status_widget.set_output_file(out_dir)
                else:
                    btn_widget.set_status("Last: Failed")
                    log_event(f"Med Record Extractor failed: {result}", "error")
                    if status_widget:
                        status_widget.append_log(f"\nFAILED: {result}\n")
                        status_widget.set_finished(False)
        except Exception as e:
            log_event(f"Error in _on_med_extract_finished: {e}", "error")

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

    def open_report_generator(self):
        """Open the report generator dialog and start pipeline if accepted."""
        if not self.file_number:
            QMessageBox.warning(self, "No Case", "Please load a case first.")
            return

        dialog = ReportGeneratorDialog(self.file_number, self)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return

        config = dialog.get_config()

        # Create status widget
        status_widget = StatusWidget(
            "Report Generator",
            f"{config['report_type']} for {self.file_number}"
        )
        self.status_list_layout.insertWidget(0, status_widget)

        # Create and start worker
        worker = ReportPipelineWorker(config, self)
        self.agent_runners.append(worker)

        # Connect signals to status widget
        worker.stage_started.connect(status_widget.update_pass_start)
        worker.stage_completed.connect(status_widget.update_pass_complete)
        worker.stage_failed.connect(status_widget.update_pass_failed)
        worker.progress_update.connect(status_widget.update_progress)
        worker.log_update.connect(lambda msg: status_widget.append_log(msg))
        worker.output_file.connect(status_widget.set_output_file)
        worker.finished_signal.connect(lambda ok: status_widget.set_finished(ok))
        worker.finished_signal.connect(lambda: self._cleanup_report_worker(worker))

        # Switch to status tab area
        self.tabs.setCurrentIndex(1)

        worker.start()

    def _cleanup_report_worker(self, worker):
        """Clean up finished report pipeline worker."""
        try:
            if worker in self.agent_runners:
                self.agent_runners.remove(worker)
        except Exception as e:
            log_event(f"Error cleaning up report worker: {e}", "error")

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
            current = self.main_splitter.sizes()
            if len(current) >= 3:
                total = sum(current)
                # Load saved preview width or use default (25% of total)
                saved_width = self._load_preview_width()
                if saved_width and saved_width > 50:
                    preview_width = min(saved_width, total - 250)  # Leave room for left and center
                else:
                    preview_width = int(total * 0.25)
                center_width = total - 180 - preview_width
                self.main_splitter.setSizes([180, center_width, preview_width])
            # Show currently selected file in preview
            selected = self.tree.selectedItems()
            if len(selected) == 1:
                item = selected[0]
                path = item.data(0, Qt.ItemDataRole.UserRole)
                item_type = item.data(0, Qt.ItemDataRole.UserRole + 1)
                if path and item_type == "file":
                    self.preview_pane.show_file(path)
        else:
            self.preview_pane.hide()
            self.preview_pane.clear()

    def _on_splitter_moved(self, pos, index):
        """Save preview pane width when splitter is moved."""
        if self.preview_pane.isVisible():
            sizes = self.main_splitter.sizes()
            if len(sizes) >= 3 and sizes[2] > 50:
                self._save_preview_width(sizes[2])

    def _save_preview_width(self, width):
        """Save preview pane width to settings."""
        settings = QSettings("iCharlotte", "iCharlotte")
        settings.setValue("preview_pane_width", width)

    def _load_preview_width(self):
        """Load preview pane width from settings."""
        settings = QSettings("iCharlotte", "iCharlotte")
        return settings.value("preview_pane_width", type=int)

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

                has_summary = any("summar" in log.get("task_type", "").lower() and log.get("status") == "success" for log in file_status)
                is_unprocessed = not file_status or not any(log.get("status") == "success" for log in file_status)

                status_match = False
                if "unprocessed" in status_filter and is_unprocessed:
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

    def _count_running_agents(self):
        """Count currently running subprocess agents (excludes queued ones)."""
        return sum(
            1 for r in self.agent_runners
            if getattr(r, 'success', None) is None and r not in self._agent_queue
        )

    def _check_system_memory_ok(self):
        """Return True if system memory is below critical threshold."""
        try:
            import psutil
            mem = psutil.virtual_memory()
            available_gb = mem.available / (1024 ** 3)
            if available_gb < 1.0:
                log_warning(f"System memory critically low: {available_gb:.1f}GB available")
                return False
            return True
        except ImportError:
            return True  # Can't check, allow it

    def _drain_agent_queue(self):
        """Start queued agents if slots are available."""
        while self._agent_queue and self._count_running_agents() < self.MAX_CONCURRENT_AGENTS:
            if not self._check_system_memory_ok():
                log_warning("Pausing agent queue drain — system memory low")
                break
            runner = self._agent_queue.pop(0)
            log_debug(f"Dequeuing agent: {runner.status_widget.task_id} — running={self._count_running_agents()}")
            runner.status_widget.update_progress(0, "Starting...")
            runner.start()

    def add_status_task(self, name, details, command, args, script_name=None, file_path=None):
        status_widget = StatusWidget(name, details)
        self.status_list_layout.insertWidget(0, status_widget) # Add to top

        runner = AgentRunner(command, args, status_widget, task_id=status_widget.task_id, file_number=self.file_number)
        self.agent_runners.append(runner) # Keep alive
        runner.finished.connect(lambda: self.cleanup_runner(runner))
        runner.finished.connect(lambda _=None: self._drain_agent_queue())

        # Wire the READY button for the deposition agent's interactive flow.
        # AgentRunner.awaiting_input fires when phase 1 emits AWAITING_INPUT:<path>,
        # which makes StatusWidget show the READY button. Clicking it emits ready_clicked.
        if script_name == "summarize_deposition.py":
            status_widget.ready_clicked.connect(
                lambda session_path, runner=runner, widget=status_widget:
                    self._open_depo_summary_dialog(session_path, runner, widget)
            )

        # Store retry info on widget for serialization
        status_widget._retry_command = command
        status_widget._retry_args = list(args)
        status_widget._retry_script_name = script_name
        status_widget._retry_file_path = file_path

        # Retry button — re-launches the same task
        status_widget.retry_requested.connect(
            lambda: self.add_status_task(name, details, command, list(args), script_name, file_path)
        )

        # Log processing entry when agent finishes
        if script_name and file_path and self.file_number:
            runner.finished.connect(lambda success: self._log_processing_entry(script_name, file_path, success))

        running = self._count_running_agents()
        log_debug(f"add_status_task: {name} — total_runners={len(self.agent_runners)}, running={running}, queued={len(self._agent_queue)}")

        # Check concurrency limit and system memory
        if running >= self.MAX_CONCURRENT_AGENTS:
            self._agent_queue.append(runner)
            status_widget.update_progress(0, f"Queued (#{len(self._agent_queue)}) — waiting for slot...")
            log_info(f"Agent '{name}' queued — {running} already running (max {self.MAX_CONCURRENT_AGENTS})")
            return runner

        if not self._check_system_memory_ok():
            self._agent_queue.append(runner)
            status_widget.update_progress(0, "Queued — system memory low")
            log_warning(f"Agent '{name}' queued due to low system memory")
            return runner

        runner.start()
        return runner

    def _open_depo_summary_dialog(self, session_path, agent_runner, status_widget):
        """Open the deposition summary config dialog. On Accept, resume phase 2."""
        from icharlotte_core.ui.depo_summary_config_dialog import DepoSummaryConfigDialog
        try:
            dlg = DepoSummaryConfigDialog(session_path, parent=self)
        except Exception as e:
            log_error(f"Failed to open deposition config dialog: {e}")
            return

        if dlg.exec() == QDialog.Accepted:
            status_widget.clear_ready_state()
            try:
                agent_runner.resume_with_config(session_path)
            except Exception as e:
                log_error(f"Failed to resume phase 2: {e}")
        # On Cancel, leave the READY button visible so the user can re-open the dialog.

    def _log_processing_entry(self, script_name, file_path, success):
        """Log a processing entry when an agent finishes.

        Instead of triggering a full tree rebuild (which destroys all QTreeWidgetItems
        and causes segfaults when other code holds stale references), we do a targeted
        update of just the processing status column for the affected file.
        """
        try:
            if not self.file_number:
                return
            proc_log = ProcessingLogDB(self.file_number)
            status = "success" if success else "failed"
            proc_log.add_entry(file_path, script_name, status)
            log_debug(f"_log_processing_entry: {script_name} {status} for {os.path.basename(file_path)} — updating status column (no tree rebuild)")

            # Refresh the cached proc log so status queries see the new entry
            if hasattr(self, '_cached_proc_log') and self._cached_proc_log is not None:
                self._cached_proc_log = ProcessingLogDB(self.file_number)

            # Targeted update: just update column 4 (processing status) for this file
            item = self.tree_item_map.get(file_path)
            if item is not None:
                try:
                    new_status = self._get_file_processing_status(file_path)
                    item.setText(4, new_status)
                    log_debug(f"_log_processing_entry: updated status for {os.path.basename(file_path)} → '{new_status}'")
                except RuntimeError:
                    log_warning(f"_log_processing_entry: stale tree item for {os.path.basename(file_path)}")
            else:
                log_debug(f"_log_processing_entry: file not in tree_item_map, skipping status update")
        except Exception as e:
            log_error(f"Error logging processing entry: {e}")
            log_event(f"Error logging processing entry: {e}", "error")

    def cleanup_runner(self, runner):
        try:
            # Only cleanup if it matches the current file number (meaning the user saw it finish)
            if getattr(runner, 'file_number', None) == self.file_number:
                if runner in self.agent_runners:
                    self.agent_runners.remove(runner)
        except Exception as e:
            log_event(f"Error in cleanup_runner: {e}", "error")

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
            if hasattr(runner, 'disconnect_widget'):
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

                # Restore retry info from serialized data, or reconstruct from agent name + filename
                retry_cmd = item_data.get("retry_command")
                retry_args = item_data.get("retry_args")
                retry_script = item_data.get("retry_script_name")
                retry_fp = item_data.get("retry_file_path")

                if not retry_cmd or not retry_args:
                    # Reconstruct from agent name and details (filename)
                    agent_name = item_data.get("agent_name", "")
                    filename = item_data.get("details", "")
                    agent = next((a for a in self.AGENTS if a["name"] == agent_name), None)
                    if agent and filename and self.case_path:
                        # Search for the file in the case folder
                        for root, dirs, files in os.walk(self.case_path):
                            if filename in files:
                                retry_fp = os.path.join(root, filename)
                                break
                        if retry_fp:
                            retry_cmd = sys.executable
                            retry_script = agent["script"]
                            retry_args = [os.path.join(SCRIPTS_DIR, retry_script), retry_fp]

                if retry_cmd and retry_args:
                    widget._retry_command = retry_cmd
                    widget._retry_args = retry_args
                    widget._retry_script_name = retry_script
                    widget._retry_file_path = retry_fp
                    name = item_data.get("agent_name", "Unknown")
                    details = item_data.get("details", "")
                    widget.retry_requested.connect(
                        lambda _n=name, _d=details, _c=retry_cmd, _a=retry_args, _s=retry_script, _f=retry_fp:
                            self.add_status_task(_n, _d, _c, list(_a), _s, _f)
                    )

                # Try to reconnect to running agent
                reconnected = False
                if not widget.is_finished:
                    task_id = getattr(widget, 'task_id', None)
                    if task_id:
                        for runner in self.agent_runners:
                            if getattr(runner, 'task_id', None) == task_id and hasattr(runner, 'reconnect_widget'):
                                runner.reconnect_widget(widget)
                                reconnected = True
                                if getattr(runner, 'success', None) is not None:
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
        # Save chat conversation before closing
        if hasattr(self, 'chat_tab') and self.chat_tab:
            self.chat_tab.save_current_state()
        # Save email update tab state before closing
        if hasattr(self, 'email_update_tab') and self.email_update_tab:
            self.email_update_tab.save_state()
        # Save discovery tab state before closing
        if hasattr(self, 'discovery_tab') and self.discovery_tab:
            self.discovery_tab.save_state()
        # Clean up global hotkeys
        if KEYBOARD_AVAILABLE:
            try:
                keyboard.unhook_all_hotkeys()
                stop_word_hotkey()
                log_event("Global hotkeys unregistered")
            except Exception as e:
                log_event(f"Error unregistering hotkeys: {e}", "error")
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
                agent = next((a for a in self.AGENTS if a["id"] == tid), None)
                if not agent: continue
                
                if tid == "separate":
                    self.run_separator_path(path)
                    count += 1
                else:
                    filename = os.path.basename(path)
                    display_name = agent["name"]
                    details = f"{filename}"
                    script_name = agent["script"]

                    args = [os.path.join(SCRIPTS_DIR, script_name), path]
                    self.add_status_task(display_name, details, sys.executable, args, script_name=script_name, file_path=path)
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
                        item.setText(2, entry.get('size_str', ''))

                        # Convert cached date to standard format
                        date_str = format_date_to_mm_dd_yyyy(entry.get('date_str', ''))
                        item.setText(3, date_str)
                    
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
        log_debug(f"save_cache: starting, map_size={len(self.tree_item_map)}, case={self.file_number}")
        data = []
        _deleted_count = 0
        for path, item in self.tree_item_map.items():
            if path == self.case_path: continue

            try:
                entry = {
                    'path': path,
                    'type': item.data(0, Qt.ItemDataRole.UserRole + 1),
                    'size_str': item.text(2),
                    'date_str': item.text(3),
                    'task_ids': item.data(0, Qt.ItemDataRole.UserRole + 2)
                }
                data.append(entry)
            except RuntimeError:
                # Item's C++ object was already deleted (e.g., parent removed)
                _deleted_count += 1
                continue

        if _deleted_count:
            log_warning(f"save_cache: {_deleted_count} deleted C++ items skipped")

        if not os.path.exists(GEMINI_DATA_DIR):
            os.makedirs(GEMINI_DATA_DIR)

        cache_path = os.path.join(GEMINI_DATA_DIR, f"{self.file_number}_tree.json")
        try:
            with open(cache_path, 'w') as f:
                json.dump(data, f)
            log_debug(f"save_cache: wrote {len(data)} items to {cache_path}")
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
                   '.txt', '.msg', '.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tiff',
                   '.mp3', '.mp4', '.avi', '.mov', '.zip', '.rar', '.7z']:
            # Use the icon provider with just the extension info
            icon = self.icon_provider.icon(QFileInfo(f"dummy{ext}"))
            self._icon_cache[ext] = icon
            return icon

        # Default file icon for unknown extensions
        self._icon_cache[ext] = self._file_icon
        return self._file_icon

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
                    if "discovery" in task or "disc" in task:
                        successful.add("S-DISC")
                    elif "deposition" in task or "depo" in task:
                        successful.add("S-DEPO")
                    elif "med_record" in task or "med_rec" in task:
                        successful.add("MED-REC")
                    elif "summar" in task:
                        successful.add("SUM")
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

    def _schedule_tree_refresh(self):
        """Debounced tree refresh — coalesces rapid calls into one."""
        self._tree_refresh_timer.start(500)

    def populate_tree(self):
        # Re-entrancy guard: if we're already inside populate_tree (e.g. via
        # processEvents re-entrancy), defer to a scheduled refresh instead.
        if self._populating_tree:
            log_debug("populate_tree: SKIPPED (re-entrant call) — scheduling deferred refresh")
            self._schedule_tree_refresh()
            return

        self._populating_tree = True
        try:
            self._populate_tree_inner()
        finally:
            self._populating_tree = False

    def _populate_tree_inner(self):
        # Increment generation so stale worker callbacks are ignored
        self._tree_generation += 1
        current_gen = self._tree_generation
        log_debug(f"populate_tree: gen={current_gen}, case={self.file_number}, path={self.case_path}")

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
            # CRITICAL: wait for the thread to actually finish before
            # replacing self.worker.  Without this, the old QThread can be
            # GC'd while its thread is still inside os.walk/emit, causing a
            # C++ segfault ("QThread: Destroyed while thread is still running").
            if self.worker.isRunning():
                if not self.worker.wait(5000):  # 5 s max; Z: walk is I/O-bound
                    log_warning("populate_tree: old worker did not stop within 5 s — detaching")
                    # Keep a reference so it isn't destroyed while running
                    self._old_worker = self.worker
                    self._old_worker.finished.connect(self._old_worker.deleteLater)

        self.worker = DirectoryTreeWorker(self.case_path)
        # Use lambdas that capture generation to ignore stale callbacks after file switch
        self.worker.data_ready.connect(lambda batch, gen=current_gen: self._on_tree_batch(gen, batch))
        self.worker.finished.connect(lambda gen=current_gen: self._on_scan_complete(gen))
        self.worker.start()
        log_debug(f"populate_tree: worker started for gen={current_gen}")

    def _on_tree_batch(self, generation, batch):
        """Wrapper that ignores stale worker callbacks from a previous file."""
        if generation != self._tree_generation:
            log_debug(f"_on_tree_batch: STALE gen={generation}, current={self._tree_generation} — ignoring")
            return
        self.add_tree_batch(batch)

    def _on_scan_complete(self, generation):
        """Wrapper that ignores stale worker callbacks from a previous file."""
        if generation != self._tree_generation:
            log_debug(f"_on_scan_complete: STALE gen={generation}, current={self._tree_generation} — ignoring")
            return
        self.on_scan_complete()

    def add_tree_batch(self, batch):
        _batch_gen = self._tree_generation
        _batch_dirs = 0
        _batch_files = 0
        try:
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
                        _batch_dirs += 1

                for f, size, mtime in files:
                    file_path = os.path.join(root, f)
                    self.visited_paths.add(file_path)

                    size_str = f"{size / 1024:.1f} KB" if size < 1024 * 1024 else f"{size / (1024 * 1024):.1f} MB"
                    date_str = format_date_to_mm_dd_yyyy(mtime)

                    # Get processing status (uses cached DB)
                    proc_status = self._get_file_processing_status(file_path) if self.file_number else ""

                    if file_path not in self.tree_item_map:
                        f_item = QTreeWidgetItem(parent_item)
                        f_item.setText(0, f)

                        # Use cached extension-based icons to avoid network access
                        f_item.setIcon(0, self._get_cached_icon(file_path))

                        f_item.setData(0, Qt.ItemDataRole.UserRole, file_path)
                        f_item.setData(0, Qt.ItemDataRole.UserRole + 1, "file")
                        f_item.setText(2, size_str)
                        f_item.setText(3, date_str)
                        f_item.setText(4, proc_status)

                        self.tree_item_map[file_path] = f_item
                        _batch_files += 1
                    else:
                        f_item = self.tree_item_map[file_path]
                        try:
                            f_item.setText(2, size_str)
                            f_item.setText(3, date_str)
                            f_item.setText(4, proc_status)
                        except RuntimeError:
                            log_warning(f"add_tree_batch: RuntimeError on existing item {file_path}")
                            continue

            log_debug(f"add_tree_batch: gen={_batch_gen}, +{_batch_dirs} dirs, +{_batch_files} files, map_size={len(self.tree_item_map)}")
        except Exception as e:
            log_error(f"add_tree_batch CRASHED: gen={_batch_gen}, +{_batch_dirs}d/+{_batch_files}f, map_size={len(self.tree_item_map)}, error={e}", exc_info=True)
            raise

    def on_scan_complete(self):
        log_debug(f"on_scan_complete: gen={self._tree_generation}, case={self.file_number}, map_size={len(self.tree_item_map)}, visited={len(self.visited_paths)}")
        try:
            # Prune items not visited (deleted files/folders)
            to_remove = []
            for path, item in self.tree_item_map.items():
                if path != self.case_path and path not in self.visited_paths:
                    to_remove.append(path)

            log_debug(f"on_scan_complete: pruning {len(to_remove)} stale items")
            for path in to_remove:
                item = self.tree_item_map.pop(path)
                try:
                    parent = item.parent()
                    if parent:
                        parent.removeChild(item)
                except RuntimeError:
                    log_warning(f"on_scan_complete: RuntimeError removing {path}")

            self.save_cache()
            self.tree.setEnabled(True)
            self.process_btn.setEnabled(True)
            self.status_label.setText(f"Scan Complete. Case: {self.file_number}")
            self.tree.setSortingEnabled(True)
            log_debug(f"on_scan_complete: DONE, final map_size={len(self.tree_item_map)}")
        except Exception as e:
            log_error(f"on_scan_complete CRASHED: gen={self._tree_generation}, case={self.file_number}, error={e}", exc_info=True)
            raise

# Note: Exception handling is now managed by app_crash_handler module
# The install_crash_handler() function sets up sys.excepthook automatically
# Legacy exception_hook removed - see icharlotte_core/app_crash_handler.py

if __name__ == "__main__":
    # Enable faulthandler to catch C-level segfaults that bypass Python exceptions
    import faulthandler
    _faulthandler_log = os.path.join(os.path.dirname(__file__), "logs", "crashes", "faulthandler.log")
    os.makedirs(os.path.dirname(_faulthandler_log), exist_ok=True)
    _faulthandler_file = open(_faulthandler_log, "a")
    _faulthandler_file.write(f"\n{'='*60}\nSession start: {__import__('datetime').datetime.now().isoformat()}\n{'='*60}\n")
    _faulthandler_file.flush()
    faulthandler.enable(file=_faulthandler_file, all_threads=True)

    # Install crash handler FIRST before anything else
    crash_handler = install_crash_handler()
    checkpoint("Application starting")

    try:
        # Parse command-line arguments for restart state restoration
        checkpoint("Parsing command-line arguments")
        parser = argparse.ArgumentParser(description='iCharlotte Legal Document Management Suite')
        parser.add_argument('--file-number', type=str, help='File number to load on startup')
        parser.add_argument('--case-path', type=str, help='Case path to load on startup')
        parser.add_argument('--tab', type=int, help='Tab index to open on startup')
        args, remaining = parser.parse_known_args()

        # Add startup args to crash context
        add_context('startup_file_number', args.file_number)
        add_context('startup_case_path', args.case_path)
        add_context('startup_tab', args.tab)

        # Debug: Log received arguments
        log_debug(f"sys.argv = {sys.argv}")
        log_debug(f"Parsed args = file_number={args.file_number}, case_path={args.case_path}, tab={args.tab}")

        # Disable Chromium sandbox and security for local file editing
        checkpoint("Configuring Qt WebEngine")
        os.environ["QTWEBENGINE_CHROMIUM_FLAGS"] = "--no-sandbox --disable-web-security"

        checkpoint("Creating QApplication")
        app = QApplication(sys.argv)

        # Set application icon (for taskbar)
        icon_path = os.path.join(os.path.dirname(__file__), 'icharlotte.ico')
        if os.path.exists(icon_path):
            app.setWindowIcon(QIcon(icon_path))

        # Update crash handler with Qt references
        crash_handler.set_qt_references(qt_app=app)

        # Log startup parameters
        log_info(f"Starting with args: file_number={args.file_number}, case_path={args.case_path}, tab={args.tab}")

        # Launch Main Window with restored state if provided
        checkpoint("Creating MainWindow")
        window = MainWindow(
            file_number=args.file_number,
            case_path=args.case_path,
            initial_tab=args.tab
        )

        # Update crash handler with window reference
        crash_handler.set_qt_references(main_window=window)

        checkpoint("Showing MainWindow")
        window.show()

        # Start periodic health monitor to capture state before silent crashes
        from PySide6.QtCore import QTimer as _QTimer
        _health_timer = _QTimer()
        def _health_check():
            try:
                import psutil
                proc = psutil.Process(os.getpid())
                mem = proc.memory_info()
                children = proc.children()
                threads = proc.num_threads()
                # Count running agents
                running_agents = 0
                agent_names = []
                if hasattr(window, 'agent_runners'):
                    for runner in window.agent_runners:
                        if getattr(runner, 'success', None) is None:  # Still running
                            running_agents += 1
                            name = os.path.basename(runner.args[0]) if hasattr(runner, 'args') and runner.args else "?"
                            agent_names.append(name)
                log_debug(
                    f"HEALTH: RSS={mem.rss // 1024 // 1024}MB, "
                    f"threads={threads}, child_procs={len(children)}, "
                    f"running_agents={running_agents}"
                    + (f" [{', '.join(agent_names)}]" if agent_names else "")
                )
                # Warn if memory is getting high (>1.5GB)
                if mem.rss > 1_500_000_000:
                    log_warning(f"HIGH MEMORY: {mem.rss // 1024 // 1024}MB RSS — crash risk")
                # Dump faulthandler traceback periodically when agents are running
                if running_agents > 0:
                    _faulthandler_file.write(f"\n--- Health check {__import__('datetime').datetime.now().isoformat()} agents={running_agents} ---\n")
                    faulthandler.dump_traceback(file=_faulthandler_file, all_threads=True)
                    _faulthandler_file.flush()
            except Exception:
                pass
        _health_timer.timeout.connect(_health_check)
        _health_timer.start(30000)  # Every 30 seconds

        checkpoint("Entering Qt event loop")
        sys.exit(app.exec())

    except Exception as e:
        log_error(f"CRITICAL MAIN ERROR: {e}", exc_info=True)
        print(f"CRITICAL MAIN ERROR: {e}")
        import traceback
        traceback.print_exc()
        # Keep window open if possible or wait for input so user can see error
        input("Press Enter to exit...")
