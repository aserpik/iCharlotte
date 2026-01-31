"""
Office Hotkey Integration - Global hotkey to trigger LLM processing on Word/Outlook selections.

Press Win+V to show a popup with prompt options when Word or Outlook compose window is active.
The selected/entered prompt is sent to the LLM along with any selected text, and the result
replaces the selection (or inserts at cursor if nothing selected).

Supports:
- Microsoft Word documents
- Microsoft Outlook email compose windows (new email, reply, forward)

The popup automatically detects which application is active and shows context-appropriate
prompts (document prompts for Word, email prompts for Outlook).
"""

import os
import json
import threading
from typing import Optional, Callable

from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QComboBox,
    QLineEdit, QPushButton, QMessageBox, QApplication,
    QTextEdit, QFrame, QSizePolicy, QCheckBox, QSpinBox,
    QGroupBox, QGridLayout, QFontComboBox, QDoubleSpinBox,
    QTabWidget, QWidget, QScrollArea
)
from PySide6.QtCore import Qt, Signal, QObject, QTimer, QThread
from PySide6.QtGui import QFont, QKeySequence, QShortcut
import re

# Import MasterCaseDatabase for case variable insertion
try:
    from .master_db import MasterCaseDatabase
    HAS_MASTER_DB = True
except ImportError:
    HAS_MASTER_DB = False

try:
    import keyboard
    HAS_KEYBOARD = True
except ImportError:
    HAS_KEYBOARD = False

try:
    import win32com.client
    import pythoncom
    import win32gui
    import win32process
    import subprocess
    HAS_WIN32 = True
except ImportError:
    HAS_WIN32 = False

# Application context types
APP_CONTEXT_UNKNOWN = "unknown"
APP_CONTEXT_WORD = "word"
APP_CONTEXT_OUTLOOK = "outlook"

# Window class names
WORD_WINDOW_CLASS = "OpusApp"
OUTLOOK_INSPECTOR_CLASS = "rctrl_renwnd32"

# Default email prompts
DEFAULT_OUTLOOK_PROMPTS = [
    {"name": "Make Professional", "prompt": "Rewrite this email in a professional, courteous tone while maintaining the original meaning and intent."},
    {"name": "Shorten Email", "prompt": "Condense this email to be more concise while preserving all key points and action items."},
    {"name": "Fix Grammar", "prompt": "Fix any grammar, spelling, or punctuation errors in this email."},
    {"name": "Soften Tone", "prompt": "Rewrite this email with a softer, more diplomatic tone while keeping the message clear."},
    {"name": "Add Clarity", "prompt": "Improve clarity and structure of this email. Ensure requests and next steps are clearly stated."},
]

# Email-specific system prompt
EMAIL_SYSTEM_PROMPT = (
    "You are a helpful email writing assistant. Follow the user's instructions precisely. "
    "Output ONLY the processed body text - never include Subject lines, To/From/CC headers, "
    "greetings, signatures, or any email metadata. Just output the revised text content directly. "
    "Do not add any preamble or explanation."
)


def detect_active_app_context() -> tuple:
    """
    Detect whether Word or Outlook compose window is active.
    Returns: (context_type, inspector_object_or_None)
    """
    if not HAS_WIN32:
        return APP_CONTEXT_UNKNOWN, None

    try:
        # Get the foreground window
        hwnd = win32gui.GetForegroundWindow()
        if not hwnd:
            return APP_CONTEXT_UNKNOWN, None

        class_name = win32gui.GetClassName(hwnd)

        # Check for Word
        if class_name == WORD_WINDOW_CLASS:
            return APP_CONTEXT_WORD, None

        # Check for Outlook Inspector (compose/reply window)
        if class_name == OUTLOOK_INSPECTOR_CLASS:
            # Verify it's a compose window via COM
            try:
                pythoncom.CoInitialize()
                outlook = win32com.client.GetActiveObject("Outlook.Application")
                inspector = outlook.ActiveInspector()
                if inspector:
                    # Check if the inspector has a WordEditor (compose mode)
                    try:
                        word_editor = inspector.WordEditor
                        if word_editor:
                            return APP_CONTEXT_OUTLOOK, inspector
                    except:
                        pass
            except:
                pass

        return APP_CONTEXT_UNKNOWN, None
    except Exception as e:
        try:
            print(f"Error detecting app context: {e}")
        except OSError:
            pass
        return APP_CONTEXT_UNKNOWN, None


def detect_case_from_document() -> tuple:
    """
    Detect case info from the currently open Word document's file path.
    Returns: (case_dict_or_None, file_number_or_None, error_message_or_None)

    Parses the document path for file number pattern (####.###) and looks up
    the case in MasterCaseDatabase.
    """
    if not HAS_WIN32:
        return None, None, "Win32 not available"

    if not HAS_MASTER_DB:
        return None, None, "Database not available"

    try:
        pythoncom.CoInitialize()

        # Try to connect to Word
        word = None
        try:
            word = win32com.client.GetActiveObject("Word.Application")
        except:
            try:
                word = win32com.client.GetObject(None, "Word.Application")
            except:
                pass

        if not word:
            return None, None, "Word not running"

        doc = word.ActiveDocument
        if not doc:
            return None, None, "No active document"

        # Get the document's full path
        try:
            doc_path = doc.FullName
        except:
            return None, None, "Document not saved"

        if not doc_path or doc_path.startswith("Document"):
            return None, None, "Document not saved"

        # Try multiple patterns to extract file number from path
        file_number = None

        # Pattern 1: Direct match for ####.### format (e.g., "1280.001 - Name")
        direct_pattern = r'(\d{4}\.\d{3})'
        direct_matches = re.findall(direct_pattern, doc_path)
        if direct_matches:
            file_number = direct_matches[0]

        # Pattern 2: Split folder structure (e.g., "3800- NATIONWIDE\167 - Nikitina")
        # Carrier folder: 4 digits followed by optional space and hyphen
        # Case folder: 1-3 digits followed by space and hyphen
        if not file_number:
            # Look for carrier folder pattern: ####- or #### -
            carrier_pattern = r'[\\/](\d{4})\s*-\s*[^\\/]+'
            # Look for case subfolder pattern: ### - (1-3 digits)
            case_pattern = r'[\\/](\d{1,3})\s*-\s*[^\\/]+'

            carrier_matches = re.findall(carrier_pattern, doc_path)
            case_matches = re.findall(case_pattern, doc_path)

            if carrier_matches and case_matches:
                # Use the last carrier and case match (most specific to the document)
                carrier = carrier_matches[-1]
                case_num = case_matches[-1].zfill(3)  # Pad to 3 digits
                file_number = f"{carrier}.{case_num}"

        if not file_number:
            # Show truncated path for debugging
            short_path = doc_path if len(doc_path) <= 60 else "..." + doc_path[-57:]
            return None, None, f"No case number in: {short_path}"

        # Look up in database
        try:
            db = MasterCaseDatabase()
            case = db.get_case(file_number)
            if case:
                return case, file_number, None
            else:
                return None, file_number, f"Case {file_number} not in database"
        except Exception as e:
            return None, file_number, f"Database error: {str(e)}"

    except Exception as e:
        return None, None, f"Error: {str(e)}"


def kill_zombie_word_processes():
    """Kill Word processes that don't have visible windows (zombies)."""
    if not HAS_WIN32:
        return

    try:
        # Find PIDs of visible Word windows
        visible_pids = set()

        def enum_callback(hwnd, results):
            if win32gui.IsWindowVisible(hwnd):
                class_name = win32gui.GetClassName(hwnd)
                if class_name == "OpusApp":
                    _, pid = win32process.GetWindowThreadProcessId(hwnd)
                    results.add(pid)
            return True

        win32gui.EnumWindows(enum_callback, visible_pids)

        if not visible_pids:
            # No visible Word windows, don't kill anything
            return

        # Find all Word processes
        result = subprocess.run(
            ['tasklist', '/FI', 'IMAGENAME eq WINWORD.EXE', '/FO', 'CSV'],
            capture_output=True, text=True, timeout=5
        )

        all_pids = set()
        lines = result.stdout.strip().split('\n')
        for line in lines[1:]:  # Skip header
            parts = line.replace('"', '').split(',')
            if len(parts) >= 2:
                try:
                    all_pids.add(int(parts[1]))
                except:
                    pass

        # Kill zombie processes (those without visible windows)
        zombie_pids = all_pids - visible_pids
        for pid in zombie_pids:
            try:
                subprocess.run(
                    ['taskkill', '/PID', str(pid), '/F'],
                    capture_output=True, timeout=5
                )
                print(f"Killed zombie Word process: PID {pid}")
            except Exception as e:
                print(f"Failed to kill PID {pid}: {e}")

    except Exception as e:
        print(f"Error killing zombie processes: {e}")

from icharlotte_core.config import GEMINI_DATA_DIR


# Format options
FORMAT_PLAIN = "Plain Text"
FORMAT_MATCH = "Match Selection"
FORMAT_MARKDOWN = "Parse Markdown"
FORMAT_DEFAULT = "Document Default"
FORMAT_CUSTOM = "Custom..."

# Available LLM models for selection
# Format: (display_name, provider, model_id)
AVAILABLE_MODELS = [
    ("Gemini 2.5 Flash (Fast)", "Gemini", "gemini-2.5-flash"),
    ("Gemini 2.5 Pro", "Gemini", "gemini-2.5-pro"),
    ("Gemini 2.0 Flash", "Gemini", "models/gemini-2.0-flash"),
    ("Gemini 3 Flash Preview", "Gemini", "gemini-3-flash-preview"),
    ("Gemini 3 Pro Preview", "Gemini", "gemini-3-pro-preview"),
    ("Claude Sonnet 4", "Claude", "claude-sonnet-4-20250514"),
    ("Claude Haiku 4 (Fast)", "Claude", "claude-haiku-4-20250514"),
    ("GPT-4o", "OpenAI", "gpt-4o"),
    ("GPT-4o Mini (Fast)", "OpenAI", "gpt-4o-mini"),
]

DEFAULT_MODEL_INDEX = 0  # Gemini 2.5 Flash


class CustomFormatDialog(QDialog):
    """Dialog for specifying custom text formatting."""

    def __init__(self, parent=None, current_settings=None):
        super().__init__(parent)
        self.setWindowTitle("Custom Format Settings")
        self.setMinimumWidth(350)
        self.settings = current_settings or {}

        self.setStyleSheet("""
            QDialog {
                background-color: #1e1e2e;
            }
            QLabel {
                color: #cdd6f4;
                font-size: 12px;
            }
            QGroupBox {
                color: #cdd6f4;
                border: 1px solid #45475a;
                border-radius: 5px;
                margin-top: 10px;
                padding-top: 10px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px;
            }
            QComboBox, QSpinBox, QDoubleSpinBox, QFontComboBox {
                background-color: #313244;
                color: #cdd6f4;
                border: 1px solid #45475a;
                border-radius: 5px;
                padding: 5px;
            }
            QCheckBox {
                color: #cdd6f4;
            }
            QCheckBox::indicator {
                width: 16px;
                height: 16px;
            }
            QPushButton {
                background-color: #6c63ff;
                color: white;
                border: none;
                border-radius: 5px;
                padding: 8px 16px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #7c73ff;
            }
            QPushButton#cancelBtn {
                background-color: #45475a;
            }
        """)

        self.setup_ui()
        self.load_settings()

    def setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(15, 15, 15, 15)
        layout.setSpacing(10)

        # Font Group
        font_group = QGroupBox("Font")
        font_layout = QGridLayout(font_group)

        font_layout.addWidget(QLabel("Font:"), 0, 0)
        self.font_combo = QFontComboBox()
        self.font_combo.setCurrentFont(QFont("Times New Roman"))
        font_layout.addWidget(self.font_combo, 0, 1)

        font_layout.addWidget(QLabel("Size:"), 1, 0)
        self.size_spin = QSpinBox()
        self.size_spin.setRange(6, 72)
        self.size_spin.setValue(12)
        font_layout.addWidget(self.size_spin, 1, 1)

        layout.addWidget(font_group)

        # Style Group
        style_group = QGroupBox("Style")
        style_layout = QHBoxLayout(style_group)

        self.bold_check = QCheckBox("Bold")
        self.italic_check = QCheckBox("Italic")
        self.underline_check = QCheckBox("Underline")

        style_layout.addWidget(self.bold_check)
        style_layout.addWidget(self.italic_check)
        style_layout.addWidget(self.underline_check)
        style_layout.addStretch()

        layout.addWidget(style_group)

        # Paragraph Group
        para_group = QGroupBox("Paragraph")
        para_layout = QGridLayout(para_group)

        para_layout.addWidget(QLabel("First Line Indent:"), 0, 0)
        self.first_indent_spin = QDoubleSpinBox()
        self.first_indent_spin.setRange(0, 2)
        self.first_indent_spin.setSingleStep(0.25)
        self.first_indent_spin.setValue(0)
        self.first_indent_spin.setSuffix(" in")
        para_layout.addWidget(self.first_indent_spin, 0, 1)

        para_layout.addWidget(QLabel("Left Indent:"), 1, 0)
        self.left_indent_spin = QDoubleSpinBox()
        self.left_indent_spin.setRange(0, 4)
        self.left_indent_spin.setSingleStep(0.25)
        self.left_indent_spin.setValue(0)
        self.left_indent_spin.setSuffix(" in")
        para_layout.addWidget(self.left_indent_spin, 1, 1)

        para_layout.addWidget(QLabel("Line Spacing:"), 2, 0)
        self.line_spacing_combo = QComboBox()
        self.line_spacing_combo.addItems(["Single", "1.5 Lines", "Double"])
        para_layout.addWidget(self.line_spacing_combo, 2, 1)

        layout.addWidget(para_group)

        # Buttons
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()

        cancel_btn = QPushButton("Cancel")
        cancel_btn.setObjectName("cancelBtn")
        cancel_btn.clicked.connect(self.reject)
        btn_layout.addWidget(cancel_btn)

        ok_btn = QPushButton("OK")
        ok_btn.clicked.connect(self.accept)
        btn_layout.addWidget(ok_btn)

        layout.addLayout(btn_layout)

    def load_settings(self):
        """Load settings from dictionary."""
        if not self.settings:
            return

        if 'font_name' in self.settings:
            self.font_combo.setCurrentFont(QFont(self.settings['font_name']))
        if 'font_size' in self.settings:
            self.size_spin.setValue(self.settings['font_size'])
        if 'bold' in self.settings:
            self.bold_check.setChecked(self.settings['bold'])
        if 'italic' in self.settings:
            self.italic_check.setChecked(self.settings['italic'])
        if 'underline' in self.settings:
            self.underline_check.setChecked(self.settings['underline'])
        if 'first_indent' in self.settings:
            self.first_indent_spin.setValue(self.settings['first_indent'])
        if 'left_indent' in self.settings:
            self.left_indent_spin.setValue(self.settings['left_indent'])
        if 'line_spacing' in self.settings:
            idx = self.line_spacing_combo.findText(self.settings['line_spacing'])
            if idx >= 0:
                self.line_spacing_combo.setCurrentIndex(idx)

    def get_settings(self) -> dict:
        """Get the format settings as a dictionary."""
        return {
            'font_name': self.font_combo.currentFont().family(),
            'font_size': self.size_spin.value(),
            'bold': self.bold_check.isChecked(),
            'italic': self.italic_check.isChecked(),
            'underline': self.underline_check.isChecked(),
            'first_indent': self.first_indent_spin.value(),
            'left_indent': self.left_indent_spin.value(),
            'line_spacing': self.line_spacing_combo.currentText()
        }


class HotkeySignals(QObject):
    """Signals for cross-thread communication."""
    show_popup = Signal()


class LLMWorkerThread(QThread):
    """Worker thread for running LLM calls without blocking the UI."""
    finished = Signal(str, str)  # (result, error_message)

    def __init__(self, llm_func, prompt: str, parent=None):
        super().__init__(parent)
        self.llm_func = llm_func
        self.prompt = prompt
        self._cancelled = False

    def run(self):
        """Execute the LLM call in a separate thread."""
        if self._cancelled:
            self.finished.emit("", "cancelled")
            return

        try:
            result = self.llm_func(self.prompt)
            if self._cancelled:
                self.finished.emit("", "cancelled")
            else:
                self.finished.emit(result or "", "")
        except Exception as e:
            if not self._cancelled:
                self.finished.emit("", str(e))

    def cancel(self):
        """Mark the thread as cancelled."""
        self._cancelled = True


class WordLLMPopup(QDialog):
    """Popup dialog for LLM processing of Word content."""

    def __init__(self, parent=None, llm_callback: Optional[Callable] = None, cursor_pos=None):
        super().__init__(parent)
        self.llm_callback = llm_callback
        self.cursor_pos = cursor_pos  # Position to show popup at (QPoint or None)
        self.prompts_path = os.path.join(GEMINI_DATA_DIR, "word_llm_prompts.json")
        self.outlook_prompts_path = os.path.join(GEMINI_DATA_DIR, "outlook_llm_prompts.json")
        self.format_settings_path = os.path.join(GEMINI_DATA_DIR, "word_format_settings.json")
        self.model_settings_path = os.path.join(GEMINI_DATA_DIR, "word_model_settings.json")
        self.prompts = []  # Word prompts
        self.outlook_prompts = []  # Outlook/email prompts
        self.custom_format_settings = {}  # Custom format settings
        self.selected_model_index = DEFAULT_MODEL_INDEX  # Selected model index
        self._word_app = None  # Stored Word COM reference during execution
        self._captured_format = None  # Captured format from selection

        # App context for Word vs Outlook
        self.app_context = APP_CONTEXT_WORD  # Default context
        self.active_inspector = None  # Outlook Inspector reference when in email mode

        # Worker thread for LLM calls
        self._worker_thread = None
        self._pending_format_type = None  # Format type for pending task
        self._is_outlook_task = False  # Whether pending task is for Outlook

        # Case detection for variable insertion
        self._detected_case = None  # Case dict from database
        self._detected_file_number = None  # File number extracted from path
        self._case_detection_error = None  # Error message if detection failed

        # Drag/move support for independent positioning
        self._drag_pos = None  # Starting position for drag
        self._is_dragging = False

        self.setWindowTitle("AI Assistant")
        self.setWindowFlags(
            Qt.WindowType.Dialog |
            Qt.WindowType.FramelessWindowHint |
            Qt.WindowType.WindowStaysOnTopHint
        )
        self.setMinimumWidth(450)
        self.setStyleSheet("""
            QDialog {
                background-color: #1e1e2e;
                border: 2px solid #6c63ff;
                border-radius: 10px;
            }
            QLabel {
                color: #cdd6f4;
                font-size: 12px;
            }
            QComboBox {
                background-color: #313244;
                color: #cdd6f4;
                border: 1px solid #45475a;
                border-radius: 5px;
                padding: 8px;
                font-size: 12px;
            }
            QComboBox:hover {
                border-color: #6c63ff;
            }
            QComboBox::drop-down {
                border: none;
                width: 30px;
            }
            QComboBox QAbstractItemView {
                background-color: #313244;
                color: #cdd6f4;
                selection-background-color: #6c63ff;
            }
            QLineEdit, QTextEdit {
                background-color: #313244;
                color: #cdd6f4;
                border: 1px solid #45475a;
                border-radius: 5px;
                padding: 8px;
                font-size: 12px;
            }
            QLineEdit:focus, QTextEdit:focus {
                border-color: #6c63ff;
            }
            QPushButton {
                background-color: #6c63ff;
                color: white;
                border: none;
                border-radius: 5px;
                padding: 8px 16px;
                font-size: 12px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #7c73ff;
            }
            QPushButton:pressed {
                background-color: #5c53ef;
            }
            QPushButton#cancelBtn {
                background-color: #45475a;
            }
            QPushButton#cancelBtn:hover {
                background-color: #585b70;
            }
            QPushButton#deleteBtn {
                background-color: #f38ba8;
            }
            QPushButton#deleteBtn:hover {
                background-color: #f5a0b8;
            }
        """)

        self.setup_ui()
        self.load_prompts()
        self._load_format_settings()
        self._load_model_settings()
        self._update_format_preview()

        # Close on Escape
        self.shortcut_escape = QShortcut(QKeySequence(Qt.Key.Key_Escape), self)
        self.shortcut_escape.activated.connect(self.close)

        # Execute on Enter (when in custom prompt field)
        self.shortcut_enter = QShortcut(QKeySequence(Qt.Key.Key_Return), self)
        self.shortcut_enter.activated.connect(self.execute)

    def setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(15, 15, 15, 15)
        layout.setSpacing(10)

        # Header row with drag handle and title
        header_widget = QWidget()
        header_widget.setCursor(Qt.CursorShape.SizeAllCursor)
        header_layout = QHBoxLayout(header_widget)
        header_layout.setContentsMargins(0, 0, 0, 0)
        header_layout.setSpacing(8)

        # Drag handle indicator (grip dots)
        drag_handle = QLabel("\u2630")  # Unicode hamburger menu / grip icon
        drag_handle.setStyleSheet("font-size: 14px; color: #6c7086;")
        drag_handle.setToolTip("Drag to move window")
        header_layout.addWidget(drag_handle)

        # Title (dynamic based on context)
        self.title_label = QLabel("AI Assistant for Word")
        self.title_label.setStyleSheet("font-size: 16px; font-weight: bold; color: #6c63ff;")
        self.title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        header_layout.addWidget(self.title_label, 1)  # Stretch to fill

        # Spacer to balance the drag handle
        spacer_label = QLabel("")
        spacer_label.setFixedWidth(20)
        header_layout.addWidget(spacer_label)

        self.header_widget = header_widget
        layout.addWidget(header_widget)

        # Case info header (shows detected case)
        self.case_info_label = QLabel("")
        self.case_info_label.setStyleSheet("color: #a6e3a1; font-size: 11px; padding: 4px;")
        self.case_info_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.case_info_label.setWordWrap(True)
        layout.addWidget(self.case_info_label)

        # Tab widget for AI Prompt vs Case Variables
        self.tab_widget = QTabWidget()
        self.tab_widget.setStyleSheet("""
            QTabWidget::pane {
                border: 1px solid #45475a;
                border-radius: 5px;
                background-color: #1e1e2e;
            }
            QTabBar::tab {
                background-color: #313244;
                color: #cdd6f4;
                padding: 8px 16px;
                border: 1px solid #45475a;
                border-bottom: none;
                border-top-left-radius: 5px;
                border-top-right-radius: 5px;
                margin-right: 2px;
            }
            QTabBar::tab:selected {
                background-color: #6c63ff;
                color: white;
            }
            QTabBar::tab:hover:!selected {
                background-color: #45475a;
            }
        """)
        layout.addWidget(self.tab_widget)

        # Create AI Prompt tab
        self._setup_ai_prompt_tab()

        # Create Case Variables tab
        self._setup_case_variables_tab()

        # Status label (shared across tabs)
        self.status_label = QLabel("")
        self.status_label.setStyleSheet("color: #a6adc8; font-style: italic;")
        layout.addWidget(self.status_label)

        # Buttons (shared across tabs)
        btn_row = QHBoxLayout()
        btn_row.addStretch()

        self.cancel_btn = QPushButton("Cancel")
        self.cancel_btn.setObjectName("cancelBtn")
        self.cancel_btn.clicked.connect(self._on_cancel_clicked)
        btn_row.addWidget(self.cancel_btn)

        self.execute_btn = QPushButton("Execute (Enter)")
        self.execute_btn.clicked.connect(self.execute)
        btn_row.addWidget(self.execute_btn)

        layout.addLayout(btn_row)

    def _setup_ai_prompt_tab(self):
        """Setup the AI Prompt tab with existing prompt/format/model controls."""
        ai_tab = QWidget()
        ai_layout = QVBoxLayout(ai_tab)
        ai_layout.setContentsMargins(10, 10, 10, 10)
        ai_layout.setSpacing(8)

        # Saved Prompts Section
        prompt_label = QLabel("Saved Prompts:")
        ai_layout.addWidget(prompt_label)

        prompt_row = QHBoxLayout()
        self.prompt_combo = QComboBox()
        self.prompt_combo.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self.prompt_combo.currentIndexChanged.connect(self.on_prompt_selected)
        prompt_row.addWidget(self.prompt_combo)

        delete_btn = QPushButton("Delete")
        delete_btn.setObjectName("deleteBtn")
        delete_btn.setFixedWidth(70)
        delete_btn.clicked.connect(self.delete_prompt)
        prompt_row.addWidget(delete_btn)

        ai_layout.addLayout(prompt_row)

        # Separator
        line = QFrame()
        line.setFrameShape(QFrame.Shape.HLine)
        line.setStyleSheet("background-color: #45475a;")
        ai_layout.addWidget(line)

        # Custom Prompt Section
        custom_label = QLabel("Custom Prompt (or type to add new):")
        ai_layout.addWidget(custom_label)

        self.custom_input = QTextEdit()
        self.custom_input.setPlaceholderText("Enter your prompt here...\nExample: 'Convert this to formal legal language' or 'Summarize in bullet points'")
        self.custom_input.setMaximumHeight(80)
        ai_layout.addWidget(self.custom_input)

        # Save prompt checkbox row
        save_row = QHBoxLayout()
        self.save_name_input = QLineEdit()
        self.save_name_input.setPlaceholderText("Name for saving (optional)")
        self.save_name_input.setFixedWidth(200)
        save_row.addWidget(self.save_name_input)

        save_btn = QPushButton("Save Prompt")
        save_btn.setFixedWidth(100)
        save_btn.clicked.connect(self.save_prompt)
        save_row.addWidget(save_btn)
        save_row.addStretch()

        ai_layout.addLayout(save_row)

        # Separator
        line2 = QFrame()
        line2.setFrameShape(QFrame.Shape.HLine)
        line2.setStyleSheet("background-color: #45475a;")
        ai_layout.addWidget(line2)

        # Format Section
        format_label = QLabel("Output Format:")
        ai_layout.addWidget(format_label)

        format_row = QHBoxLayout()
        self.format_combo = QComboBox()
        self.format_combo.addItems([
            FORMAT_PLAIN,
            FORMAT_MATCH,
            FORMAT_MARKDOWN,
            FORMAT_DEFAULT,
            FORMAT_CUSTOM
        ])
        self.format_combo.currentTextChanged.connect(self._on_format_changed)
        self.format_combo.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        format_row.addWidget(self.format_combo)

        self.format_settings_btn = QPushButton("Settings")
        self.format_settings_btn.setFixedWidth(70)
        self.format_settings_btn.clicked.connect(self._open_format_settings)
        self.format_settings_btn.setEnabled(False)  # Only enabled for Custom
        format_row.addWidget(self.format_settings_btn)

        ai_layout.addLayout(format_row)

        # Format preview label
        self.format_preview = QLabel("")
        self.format_preview.setStyleSheet("color: #a6adc8; font-size: 10px; font-style: italic;")
        self.format_preview.setWordWrap(True)
        ai_layout.addWidget(self.format_preview)

        # Model Selection Section
        model_label = QLabel("AI Model:")
        ai_layout.addWidget(model_label)

        self.model_combo = QComboBox()
        for display_name, provider, model_id in AVAILABLE_MODELS:
            self.model_combo.addItem(f"{display_name}", (provider, model_id))
        self.model_combo.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self.model_combo.currentIndexChanged.connect(self._on_model_changed)
        ai_layout.addWidget(self.model_combo)

        # Use All Text checkbox
        self.use_all_text_check = QCheckBox("Use all document text as context (still replaces only selection)")
        self.use_all_text_check.setStyleSheet("color: #cdd6f4; font-size: 11px;")
        self.use_all_text_check.setToolTip(
            "When checked, sends the entire document to the LLM for context,\n"
            "but only replaces the selected text with the output."
        )
        ai_layout.addWidget(self.use_all_text_check)

        ai_layout.addStretch()
        self.tab_widget.addTab(ai_tab, "AI Prompt")

    def _setup_case_variables_tab(self):
        """Setup the Case Variables tab for inserting case info into documents."""
        case_tab = QWidget()
        case_layout = QVBoxLayout(case_tab)
        case_layout.setContentsMargins(10, 10, 10, 10)
        case_layout.setSpacing(8)

        # Scroll area for variables list
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setStyleSheet("""
            QScrollArea {
                border: none;
                background-color: transparent;
            }
            QScrollBar:vertical {
                background-color: #313244;
                width: 10px;
                border-radius: 5px;
            }
            QScrollBar::handle:vertical {
                background-color: #45475a;
                border-radius: 5px;
                min-height: 20px;
            }
            QScrollBar::handle:vertical:hover {
                background-color: #6c63ff;
            }
        """)

        # Container for variable rows
        self.variables_container = QWidget()
        self.variables_layout = QVBoxLayout(self.variables_container)
        self.variables_layout.setContentsMargins(0, 0, 0, 0)
        self.variables_layout.setSpacing(6)

        # Placeholder - will be populated when case is detected
        self.no_case_label = QLabel("No case detected.\nOpen a document from a case folder to see variables.")
        self.no_case_label.setStyleSheet("color: #a6adc8; font-style: italic; padding: 20px;")
        self.no_case_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.no_case_label.setWordWrap(True)
        self.variables_layout.addWidget(self.no_case_label)

        self.variables_layout.addStretch()
        scroll.setWidget(self.variables_container)
        case_layout.addWidget(scroll)

        self.tab_widget.addTab(case_tab, "Case Variables")

    def _populate_case_variables(self):
        """Populate the Case Variables tab with detected case info."""
        # Clear existing variable widgets (except stretch)
        while self.variables_layout.count() > 0:
            item = self.variables_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        if not self._detected_case:
            # Show no case message
            self.no_case_label = QLabel(
                self._case_detection_error or "No case detected.\nOpen a document from a case folder to see variables."
            )
            self.no_case_label.setStyleSheet("color: #a6adc8; font-style: italic; padding: 20px;")
            self.no_case_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
            self.no_case_label.setWordWrap(True)
            self.variables_layout.addWidget(self.no_case_label)
            self.variables_layout.addStretch()
            return

        # Define the case variables to display
        case_variables = [
            ("File Number", self._detected_case.get("file_number", "")),
            ("Plaintiff Name", self._detected_case.get("plaintiff_last_name", "")),
            ("Next Hearing", self._detected_case.get("next_hearing_date", "")),
            ("Trial Date", self._detected_case.get("trial_date", "")),
            ("Assigned Attorney", self._detected_case.get("assigned_attorney", "")),
            ("Case Summary", self._detected_case.get("case_summary", "")),
            ("Last Report", self._detected_case.get("last_report_text", "")),
        ]

        for label_text, value in case_variables:
            if not value:
                continue  # Skip empty values

            row_widget = QWidget()
            row_layout = QHBoxLayout(row_widget)
            row_layout.setContentsMargins(0, 0, 0, 0)
            row_layout.setSpacing(8)

            # Label
            label = QLabel(f"{label_text}:")
            label.setFixedWidth(120)
            label.setStyleSheet("color: #89b4fa; font-weight: bold;")
            row_layout.addWidget(label)

            # Value preview (truncated for long text)
            display_value = value if len(value) <= 50 else value[:47] + "..."
            value_label = QLabel(display_value)
            value_label.setStyleSheet("color: #ffffff; font-size: 12px;")
            value_label.setToolTip(value if len(value) > 50 else "")
            value_label.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
            row_layout.addWidget(value_label)

            # Insert button
            insert_btn = QPushButton("Insert")
            insert_btn.setFixedWidth(60)
            insert_btn.setStyleSheet("""
                QPushButton {
                    background-color: #45475a;
                    padding: 4px 8px;
                    font-size: 11px;
                }
                QPushButton:hover {
                    background-color: #6c63ff;
                }
            """)
            # Store the value in the button for the click handler
            insert_btn.setProperty("insert_value", value)
            insert_btn.clicked.connect(lambda checked, v=value: self._insert_case_variable(v))
            row_layout.addWidget(insert_btn)

            self.variables_layout.addWidget(row_widget)

        self.variables_layout.addStretch()

    def _insert_case_variable(self, value: str):
        """Insert a case variable value into Word at the current cursor position."""
        if not value:
            return

        try:
            pythoncom.CoInitialize()

            # Get Word application
            word = None
            try:
                word = win32com.client.GetActiveObject("Word.Application")
            except:
                try:
                    word = win32com.client.GetObject(None, "Word.Application")
                except:
                    pass

            if not word:
                self.status_label.setText("Error: Word not running")
                self.status_label.setStyleSheet("color: #f38ba8; font-style: italic;")
                return

            selection = word.Selection
            if not selection:
                self.status_label.setText("Error: No selection available")
                self.status_label.setStyleSheet("color: #f38ba8; font-style: italic;")
                return

            # Insert the text at cursor
            selection.TypeText(value)

            self.status_label.setText("Inserted!")
            self.status_label.setStyleSheet("color: #a6e3a1; font-style: italic;")

            # Auto-close after successful insert
            QTimer.singleShot(500, self.close)

        except Exception as e:
            self.status_label.setText(f"Error: {str(e)[:50]}")
            self.status_label.setStyleSheet("color: #f38ba8; font-style: italic;")

    def _detect_and_update_case(self):
        """Detect case from document and update UI."""
        if self.app_context != APP_CONTEXT_WORD:
            # Case detection only applies to Word documents
            self._detected_case = None
            self._detected_file_number = None
            self._case_detection_error = "Case variables only available in Word"
            self.case_info_label.setText("")
            return

        self._detected_case, self._detected_file_number, self._case_detection_error = detect_case_from_document()

        if self._detected_case:
            plaintiff = self._detected_case.get("plaintiff_last_name", "Unknown")
            file_num = self._detected_file_number or ""
            self.case_info_label.setText(f"Case: {plaintiff} ({file_num})")
            self.case_info_label.setStyleSheet("color: #a6e3a1; font-size: 11px; padding: 4px;")
        elif self._detected_file_number:
            self.case_info_label.setText(f"Case {self._detected_file_number} not in database")
            self.case_info_label.setStyleSheet("color: #fab387; font-size: 11px; padding: 4px;")
        else:
            self.case_info_label.setText("")

        # Populate the variables tab
        self._populate_case_variables()

    def load_prompts(self):
        """Load saved prompts from file (both Word and Outlook prompts)."""
        # Load Word prompts
        self.prompts = []
        if os.path.exists(self.prompts_path):
            try:
                with open(self.prompts_path, 'r', encoding='utf-8') as f:
                    self.prompts = json.load(f)
            except:
                pass

        # Add default Word prompts if none exist
        if not self.prompts:
            self.prompts = [
                {"name": "Improve Writing", "prompt": "Improve this text for clarity and professionalism while maintaining the original meaning."},
                {"name": "Convert to Narrative", "prompt": "Convert this text into a cohesive narrative paragraph format. Maintain all facts and details."},
                {"name": "Formal Legal Tone", "prompt": "Rewrite this in formal legal language appropriate for court documents."},
                {"name": "Summarize", "prompt": "Summarize this text concisely while keeping key points."},
                {"name": "Fix Grammar", "prompt": "Fix any grammar, spelling, or punctuation errors in this text."},
            ]
            self.save_prompts_to_file()

        # Load Outlook prompts
        self.outlook_prompts = []
        if os.path.exists(self.outlook_prompts_path):
            try:
                with open(self.outlook_prompts_path, 'r', encoding='utf-8') as f:
                    self.outlook_prompts = json.load(f)
            except:
                pass

        # Add default Outlook prompts if none exist
        if not self.outlook_prompts:
            self.outlook_prompts = list(DEFAULT_OUTLOOK_PROMPTS)
            self._save_outlook_prompts()

        self.refresh_combo()

    def refresh_combo(self):
        """Refresh the combo box with prompts for current context."""
        self.prompt_combo.clear()
        self.prompt_combo.addItem("-- Select a saved prompt --", None)

        # Get prompts for current context
        if self.app_context == APP_CONTEXT_OUTLOOK:
            prompts = self.outlook_prompts
        else:
            prompts = self.prompts

        for p in prompts:
            self.prompt_combo.addItem(p["name"], p["prompt"])

    def save_prompts_to_file(self):
        """Save Word prompts to file."""
        try:
            os.makedirs(os.path.dirname(self.prompts_path), exist_ok=True)
            with open(self.prompts_path, 'w', encoding='utf-8') as f:
                json.dump(self.prompts, f, indent=2, ensure_ascii=False)
        except Exception as e:
            print(f"Error saving prompts: {e}")

    def _save_outlook_prompts(self):
        """Save Outlook prompts to file."""
        try:
            os.makedirs(os.path.dirname(self.outlook_prompts_path), exist_ok=True)
            with open(self.outlook_prompts_path, 'w', encoding='utf-8') as f:
                json.dump(self.outlook_prompts, f, indent=2, ensure_ascii=False)
        except Exception as e:
            print(f"Error saving Outlook prompts: {e}")

    def on_prompt_selected(self, index):
        """When a saved prompt is selected, populate the custom input."""
        prompt_text = self.prompt_combo.currentData()
        if prompt_text:
            self.custom_input.setPlainText(prompt_text)

    def save_prompt(self):
        """Save the current custom prompt to the appropriate list based on context."""
        name = self.save_name_input.text().strip()
        prompt = self.custom_input.toPlainText().strip()

        if not name:
            QMessageBox.warning(self, "Name Required", "Please enter a name for the prompt.")
            return
        if not prompt:
            QMessageBox.warning(self, "Prompt Required", "Please enter a prompt to save.")
            return

        # Get the appropriate prompts list and save function based on context
        if self.app_context == APP_CONTEXT_OUTLOOK:
            prompts_list = self.outlook_prompts
            save_func = self._save_outlook_prompts
            context_name = "email"
        else:
            prompts_list = self.prompts
            save_func = self.save_prompts_to_file
            context_name = "document"

        # Check for duplicate name
        for p in prompts_list:
            if p["name"].lower() == name.lower():
                reply = QMessageBox.question(
                    self, "Overwrite?",
                    f"A prompt named '{name}' already exists. Overwrite?",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
                )
                if reply == QMessageBox.StandardButton.Yes:
                    p["prompt"] = prompt
                    save_func()
                    self.refresh_combo()
                    self.status_label.setText(f"Updated '{name}' ({context_name})")
                return

        prompts_list.append({"name": name, "prompt": prompt})
        save_func()
        self.refresh_combo()
        self.save_name_input.clear()
        self.status_label.setText(f"Saved '{name}' ({context_name})")

    def delete_prompt(self):
        """Delete the selected prompt from the appropriate list based on context."""
        index = self.prompt_combo.currentIndex()
        if index <= 0:
            QMessageBox.warning(self, "Select Prompt", "Please select a prompt to delete.")
            return

        name = self.prompt_combo.currentText()
        reply = QMessageBox.question(
            self, "Delete Prompt?",
            f"Delete '{name}'?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        if reply == QMessageBox.StandardButton.Yes:
            # Delete from appropriate list based on context
            if self.app_context == APP_CONTEXT_OUTLOOK:
                self.outlook_prompts = [p for p in self.outlook_prompts if p["name"] != name]
                self._save_outlook_prompts()
            else:
                self.prompts = [p for p in self.prompts if p["name"] != name]
                self.save_prompts_to_file()
            self.refresh_combo()
            self.custom_input.clear()
            self.status_label.setText(f"Deleted '{name}'")

    def set_app_context(self, context: str, inspector=None):
        """Set the application context and update UI accordingly."""
        self.app_context = context
        self.active_inspector = inspector

        if context == APP_CONTEXT_OUTLOOK:
            self.title_label.setText("AI Assistant for Outlook")
            self.setWindowTitle("AI Assistant - Email")
            # Hide case variables tab for Outlook
            self.tab_widget.setTabVisible(1, False)
        else:
            self.title_label.setText("AI Assistant for Word")
            self.setWindowTitle("AI Assistant - Document")
            # Show case variables tab for Word
            self.tab_widget.setTabVisible(1, True)
            # Detect case from document path
            self._detect_and_update_case()

        self.refresh_combo()  # Refresh prompts for context

    def _on_format_changed(self, format_type: str):
        """Handle format dropdown change."""
        self.format_settings_btn.setEnabled(format_type == FORMAT_CUSTOM)

        # Pre-capture format when Match Selection is chosen so user can see preview
        if format_type == FORMAT_MATCH:
            self._pre_capture_format()

        self._update_format_preview()

    def _pre_capture_format(self):
        """Pre-capture format from current selection to show in preview."""
        try:
            if self.app_context == APP_CONTEXT_OUTLOOK:
                self._capture_outlook_selection_format()
            else:
                word = self._get_word_app()
                if word:
                    self._capture_selection_format(word)
        except Exception as e:
            print(f"Pre-capture format error: {e}")
            self._captured_format = None

    def _update_format_preview(self):
        """Update the format preview label with enhanced formatting info."""
        format_type = self.format_combo.currentText()

        if format_type == FORMAT_PLAIN:
            self.format_preview.setText("No formatting will be applied")
        elif format_type == FORMAT_MATCH:
            # Show captured format details if available
            if self._captured_format:
                parts = []

                # Show paragraph style if captured
                style_name = self._captured_format.get('style_name')
                if style_name:
                    parts.append(f"Style: {style_name}")

                # Show font info
                font_name = self._captured_format.get('font_name')
                font_size = self._captured_format.get('font_size')
                if font_name and font_size:
                    parts.append(f"{font_name} {font_size}pt")

                # Show table context
                if self._captured_format.get('in_table'):
                    parts.append("[Table cell]")

                # Show track changes
                if self._captured_format.get('track_changes_enabled'):
                    parts.append("[Track Changes]")

                if parts:
                    self.format_preview.setText(" | ".join(parts))
                else:
                    self.format_preview.setText("Formatting will match the selected text")
            else:
                self.format_preview.setText("Formatting will match the selected text (captured on execute)")
        elif format_type == FORMAT_MARKDOWN:
            self.format_preview.setText("Markdown syntax: **bold**, *italic*, _underline_, `code`")
        elif format_type == FORMAT_DEFAULT:
            self.format_preview.setText("Uses Word's default paragraph style")
        elif format_type == FORMAT_CUSTOM:
            if self.custom_format_settings:
                s = self.custom_format_settings
                preview = f"{s.get('font_name', 'Default')}, {s.get('font_size', 12)}pt"
                styles = []
                if s.get('bold'):
                    styles.append('Bold')
                if s.get('italic'):
                    styles.append('Italic')
                if s.get('underline'):
                    styles.append('Underline')
                if styles:
                    preview += f", {'/'.join(styles)}"
                self.format_preview.setText(preview)
            else:
                self.format_preview.setText("Click Settings to configure custom format")

    def _open_format_settings(self):
        """Open the custom format settings dialog."""
        dialog = CustomFormatDialog(self, self.custom_format_settings)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.custom_format_settings = dialog.get_settings()
            self._save_format_settings()
            self._update_format_preview()

    def _load_format_settings(self):
        """Load custom format settings from file."""
        if os.path.exists(self.format_settings_path):
            try:
                with open(self.format_settings_path, 'r', encoding='utf-8') as f:
                    self.custom_format_settings = json.load(f)
            except:
                pass

    def _save_format_settings(self):
        """Save custom format settings to file."""
        try:
            os.makedirs(os.path.dirname(self.format_settings_path), exist_ok=True)
            with open(self.format_settings_path, 'w', encoding='utf-8') as f:
                json.dump(self.custom_format_settings, f, indent=2)
        except Exception as e:
            print(f"Error saving format settings: {e}")

    def _load_model_settings(self):
        """Load model selection from file."""
        if os.path.exists(self.model_settings_path):
            try:
                with open(self.model_settings_path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    self.selected_model_index = data.get('model_index', DEFAULT_MODEL_INDEX)
                    # Validate index is within range
                    if self.selected_model_index >= len(AVAILABLE_MODELS):
                        self.selected_model_index = DEFAULT_MODEL_INDEX
            except:
                pass

        # Set the combo box to the saved selection
        if hasattr(self, 'model_combo'):
            self.model_combo.setCurrentIndex(self.selected_model_index)

    def _save_model_settings(self):
        """Save model selection to file."""
        try:
            os.makedirs(os.path.dirname(self.model_settings_path), exist_ok=True)
            with open(self.model_settings_path, 'w', encoding='utf-8') as f:
                json.dump({'model_index': self.selected_model_index}, f, indent=2)
        except Exception as e:
            print(f"Error saving model settings: {e}")

    def _on_model_changed(self, index: int):
        """Handle model dropdown change."""
        self.selected_model_index = index
        self._save_model_settings()
        # Get model info for display
        if index < len(AVAILABLE_MODELS):
            display_name, provider, model_id = AVAILABLE_MODELS[index]
            print(f"Model changed to: {display_name} ({provider})")

    def _get_selected_model(self) -> tuple:
        """Get the currently selected model (provider, model_id)."""
        if self.selected_model_index < len(AVAILABLE_MODELS):
            _, provider, model_id = AVAILABLE_MODELS[self.selected_model_index]
            return provider, model_id
        # Fallback to default
        _, provider, model_id = AVAILABLE_MODELS[DEFAULT_MODEL_INDEX]
        return provider, model_id

    def _capture_selection_format(self, word):
        """Capture the formatting of the current selection in Word.

        Enhanced to capture:
        - Paragraph styles (Heading 1, Normal, Body Text, etc.)
        - Table cell context and formatting
        - Track Changes state
        - Inline font formatting
        """
        try:
            selection = word.Selection
            font = selection.Font
            para = selection.ParagraphFormat

            # Word constants for underline
            # wdUnderlineNone = 0, wdUnderlineSingle = 1, etc.
            underline_val = font.Underline
            has_underline = underline_val != 0 and underline_val != 9999999  # 9999999 = mixed

            self._captured_format = {
                'font_name': font.Name if font.Name != '' else None,
                'font_size': font.Size if font.Size != 9999999 else 12,
                'bold': font.Bold == -1,  # -1 = True in Word
                'italic': font.Italic == -1,
                'underline': has_underline,
                'first_indent': para.FirstLineIndent / 72 if para.FirstLineIndent != 9999999 else 0,
                'left_indent': para.LeftIndent / 72 if para.LeftIndent != 9999999 else 0,
            }

            # Capture paragraph style name
            try:
                style = selection.Style
                if style:
                    style_name = style.NameLocal  # Localized name (e.g., "Heading 1")
                    self._captured_format['style_name'] = style_name
                    print(f"Captured paragraph style: {style_name}")
            except Exception as e:
                print(f"Could not capture style: {e}")
                self._captured_format['style_name'] = None

            # Detect if selection is inside a table cell
            try:
                # wdWithInTable = 12
                if selection.Information(12):  # wdWithInTable
                    self._captured_format['in_table'] = True

                    # Capture table cell info
                    try:
                        cell = selection.Cells(1) if selection.Cells.Count > 0 else None
                        if cell:
                            # Capture cell shading (background color)
                            shading = cell.Shading
                            if shading:
                                self._captured_format['cell_shading_color'] = shading.BackgroundPatternColor
                                self._captured_format['cell_shading_pattern'] = shading.Texture

                            # Capture cell vertical alignment
                            self._captured_format['cell_vertical_align'] = cell.VerticalAlignment

                            # Capture cell width if available
                            try:
                                self._captured_format['cell_width'] = cell.Width
                            except:
                                pass

                            print(f"Captured table cell formatting")
                    except Exception as e:
                        print(f"Could not capture cell details: {e}")
                else:
                    self._captured_format['in_table'] = False
            except Exception as e:
                print(f"Could not detect table context: {e}")
                self._captured_format['in_table'] = False

            # Capture Track Changes state
            try:
                doc = word.ActiveDocument
                if doc:
                    # TrackRevisions property indicates if Track Changes is enabled
                    self._captured_format['track_changes_enabled'] = doc.TrackRevisions
                    if doc.TrackRevisions:
                        print("Track Changes is enabled - will preserve revision tracking")
            except Exception as e:
                print(f"Could not check Track Changes: {e}")
                self._captured_format['track_changes_enabled'] = False

            # Capture line spacing
            try:
                # wdLineSpaceSingle = 0, wdLineSpace1pt5 = 1, wdLineSpaceDouble = 2
                # wdLineSpaceAtLeast = 3, wdLineSpaceExactly = 4, wdLineSpaceMultiple = 5
                line_spacing_rule = para.LineSpacingRule
                self._captured_format['line_spacing_rule'] = line_spacing_rule
                if line_spacing_rule in [3, 4, 5]:  # wdLineSpaceAtLeast, Exactly, Multiple
                    self._captured_format['line_spacing'] = para.LineSpacing
            except:
                pass

            # Capture font color
            try:
                font_color = font.Color
                if font_color != 9999999:  # Not mixed
                    self._captured_format['font_color'] = font_color
            except:
                pass

            print(f"Captured format: {self._captured_format}")
        except Exception as e:
            print(f"Error capturing format: {e}")
            self._captured_format = None

    def _on_cancel_clicked(self):
        """Handle cancel button click - either close dialog or cancel running task."""
        if self._worker_thread and self._worker_thread.isRunning():
            # Cancel the running task
            self._worker_thread.cancel()
            self.status_label.setText("Cancelling...")
            self.cancel_btn.setEnabled(False)
            # Wait briefly for thread to finish, then close
            QTimer.singleShot(200, self._cleanup_and_close)
        else:
            # No task running, just close
            self.close()

    def _cleanup_and_close(self):
        """Clean up worker thread and close dialog."""
        if self._worker_thread:
            if self._worker_thread.isRunning():
                self._worker_thread.terminate()
                self._worker_thread.wait(500)
            self._worker_thread = None
        self.close()

    def execute(self):
        """Execute the LLM processing."""
        prompt = self.custom_input.toPlainText().strip()
        if not prompt:
            QMessageBox.warning(self, "Prompt Required", "Please enter or select a prompt.")
            return

        self.status_label.setText("Processing...")
        self.execute_btn.setEnabled(False)
        self.cancel_btn.setText("Cancel Task")
        QApplication.processEvents()

        # Run in a slight delay to allow UI to update
        QTimer.singleShot(100, lambda: self._do_execute(prompt))

    def _do_execute(self, prompt: str):
        """Actually execute the LLM call - dispatch to Word or Outlook handler."""
        try:
            # Dispatch based on context
            if self.app_context == APP_CONTEXT_OUTLOOK:
                self._do_execute_outlook(prompt)
                return

            # Default: Word context
            # Check if Word is running with a document open
            word = self._get_word_app()

            if not word:
                self.status_label.setText("Word is not running")
                self.execute_btn.setEnabled(True)
                QMessageBox.warning(self, "Word Not Found",
                    "Microsoft Word is not running or no document is open.\n\nPlease open a Word document first, then try again.")
                return

            # Capture format if "Match Selection" is chosen
            format_type = self.format_combo.currentText()
            if format_type == FORMAT_MATCH:
                self._capture_selection_format(word)

            # Get Word selection
            selected_text, has_selection = self._get_word_selection_internal()

            # Check if we should use all document text as context
            use_all_text = self.use_all_text_check.isChecked()

            if use_all_text and has_selection and selected_text:
                # Get all document text for context
                all_text = self._get_all_document_text()
                if all_text:
                    # Build prompt with full document context but mark selected text
                    full_prompt = (
                        f"{prompt}\n\n"
                        f"=== FULL DOCUMENT (for context) ===\n{all_text}\n\n"
                        f"=== SELECTED TEXT TO PROCESS ===\n{selected_text}\n\n"
                        f"Process only the SELECTED TEXT above according to the instructions. "
                        f"Use the full document for context but output only the processed version of the selected text."
                    )
                    self.status_label.setText("Processing with full document context...")
                else:
                    # Fallback to just selected text
                    full_prompt = f"{prompt}\n\nText to process:\n{selected_text}"
                    self.status_label.setText("Processing selected text...")
            elif has_selection and selected_text:
                # Build full prompt with selected text only
                full_prompt = f"{prompt}\n\nText to process:\n{selected_text}"
                self.status_label.setText("Processing selected text...")
            elif use_all_text:
                # No selection but use all text - process entire document
                all_text = self._get_all_document_text()
                if all_text:
                    full_prompt = f"{prompt}\n\nText to process:\n{all_text}"
                    self.status_label.setText("Processing entire document...")
                else:
                    full_prompt = prompt
                    self.status_label.setText("Processing prompt...")
            else:
                # No selection, just use the prompt
                full_prompt = prompt
                self.status_label.setText("Processing prompt...")

            QApplication.processEvents()

            # Store format type for callback
            self._pending_format_type = format_type
            self._is_outlook_task = False

            # Start worker thread for LLM call
            llm_func = self.llm_callback if self.llm_callback else self._default_llm_call
            self._worker_thread = LLMWorkerThread(llm_func, full_prompt, self)
            self._worker_thread.finished.connect(self._on_llm_result)
            self._worker_thread.start()

        except Exception as e:
            self.status_label.setText(f"Error: {str(e)[:50]}")
            self.execute_btn.setEnabled(True)
            self.cancel_btn.setText("Cancel")
            QMessageBox.critical(self, "Error", f"Failed to process:\n{e}")

    def _on_llm_result(self, result: str, error: str):
        """Handle the result from the LLM worker thread."""
        self._worker_thread = None

        # Check if cancelled
        if error == "cancelled":
            self.status_label.setText("Cancelled")
            self.execute_btn.setEnabled(True)
            self.cancel_btn.setText("Cancel")
            self.cancel_btn.setEnabled(True)
            return

        # Check for error
        if error:
            self.status_label.setText(f"Error: {error[:50]}")
            self.execute_btn.setEnabled(True)
            self.cancel_btn.setText("Cancel")
            self.cancel_btn.setEnabled(True)
            QMessageBox.critical(self, "Error", f"Failed to process:\n{error}")
            return

        # Check for empty result
        if not result:
            self.status_label.setText("No response from LLM")
            self.execute_btn.setEnabled(True)
            self.cancel_btn.setText("Cancel")
            self.cancel_btn.setEnabled(True)
            return

        # Insert result into the appropriate application
        try:
            if self._is_outlook_task:
                self._set_outlook_text(result, self._pending_format_type)
            else:
                self._set_word_text_internal(result, self._pending_format_type)

            self.status_label.setText("Done!")
            QTimer.singleShot(500, self.close)
        except Exception as e:
            self.status_label.setText(f"Error inserting: {str(e)[:50]}")
            self.execute_btn.setEnabled(True)
            self.cancel_btn.setText("Cancel")
            self.cancel_btn.setEnabled(True)
            QMessageBox.critical(self, "Error", f"Failed to insert result:\n{e}")

    def _get_word_app(self):
        """Get a connection to the running Word application with active document."""
        word = None

        # Kill any zombie Word processes first
        kill_zombie_word_processes()

        # Try to connect to Word
        try:
            word = win32com.client.GetActiveObject("Word.Application")
        except:
            try:
                word = win32com.client.GetObject(Class="Word.Application")
            except:
                try:
                    word = win32com.client.Dispatch("Word.Application")
                except:
                    print("Could not connect to Word at all")
                    return None

        if not word:
            print("No Word connection")
            return None

        # Check for regular documents first
        try:
            if word.Documents.Count > 0:
                print(f"Connected to Word - {word.Documents.Count} docs open")
                return word
        except:
            pass

        # Check for Protected View windows
        try:
            pv_count = word.ProtectedViewWindows.Count
            if pv_count > 0:
                print(f"Found {pv_count} Protected View window(s)")
                # Try to activate/edit the first Protected View window
                # This will convert it to a regular document
                try:
                    pv_window = word.ProtectedViewWindows(1)
                    print(f"Protected View document: {pv_window.Document.Name}")
                    # Note: We can still work with Selection even in Protected View
                    return word
                except Exception as e:
                    print(f"Could not access Protected View window: {e}")
        except Exception as e:
            print(f"Error checking Protected View: {e}")

        # Check if Word has any windows at all via Windows API
        try:
            word_windows = []
            def enum_callback(hwnd, results):
                if win32gui.IsWindowVisible(hwnd):
                    class_name = win32gui.GetClassName(hwnd)
                    if class_name == "OpusApp":
                        title = win32gui.GetWindowText(hwnd)
                        results.append((hwnd, title))
                return True
            win32gui.EnumWindows(enum_callback, word_windows)

            if word_windows:
                print(f"Found {len(word_windows)} Word window(s) via Windows API")
                for hwnd, title in word_windows:
                    print(f"   - {title}")
                # Word has windows, so let's try to use Selection anyway
                # Sometimes Documents.Count is 0 but Selection still works
                return word
        except Exception as e:
            print(f"Window enumeration failed: {e}")

        print("Word running but no accessible documents")
        return None

    def _get_word_selection_internal(self) -> tuple:
        """Get selected text from Word. Returns (text, has_selection)."""
        try:
            word = self._get_word_app()

            if not word:
                print("Could not connect to Word")
                return "", False

            # Try to get selection - this might work even if Documents.Count is 0
            selection = None
            try:
                selection = word.Selection
            except Exception as e:
                print(f"Could not get Selection object: {e}")

            if not selection:
                print("Selection object is None")
                return "", False

            text = ""
            try:
                raw_text = selection.Text
                text = raw_text.strip() if raw_text else ""
                if text:
                    print(f"Got selection text: '{text[:50]}...' ({len(text)} chars)")
                else:
                    print("No text selected (cursor at insertion point)")
            except Exception as e:
                print(f"Error getting selection text: {e}")
                return "", False

            # wdSelectionIP = 1 (insertion point, no selection)
            has_selection = False
            try:
                sel_type = selection.Type
                has_selection = sel_type != 1 and len(text) > 0
                print(f"Selection type: {sel_type}, has_selection: {has_selection}")
            except:
                has_selection = len(text) > 0

            return text, has_selection
        except Exception as e:
            print(f"Could not get Word selection: {e}")
            return "", False

    def _get_all_document_text(self) -> str:
        """Get all text from the active Word document."""
        try:
            word = self._get_word_app()
            if not word:
                return ""

            # Try to get the active document's content
            try:
                doc = word.ActiveDocument
                if doc:
                    # Get all text from the document
                    content = doc.Content.Text
                    return content.strip() if content else ""
            except Exception as e:
                print(f"Error getting document content: {e}")

            return ""
        except Exception as e:
            print(f"Could not get all document text: {e}")
            return ""

    def _set_word_text_internal(self, text: str, format_type: str = FORMAT_PLAIN):
        """Set text in Word with formatting - re-establish connection to ensure it's fresh.

        Enhanced to:
        - Support Track Changes for all format modes
        - Handle table cell context properly
        - Preserve paragraph styles
        """
        try:
            word = self._get_word_app()

            if not word:
                raise Exception("Could not connect to Word")

            selection = word.Selection
            if not selection:
                raise Exception("Could not access Word selection")

            doc = word.ActiveDocument

            # Check if Track Changes is enabled and preserve it for all format modes
            track_changes_active = False
            try:
                track_changes_active = doc.TrackRevisions if doc else False
                if track_changes_active:
                    print("Track Changes is active - changes will be tracked")
            except:
                pass

            if format_type == FORMAT_PLAIN:
                # Insert text, but still handle bullet points properly
                self._insert_with_bullets(word, selection, text)

            elif format_type == FORMAT_MATCH:
                # Apply captured format (includes Track Changes handling)
                self._insert_with_format(selection, text, self._captured_format)

            elif format_type == FORMAT_MARKDOWN:
                # Parse markdown and apply formatting
                self._insert_with_markdown(word, selection, text)

            elif format_type == FORMAT_DEFAULT:
                # Reset to Normal style, then insert
                selection.Style = doc.Styles("Normal")
                selection.TypeText(text)

            elif format_type == FORMAT_CUSTOM:
                # Apply custom format settings with Track Changes awareness
                custom_settings = dict(self.custom_format_settings) if self.custom_format_settings else {}
                custom_settings['track_changes_enabled'] = track_changes_active
                self._insert_with_format(selection, text, custom_settings)

            else:
                # Fallback to plain text
                selection.TypeText(text)

            print(f"Successfully inserted {len(text)} characters into Word with format: {format_type}")

        except Exception as e:
            print(f"_set_word_text_internal error: {e}")
            raise

    def _insert_with_format(self, selection, text: str, format_settings: dict):
        """Insert text with specific formatting.

        Enhanced to apply:
        - Paragraph styles (Heading 1, Normal, Body Text, etc.)
        - Table cell formatting when inside tables
        - Track Changes support (applies changes as tracked edits)
        - Inline font formatting
        """
        if not format_settings:
            selection.TypeText(text)
            return

        try:
            word = selection.Application
            doc = word.ActiveDocument

            # Handle Track Changes - ensure revisions are tracked if originally enabled
            track_changes_was_enabled = format_settings.get('track_changes_enabled', False)
            original_track_state = None
            if track_changes_was_enabled:
                try:
                    original_track_state = doc.TrackRevisions
                    # Ensure Track Changes is on so our edits are tracked
                    if not doc.TrackRevisions:
                        doc.TrackRevisions = True
                        print("Temporarily enabled Track Changes for this edit")
                except Exception as e:
                    print(f"Could not set Track Changes: {e}")

            # Apply paragraph style first (if captured)
            style_name = format_settings.get('style_name')
            if style_name:
                try:
                    # Try to apply the style by name
                    selection.Style = doc.Styles(style_name)
                    print(f"Applied paragraph style: {style_name}")
                except Exception as e:
                    print(f"Could not apply style '{style_name}': {e}")
                    # Style might not exist - continue with other formatting

            font = selection.Font
            para = selection.ParagraphFormat

            # Apply font settings
            if format_settings.get('font_name'):
                font.Name = format_settings['font_name']
            if format_settings.get('font_size'):
                font.Size = format_settings['font_size']

            # Apply font color if captured
            if 'font_color' in format_settings:
                try:
                    font.Color = format_settings['font_color']
                except:
                    pass

            # Apply style (True = -1 in Word COM, False = 0)
            font.Bold = -1 if format_settings.get('bold') else 0
            font.Italic = -1 if format_settings.get('italic') else 0

            # Underline: wdUnderlineSingle = 1, wdUnderlineNone = 0
            font.Underline = 1 if format_settings.get('underline') else 0

            # Apply paragraph settings (convert inches to points: 1 inch = 72 points)
            if 'first_indent' in format_settings:
                para.FirstLineIndent = format_settings['first_indent'] * 72
            if 'left_indent' in format_settings:
                para.LeftIndent = format_settings['left_indent'] * 72

            # Line spacing - handle both simple string values and captured rule values
            line_spacing = format_settings.get('line_spacing', None)
            line_spacing_rule = format_settings.get('line_spacing_rule', None)

            if line_spacing_rule is not None:
                try:
                    para.LineSpacingRule = line_spacing_rule
                    # If it's a specific spacing value, set that too
                    if line_spacing_rule in [3, 4, 5] and 'line_spacing' in format_settings:
                        para.LineSpacing = format_settings['line_spacing']
                except:
                    pass
            elif line_spacing:
                if line_spacing == 'Single':
                    para.LineSpacingRule = 0  # wdLineSpaceSingle
                elif line_spacing == '1.5 Lines':
                    para.LineSpacingRule = 1  # wdLineSpace1pt5
                elif line_spacing == 'Double':
                    para.LineSpacingRule = 2  # wdLineSpaceDouble

            # Insert the text
            selection.TypeText(text)

            # Handle table cell formatting after text insertion
            if format_settings.get('in_table'):
                try:
                    # Re-check if we're still in a table after insertion
                    if selection.Information(12):  # wdWithInTable
                        cell = selection.Cells(1) if selection.Cells.Count > 0 else None
                        if cell:
                            # Apply cell shading if it was captured
                            if 'cell_shading_color' in format_settings:
                                cell.Shading.BackgroundPatternColor = format_settings['cell_shading_color']
                            if 'cell_shading_pattern' in format_settings:
                                cell.Shading.Texture = format_settings['cell_shading_pattern']

                            # Apply vertical alignment
                            if 'cell_vertical_align' in format_settings:
                                cell.VerticalAlignment = format_settings['cell_vertical_align']

                            print("Applied table cell formatting")
                except Exception as e:
                    print(f"Could not apply table cell formatting: {e}")

            # Restore original Track Changes state if we modified it
            if original_track_state is not None and not original_track_state:
                try:
                    doc.TrackRevisions = False
                    print("Restored Track Changes to original state (off)")
                except:
                    pass

        except Exception as e:
            print(f"Error applying format: {e}")
            # Fallback to plain text
            selection.TypeText(text)

    def _is_in_table(self, selection) -> bool:
        """Check if the current selection is inside a table cell."""
        try:
            return selection.Information(12)  # wdWithInTable = 12
        except:
            return False

    def _insert_with_bullets(self, word, selection, text: str):
        """Insert text with proper Word bullet formatting (no inline markdown).

        Enhanced to handle table cell context - bullets work differently in tables.
        """
        # Check if we're in a table - bullet handling may differ
        in_table = self._is_in_table(selection)
        if in_table:
            print("Inserting bullets within table cell")

        # First pass: group lines into items (bullet items include their continuations)
        items = self._parse_bullet_items(text)

        in_bullet_list = False
        for i, item in enumerate(items):
            if item['is_bullet']:
                if not in_bullet_list:
                    # First bullet - apply formatting and set up list template
                    in_bullet_list = True

                    # Apply Word's bullet formatting (enables auto-continuation on Enter)
                    try:
                        selection.Range.ListFormat.ApplyBulletDefault()

                        # Modify the ListTemplate level settings for correct indentation
                        # This is what actually controls bullet/text positioning in Word lists
                        # In tables, we use smaller indents to fit within cells
                        list_fmt = selection.Range.ListFormat
                        if list_fmt.ListTemplate and list_fmt.ListLevelNumber > 0:
                            level = list_fmt.ListTemplate.ListLevels(list_fmt.ListLevelNumber)
                            if in_table:
                                # Smaller indents for table cells
                                level.NumberPosition = 18  # 0.25 inch
                                level.TextPosition = 36    # 0.5 inch
                                level.TabPosition = 36     # 0.5 inch
                            else:
                                level.NumberPosition = 36  # 0.5 inch - where bullet sits
                                level.TextPosition = 54    # 0.75 inch - where text starts
                                level.TabPosition = 54     # 0.75 inch - tab stop
                    except Exception as e:
                        print(f"Could not modify list template: {e}")

                # For subsequent bullets, Word's list continuation handles formatting
                selection.TypeText(item['text'])

                if i < len(items) - 1:
                    selection.TypeParagraph()
            else:
                # Non-bullet item
                if in_bullet_list:
                    in_bullet_list = False
                    try:
                        selection.Range.ListFormat.RemoveNumbers()
                        para = selection.ParagraphFormat
                        para.LeftIndent = 0
                        para.FirstLineIndent = 0
                    except Exception as e:
                        print(f"Could not remove bullet formatting: {e}")

                if item['text']:
                    selection.TypeText(item['text'])

                if i < len(items) - 1:
                    selection.TypeParagraph()

    def _parse_bullet_items(self, text: str) -> list:
        """Parse text into items, combining bullet lines with their continuations."""
        lines = text.split('\n')
        bullet_pattern = re.compile(r'^[\s]*[-*•]\s+(.*)$')
        numbered_pattern = re.compile(r'^[\s]*\d+[.)]\s+(.*)$')

        items = []
        current_bullet = None

        for line in lines:
            bullet_match = bullet_pattern.match(line)
            numbered_match = numbered_pattern.match(line)

            if bullet_match or numbered_match:
                # Save previous bullet if exists
                if current_bullet is not None:
                    items.append({'is_bullet': True, 'text': current_bullet})

                # Start new bullet
                current_bullet = bullet_match.group(1) if bullet_match else numbered_match.group(1)

            elif current_bullet is not None and line.strip():
                # Continuation of current bullet - append with space
                current_bullet += " " + line.strip()

            else:
                # Non-bullet line or empty line
                if current_bullet is not None:
                    items.append({'is_bullet': True, 'text': current_bullet})
                    current_bullet = None

                # Add non-bullet item (even if empty, to preserve paragraph breaks)
                if line.strip() or (items and items[-1]['is_bullet']):
                    items.append({'is_bullet': False, 'text': line.strip()})

        # Don't forget the last bullet
        if current_bullet is not None:
            items.append({'is_bullet': True, 'text': current_bullet})

        return items

    def _insert_with_markdown(self, word, selection, text: str):
        """Parse markdown and insert with Word formatting, including proper bullet lists.

        Enhanced to handle table cell context with appropriate indentation.
        """
        # Check if we're in a table
        in_table = self._is_in_table(selection)
        if in_table:
            print("Inserting markdown within table cell")

        # First pass: group lines into items (bullet items include their continuations)
        items = self._parse_bullet_items(text)

        in_bullet_list = False
        for i, item in enumerate(items):
            if item['is_bullet']:
                if not in_bullet_list:
                    # First bullet - apply formatting and set up list template
                    in_bullet_list = True

                    # Apply Word's bullet formatting (enables auto-continuation on Enter)
                    try:
                        selection.Range.ListFormat.ApplyBulletDefault()

                        # Modify the ListTemplate level settings for correct indentation
                        list_fmt = selection.Range.ListFormat
                        if list_fmt.ListTemplate and list_fmt.ListLevelNumber > 0:
                            level = list_fmt.ListTemplate.ListLevels(list_fmt.ListLevelNumber)
                            if in_table:
                                # Smaller indents for table cells
                                level.NumberPosition = 18  # 0.25 inch
                                level.TextPosition = 36    # 0.5 inch
                                level.TabPosition = 36     # 0.5 inch
                            else:
                                level.NumberPosition = 36  # 0.5 inch - where bullet sits
                                level.TextPosition = 54    # 0.75 inch - where text starts
                                level.TabPosition = 54     # 0.75 inch - tab stop
                    except Exception as e:
                        print(f"Could not modify list template: {e}")

                # For subsequent bullets, Word's list continuation handles formatting
                # Insert the text (parse for inline markdown formatting)
                self._insert_formatted_text(selection, item['text'])

                if i < len(items) - 1:
                    selection.TypeParagraph()
            else:
                # Non-bullet item
                if in_bullet_list:
                    in_bullet_list = False
                    try:
                        selection.Range.ListFormat.RemoveNumbers()
                        para = selection.ParagraphFormat
                        para.LeftIndent = 0
                        para.FirstLineIndent = 0
                    except Exception as e:
                        print(f"Could not remove bullet formatting: {e}")

                if item['text']:
                    self._insert_formatted_text(selection, item['text'])

                if i < len(items) - 1:
                    selection.TypeParagraph()

        # End bullet list if we finished in one
        if in_bullet_list:
            # Move to end and remove list format for next content
            pass  # List ends naturally

    def _insert_formatted_text(self, selection, text: str):
        """Insert text with inline markdown formatting (bold, italic, etc.)."""
        segments = self._parse_markdown_segments(text)

        for segment in segments:
            seg_text = segment['text']
            if not seg_text:
                continue

            # Set formatting based on segment type
            font = selection.Font

            # Reset to defaults first
            font.Bold = 0
            font.Italic = 0
            font.Underline = 0

            if segment.get('bold'):
                font.Bold = -1
            if segment.get('italic'):
                font.Italic = -1
            if segment.get('underline'):
                font.Underline = 1
            if segment.get('code'):
                # Use monospace font for code
                font.Name = "Consolas"

            selection.TypeText(seg_text)

            # Reset font name if we changed it for code
            if segment.get('code'):
                font.Name = "Times New Roman"

    def _parse_markdown_segments(self, text: str) -> list:
        """Parse markdown text into segments with formatting info."""
        segments = []
        current_pos = 0

        # Combined pattern for markdown formatting
        # Order matters: check ** before *, __ before _
        pattern = r'(\*\*(.+?)\*\*|__(.+?)__|(?<!\*)\*(?!\*)(.+?)(?<!\*)\*(?!\*)|(?<!_)_(?!_)(.+?)(?<!_)_(?!_)|~~(.+?)~~|`([^`]+)`)'

        for match in re.finditer(pattern, text):
            start, end = match.span()

            # Add plain text before this match
            if start > current_pos:
                segments.append({
                    'text': text[current_pos:start],
                    'bold': False, 'italic': False, 'underline': False, 'code': False
                })

            # Determine which group matched
            groups = match.groups()
            if groups[1]:  # **bold**
                segments.append({'text': groups[1], 'bold': True, 'italic': False, 'underline': False, 'code': False})
            elif groups[2]:  # __bold__
                segments.append({'text': groups[2], 'bold': True, 'italic': False, 'underline': False, 'code': False})
            elif groups[3]:  # *italic*
                segments.append({'text': groups[3], 'bold': False, 'italic': True, 'underline': False, 'code': False})
            elif groups[4]:  # _underline_ (we'll use single underscore for underline)
                segments.append({'text': groups[4], 'bold': False, 'italic': False, 'underline': True, 'code': False})
            elif groups[5]:  # ~~strikethrough~~ - we'll treat as italic for now
                segments.append({'text': groups[5], 'bold': False, 'italic': True, 'underline': False, 'code': False})
            elif groups[6]:  # `code`
                segments.append({'text': groups[6], 'bold': False, 'italic': False, 'underline': False, 'code': True})

            current_pos = end

        # Add remaining plain text
        if current_pos < len(text):
            segments.append({
                'text': text[current_pos:],
                'bold': False, 'italic': False, 'underline': False, 'code': False
            })

        return segments if segments else [{'text': text, 'bold': False, 'italic': False, 'underline': False, 'code': False}]

    def _is_word_running(self) -> bool:
        """Check if Word is running."""
        if not HAS_WIN32:
            return False
        try:
            pythoncom.CoInitialize()
            word = win32com.client.GetActiveObject("Word.Application")
            return word is not None
        except:
            return False

    def _get_word_selection(self) -> tuple:
        """Get selected text from Word. Returns (text, has_selection)."""
        if not HAS_WIN32:
            return "", False

        try:
            pythoncom.CoInitialize()
            try:
                word = win32com.client.GetActiveObject("Word.Application")
            except Exception:
                # Word not running
                return "", False

            if not word:
                return "", False

            selection = word.Selection
            if not selection:
                return "", False

            text = ""
            try:
                text = selection.Text.strip() if selection.Text else ""
            except:
                pass

            # Check if there's actually a selection (not just cursor position)
            # wdSelectionIP = 1 (insertion point, no selection)
            has_selection = False
            try:
                has_selection = selection.Type != 1
            except:
                has_selection = len(text) > 0

            return text, has_selection
        except Exception as e:
            print(f"Could not get Word selection: {e}")
            return "", False

    def _set_word_text(self, text: str, replace: bool = True):
        """Set text in Word (replace selection or insert at cursor)."""
        if not HAS_WIN32:
            raise Exception("win32com not available")

        try:
            pythoncom.CoInitialize()
            word = win32com.client.GetActiveObject("Word.Application")

            if not word:
                raise Exception("Could not connect to Word")

            selection = word.Selection
            if not selection:
                raise Exception("Could not get Word selection")

            # TypeText replaces selection if there is one, or inserts at cursor
            selection.TypeText(text)

        except Exception as e:
            print(f"Could not set Word text: {e}")
            raise

    def _default_llm_call(self, prompt: str) -> str:
        """Default LLM call using the app's LLM infrastructure."""
        try:
            from icharlotte_core.llm import LLMHandler

            # Get the selected model
            provider, model_id = self._get_selected_model()
            print(f"Using model: {provider} / {model_id}")

            # Use LLMHandler.generate with proper parameters
            settings = {
                'temperature': 0.7,
                'top_p': 0.95,
                'max_tokens': 4096,
                'stream': False,
                'thinking_level': 'None'
            }

            result = LLMHandler.generate(
                provider=provider,
                model=model_id,
                system_prompt="You are a helpful writing assistant. Follow the user's instructions precisely. Output only the requested text without any preamble, explanation, or markdown formatting unless specifically asked.",
                user_prompt=prompt,
                file_contents="",
                settings=settings
            )

            return result.strip() if result else ""
        except Exception as e:
            print(f"LLM call failed: {e}")
            raise

    # ========== Outlook-specific methods ==========

    def _do_execute_outlook(self, prompt: str):
        """Execute LLM processing for Outlook email context."""
        try:
            # Verify inspector is still valid
            if not self.active_inspector:
                self.status_label.setText("Outlook compose window not found")
                self.execute_btn.setEnabled(True)
                self.cancel_btn.setText("Cancel")
                QMessageBox.warning(self, "Outlook Not Found",
                    "No active Outlook compose window found.\n\nPlease open an email compose window first, then try again.")
                return

            # Capture format if "Match Selection" is chosen
            format_type = self.format_combo.currentText()
            if format_type == FORMAT_MATCH:
                self._capture_outlook_selection_format()

            # Get selection from Outlook
            selected_text, has_selection = self._get_outlook_selection()

            if has_selection and selected_text:
                full_prompt = f"{prompt}\n\nEmail text to process:\n{selected_text}"
                self.status_label.setText("Processing selected email text...")
            else:
                full_prompt = prompt
                self.status_label.setText("Processing prompt...")

            QApplication.processEvents()

            # Store format type for callback
            self._pending_format_type = format_type
            self._is_outlook_task = True

            # Start worker thread for LLM call (using email-specific function)
            self._worker_thread = LLMWorkerThread(self._call_llm_for_email, full_prompt, self)
            self._worker_thread.finished.connect(self._on_llm_result)
            self._worker_thread.start()

        except Exception as e:
            self.status_label.setText(f"Error: {str(e)[:50]}")
            self.execute_btn.setEnabled(True)
            self.cancel_btn.setText("Cancel")
            QMessageBox.critical(self, "Error", f"Failed to process:\n{e}")

    def _get_outlook_selection(self) -> tuple:
        """Get selected text from Outlook compose window. Returns (text, has_selection)."""
        if not self.active_inspector:
            return "", False

        try:
            word_editor = self.active_inspector.WordEditor
            if not word_editor:
                print("WordEditor not available in Outlook Inspector")
                return "", False

            selection = word_editor.Application.Selection
            if not selection:
                return "", False

            text = ""
            try:
                raw_text = selection.Text
                text = raw_text.strip() if raw_text else ""
                if text:
                    print(f"Got Outlook selection text: '{text[:50]}...' ({len(text)} chars)")
                else:
                    print("No text selected in Outlook (cursor at insertion point)")
            except Exception as e:
                print(f"Error getting Outlook selection text: {e}")
                return "", False

            # wdSelectionIP = 1 (insertion point, no selection)
            has_selection = False
            try:
                sel_type = selection.Type
                has_selection = sel_type != 1 and len(text) > 0
                print(f"Outlook selection type: {sel_type}, has_selection: {has_selection}")
            except:
                has_selection = len(text) > 0

            return text, has_selection
        except Exception as e:
            print(f"Could not get Outlook selection: {e}")
            return "", False

    def _set_outlook_text(self, text: str, format_type: str = FORMAT_PLAIN):
        """Insert text into Outlook compose window with formatting."""
        if not self.active_inspector:
            raise Exception("No active Outlook inspector")

        try:
            word_editor = self.active_inspector.WordEditor
            if not word_editor:
                raise Exception("Could not access Outlook WordEditor")

            selection = word_editor.Application.Selection
            if not selection:
                raise Exception("Could not get selection in Outlook")

            # Use same formatting logic as Word since WordEditor provides full Word DOM access
            if format_type == FORMAT_PLAIN:
                self._insert_with_bullets(word_editor.Application, selection, text)

            elif format_type == FORMAT_MATCH:
                self._insert_with_format(selection, text, self._captured_format)

            elif format_type == FORMAT_MARKDOWN:
                self._insert_with_markdown(word_editor.Application, selection, text)

            elif format_type == FORMAT_DEFAULT:
                # For email, just insert plain text
                selection.TypeText(text)

            elif format_type == FORMAT_CUSTOM:
                self._insert_with_format(selection, text, self.custom_format_settings)

            else:
                selection.TypeText(text)

            print(f"Successfully inserted {len(text)} characters into Outlook with format: {format_type}")

        except Exception as e:
            print(f"_set_outlook_text error: {e}")
            raise

    def _capture_outlook_selection_format(self):
        """Capture the formatting of the current selection in Outlook compose window.

        Enhanced to capture paragraph styles and table context in email composition.
        """
        if not self.active_inspector:
            self._captured_format = None
            return

        try:
            word_editor = self.active_inspector.WordEditor
            if not word_editor:
                self._captured_format = None
                return

            selection = word_editor.Application.Selection
            font = selection.Font
            para = selection.ParagraphFormat

            underline_val = font.Underline
            has_underline = underline_val != 0 and underline_val != 9999999

            self._captured_format = {
                'font_name': font.Name if font.Name != '' else None,
                'font_size': font.Size if font.Size != 9999999 else 11,  # Default 11 for email
                'bold': font.Bold == -1,
                'italic': font.Italic == -1,
                'underline': has_underline,
                'first_indent': para.FirstLineIndent / 72 if para.FirstLineIndent != 9999999 else 0,
                'left_indent': para.LeftIndent / 72 if para.LeftIndent != 9999999 else 0,
            }

            # Capture paragraph style name
            try:
                style = selection.Style
                if style:
                    style_name = style.NameLocal
                    self._captured_format['style_name'] = style_name
                    print(f"Captured Outlook paragraph style: {style_name}")
            except Exception as e:
                print(f"Could not capture Outlook style: {e}")
                self._captured_format['style_name'] = None

            # Detect if selection is inside a table (emails can have tables)
            try:
                if selection.Information(12):  # wdWithInTable
                    self._captured_format['in_table'] = True
                    try:
                        cell = selection.Cells(1) if selection.Cells.Count > 0 else None
                        if cell:
                            shading = cell.Shading
                            if shading:
                                self._captured_format['cell_shading_color'] = shading.BackgroundPatternColor
                                self._captured_format['cell_shading_pattern'] = shading.Texture
                            self._captured_format['cell_vertical_align'] = cell.VerticalAlignment
                            print("Captured Outlook table cell formatting")
                    except:
                        pass
                else:
                    self._captured_format['in_table'] = False
            except:
                self._captured_format['in_table'] = False

            # Capture font color
            try:
                font_color = font.Color
                if font_color != 9999999:
                    self._captured_format['font_color'] = font_color
            except:
                pass

            # Capture line spacing
            try:
                line_spacing_rule = para.LineSpacingRule
                self._captured_format['line_spacing_rule'] = line_spacing_rule
                if line_spacing_rule in [3, 4, 5]:
                    self._captured_format['line_spacing'] = para.LineSpacing
            except:
                pass

            # No Track Changes in Outlook email composition
            self._captured_format['track_changes_enabled'] = False

            print(f"Captured Outlook format: {self._captured_format}")
        except Exception as e:
            print(f"Error capturing Outlook format: {e}")
            self._captured_format = None

    def _call_llm_for_email(self, prompt: str) -> str:
        """LLM call with email-specific system prompt."""
        try:
            from icharlotte_core.llm import LLMHandler

            provider, model_id = self._get_selected_model()
            print(f"Using model for email: {provider} / {model_id}")

            settings = {
                'temperature': 0.7,
                'top_p': 0.95,
                'max_tokens': 4096,
                'stream': False,
                'thinking_level': 'None'
            }

            result = LLMHandler.generate(
                provider=provider,
                model=model_id,
                system_prompt=EMAIL_SYSTEM_PROMPT,
                user_prompt=prompt,
                file_contents="",
                settings=settings
            )

            return result.strip() if result else ""
        except Exception as e:
            print(f"Email LLM call failed: {e}")
            raise

    # ========== End Outlook-specific methods ==========

    def showEvent(self, event):
        """Position the dialog when shown."""
        super().showEvent(event)

        if self.cursor_pos:
            # Position at cursor location
            x = self.cursor_pos.x()
            y = self.cursor_pos.y()

            # Get the screen that contains the cursor
            screen = QApplication.screenAt(self.cursor_pos)
            if screen:
                screen_geo = screen.availableGeometry()
            else:
                screen_geo = QApplication.primaryScreen().availableGeometry()

            # Ensure the popup stays within screen bounds
            # Adjust if popup would go off the right edge
            if x + self.width() > screen_geo.right():
                x = screen_geo.right() - self.width()
            # Adjust if popup would go off the bottom edge
            if y + self.height() > screen_geo.bottom():
                y = screen_geo.bottom() - self.height()
            # Ensure not off left or top edge
            if x < screen_geo.left():
                x = screen_geo.left()
            if y < screen_geo.top():
                y = screen_geo.top()

            self.move(x, y)
        else:
            # Fallback: center on screen
            screen = QApplication.primaryScreen().geometry()
            self.move(
                (screen.width() - self.width()) // 2,
                (screen.height() - self.height()) // 3
            )

        # Focus the custom input
        self.custom_input.setFocus()

    def closeEvent(self, event):
        """Clean up worker thread when dialog is closed."""
        if self._worker_thread and self._worker_thread.isRunning():
            self._worker_thread.cancel()
            self._worker_thread.terminate()
            self._worker_thread.wait(500)
            self._worker_thread = None
        super().closeEvent(event)

    def keyPressEvent(self, event):
        """Handle key presses."""
        if event.key() == Qt.Key.Key_Escape:
            self._on_cancel_clicked()
        else:
            super().keyPressEvent(event)

    def mousePressEvent(self, event):
        """Start dragging if clicking on header area."""
        if event.button() == Qt.MouseButton.LeftButton:
            # Check if click is in the header/title area (top 50 pixels)
            if event.position().y() <= 50:
                self._is_dragging = True
                self._drag_pos = event.globalPosition().toPoint() - self.frameGeometry().topLeft()
                event.accept()
                return
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        """Move window while dragging."""
        if self._is_dragging and self._drag_pos is not None:
            new_pos = event.globalPosition().toPoint() - self._drag_pos
            self.move(new_pos)
            event.accept()
            return
        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event):
        """Stop dragging."""
        if event.button() == Qt.MouseButton.LeftButton and self._is_dragging:
            self._is_dragging = False
            self._drag_pos = None
            event.accept()
            return
        super().mouseReleaseEvent(event)


class WordHotkeyManager:
    """Manages global hotkey registration and popup display."""

    def __init__(self, main_window=None):
        self.main_window = main_window
        self.popup = None
        self.signals = HotkeySignals()
        self.signals.show_popup.connect(self._show_popup)
        self._hotkey_registered = False

    def start(self):
        """Start listening for the global hotkey."""
        if not HAS_KEYBOARD:
            print("keyboard module not available - hotkey disabled")
            return False

        if self._hotkey_registered:
            return True

        try:
            # Register Win+V hotkey
            keyboard.add_hotkey('win+v', self._on_hotkey, suppress=True)
            self._hotkey_registered = True
            print("Global hotkey Win+V registered")
            return True
        except Exception as e:
            print(f"Failed to register hotkey: {e}")
            return False

    def stop(self):
        """Stop listening for the global hotkey."""
        if HAS_KEYBOARD and self._hotkey_registered:
            try:
                keyboard.remove_hotkey('win+v')
                self._hotkey_registered = False
                print("Global hotkey unregistered")
            except:
                pass

    def _on_hotkey(self):
        """Called when the hotkey is pressed (from keyboard thread)."""
        # Emit signal to show popup on main thread
        self.signals.show_popup.emit()

    def _show_popup(self):
        """Show the popup dialog (on main thread)."""
        if self.popup is None or not self.popup.isVisible():
            # Detect active application context BEFORE creating popup
            app_context, inspector = detect_active_app_context()

            # Only show if Word or Outlook compose is active
            if app_context == APP_CONTEXT_UNKNOWN:
                try:
                    print("Win+V pressed but neither Word nor Outlook compose is active")
                except OSError:
                    pass
                return

            # Capture cursor position on main thread
            from PySide6.QtGui import QCursor
            cursor_pos = QCursor.pos()

            self.popup = WordLLMPopup(
                parent=None,  # No parent so it's a top-level window
                llm_callback=self._get_llm_callback(),
                cursor_pos=cursor_pos
            )

            # Set the detected context before showing
            self.popup.set_app_context(app_context, inspector)

            self.popup.show()

            # Force focus to the popup window (Windows focus stealing prevention workaround)
            self._force_focus(self.popup)

    def _force_focus(self, window):
        """Force focus to the window using Windows API."""
        try:
            # First, use Qt methods
            window.raise_()
            window.activateWindow()
            QApplication.setActiveWindow(window)

            # Then use Windows API for guaranteed focus
            if HAS_WIN32:
                import ctypes
                hwnd = int(window.winId())

                # Allow our process to set foreground window
                ctypes.windll.user32.AllowSetForegroundWindow(-1)  # ASFW_ANY

                # Attach to the foreground thread to steal focus
                foreground_hwnd = ctypes.windll.user32.GetForegroundWindow()
                foreground_thread = ctypes.windll.user32.GetWindowThreadProcessId(foreground_hwnd, None)
                current_thread = ctypes.windll.kernel32.GetCurrentThreadId()

                if foreground_thread != current_thread:
                    ctypes.windll.user32.AttachThreadInput(foreground_thread, current_thread, True)

                ctypes.windll.user32.SetForegroundWindow(hwnd)
                ctypes.windll.user32.BringWindowToTop(hwnd)
                ctypes.windll.user32.SetFocus(hwnd)

                if foreground_thread != current_thread:
                    ctypes.windll.user32.AttachThreadInput(foreground_thread, current_thread, False)

            # Schedule focus to the input field after a brief delay
            QTimer.singleShot(50, lambda: self._focus_input(window))

        except Exception as e:
            print(f"Force focus error: {e}")
            # Fallback
            window.activateWindow()
            window.raise_()

    def _focus_input(self, window):
        """Set focus to the custom input field."""
        try:
            if window and hasattr(window, 'custom_input'):
                window.custom_input.setFocus()
                window.custom_input.activateWindow()
        except Exception as e:
            print(f"Focus input error: {e}")

    def _get_llm_callback(self):
        """Get the LLM callback function."""
        # Could customize this based on main_window settings
        return None  # Use default LLM call in popup


# Global instance
_hotkey_manager: Optional[WordHotkeyManager] = None


def init_word_hotkey(main_window=None) -> bool:
    """Initialize the Word hotkey manager."""
    global _hotkey_manager

    if _hotkey_manager is not None:
        return True

    _hotkey_manager = WordHotkeyManager(main_window)
    return _hotkey_manager.start()


def stop_word_hotkey():
    """Stop the Word hotkey manager."""
    global _hotkey_manager

    if _hotkey_manager is not None:
        _hotkey_manager.stop()
        _hotkey_manager = None
