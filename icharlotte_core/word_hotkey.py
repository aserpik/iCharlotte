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
    QGroupBox, QGridLayout, QFontComboBox, QDoubleSpinBox
)
from PySide6.QtCore import Qt, Signal, QObject, QTimer
from PySide6.QtGui import QFont, QKeySequence, QShortcut
import re

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

        # Title (dynamic based on context)
        self.title_label = QLabel("AI Assistant for Word")
        self.title_label.setStyleSheet("font-size: 16px; font-weight: bold; color: #6c63ff;")
        self.title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.title_label)

        # Saved Prompts Section
        prompt_label = QLabel("Saved Prompts:")
        layout.addWidget(prompt_label)

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

        layout.addLayout(prompt_row)

        # Separator
        line = QFrame()
        line.setFrameShape(QFrame.Shape.HLine)
        line.setStyleSheet("background-color: #45475a;")
        layout.addWidget(line)

        # Custom Prompt Section
        custom_label = QLabel("Custom Prompt (or type to add new):")
        layout.addWidget(custom_label)

        self.custom_input = QTextEdit()
        self.custom_input.setPlaceholderText("Enter your prompt here...\nExample: 'Convert this to formal legal language' or 'Summarize in bullet points'")
        self.custom_input.setMaximumHeight(80)
        layout.addWidget(self.custom_input)

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

        layout.addLayout(save_row)

        # Separator
        line2 = QFrame()
        line2.setFrameShape(QFrame.Shape.HLine)
        line2.setStyleSheet("background-color: #45475a;")
        layout.addWidget(line2)

        # Format Section
        format_label = QLabel("Output Format:")
        layout.addWidget(format_label)

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

        layout.addLayout(format_row)

        # Format preview label
        self.format_preview = QLabel("")
        self.format_preview.setStyleSheet("color: #a6adc8; font-size: 10px; font-style: italic;")
        self.format_preview.setWordWrap(True)
        layout.addWidget(self.format_preview)

        # Model Selection Section
        model_label = QLabel("AI Model:")
        layout.addWidget(model_label)

        self.model_combo = QComboBox()
        for display_name, provider, model_id in AVAILABLE_MODELS:
            self.model_combo.addItem(f"{display_name}", (provider, model_id))
        self.model_combo.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self.model_combo.currentIndexChanged.connect(self._on_model_changed)
        layout.addWidget(self.model_combo)

        # Use All Text checkbox
        self.use_all_text_check = QCheckBox("Use all document text as context (still replaces only selection)")
        self.use_all_text_check.setStyleSheet("color: #cdd6f4; font-size: 11px;")
        self.use_all_text_check.setToolTip(
            "When checked, sends the entire document to the LLM for context,\n"
            "but only replaces the selected text with the output."
        )
        layout.addWidget(self.use_all_text_check)

        # Status label
        self.status_label = QLabel("")
        self.status_label.setStyleSheet("color: #a6adc8; font-style: italic;")
        layout.addWidget(self.status_label)

        # Buttons
        btn_row = QHBoxLayout()
        btn_row.addStretch()

        cancel_btn = QPushButton("Cancel")
        cancel_btn.setObjectName("cancelBtn")
        cancel_btn.clicked.connect(self.close)
        btn_row.addWidget(cancel_btn)

        self.execute_btn = QPushButton("Execute (Enter)")
        self.execute_btn.clicked.connect(self.execute)
        btn_row.addWidget(self.execute_btn)

        layout.addLayout(btn_row)

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
        else:
            self.title_label.setText("AI Assistant for Word")
            self.setWindowTitle("AI Assistant - Document")

        self.refresh_combo()  # Refresh prompts for context

    def _on_format_changed(self, format_type: str):
        """Handle format dropdown change."""
        self.format_settings_btn.setEnabled(format_type == FORMAT_CUSTOM)
        self._update_format_preview()

    def _update_format_preview(self):
        """Update the format preview label."""
        format_type = self.format_combo.currentText()

        if format_type == FORMAT_PLAIN:
            self.format_preview.setText("No formatting will be applied")
        elif format_type == FORMAT_MATCH:
            self.format_preview.setText("Formatting will match the selected text")
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
        """Capture the formatting of the current selection in Word."""
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
            print(f"Captured format: {self._captured_format}")
        except Exception as e:
            print(f"Error capturing format: {e}")
            self._captured_format = None

    def execute(self):
        """Execute the LLM processing."""
        prompt = self.custom_input.toPlainText().strip()
        if not prompt:
            QMessageBox.warning(self, "Prompt Required", "Please enter or select a prompt.")
            return

        self.status_label.setText("Processing...")
        self.execute_btn.setEnabled(False)
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

            # Call LLM
            if self.llm_callback:
                result = self.llm_callback(full_prompt)
            else:
                result = self._default_llm_call(full_prompt)

            if result:
                # Insert/replace in Word with formatting
                self._set_word_text_internal(result, format_type)
                self.status_label.setText("Done!")
                QTimer.singleShot(500, self.close)
            else:
                self.status_label.setText("No response from LLM")
                self.execute_btn.setEnabled(True)

        except Exception as e:
            self.status_label.setText(f"Error: {str(e)[:50]}")
            self.execute_btn.setEnabled(True)
            QMessageBox.critical(self, "Error", f"Failed to process:\n{e}")

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
        """Set text in Word with formatting - re-establish connection to ensure it's fresh."""
        try:
            word = self._get_word_app()

            if not word:
                raise Exception("Could not connect to Word")

            selection = word.Selection
            if not selection:
                raise Exception("Could not access Word selection")

            if format_type == FORMAT_PLAIN:
                # Insert text, but still handle bullet points properly
                self._insert_with_bullets(word, selection, text)

            elif format_type == FORMAT_MATCH:
                # Apply captured format
                self._insert_with_format(selection, text, self._captured_format)

            elif format_type == FORMAT_MARKDOWN:
                # Parse markdown and apply formatting
                self._insert_with_markdown(word, selection, text)

            elif format_type == FORMAT_DEFAULT:
                # Reset to Normal style, then insert
                selection.Style = word.ActiveDocument.Styles("Normal")
                selection.TypeText(text)

            elif format_type == FORMAT_CUSTOM:
                # Apply custom format settings
                self._insert_with_format(selection, text, self.custom_format_settings)

            else:
                # Fallback to plain text
                selection.TypeText(text)

            print(f"Successfully inserted {len(text)} characters into Word with format: {format_type}")

        except Exception as e:
            print(f"_set_word_text_internal error: {e}")
            raise

    def _insert_with_format(self, selection, text: str, format_settings: dict):
        """Insert text with specific formatting."""
        if not format_settings:
            selection.TypeText(text)
            return

        try:
            font = selection.Font
            para = selection.ParagraphFormat

            # Apply font settings
            if format_settings.get('font_name'):
                font.Name = format_settings['font_name']
            if format_settings.get('font_size'):
                font.Size = format_settings['font_size']

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

            # Line spacing
            line_spacing = format_settings.get('line_spacing', 'Single')
            if line_spacing == 'Single':
                para.LineSpacingRule = 0  # wdLineSpaceSingle
            elif line_spacing == '1.5 Lines':
                para.LineSpacingRule = 1  # wdLineSpace1pt5
            elif line_spacing == 'Double':
                para.LineSpacingRule = 2  # wdLineSpaceDouble

            # Insert the text
            selection.TypeText(text)

        except Exception as e:
            print(f"Error applying format: {e}")
            # Fallback to plain text
            selection.TypeText(text)

    def _insert_with_bullets(self, word, selection, text: str):
        """Insert text with proper Word bullet formatting (no inline markdown)."""
        # First pass: group lines into items (bullet items include their continuations)
        items = self._parse_bullet_items(text)

        in_bullet_list = False
        for i, item in enumerate(items):
            if item['is_bullet']:
                if not in_bullet_list:
                    # First bullet - apply formatting and set up list template
                    in_bullet_list = True

                    # Apply Word's bullet formatting (enables auto-continuation on Enter)
                    selection.Range.ListFormat.ApplyBulletDefault()

                    # Modify the ListTemplate level settings for correct indentation
                    # This is what actually controls bullet/text positioning in Word lists
                    try:
                        list_fmt = selection.Range.ListFormat
                        if list_fmt.ListTemplate and list_fmt.ListLevelNumber > 0:
                            level = list_fmt.ListTemplate.ListLevels(list_fmt.ListLevelNumber)
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
                    selection.Range.ListFormat.RemoveNumbers()
                    para = selection.ParagraphFormat
                    para.LeftIndent = 0
                    para.FirstLineIndent = 0

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
        """Parse markdown and insert with Word formatting, including proper bullet lists."""
        # First pass: group lines into items (bullet items include their continuations)
        items = self._parse_bullet_items(text)

        in_bullet_list = False
        for i, item in enumerate(items):
            if item['is_bullet']:
                if not in_bullet_list:
                    # First bullet - apply formatting and set up list template
                    in_bullet_list = True

                    # Apply Word's bullet formatting (enables auto-continuation on Enter)
                    selection.Range.ListFormat.ApplyBulletDefault()

                    # Modify the ListTemplate level settings for correct indentation
                    # This is what actually controls bullet/text positioning in Word lists
                    try:
                        list_fmt = selection.Range.ListFormat
                        if list_fmt.ListTemplate and list_fmt.ListLevelNumber > 0:
                            level = list_fmt.ListTemplate.ListLevels(list_fmt.ListLevelNumber)
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
                    selection.Range.ListFormat.RemoveNumbers()
                    para = selection.ParagraphFormat
                    para.LeftIndent = 0
                    para.FirstLineIndent = 0

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

            # Call LLM with email-specific system prompt
            result = self._call_llm_for_email(full_prompt)

            if result:
                self._set_outlook_text(result, format_type)
                self.status_label.setText("Done!")
                QTimer.singleShot(500, self.close)
            else:
                self.status_label.setText("No response from LLM")
                self.execute_btn.setEnabled(True)

        except Exception as e:
            self.status_label.setText(f"Error: {str(e)[:50]}")
            self.execute_btn.setEnabled(True)
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
        """Capture the formatting of the current selection in Outlook compose window."""
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

    def keyPressEvent(self, event):
        """Handle key presses."""
        if event.key() == Qt.Key.Key_Escape:
            self.close()
        else:
            super().keyPressEvent(event)


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
