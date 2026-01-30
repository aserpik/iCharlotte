import os
import json
from datetime import datetime
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QFormLayout, QComboBox, QDialogButtonBox,
    QTableWidget, QHeaderView, QPushButton, QHBoxLayout, QMessageBox,
    QTableWidgetItem, QDoubleSpinBox, QSpinBox, QTextEdit, QLabel,
    QGroupBox, QTabWidget, QWidget, QListWidget, QListWidgetItem,
    QGridLayout, QCheckBox, QFrame, QScrollArea, QSplitter,
    QProgressBar, QLineEdit, QFileDialog, QInputDialog, QToolButton,
    QRadioButton, QButtonGroup
)
from PySide6.QtCore import Qt, QByteArray, Signal
from PySide6.QtGui import QFont, QColor, QTextCharFormat, QSyntaxHighlighter
from ..config import GEMINI_DATA_DIR
from ..utils import log_event
from ..llm_config import LLMConfig, ModelSpec, TaskConfig
from ..prompt_manager import get_prompt_manager, PromptManager, PromptVersion, PROMPTS_DIR
from ..llm import LLMWorker
from ..chat.token_counter import TokenCounter

# Available models per provider (must be defined before PromptsDialog class)
AVAILABLE_MODELS = {
    "Gemini": [
        ("gemini-3-pro-preview", "Gemini 3 Pro Preview"),
        ("gemini-3-flash-preview", "Gemini 3 Flash Preview"),
        ("gemini-2.5-pro", "Gemini 2.5 Pro"),
        ("gemini-2.5-flash", "Gemini 2.5 Flash"),
        ("gemini-2.0-flash", "Gemini 2.0 Flash"),
        ("gemini-1.5-pro", "Gemini 1.5 Pro"),
        ("gemini-1.5-flash", "Gemini 1.5 Flash"),
    ],
    "Claude": [
        ("claude-opus-4-20250514", "Claude Opus 4"),
        ("claude-sonnet-4-20250514", "Claude Sonnet 4"),
        ("claude-haiku-4-20250514", "Claude Haiku 4"),
        ("claude-3-5-sonnet-20241022", "Claude 3.5 Sonnet"),
        ("claude-3-opus-20240229", "Claude 3 Opus"),
    ],
    "OpenAI": [
        ("gpt-5.2-thinking", "GPT-5.2 Thinking"),
        ("gpt-5.2-instant", "GPT-5.2 Instant"),
        ("gpt-5.1-thinking", "GPT-5.1 Thinking"),
        ("gpt-5.1-instant", "GPT-5.1 Instant"),
        ("gpt-4o", "GPT-4o"),
        ("gpt-4o-mini", "GPT-4o Mini"),
        ("gpt-4-turbo", "GPT-4 Turbo"),
        ("o1", "O1"),
        ("o1-mini", "O1 Mini"),
    ],
}

class FileNumberDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("iCharlotte - Load Case")
        self.setFixedSize(300, 120)
        self.recent_file = os.path.join(GEMINI_DATA_DIR, "recent_cases.json")
        self.recent_cases = self.load_recent_cases()
        
        layout = QVBoxLayout(self)
        
        form = QFormLayout()
        self.file_input = QComboBox()
        self.file_input.setEditable(True)
        self.file_input.lineEdit().setPlaceholderText("####.###")
        self.file_input.addItems(self.recent_cases)
        # Default to most recent if available, else empty
        if self.recent_cases:
             self.file_input.setCurrentIndex(0)
        else:
             self.file_input.setCurrentIndex(-1)

        form.addRow("Enter File Number:", self.file_input)
        
        layout.addLayout(form)
        
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
        
    def load_recent_cases(self):
        if os.path.exists(self.recent_file):
            try:
                with open(self.recent_file, 'r') as f:
                    return json.load(f)
            except:
                return []
        return []

    def save_recent_case(self, file_num):
        if not file_num: return
        if file_num in self.recent_cases:
            self.recent_cases.remove(file_num)
        self.recent_cases.insert(0, file_num)
        self.recent_cases = self.recent_cases[:10] # Keep last 10
        
        if not os.path.exists(GEMINI_DATA_DIR):
            os.makedirs(GEMINI_DATA_DIR)
            
        try:
            with open(self.recent_file, 'w') as f:
                json.dump(self.recent_cases, f)
        except Exception as e:
            log_event(f"Error saving recent cases: {e}", "error")

    def get_file_number(self):
        val = self.file_input.currentText().strip()
        if val:
            self.save_recent_case(val)
        return val

class VariablesDialog(QDialog):
    def __init__(self, file_number, parent=None):
        super().__init__(parent)
        self.file_number = file_number
        self.json_path = os.path.join(GEMINI_DATA_DIR, f"{file_number}.json")
        self.settings_dir = os.path.join(GEMINI_DATA_DIR, "..", "config")
        self.settings_path = os.path.join(self.settings_dir, "variables_settings.json")
        self.raw_data = {}
        
        self.setWindowTitle(f"Variables - {file_number}")
        self.resize(600, 500)
        
        layout = QVBoxLayout(self)
        
        # Table
        self.table = QTableWidget()
        self.table.setColumnCount(2)
        self.table.setHorizontalHeaderLabels(["Variable Name", "Value"])
        
        # Make columns resizable by user
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        header.setStretchLastSection(True)
        
        layout.addWidget(self.table)
        
        # Buttons
        btn_layout = QHBoxLayout()
        add_btn = QPushButton("Add Variable")
        add_btn.clicked.connect(self.add_row)
        save_btn = QPushButton("Save")
        save_btn.clicked.connect(self.save_data)
        cancel_btn = QPushButton("Cancel")
        cancel_btn.clicked.connect(self.reject)
        
        btn_layout.addWidget(add_btn)
        btn_layout.addStretch()
        btn_layout.addWidget(save_btn)
        btn_layout.addWidget(cancel_btn)
        layout.addLayout(btn_layout)
        
        self.load_data()
        
        # Initial sizing if no state saved
        if not os.path.exists(self.settings_path):
            self.table.resizeColumnsToContents()
            # Ensure the second column still has some space if first is very wide
            if self.table.columnWidth(0) > 300:
                self.table.setColumnWidth(0, 300)
                
        self.load_header_state()

    def load_header_state(self):
        if os.path.exists(self.settings_path):
            try:
                with open(self.settings_path, 'r') as f:
                    settings = json.load(f)
                    header_state = settings.get("header_state")
                    if header_state:
                        self.table.horizontalHeader().restoreState(QByteArray.fromHex(header_state.encode()))
            except Exception as e:
                log_event(f"Error loading variables header state: {e}", "error")

    def save_header_state(self):
        if not os.path.exists(self.settings_dir):
            try:
                os.makedirs(self.settings_dir, exist_ok=True)
            except:
                pass
            
        try:
            state = self.table.horizontalHeader().saveState().toHex().data().decode()
            settings = {"header_state": state}
            with open(self.settings_path, 'w') as f:
                json.dump(settings, f)
        except Exception as e:
            log_event(f"Error saving variables header state: {e}", "error")

    def done(self, result):
        self.save_header_state()
        super().done(result)

    def load_data(self):
        if os.path.exists(self.json_path):
            try:
                with open(self.json_path, 'r', encoding='utf-8') as f:
                    self.raw_data = json.load(f)
            except Exception as e:
                QMessageBox.warning(self, "Error", f"Failed to load variables: {e}")
                self.raw_data = {}
        
        self.table.setRowCount(0)
        for key, value in self.raw_data.items():
            row = self.table.rowCount()
            self.table.insertRow(row)
            
            key_item = QTableWidgetItem(key)
            self.table.setItem(row, 0, key_item)
            
            val_str = ""
            if isinstance(value, dict) and "value" in value:
                val_str = str(value["value"])
            else:
                val_str = str(value)
                
            val_item = QTableWidgetItem(val_str)
            self.table.setItem(row, 1, val_item)

    def add_row(self):
        row = self.table.rowCount()
        self.table.insertRow(row)
        self.table.setItem(row, 0, QTableWidgetItem("New_Variable"))
        self.table.setItem(row, 1, QTableWidgetItem(""))
        self.table.scrollToBottom()

    def save_data(self):
        new_data = {}
        for row in range(self.table.rowCount()):
            key_item = self.table.item(row, 0)
            val_item = self.table.item(row, 1)
            
            if not key_item or not key_item.text().strip():
                continue
                
            key = key_item.text().strip()
            val = val_item.text().strip()
            
            if key in self.raw_data:
                existing = self.raw_data[key]
                if isinstance(existing, dict) and "value" in existing:
                    existing["value"] = val
                    new_data[key] = existing
                else:
                    new_data[key] = val
            else:
                new_data[key] = {"value": val, "source": "user_edit", "tags": []}
        
        try:
            if not os.path.exists(GEMINI_DATA_DIR):
                os.makedirs(GEMINI_DATA_DIR)
            with open(self.json_path, 'w', encoding='utf-8') as f:
                json.dump(new_data, f, indent=4)
            self.accept()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to save: {e}")

class SettingsDialog(QDialog):
    def __init__(self, settings, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Model Settings")
        self.settings = settings
        layout = QFormLayout(self)
        
        self.temp_spin = QDoubleSpinBox()
        self.temp_spin.setRange(0.0, 2.0)
        self.temp_spin.setSingleStep(0.1)
        self.temp_spin.setValue(settings.get('temperature', 1.0))
        layout.addRow("Temperature:", self.temp_spin)
        
        self.top_p_spin = QDoubleSpinBox()
        self.top_p_spin.setRange(0.0, 1.0)
        self.top_p_spin.setSingleStep(0.05)
        self.top_p_spin.setValue(settings.get('top_p', 0.95))
        layout.addRow("Top P:", self.top_p_spin)
        
        self.tokens_spin = QSpinBox()
        self.tokens_spin.setRange(-1, 2000000000)
        self.tokens_spin.setSpecialValueText("Unlimited")
        self.tokens_spin.setValue(settings.get('max_tokens', -1))
        layout.addRow("Max Tokens:", self.tokens_spin)

        self.thinking_combo = QComboBox()
        self.thinking_combo.addItems(["None", "Minimal", "Low", "Medium", "High"])
        
        current_level = settings.get('thinking_level', "None")
        index = self.thinking_combo.findText(current_level, Qt.MatchFlag.MatchFixedString)
        if index >= 0:
            self.thinking_combo.setCurrentIndex(index)
            
        self.thinking_combo.setToolTip("Gemini 3.0 Only. Pro supports High/Low. Flash supports all.")
        layout.addRow("Thinking Level:", self.thinking_combo)
        
        btns = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        btns.accepted.connect(self.accept)
        btns.rejected.connect(self.reject)
        layout.addWidget(btns)
        
    def get_settings(self):
        return {
            'temperature': self.temp_spin.value(),
            'top_p': self.top_p_spin.value(),
            'max_tokens': self.tokens_spin.value(),
            'thinking_level': self.thinking_combo.currentText()
        }

class SystemPromptDialog(QDialog):
    def __init__(self, current_prompt, parent=None):
        super().__init__(parent)
        self.setWindowTitle("System Instructions")
        self.resize(400, 300)
        layout = QVBoxLayout(self)
        
        self.text_edit = QTextEdit()
        self.text_edit.setPlainText(current_prompt)
        layout.addWidget(self.text_edit)
        
        btns = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        btns.accepted.connect(self.accept)
        btns.rejected.connect(self.reject)
        layout.addWidget(btns)
        
    def get_prompt(self):
        return self.text_edit.toPlainText()

# =============================================================================
# Prompt Variable Highlighter
# =============================================================================

class PromptHighlighter(QSyntaxHighlighter):
    """Syntax highlighter for prompt templates - highlights {variable} placeholders."""

    def __init__(self, parent=None):
        super().__init__(parent)

        # Variable format: {variable_name}
        self.variable_format = QTextCharFormat()
        self.variable_format.setForeground(QColor(0, 100, 200))  # Blue
        self.variable_format.setFontWeight(QFont.Weight.Bold)

        # Section headers format
        self.header_format = QTextCharFormat()
        self.header_format.setForeground(QColor(139, 0, 139))  # Dark magenta
        self.header_format.setFontWeight(QFont.Weight.Bold)

    def highlightBlock(self, text: str):
        if not text:
            return

        # Highlight {variables}
        import re
        for match in re.finditer(r'\{[^}]+\}', text):
            self.setFormat(match.start(), match.end() - match.start(), self.variable_format)

        # Highlight common section headers
        headers = ['Instructions:', 'Output:', 'Format:', 'Example:', 'Context:', 'Rules:']
        for header in headers:
            idx = text.find(header)
            if idx >= 0:
                self.setFormat(idx, len(header), self.header_format)


# =============================================================================
# Legacy Prompt Mapping
# =============================================================================

LEGACY_PROMPT_MAP = {
    "SUMMARIZE_PROMPT.txt": ("summarize", "main"),
    "SUMMARIZE_CROSS_CHECK_PROMPT.txt": ("summarize", "cross_check"),
    "CONSOLIDATE_DISCOVERY_PROMPT.txt": ("discovery", "consolidate"),
    "CROSS_CHECK_PROMPT.txt": ("discovery", "cross_check"),
    "SUMMARIZE_DEPOSITION_PROMPT.txt": ("deposition", "main"),
    "DEPOSITION_EXTRACTION_PROMPT.txt": ("deposition", "extraction"),
    "DEPOSITION_CROSS_CHECK_PROMPT.txt": ("deposition", "cross_check"),
    "TIMELINE_EXTRACTION_PROMPT.txt": ("timeline", "extraction"),
    "CONTRADICTION_DETECTION_PROMPT.txt": ("contradiction", "detection"),
    "LIABILITY_PROMPT.txt": ("liability", "main"),
    "EXPOSURE_PROMPT.txt": ("exposure", "main"),
    "MED_RECORD_PROMPT.txt": ("med_record", "main"),
    "MED_CHRON_PROMPT.txt": ("med_chron", "main"),
    "EXTRACTION_PASS_PROMPT.txt": ("extraction", "main"),
    "SUMMARIZE_DISCOVERY_PROMPT.txt": ("discovery", "main"),
}

# Mapping from workbench agent names to LLMConfig agent IDs
WORKBENCH_TO_AGENT_ID = {
    "summarize": "agent_summarize",
    "discovery": "agent_sum_disc",
    "deposition": "agent_sum_depo",
    "timeline": "agent_timeline",
    "contradiction": "agent_contradict",
    "liability": "agent_liability",
    "exposure": "agent_exposure",
    "med_record": "agent_med_rec",
    "med_chron": "agent_med_chron",
    "extraction": "agent_separate",
    "email_update": "func_email_compose",
    "chat": "func_chat",
}

DEFAULT_IMPROVEMENT_PROMPTS = {
    "specific": """Make this prompt more specific and detailed. Add concrete constraints,
expected output format, and edge case handling. Keep the core purpose but add clarity.""",

    "examples": """Add 2-3 concrete examples to this prompt showing the expected
input-output relationship. The examples should illustrate edge cases and ideal outputs.""",

    "clarity": """Improve the clarity and structure of this prompt. Use clear sections,
numbered steps where appropriate, and remove any ambiguity. Make instructions explicit.""",

    "shorten": """Condense this prompt to be more concise while preserving its core
meaning and effectiveness. Remove redundancy and verbose phrasing.""",
}

IMPROVEMENT_PROMPTS_FILE = os.path.join(GEMINI_DATA_DIR, "improvement_prompts.json")


def load_improvement_prompts() -> dict:
    """Load custom improvement prompts from config file."""
    if os.path.exists(IMPROVEMENT_PROMPTS_FILE):
        try:
            with open(IMPROVEMENT_PROMPTS_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception as e:
            log_event(f"Error loading improvement prompts: {e}", "error")
    return DEFAULT_IMPROVEMENT_PROMPTS.copy()


def save_improvement_prompts(prompts: dict):
    """Save custom improvement prompts to config file."""
    try:
        os.makedirs(os.path.dirname(IMPROVEMENT_PROMPTS_FILE), exist_ok=True)
        with open(IMPROVEMENT_PROMPTS_FILE, 'w', encoding='utf-8') as f:
            json.dump(prompts, f, indent=2)
    except Exception as e:
        log_event(f"Error saving improvement prompts: {e}", "error")


# =============================================================================
# Improvement Prompts Settings Dialog
# =============================================================================

class ImprovementPromptsDialog(QDialog):
    """Dialog for editing the LLM improvement action prompts."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Edit Improvement Prompts")
        self.resize(700, 500)

        self.prompts = load_improvement_prompts()
        self._setup_ui()

    def _setup_ui(self):
        layout = QVBoxLayout(self)

        info_label = QLabel("Customize the prompts used by the quick improvement buttons:")
        info_label.setStyleSheet("color: #666; margin-bottom: 10px;")
        layout.addWidget(info_label)

        # Create tabs for each prompt
        self.tabs = QTabWidget()

        self.editors = {}
        labels = {
            "specific": "Make Specific",
            "examples": "Add Examples",
            "clarity": "Improve Clarity",
            "shorten": "Shorten"
        }

        for key, label in labels.items():
            tab = QWidget()
            tab_layout = QVBoxLayout(tab)

            editor = QTextEdit()
            editor.setPlainText(self.prompts.get(key, DEFAULT_IMPROVEMENT_PROMPTS.get(key, "")))
            editor.setStyleSheet("font-family: Consolas; font-size: 11px;")
            tab_layout.addWidget(editor)

            # Reset button for this prompt
            reset_btn = QPushButton("Reset to Default")
            reset_btn.clicked.connect(lambda checked, k=key, e=editor: self._reset_prompt(k, e))
            tab_layout.addWidget(reset_btn)

            self.editors[key] = editor
            self.tabs.addTab(tab, label)

        layout.addWidget(self.tabs)

        # Buttons
        btn_layout = QHBoxLayout()

        reset_all_btn = QPushButton("Reset All to Defaults")
        reset_all_btn.clicked.connect(self._reset_all)
        btn_layout.addWidget(reset_all_btn)

        btn_layout.addStretch()

        save_btn = QPushButton("Save")
        save_btn.setStyleSheet("background-color: #4CAF50; color: white; font-weight: bold;")
        save_btn.clicked.connect(self._save)
        btn_layout.addWidget(save_btn)

        cancel_btn = QPushButton("Cancel")
        cancel_btn.clicked.connect(self.reject)
        btn_layout.addWidget(cancel_btn)

        layout.addLayout(btn_layout)

    def _reset_prompt(self, key: str, editor: QTextEdit):
        """Reset a single prompt to default."""
        default = DEFAULT_IMPROVEMENT_PROMPTS.get(key, "")
        editor.setPlainText(default)

    def _reset_all(self):
        """Reset all prompts to defaults."""
        for key, editor in self.editors.items():
            default = DEFAULT_IMPROVEMENT_PROMPTS.get(key, "")
            editor.setPlainText(default)

    def _save(self):
        """Save prompts and close."""
        for key, editor in self.editors.items():
            self.prompts[key] = editor.toPlainText()

        save_improvement_prompts(self.prompts)
        self.accept()

    def get_prompts(self) -> dict:
        """Get the current prompts."""
        return self.prompts


# =============================================================================
# Advanced Prompts Dialog
# =============================================================================

class PromptsDialog(QDialog):
    """
    Advanced Prompt Engineering Workbench.

    Features:
    - Version control for prompts
    - LLM-assisted prompt improvement
    - A/B testing with manual and LLM-as-judge rating
    - Performance metrics tracking
    """

    # Mapping from tab names to default agent selections
    TAB_TO_AGENT_MAP = {
        "Liability & Exposure": "liability",
        "Email Update": "email_update",
        "Chat": "chat",
        "Index": "summarize",
        "Report": "summarize",
    }

    def __init__(self, parent=None, current_tab: str = None):
        super().__init__(parent)
        self.setWindowTitle("Prompt Engineering Workbench")
        self.resize(1100, 750)

        self.scripts_dir = os.path.join(os.getcwd(), "Scripts")
        self.prompt_manager = get_prompt_manager()
        self.llm_config = LLMConfig()
        self.initial_tab = current_tab  # Store for auto-selection after UI setup

        # Current state
        self.current_agent = None
        self.current_pass = None
        self.current_version = None
        self.llm_worker = None
        self.ab_workers = []

        # Model settings state
        self.model_rows = []  # List of (provider_combo, model_combo, remove_btn) tuples

        # A/B testing state
        self.ab_mode = "prompts"  # "prompts", "models", or "both"
        self._ab_model_a = None
        self._ab_model_b = None

        self._setup_ui()
        self._populate_agents()
        self._migrate_if_needed()

        # Auto-select agent based on the tab that was open when dialog was launched
        self._auto_select_agent_for_tab()

    def _setup_ui(self):
        """Set up the main UI."""
        layout = QVBoxLayout(self)
        layout.setSpacing(10)

        # Header: Agent/Pass/Version selectors + Token count
        header = self._create_header()
        layout.addLayout(header)

        # Model settings panel (collapsible)
        self.model_settings_panel = self._create_model_settings_panel()
        layout.addWidget(self.model_settings_panel)

        # Main tabs
        self.tabs = QTabWidget()
        self.tabs.addTab(self._create_editor_tab(), "Editor")
        self.tabs.addTab(self._create_llm_assistant_tab(), "LLM Assistant")
        self.tabs.addTab(self._create_ab_testing_tab(), "A/B Testing")
        self.tabs.addTab(self._create_history_tab(), "Version History")
        self.tabs.addTab(self._create_dashboard_tab(), "Dashboard")
        layout.addWidget(self.tabs)

        # Bottom buttons
        btn_layout = QHBoxLayout()

        self.save_new_btn = QPushButton("Save as New Version")
        self.save_new_btn.setStyleSheet("background-color: #4CAF50; color: white; font-weight: bold; padding: 8px 16px;")
        self.save_new_btn.clicked.connect(self._save_as_new_version)
        btn_layout.addWidget(self.save_new_btn)

        self.save_current_btn = QPushButton("Save to Current")
        self.save_current_btn.clicked.connect(self._save_to_current)
        btn_layout.addWidget(self.save_current_btn)

        btn_layout.addStretch()

        close_btn = QPushButton("Close")
        close_btn.clicked.connect(self.accept)
        btn_layout.addWidget(close_btn)

        layout.addLayout(btn_layout)

    def _create_header(self) -> QHBoxLayout:
        """Create the header with selectors."""
        header = QHBoxLayout()

        # Agent selector
        header.addWidget(QLabel("Agent:"))
        self.agent_combo = QComboBox()
        self.agent_combo.setMinimumWidth(150)
        self.agent_combo.currentTextChanged.connect(self._on_agent_changed)
        header.addWidget(self.agent_combo)

        # Pass selector
        header.addWidget(QLabel("Pass:"))
        self.pass_combo = QComboBox()
        self.pass_combo.setMinimumWidth(120)
        self.pass_combo.currentTextChanged.connect(self._on_pass_changed)
        header.addWidget(self.pass_combo)

        # Version selector
        header.addWidget(QLabel("Version:"))
        self.version_combo = QComboBox()
        self.version_combo.setMinimumWidth(100)
        self.version_combo.currentTextChanged.connect(self._on_version_changed)
        header.addWidget(self.version_combo)

        header.addStretch()

        # Token count
        self.token_label = QLabel("Tokens: --")
        self.token_label.setStyleSheet("color: #666; font-weight: bold;")
        header.addWidget(self.token_label)

        return header

    def _create_model_settings_panel(self) -> QGroupBox:
        """Create the collapsible model settings panel."""
        group = QGroupBox("Model Settings")
        group.setCheckable(True)
        group.setChecked(False)  # Collapsed by default
        group.setStyleSheet("""
            QGroupBox {
                font-weight: bold;
                border: 1px solid #ccc;
                border-radius: 4px;
                margin-top: 10px;
                padding-top: 10px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px;
            }
        """)

        layout = QVBoxLayout(group)
        layout.setSpacing(8)

        # Use Default checkbox with effective model display
        default_row = QHBoxLayout()
        self.use_default_checkbox = QCheckBox("Use Default")
        self.use_default_checkbox.setChecked(True)
        self.use_default_checkbox.stateChanged.connect(self._on_use_default_changed)
        default_row.addWidget(self.use_default_checkbox)

        self.effective_model_label = QLabel("(inherits from: general)")
        self.effective_model_label.setStyleSheet("color: #666; font-style: italic;")
        default_row.addWidget(self.effective_model_label)
        default_row.addStretch()
        layout.addLayout(default_row)

        # Model sequence container
        self.model_sequence_container = QVBoxLayout()
        self.model_sequence_container.setSpacing(4)
        layout.addLayout(self.model_sequence_container)

        # Add model button and save button
        btn_row = QHBoxLayout()
        self.add_model_btn = QPushButton("+ Add Fallback Model")
        self.add_model_btn.setStyleSheet("padding: 4px 12px;")
        self.add_model_btn.clicked.connect(self._add_model_row)
        self.add_model_btn.setEnabled(False)
        btn_row.addWidget(self.add_model_btn)

        btn_row.addStretch()

        self.save_model_btn = QPushButton("Save Model Settings")
        self.save_model_btn.setStyleSheet("background-color: #4CAF50; color: white; padding: 4px 12px;")
        self.save_model_btn.clicked.connect(self._save_agent_model_settings)
        btn_row.addWidget(self.save_model_btn)

        layout.addLayout(btn_row)

        return group

    def _add_model_row(self, provider: str = None, model: str = None):
        """Add a model row to the sequence."""
        if len(self.model_rows) >= 4:
            QMessageBox.warning(self, "Limit Reached", "Maximum 4 fallback models allowed.")
            return

        row_widget = QWidget()
        row_layout = QHBoxLayout(row_widget)
        row_layout.setContentsMargins(0, 0, 0, 0)
        row_layout.setSpacing(8)

        # Priority number
        priority = len(self.model_rows) + 1
        priority_label = QLabel(f"{priority}.")
        priority_label.setFixedWidth(20)
        row_layout.addWidget(priority_label)

        # Provider combo
        provider_combo = QComboBox()
        provider_combo.setMinimumWidth(80)
        for p in AVAILABLE_MODELS.keys():
            provider_combo.addItem(p)
        if provider:
            idx = provider_combo.findText(provider)
            if idx >= 0:
                provider_combo.setCurrentIndex(idx)
        row_layout.addWidget(provider_combo)

        # Model combo
        model_combo = QComboBox()
        model_combo.setMinimumWidth(180)
        self._populate_model_combo_for_provider(model_combo, provider_combo.currentText())
        if model:
            for i in range(model_combo.count()):
                if model_combo.itemData(i) == model:
                    model_combo.setCurrentIndex(i)
                    break
        row_layout.addWidget(model_combo)

        # Connect provider change to update model combo
        provider_combo.currentTextChanged.connect(
            lambda p, mc=model_combo: self._populate_model_combo_for_provider(mc, p)
        )

        # Remove button
        remove_btn = QPushButton("X")
        remove_btn.setFixedWidth(30)
        remove_btn.setStyleSheet("background-color: #f44336; color: white;")
        remove_btn.clicked.connect(lambda: self._remove_model_row(row_widget))
        row_layout.addWidget(remove_btn)

        row_layout.addStretch()

        # Store references (including remove button for visibility control)
        self.model_rows.append((row_widget, provider_combo, model_combo, priority_label, remove_btn))
        self.model_sequence_container.addWidget(row_widget)

    def _remove_model_row(self, row_widget: QWidget):
        """Remove a model row from the sequence."""
        for i, (widget, _, _, _, _) in enumerate(self.model_rows):
            if widget == row_widget:
                self.model_rows.pop(i)
                widget.deleteLater()
                break

        # Renumber remaining rows
        for i, (_, _, _, priority_label, _) in enumerate(self.model_rows):
            priority_label.setText(f"{i + 1}.")

    def _populate_model_combo_for_provider(self, combo: QComboBox, provider: str):
        """Populate model combo for a specific provider."""
        combo.clear()
        models = AVAILABLE_MODELS.get(provider, [])
        for model_id, display_name in models:
            combo.addItem(display_name, model_id)

    def _on_use_default_changed(self, state: int):
        """Handle Use Default checkbox change."""
        use_default = state == Qt.CheckState.Checked.value
        self.add_model_btn.setEnabled(not use_default)

        # Clear existing rows
        for widget, _, _, _, _ in self.model_rows:
            widget.deleteLater()
        self.model_rows.clear()

        if use_default:
            # Show default models as read-only
            self._show_default_models_readonly()
            self.effective_model_label.setVisible(True)
            self._update_effective_model_label()
        else:
            # Show editable model rows - start with first default model
            self.effective_model_label.setVisible(False)
            self._add_default_model_row()

    def _show_default_models_readonly(self):
        """Show default model sequence as read-only rows."""
        if not self.current_agent:
            return

        agent_id = WORKBENCH_TO_AGENT_ID.get(self.current_agent, f"agent_{self.current_agent}")
        info = self.llm_config.get_agent_info(agent_id)
        default_task = info.get("default_task", "general")
        sequence = self.llm_config.get_model_sequence(default_task)

        for spec in sequence[:4]:  # Max 4
            self._add_model_row(spec.provider, spec.model)

        # Disable all rows and hide remove buttons (read-only)
        for widget, provider_combo, model_combo, _, remove_btn in self.model_rows:
            provider_combo.setEnabled(False)
            model_combo.setEnabled(False)
            remove_btn.setVisible(False)

    def _add_default_model_row(self):
        """Add a model row with the first default model selected."""
        if not self.current_agent:
            self._add_model_row()
            return

        agent_id = WORKBENCH_TO_AGENT_ID.get(self.current_agent, f"agent_{self.current_agent}")
        info = self.llm_config.get_agent_info(agent_id)
        default_task = info.get("default_task", "general")
        sequence = self.llm_config.get_model_sequence(default_task)

        if sequence:
            self._add_model_row(sequence[0].provider, sequence[0].model)
        else:
            self._add_model_row("Gemini", "gemini-2.5-flash")

    def _update_effective_model_label(self):
        """Update the effective model label based on current agent."""
        if not self.current_agent:
            self.effective_model_label.setText("(no agent selected)")
            return

        agent_id = WORKBENCH_TO_AGENT_ID.get(self.current_agent, f"agent_{self.current_agent}")
        info = self.llm_config.get_agent_info(agent_id)
        default_task = info.get("default_task", "general")

        # Get first model from task config
        sequence = self.llm_config.get_model_sequence(default_task)
        if sequence:
            first_model = sequence[0]
            self.effective_model_label.setText(
                f"(inherits from: {default_task} -> {first_model.provider}: {first_model.model})"
            )
        else:
            self.effective_model_label.setText(f"(inherits from: {default_task})")

    def _load_agent_model_settings(self):
        """Load model settings for the current agent."""
        # Clear existing rows
        for widget, _, _, _, _ in self.model_rows:
            widget.deleteLater()
        self.model_rows.clear()

        if not self.current_agent:
            return

        agent_id = WORKBENCH_TO_AGENT_ID.get(self.current_agent, f"agent_{self.current_agent}")
        config = self.llm_config.get_agent_config(agent_id)

        # Set use default checkbox
        self.use_default_checkbox.blockSignals(True)
        self.use_default_checkbox.setChecked(config.use_default)
        self.use_default_checkbox.blockSignals(False)

        self.add_model_btn.setEnabled(not config.use_default)

        if config.use_default:
            # Show default models as read-only
            self.effective_model_label.setVisible(True)
            self._update_effective_model_label()
            self._show_default_models_readonly()
        else:
            # Show custom model sequence (editable)
            self.effective_model_label.setVisible(False)
            if config.model_sequence:
                for spec in config.model_sequence:
                    self._add_model_row(spec.provider, spec.model)
            else:
                # No custom models, add one default to start
                self._add_default_model_row()

    def _save_agent_model_settings(self):
        """Save model settings for the current agent."""
        if not self.current_agent:
            QMessageBox.warning(self, "Warning", "No agent selected.")
            return

        agent_id = WORKBENCH_TO_AGENT_ID.get(self.current_agent, f"agent_{self.current_agent}")
        use_default = self.use_default_checkbox.isChecked()

        model_sequence = []
        if not use_default:
            for _, provider_combo, model_combo, _, _ in self.model_rows:
                provider = provider_combo.currentText()
                model = model_combo.currentData()
                if model:
                    model_sequence.append(ModelSpec(provider=provider, model=model))

        self.llm_config.update_agent_config(
            agent_id,
            model_sequence=model_sequence if model_sequence else None,
            use_default=use_default
        )

        QMessageBox.information(self, "Saved", f"Model settings saved for {self.current_agent}.")

    def _create_editor_tab(self) -> QWidget:
        """Create the main editor tab."""
        tab = QWidget()
        layout = QVBoxLayout(tab)

        # Toolbar with toggle
        toolbar = QHBoxLayout()

        # Raw/Rendered toggle button
        self.editor_mode_toggle = QPushButton("Switch to Rendered")
        self.editor_mode_toggle.setCheckable(True)
        self.editor_mode_toggle.setStyleSheet("""
            QPushButton {
                background-color: #f5f5f5;
                border: 1px solid #ccc;
                border-radius: 4px;
                padding: 6px 12px;
                font-weight: 500;
            }
            QPushButton:checked {
                background-color: #e3f2fd;
                border-color: #2196F3;
                color: #1976D2;
            }
            QPushButton:hover {
                background-color: #e8e8e8;
            }
            QPushButton:checked:hover {
                background-color: #bbdefb;
            }
        """)
        self.editor_mode_toggle.clicked.connect(self._toggle_editor_mode)
        toolbar.addWidget(self.editor_mode_toggle)

        # Mode indicator label
        self.mode_label = QLabel("Mode: Raw Markdown")
        self.mode_label.setStyleSheet("color: #666; margin-left: 10px;")
        toolbar.addWidget(self.mode_label)

        toolbar.addStretch()
        layout.addLayout(toolbar)

        # Editor with syntax highlighting
        self.editor = QTextEdit()
        self.editor.setStyleSheet("font-family: Consolas, monospace; font-size: 12px;")
        self.editor.textChanged.connect(self._update_token_count)
        self.highlighter = PromptHighlighter(self.editor.document())
        self.editor_raw_mode = True  # Track current mode
        layout.addWidget(self.editor)

        return tab

    def _toggle_editor_mode(self):
        """Toggle between raw markdown and rendered markdown editing."""
        if self.editor_mode_toggle.isChecked():
            # Switching TO rendered mode
            raw_text = self.editor.toPlainText()
            self.editor.blockSignals(True)
            self.editor.setMarkdown(raw_text)
            self.editor.blockSignals(False)
            self.editor_raw_mode = False
            self.editor_mode_toggle.setText("Switch to Raw")
            self.mode_label.setText("Mode: Rendered (WYSIWYG)")
            self.mode_label.setStyleSheet("color: #1976D2; margin-left: 10px; font-weight: bold;")
            # Disable syntax highlighter in rendered mode
            self.highlighter.setDocument(None)
        else:
            # Switching TO raw mode - convert rendered back to markdown
            markdown_text = self.editor.toMarkdown()
            self.editor.blockSignals(True)
            self.editor.setPlainText(markdown_text)
            self.editor.blockSignals(False)
            self.editor_raw_mode = True
            self.editor_mode_toggle.setText("Switch to Rendered")
            self.mode_label.setText("Mode: Raw Markdown")
            self.mode_label.setStyleSheet("color: #666; margin-left: 10px;")
            # Re-enable syntax highlighter
            self.highlighter.setDocument(self.editor.document())

        self._update_token_count()

    def _create_llm_assistant_tab(self) -> QWidget:
        """Create the LLM Assistant tab for prompt improvement."""
        tab = QWidget()
        layout = QVBoxLayout(tab)

        # Model selector and quick actions
        controls = QHBoxLayout()

        controls.addWidget(QLabel("Model:"))
        self.improve_model_combo = QComboBox()
        self._populate_model_combo(self.improve_model_combo)
        controls.addWidget(self.improve_model_combo)

        controls.addStretch()

        # Gear button for editing improvement prompts
        self.improvement_settings_btn = QToolButton()
        self.improvement_settings_btn.setText("\u2699")  # Gear unicode
        self.improvement_settings_btn.setToolTip("Edit improvement prompts")
        self.improvement_settings_btn.setStyleSheet("""
            QToolButton {
                font-size: 16px;
                padding: 4px 8px;
                border: 1px solid #ccc;
                border-radius: 4px;
                background: #f5f5f5;
            }
            QToolButton:hover {
                background: #e0e0e0;
            }
        """)
        self.improvement_settings_btn.clicked.connect(self._open_improvement_prompts_dialog)
        controls.addWidget(self.improvement_settings_btn)

        # Quick action buttons
        for action, label in [("specific", "Make Specific"), ("examples", "Add Examples"),
                               ("clarity", "Improve Clarity"), ("shorten", "Shorten")]:
            btn = QPushButton(label)
            btn.clicked.connect(lambda checked, a=action: self._run_quick_improvement(a))
            btn.setStyleSheet("padding: 6px 12px;")
            controls.addWidget(btn)

        layout.addLayout(controls)

        # Custom instruction
        instruction_layout = QHBoxLayout()
        instruction_layout.addWidget(QLabel("Custom Instruction:"))
        self.custom_instruction = QLineEdit()
        self.custom_instruction.setPlaceholderText("Enter custom improvement instruction...")
        instruction_layout.addWidget(self.custom_instruction)

        self.run_custom_btn = QPushButton("Improve")
        self.run_custom_btn.setStyleSheet("background-color: #2196F3; color: white; font-weight: bold;")
        self.run_custom_btn.clicked.connect(self._run_custom_improvement)
        instruction_layout.addWidget(self.run_custom_btn)

        layout.addLayout(instruction_layout)

        # Progress bar
        self.improve_progress = QProgressBar()
        self.improve_progress.setRange(0, 0)  # Indeterminate
        self.improve_progress.setVisible(False)
        layout.addWidget(self.improve_progress)

        # Side-by-side comparison
        splitter = QSplitter(Qt.Orientation.Horizontal)

        # Original panel
        original_widget = QWidget()
        original_layout = QVBoxLayout(original_widget)
        original_layout.setContentsMargins(0, 0, 0, 0)
        original_header = QLabel("Original Prompt")
        original_header.setStyleSheet("font-weight: bold; background-color: #fff0f0; padding: 5px;")
        original_layout.addWidget(original_header)
        self.original_text = QTextEdit()
        self.original_text.setReadOnly(True)
        self.original_text.setStyleSheet("font-family: Consolas; background-color: #fffafa;")
        original_layout.addWidget(self.original_text)
        splitter.addWidget(original_widget)

        # Improved panel
        improved_widget = QWidget()
        improved_layout = QVBoxLayout(improved_widget)
        improved_layout.setContentsMargins(0, 0, 0, 0)
        improved_header = QLabel("Improved Prompt")
        improved_header.setStyleSheet("font-weight: bold; background-color: #f0fff0; padding: 5px;")
        improved_layout.addWidget(improved_header)
        self.improved_text = QTextEdit()
        self.improved_text.setReadOnly(True)
        self.improved_text.setStyleSheet("font-family: Consolas; background-color: #f0fff0;")
        improved_layout.addWidget(self.improved_text)
        splitter.addWidget(improved_widget)

        layout.addWidget(splitter)

        # Accept/Reject buttons
        action_layout = QHBoxLayout()
        action_layout.addStretch()

        self.accept_btn = QPushButton("Accept Improvement")
        self.accept_btn.setStyleSheet("background-color: #4CAF50; color: white; font-weight: bold; padding: 8px 20px;")
        self.accept_btn.clicked.connect(self._accept_improvement)
        self.accept_btn.setEnabled(False)
        action_layout.addWidget(self.accept_btn)

        self.reject_btn = QPushButton("Reject")
        self.reject_btn.setStyleSheet("background-color: #f44336; color: white; padding: 8px 20px;")
        self.reject_btn.clicked.connect(self._reject_improvement)
        self.reject_btn.setEnabled(False)
        action_layout.addWidget(self.reject_btn)

        layout.addLayout(action_layout)

        return tab

    def _create_ab_testing_tab(self) -> QWidget:
        """Create the A/B Testing tab."""
        tab = QWidget()
        layout = QVBoxLayout(tab)

        # Comparison mode selector
        mode_group = QGroupBox("Comparison Mode")
        mode_layout = QHBoxLayout(mode_group)

        self.ab_mode_group = QButtonGroup(self)
        self.ab_mode_prompts = QRadioButton("Compare Prompts")
        self.ab_mode_prompts.setToolTip("Different prompts, same model")
        self.ab_mode_prompts.setChecked(True)
        self.ab_mode_group.addButton(self.ab_mode_prompts, 0)
        mode_layout.addWidget(self.ab_mode_prompts)

        self.ab_mode_models = QRadioButton("Compare Models")
        self.ab_mode_models.setToolTip("Same prompt, different models")
        self.ab_mode_group.addButton(self.ab_mode_models, 1)
        mode_layout.addWidget(self.ab_mode_models)

        self.ab_mode_both = QRadioButton("Compare Both")
        self.ab_mode_both.setToolTip("Different prompts AND different models")
        self.ab_mode_group.addButton(self.ab_mode_both, 2)
        mode_layout.addWidget(self.ab_mode_both)

        mode_layout.addStretch()
        layout.addWidget(mode_group)

        # Connect mode changes
        self.ab_mode_group.buttonClicked.connect(self._on_ab_mode_changed)

        # Selection area with A and B columns
        selection_group = QGroupBox("Test Configuration")
        selection_layout = QGridLayout(selection_group)

        # Headers
        selection_layout.addWidget(QLabel(""), 0, 0)
        header_a = QLabel("A")
        header_a.setStyleSheet("font-weight: bold; color: #1976D2;")
        selection_layout.addWidget(header_a, 0, 1)
        header_b = QLabel("B")
        header_b.setStyleSheet("font-weight: bold; color: #F57C00;")
        selection_layout.addWidget(header_b, 0, 2)

        # Prompt version row
        selection_layout.addWidget(QLabel("Prompt:"), 1, 0)
        self.ab_version_a = QComboBox()
        self.ab_version_a.setMinimumWidth(150)
        selection_layout.addWidget(self.ab_version_a, 1, 1)
        self.ab_version_b = QComboBox()
        self.ab_version_b.setMinimumWidth(150)
        selection_layout.addWidget(self.ab_version_b, 1, 2)

        # Model row
        selection_layout.addWidget(QLabel("Model:"), 2, 0)
        self.ab_model_a = QComboBox()
        self.ab_model_a.setMinimumWidth(150)
        self._populate_model_combo(self.ab_model_a)
        selection_layout.addWidget(self.ab_model_a, 2, 1)
        self.ab_model_b = QComboBox()
        self.ab_model_b.setMinimumWidth(150)
        self._populate_model_combo(self.ab_model_b)
        selection_layout.addWidget(self.ab_model_b, 2, 2)

        # Initial state for "prompts" mode: model A enabled, model B disabled (copies A)
        self.ab_model_a.setEnabled(True)
        self.ab_model_b.setEnabled(False)

        layout.addWidget(selection_group)

        # Run test button
        run_layout = QHBoxLayout()
        run_layout.addStretch()
        self.run_test_btn = QPushButton("Run A/B Test")
        self.run_test_btn.setStyleSheet("background-color: #FF9800; color: white; font-weight: bold; padding: 8px 20px;")
        self.run_test_btn.clicked.connect(self._run_ab_test)
        run_layout.addWidget(self.run_test_btn)
        run_layout.addStretch()
        layout.addLayout(run_layout)

        # Test input
        input_group = QGroupBox("Test Input")
        input_layout = QVBoxLayout(input_group)

        input_controls = QHBoxLayout()
        self.load_test_file_btn = QPushButton("Load from File")
        self.load_test_file_btn.clicked.connect(self._load_test_input)
        input_controls.addWidget(self.load_test_file_btn)
        input_controls.addStretch()
        input_layout.addLayout(input_controls)

        self.test_input = QTextEdit()
        self.test_input.setPlaceholderText("Paste test input here or load from file...")
        self.test_input.setMaximumHeight(120)
        input_layout.addWidget(self.test_input)

        layout.addWidget(input_group)

        # Progress
        self.ab_progress = QProgressBar()
        self.ab_progress.setRange(0, 0)
        self.ab_progress.setVisible(False)
        layout.addWidget(self.ab_progress)

        # Results comparison
        results_splitter = QSplitter(Qt.Orientation.Horizontal)

        # Version A result
        result_a = QWidget()
        result_a_layout = QVBoxLayout(result_a)
        result_a_layout.setContentsMargins(0, 0, 0, 0)
        self.result_a_header = QLabel("A Output")
        self.result_a_header.setStyleSheet("font-weight: bold; background-color: #e3f2fd; padding: 5px;")
        result_a_layout.addWidget(self.result_a_header)
        self.result_a_text = QTextEdit()
        self.result_a_text.setReadOnly(True)
        result_a_layout.addWidget(self.result_a_text)
        results_splitter.addWidget(result_a)

        # Version B result
        result_b = QWidget()
        result_b_layout = QVBoxLayout(result_b)
        result_b_layout.setContentsMargins(0, 0, 0, 0)
        self.result_b_header = QLabel("B Output")
        self.result_b_header.setStyleSheet("font-weight: bold; background-color: #fff3e0; padding: 5px;")
        result_b_layout.addWidget(self.result_b_header)
        self.result_b_text = QTextEdit()
        self.result_b_text.setReadOnly(True)
        result_b_layout.addWidget(self.result_b_text)
        results_splitter.addWidget(result_b)

        layout.addWidget(results_splitter)

        # Rating buttons
        rating_layout = QHBoxLayout()
        rating_layout.addStretch()

        self.rate_a_btn = QPushButton("A Wins")
        self.rate_a_btn.setStyleSheet("background-color: #2196F3; color: white; padding: 8px 20px;")
        self.rate_a_btn.clicked.connect(lambda: self._record_ab_rating("A"))
        self.rate_a_btn.setEnabled(False)
        rating_layout.addWidget(self.rate_a_btn)

        self.rate_tie_btn = QPushButton("Tie")
        self.rate_tie_btn.setStyleSheet("padding: 8px 20px;")
        self.rate_tie_btn.clicked.connect(lambda: self._record_ab_rating("Tie"))
        self.rate_tie_btn.setEnabled(False)
        rating_layout.addWidget(self.rate_tie_btn)

        self.rate_b_btn = QPushButton("B Wins")
        self.rate_b_btn.setStyleSheet("background-color: #FF9800; color: white; padding: 8px 20px;")
        self.rate_b_btn.clicked.connect(lambda: self._record_ab_rating("B"))
        self.rate_b_btn.setEnabled(False)
        rating_layout.addWidget(self.rate_b_btn)

        rating_layout.addStretch()
        layout.addLayout(rating_layout)

        return tab

    def _on_ab_mode_changed(self, button):
        """Handle A/B test mode change."""
        mode_id = self.ab_mode_group.id(button)

        if mode_id == 0:  # Compare Prompts
            self.ab_mode = "prompts"
            # Version dropdowns enabled, model A enabled, model B disabled (copies A)
            self.ab_version_a.setEnabled(True)
            self.ab_version_b.setEnabled(True)
            self.ab_model_a.setEnabled(True)
            self.ab_model_b.setEnabled(False)
        elif mode_id == 1:  # Compare Models
            self.ab_mode = "models"
            # Version B disabled (copies A), both models enabled
            self.ab_version_a.setEnabled(True)
            self.ab_version_b.setEnabled(False)
            self.ab_model_a.setEnabled(True)
            self.ab_model_b.setEnabled(True)
        else:  # Compare Both
            self.ab_mode = "both"
            # Everything enabled
            self.ab_version_a.setEnabled(True)
            self.ab_version_b.setEnabled(True)
            self.ab_model_a.setEnabled(True)
            self.ab_model_b.setEnabled(True)

    def _create_history_tab(self) -> QWidget:
        """Create the Version History tab."""
        tab = QWidget()
        layout = QHBoxLayout(tab)

        # Left: Version list
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        left_layout.setContentsMargins(0, 0, 0, 0)

        left_layout.addWidget(QLabel("Versions:"))

        self.version_table = QTableWidget()
        self.version_table.setColumnCount(5)
        self.version_table.setHorizontalHeaderLabels(["Version", "Date", "Score", "Uses", "Current"])
        self.version_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.version_table.setSelectionMode(QTableWidget.SelectionMode.SingleSelection)
        self.version_table.currentItemChanged.connect(self._on_history_version_selected)
        header = self.version_table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        header.setStretchLastSection(True)
        left_layout.addWidget(self.version_table)

        # Action buttons
        btn_layout = QHBoxLayout()

        self.set_current_btn = QPushButton("Set as Current")
        self.set_current_btn.clicked.connect(self._set_version_as_current)
        btn_layout.addWidget(self.set_current_btn)

        self.compare_btn = QPushButton("Compare")
        self.compare_btn.clicked.connect(self._compare_versions)
        btn_layout.addWidget(self.compare_btn)

        left_layout.addLayout(btn_layout)

        layout.addWidget(left_widget, 1)

        # Right: Version details
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        right_layout.setContentsMargins(0, 0, 0, 0)

        right_layout.addWidget(QLabel("Version Details:"))

        self.version_details = QTextEdit()
        self.version_details.setReadOnly(True)
        self.version_details.setStyleSheet("font-family: Consolas;")
        right_layout.addWidget(self.version_details)

        layout.addWidget(right_widget, 2)

        return tab

    def _create_dashboard_tab(self) -> QWidget:
        """Create the Performance Dashboard tab."""
        tab = QWidget()
        layout = QVBoxLayout(tab)

        # Summary stats
        stats_group = QGroupBox("Summary")
        stats_layout = QHBoxLayout(stats_group)

        self.stat_total_versions = QLabel("Total Versions: --")
        self.stat_total_versions.setStyleSheet("font-size: 14px; font-weight: bold;")
        stats_layout.addWidget(self.stat_total_versions)

        self.stat_current_score = QLabel("Current Score: --")
        self.stat_current_score.setStyleSheet("font-size: 14px; font-weight: bold;")
        stats_layout.addWidget(self.stat_current_score)

        self.stat_total_usage = QLabel("Total Usage: --")
        self.stat_total_usage.setStyleSheet("font-size: 14px; font-weight: bold;")
        stats_layout.addWidget(self.stat_total_usage)

        stats_layout.addStretch()
        layout.addWidget(stats_group)

        # Metrics table
        layout.addWidget(QLabel("All Versions:"))

        self.metrics_table = QTableWidget()
        self.metrics_table.setColumnCount(6)
        self.metrics_table.setHorizontalHeaderLabels(["Version", "Created", "Author", "Description", "Score", "Usage"])
        header = self.metrics_table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        header.setStretchLastSection(True)
        layout.addWidget(self.metrics_table)

        # Refresh button
        refresh_btn = QPushButton("Refresh Stats")
        refresh_btn.clicked.connect(self._refresh_dashboard)
        layout.addWidget(refresh_btn)

        return tab

    # =========================================================================
    # Population Methods
    # =========================================================================

    def _populate_agents(self):
        """Populate the agent dropdown from legacy files and prompt manager."""
        self.agent_combo.blockSignals(True)
        self.agent_combo.clear()

        agents = set()

        # Get agents from legacy files
        if os.path.exists(self.scripts_dir):
            for filename in os.listdir(self.scripts_dir):
                if filename.endswith("_PROMPT.txt"):
                    if filename in LEGACY_PROMPT_MAP:
                        agent, _ = LEGACY_PROMPT_MAP[filename]
                        agents.add(agent)

        # Get agents from prompt manager registry
        registry = self.prompt_manager._registry.get("prompts", {})
        for key in registry.keys():
            agent = key.split(":")[0]
            agents.add(agent)

        # Add predefined agents
        for agent in ['summarize', 'discovery', 'deposition', 'timeline', 'contradiction',
                      'liability', 'exposure', 'med_record', 'med_chron', 'extraction',
                      'email_update', 'chat']:
            agents.add(agent)

        for agent in sorted(agents):
            self.agent_combo.addItem(agent)

        self.agent_combo.blockSignals(False)

        if self.agent_combo.count() > 0:
            self._on_agent_changed(self.agent_combo.currentText())

    def _populate_model_combo(self, combo: QComboBox):
        """Populate a model selector combo box."""
        combo.clear()

        for provider, models in AVAILABLE_MODELS.items():
            for model_id, display_name in models:
                combo.addItem(f"{provider}: {display_name}", (provider, model_id))

    def _migrate_if_needed(self):
        """Run migration if the registry doesn't exist."""
        if not os.path.exists(os.path.join(PROMPTS_DIR, "registry.json")):
            try:
                migrated = self.prompt_manager.migrate_legacy_prompts()
                if migrated > 0:
                    log_event(f"Migrated {migrated} legacy prompts to versioned storage")
            except Exception as e:
                log_event(f"Error migrating prompts: {e}", "error")

    def _auto_select_agent_for_tab(self):
        """Auto-select the most relevant agent based on the tab that was open."""
        if not self.initial_tab:
            return

        # Look up the agent for this tab
        target_agent = self.TAB_TO_AGENT_MAP.get(self.initial_tab)
        if not target_agent:
            return

        # Find and select the agent in the combo box
        for i in range(self.agent_combo.count()):
            if self.agent_combo.itemText(i) == target_agent:
                self.agent_combo.setCurrentIndex(i)
                log_event(f"PromptsDialog: Auto-selected agent '{target_agent}' for tab '{self.initial_tab}'")
                break

    # =========================================================================
    # Event Handlers
    # =========================================================================

    def _on_agent_changed(self, agent: str):
        """Handle agent selection change."""
        self.current_agent = agent
        self.pass_combo.blockSignals(True)
        self.pass_combo.clear()

        passes = set()

        # Special handling for email_update agent - its prompts are stored differently
        if agent == "email_update":
            passes.add("system_prompt")
            passes.add("topic_instruction")
            passes.add("handling_instruction")

        # Get passes from legacy files for this agent
        for filename, (file_agent, file_pass) in LEGACY_PROMPT_MAP.items():
            if file_agent == agent:
                passes.add(file_pass)

        # Get passes from prompt manager
        registry = self.prompt_manager._registry.get("prompts", {})
        for key in registry.keys():
            if key.startswith(f"{agent}:"):
                _, pass_name = key.split(":", 1)
                passes.add(pass_name)

        for p in sorted(passes):
            self.pass_combo.addItem(p)

        self.pass_combo.blockSignals(False)

        # Load model settings for this agent
        self._load_agent_model_settings()

        if self.pass_combo.count() > 0:
            self._on_pass_changed(self.pass_combo.currentText())

    def _on_pass_changed(self, pass_name: str):
        """Handle pass selection change."""
        self.current_pass = pass_name
        self._populate_versions()

    def _populate_versions(self):
        """Populate the version dropdown."""
        self.version_combo.blockSignals(True)
        self.version_combo.clear()

        if not self.current_agent or not self.current_pass:
            self.version_combo.blockSignals(False)
            return

        # Get versions from prompt manager
        versions = self.prompt_manager.list_versions(self.current_agent, self.current_pass)

        if versions:
            for v in versions:
                suffix = " (current)" if v.is_current else ""
                self.version_combo.addItem(f"{v.version}{suffix}", v.version)
        else:
            # No versions yet - check for legacy file
            self.version_combo.addItem("legacy", "legacy")

        # Also update A/B version selectors
        self.ab_version_a.clear()
        self.ab_version_b.clear()
        for i in range(self.version_combo.count()):
            text = self.version_combo.itemText(i)
            data = self.version_combo.itemData(i)
            self.ab_version_a.addItem(text, data)
            self.ab_version_b.addItem(text, data)

        self.version_combo.blockSignals(False)

        if self.version_combo.count() > 0:
            self._on_version_changed(self.version_combo.currentText())

        # Refresh history tab
        self._refresh_history_table()
        self._refresh_dashboard()

    def _on_version_changed(self, version_text: str):
        """Handle version selection change."""
        if not version_text:
            return

        version = self.version_combo.currentData() or "current"
        self.current_version = version

        # Load prompt content
        content = None
        if version == "legacy":
            # Load from legacy file
            content = self._load_legacy_prompt()
        else:
            content = self.prompt_manager.get_prompt(self.current_agent, self.current_pass, version)

        if content:
            self._set_editor_content(content)
            self._update_token_count()

        # Update original text in LLM Assistant tab
        self.original_text.setMarkdown(content or "")

    def _load_legacy_prompt(self) -> str:
        """Load a legacy prompt file."""
        # Special handling for email_update agent
        if self.current_agent == "email_update":
            return self._load_email_update_prompt()

        for filename, (agent, pass_name) in LEGACY_PROMPT_MAP.items():
            if agent == self.current_agent and pass_name == self.current_pass:
                path = os.path.join(self.scripts_dir, filename)
                if os.path.exists(path):
                    with open(path, 'r', encoding='utf-8') as f:
                        return f.read()
        return ""

    def _load_email_update_prompt(self) -> str:
        """Load email update prompts from their JSON file."""
        json_path = os.path.join(GEMINI_DATA_DIR, "email_update_prompts.json")

        # Default prompts (same as in email_update_tab.py)
        defaults = {
            'system_prompt': (
                "You are an expert legal assistant writing specific sections of a formal status update email to an insurance claims adjuster. "
                "Your tone should be professional, objective, and strategic (defense-oriented). "
                "CRITICAL: You are only generating SPECIFIC SNIPPETS for an email, not the whole email. "
                "DO NOT include subject lines, factual backgrounds, salutations, or opening sentences. "
                "Use the provided STYLE EXAMPLES for structure and formatting within your assigned section, "
                "but DO NOT use any substantive facts from the examples. "
                "Use ONLY the provided CURRENT CASE DATA for facts. "
                "IMPORTANT: Do not wrap your entire response in bold markers (**). Only bold specific terms if necessary for emphasis. "
                "Do not provide any disclaimers about being an AI or not being an attorney."
            ),
            'topic_instruction': (
                "INSTRUCTION: Generate ONLY the 'New Topic' section (Name and Paragraph) based on the input.\n"
                "1. **Topic Name**: Create a concise header (e.g., 'Deposition of [Name]', 'Settlement Status').\n"
                "2. **Topic Paragraph**: Summarize the event or document. If a deposition, include demeanor and strategic analysis. Use bullet points for key facts where appropriate.\n"
                "STRICT REQUIREMENT: Output NOTHING besides the Topic Name and Topic Content labels. "
                "Ensure the content of the paragraph starts immediately after the label."
            ),
            'handling_instruction': (
                "INSTRUCTION: Generate ONLY the 'Further Case Handling' paragraph.\n"
                "   - **Start** with a phrase like 'Moving forward, we will...', 'Going forward, we will...', or 'In the meantime, we will...'.\n"
                "   - **Content**: List specific, actionable next steps (e.g., noticing depositions, filing motions, reviewing medical records, retaining experts).\n"
                "   - **Closing**: End with a standard closing like 'We will provide you with a further update upon completion of the same.' or 'Please advise if you have any objection...'.\n"
                "   - **Style**: Proactive and concise.\n"
                "STRICT REQUIREMENT: Output NOTHING besides the Further Handling label and paragraph."
            )
        }

        # Try to load from file
        if os.path.exists(json_path):
            try:
                with open(json_path, 'r', encoding='utf-8') as f:
                    saved = json.load(f)
                    defaults.update(saved)
            except Exception as e:
                log_event(f"Error loading email update prompts: {e}", "warning")

        # Return the prompt for the current pass
        return defaults.get(self.current_pass, "")

    def _save_email_update_prompt(self):
        """Save email update prompt to their JSON file."""
        json_path = os.path.join(GEMINI_DATA_DIR, "email_update_prompts.json")
        content = self.editor.toPlainText()

        # Load existing prompts
        prompts = {}
        if os.path.exists(json_path):
            try:
                with open(json_path, 'r', encoding='utf-8') as f:
                    prompts = json.load(f)
            except Exception:
                pass

        # Update the specific prompt
        prompts[self.current_pass] = content

        # Save back
        try:
            if not os.path.exists(GEMINI_DATA_DIR):
                os.makedirs(GEMINI_DATA_DIR)
            with open(json_path, 'w', encoding='utf-8') as f:
                json.dump(prompts, f, indent=4)
            QMessageBox.information(self, "Success", f"Saved {self.current_pass} prompt.")
            log_event(f"Saved email_update prompt: {self.current_pass}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to save: {e}")

    def _get_editor_raw_content(self) -> str:
        """Get the raw markdown content from editor, regardless of current mode."""
        if self.editor_raw_mode:
            return self.editor.toPlainText()
        else:
            return self.editor.toMarkdown()

    def _set_editor_content(self, content: str):
        """Set editor content, respecting current mode."""
        self.editor.blockSignals(True)
        if self.editor_raw_mode:
            self.editor.setPlainText(content)
        else:
            self.editor.setMarkdown(content)
        self.editor.blockSignals(False)

    def _update_token_count(self):
        """Update the token count display."""
        text = self._get_editor_raw_content()
        tokens = TokenCounter.estimate_tokens(text)

        if tokens >= 1000:
            self.token_label.setText(f"Tokens: {tokens/1000:.1f}k")
        else:
            self.token_label.setText(f"Tokens: {tokens}")

    # =========================================================================
    # Save Methods
    # =========================================================================

    def _save_as_new_version(self):
        """Save the current content as a new version."""
        if not self.current_agent or not self.current_pass:
            QMessageBox.warning(self, "Error", "Please select an agent and pass first.")
            return

        description, ok = QInputDialog.getText(
            self, "New Version", "Enter version description:",
            QLineEdit.EchoMode.Normal, ""
        )

        if not ok:
            return

        content = self._get_editor_raw_content()

        try:
            version = self.prompt_manager.create_version(
                self.current_agent,
                self.current_pass,
                content,
                description=description,
                author="user",
                set_as_current=True
            )

            QMessageBox.information(self, "Success", f"Created version {version}")
            self._populate_versions()

        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to save: {e}")

    def _save_to_current(self):
        """Save to the current legacy file (for backward compatibility)."""
        if not self.current_agent or not self.current_pass:
            return

        # Special handling for email_update agent
        if self.current_agent == "email_update":
            self._save_email_update_prompt()
            return

        # Find the legacy filename
        legacy_file = None
        for filename, (agent, pass_name) in LEGACY_PROMPT_MAP.items():
            if agent == self.current_agent and pass_name == self.current_pass:
                legacy_file = filename
                break

        if not legacy_file:
            # No legacy file, save as new version instead
            self._save_as_new_version()
            return

        path = os.path.join(self.scripts_dir, legacy_file)
        content = self._get_editor_raw_content()

        try:
            with open(path, 'w', encoding='utf-8') as f:
                f.write(content)
            QMessageBox.information(self, "Success", f"Saved to {legacy_file}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to save: {e}")

    # =========================================================================
    # LLM Assistant Methods
    # =========================================================================

    def _open_improvement_prompts_dialog(self):
        """Open dialog to edit improvement prompts."""
        dialog = ImprovementPromptsDialog(self)
        dialog.exec()

    def _run_quick_improvement(self, action: str):
        """Run a quick improvement action."""
        prompts = load_improvement_prompts()
        instruction = prompts.get(action, DEFAULT_IMPROVEMENT_PROMPTS.get(action, ""))
        self._run_improvement(instruction)

    def _run_custom_improvement(self):
        """Run custom improvement instruction."""
        instruction = self.custom_instruction.text().strip()
        if not instruction:
            QMessageBox.warning(self, "Warning", "Please enter an instruction.")
            return
        self._run_improvement(instruction)

    def _run_improvement(self, instruction: str):
        """Run an LLM improvement on the current prompt."""
        current_prompt = self._get_editor_raw_content()
        if not current_prompt.strip():
            QMessageBox.warning(self, "Warning", "No prompt to improve.")
            return

        # Get selected model
        model_data = self.improve_model_combo.currentData()
        if not model_data:
            QMessageBox.warning(self, "Warning", "Please select a model.")
            return

        provider, model = model_data

        # Update original text
        self.original_text.setMarkdown(current_prompt)
        self.improved_text.clear()

        # Show progress
        self.improve_progress.setVisible(True)
        self.accept_btn.setEnabled(False)
        self.reject_btn.setEnabled(False)

        # Build the improvement request
        system_prompt = """You are a prompt engineering expert. Your task is to improve prompts for LLM systems.
When improving a prompt:
- Keep the core purpose and intent
- Make instructions clearer and more explicit
- Add structure where helpful
- Remove ambiguity
- Ensure the output format is well-defined

Return ONLY the improved prompt, no explanations or commentary."""

        user_prompt = f"""Instruction: {instruction}

Original Prompt:
{current_prompt}

Provide the improved prompt:"""

        # Create and start worker
        self.llm_worker = LLMWorker(
            provider=provider,
            model=model,
            system=system_prompt,
            user=user_prompt,
            files="",
            settings={'stream': True, 'temperature': 0.7}
        )

        self.llm_worker.new_token.connect(self._on_improvement_token)
        self.llm_worker.finished.connect(self._on_improvement_finished)
        self.llm_worker.error.connect(self._on_improvement_error)
        self.llm_worker.start()

    def _on_improvement_token(self, token: str):
        """Handle streaming token from improvement."""
        cursor = self.improved_text.textCursor()
        cursor.movePosition(cursor.MoveOperation.End)
        cursor.insertText(token)
        self.improved_text.setTextCursor(cursor)

    def _on_improvement_finished(self, full_text: str):
        """Handle improvement completion."""
        self.improve_progress.setVisible(False)
        # Re-render with markdown now that streaming is complete
        self.improved_text.setMarkdown(full_text)
        self.accept_btn.setEnabled(True)
        self.reject_btn.setEnabled(True)

    def _on_improvement_error(self, error: str):
        """Handle improvement error."""
        self.improve_progress.setVisible(False)
        QMessageBox.critical(self, "Error", f"Improvement failed: {error}")

    def _accept_improvement(self):
        """Accept the improved prompt."""
        # Get raw markdown from the rendered improved text
        improved = self.improved_text.toMarkdown()
        if improved.strip():
            self._set_editor_content(improved)
            self.improved_text.clear()
            self.accept_btn.setEnabled(False)
            self.reject_btn.setEnabled(False)
            self.tabs.setCurrentIndex(0)  # Switch to Editor tab

    def _reject_improvement(self):
        """Reject the improved prompt."""
        self.improved_text.clear()
        self.accept_btn.setEnabled(False)
        self.reject_btn.setEnabled(False)

    # =========================================================================
    # A/B Testing Methods
    # =========================================================================

    def _load_test_input(self):
        """Load test input from a file."""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Load Test Input", "", "Text Files (*.txt);;All Files (*)"
        )
        if file_path:
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    self.test_input.setPlainText(f.read())
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to load file: {e}")

    def _run_ab_test(self):
        """Run A/B test with selected versions and/or models."""
        test_text = self.test_input.toPlainText().strip()
        if not test_text:
            QMessageBox.warning(self, "Warning", "Please enter test input.")
            return

        # Determine prompts based on mode
        version_a = self.ab_version_a.currentData()

        if self.ab_mode == "models":
            # Same prompt for both
            version_b = version_a
        else:
            version_b = self.ab_version_b.currentData()

        if not version_a:
            QMessageBox.warning(self, "Warning", "Please select a prompt version.")
            return

        if self.ab_mode == "prompts" and version_a == version_b:
            QMessageBox.warning(self, "Warning", "Please select different prompt versions.")
            return

        # Get prompts
        prompt_a = self.prompt_manager.get_prompt(self.current_agent, self.current_pass, version_a)
        if not prompt_a:
            prompt_a = self._load_legacy_prompt() if version_a == "legacy" else ""

        if self.ab_mode == "models":
            prompt_b = prompt_a
        else:
            prompt_b = self.prompt_manager.get_prompt(self.current_agent, self.current_pass, version_b)
            if not prompt_b:
                prompt_b = self._load_legacy_prompt() if version_b == "legacy" else ""

        if not prompt_a or not prompt_b:
            QMessageBox.warning(self, "Warning", "Could not load prompt(s).")
            return

        # Determine models based on mode
        model_data_a = self.ab_model_a.currentData()
        if not model_data_a:
            model_data_a = ("Gemini", "gemini-2.5-flash")

        if self.ab_mode == "prompts":
            # Same model for both
            model_data_b = model_data_a
        else:
            model_data_b = self.ab_model_b.currentData()
            if not model_data_b:
                model_data_b = ("Gemini", "gemini-2.5-flash")

        if self.ab_mode == "models" and model_data_a == model_data_b:
            QMessageBox.warning(self, "Warning", "Please select different models.")
            return

        provider_a, model_a = model_data_a
        provider_b, model_b = model_data_b

        # Clear results
        self.result_a_text.clear()
        self.result_b_text.clear()
        self.ab_progress.setVisible(True)
        self._disable_rating_buttons()

        # Store for later
        self._ab_version_a = version_a
        self._ab_version_b = version_b
        self._ab_model_a = model_data_a
        self._ab_model_b = model_data_b
        self._ab_results = {"A": "", "B": ""}

        # Update headers to show what's being compared
        if self.ab_mode == "prompts":
            self.result_a_header.setText(f"A: {version_a} ({provider_a}: {model_a})")
            self.result_b_header.setText(f"B: {version_b} ({provider_a}: {model_a})")
        elif self.ab_mode == "models":
            self.result_a_header.setText(f"A: {provider_a}: {model_a}")
            self.result_b_header.setText(f"B: {provider_b}: {model_b}")
        else:  # both
            self.result_a_header.setText(f"A: {version_a} / {provider_a}: {model_a}")
            self.result_b_header.setText(f"B: {version_b} / {provider_b}: {model_b}")

        # Run test A
        worker_a = LLMWorker(provider_a, model_a, prompt_a, test_text, "", {'stream': False})
        worker_a.finished.connect(lambda r: self._on_ab_result("A", r))
        worker_a.error.connect(lambda e: self._on_ab_error("A", e))
        worker_a.start()
        self.ab_workers.append(worker_a)

        # Run test B
        worker_b = LLMWorker(provider_b, model_b, prompt_b, test_text, "", {'stream': False})
        worker_b.finished.connect(lambda r: self._on_ab_result("B", r))
        worker_b.error.connect(lambda e: self._on_ab_error("B", e))
        worker_b.start()
        self.ab_workers.append(worker_b)

    def _on_ab_result(self, version: str, result: str):
        """Handle A/B test result."""
        self._ab_results[version] = result

        if version == "A":
            self.result_a_text.setMarkdown(result)
        else:
            self.result_b_text.setMarkdown(result)

        # Check if both done
        if self._ab_results["A"] and self._ab_results["B"]:
            self.ab_progress.setVisible(False)
            self._enable_rating_buttons()

    def _on_ab_error(self, version: str, error: str):
        """Handle A/B test error."""
        self.ab_progress.setVisible(False)
        QMessageBox.critical(self, "Error", f"Version {version} failed: {error}")

    def _enable_rating_buttons(self):
        self.rate_a_btn.setEnabled(True)
        self.rate_tie_btn.setEnabled(True)
        self.rate_b_btn.setEnabled(True)

    def _disable_rating_buttons(self):
        self.rate_a_btn.setEnabled(False)
        self.rate_tie_btn.setEnabled(False)
        self.rate_b_btn.setEnabled(False)

    def _record_ab_rating(self, winner: str):
        """Record A/B test rating."""
        # Calculate scores
        if winner == "A":
            score_a, score_b = 1.0, 0.0
        elif winner == "B":
            score_a, score_b = 0.0, 1.0
        else:  # Tie
            score_a, score_b = 0.5, 0.5

        # Record performance
        if hasattr(self, '_ab_version_a') and self._ab_version_a != "legacy":
            self.prompt_manager.record_usage(self.current_agent, self.current_pass, self._ab_version_a)
            self.prompt_manager.record_performance(self.current_agent, self.current_pass, score_a, self._ab_version_a)

        if hasattr(self, '_ab_version_b') and self._ab_version_b != "legacy":
            self.prompt_manager.record_usage(self.current_agent, self.current_pass, self._ab_version_b)
            self.prompt_manager.record_performance(self.current_agent, self.current_pass, score_b, self._ab_version_b)

        self._disable_rating_buttons()
        QMessageBox.information(self, "Recorded", f"Rating recorded: {winner} wins!")

        # Refresh dashboard
        self._refresh_dashboard()

    # =========================================================================
    # History Tab Methods
    # =========================================================================

    def _refresh_history_table(self):
        """Refresh the version history table."""
        self.version_table.setRowCount(0)

        if not self.current_agent or not self.current_pass:
            return

        versions = self.prompt_manager.list_versions(self.current_agent, self.current_pass)

        for v in versions:
            row = self.version_table.rowCount()
            self.version_table.insertRow(row)

            self.version_table.setItem(row, 0, QTableWidgetItem(v.version))

            # Format date
            try:
                dt = datetime.fromisoformat(v.created)
                date_str = dt.strftime("%Y-%m-%d %H:%M")
            except:
                date_str = v.created[:16] if len(v.created) > 16 else v.created
            self.version_table.setItem(row, 1, QTableWidgetItem(date_str))

            score_str = f"{v.performance_score:.2f}" if v.performance_score else "--"
            self.version_table.setItem(row, 2, QTableWidgetItem(score_str))

            self.version_table.setItem(row, 3, QTableWidgetItem(str(v.usage_count)))

            current_str = "Yes" if v.is_current else ""
            self.version_table.setItem(row, 4, QTableWidgetItem(current_str))

    def _on_history_version_selected(self, current, previous):
        """Handle version selection in history table."""
        if not current:
            return

        row = current.row()
        version_item = self.version_table.item(row, 0)
        if not version_item:
            return

        version = version_item.text()

        # Load and display version content
        content = self.prompt_manager.get_prompt(self.current_agent, self.current_pass, version)
        if content:
            self.version_details.setPlainText(content)

    def _set_version_as_current(self):
        """Set selected version as current."""
        current_row = self.version_table.currentRow()
        if current_row < 0:
            QMessageBox.warning(self, "Warning", "Please select a version.")
            return

        version_item = self.version_table.item(current_row, 0)
        if not version_item:
            return

        version = version_item.text()

        if self.prompt_manager.set_current(self.current_agent, self.current_pass, version):
            QMessageBox.information(self, "Success", f"Set {version} as current.")
            self._populate_versions()
        else:
            QMessageBox.critical(self, "Error", "Failed to set current version.")

    def _compare_versions(self):
        """Compare selected version with current."""
        current_row = self.version_table.currentRow()
        if current_row < 0:
            QMessageBox.warning(self, "Warning", "Please select a version to compare.")
            return

        version_item = self.version_table.item(current_row, 0)
        if not version_item:
            return

        selected_version = version_item.text()

        # Get current version content
        current_content = self.prompt_manager.get_prompt(self.current_agent, self.current_pass, "current")
        if not current_content:
            current_content = self._load_legacy_prompt()

        # Get selected version content
        selected_content = self.prompt_manager.get_prompt(self.current_agent, self.current_pass, selected_version)

        if not current_content or not selected_content:
            QMessageBox.warning(self, "Warning", "Could not load version content.")
            return

        # Show diff dialog
        from .diff_viewer import show_diff_dialog
        show_diff_dialog(
            current_content, selected_content,
            title=f"Compare: Current vs {selected_version}",
            old_label="Current Version",
            new_label=f"Version {selected_version}",
            parent=self
        )

    # =========================================================================
    # Dashboard Methods
    # =========================================================================

    def _refresh_dashboard(self):
        """Refresh the dashboard statistics."""
        if not self.current_agent or not self.current_pass:
            self.stat_total_versions.setText("Total Versions: --")
            self.stat_current_score.setText("Current Score: --")
            self.stat_total_usage.setText("Total Usage: --")
            self.metrics_table.setRowCount(0)
            return

        versions = self.prompt_manager.list_versions(self.current_agent, self.current_pass)

        # Summary stats
        self.stat_total_versions.setText(f"Total Versions: {len(versions)}")

        current_version = next((v for v in versions if v.is_current), None)
        if current_version and current_version.performance_score:
            self.stat_current_score.setText(f"Current Score: {current_version.performance_score:.2f}")
        else:
            self.stat_current_score.setText("Current Score: --")

        total_usage = sum(v.usage_count for v in versions)
        self.stat_total_usage.setText(f"Total Usage: {total_usage}")

        # Metrics table
        self.metrics_table.setRowCount(0)

        for v in versions:
            row = self.metrics_table.rowCount()
            self.metrics_table.insertRow(row)

            self.metrics_table.setItem(row, 0, QTableWidgetItem(v.version))

            try:
                dt = datetime.fromisoformat(v.created)
                date_str = dt.strftime("%Y-%m-%d %H:%M")
            except:
                date_str = v.created[:16] if len(v.created) > 16 else v.created
            self.metrics_table.setItem(row, 1, QTableWidgetItem(date_str))

            self.metrics_table.setItem(row, 2, QTableWidgetItem(v.author or "--"))
            self.metrics_table.setItem(row, 3, QTableWidgetItem(v.description or "--"))

            score_str = f"{v.performance_score:.2f}" if v.performance_score else "--"
            self.metrics_table.setItem(row, 4, QTableWidgetItem(score_str))

            self.metrics_table.setItem(row, 5, QTableWidgetItem(str(v.usage_count)))


# =============================================================================
# LLM Settings Dialog
# =============================================================================

TASK_TYPE_DESCRIPTIONS = {
    "general": "Default for general-purpose LLM calls",
    "extraction": "Document data extraction (entities, dates, facts)",
    "summary": "Document summarization and condensing",
    "cross_check": "Verification and consistency checking",
    "classification": "Quick categorization and labeling",
    "quick": "Fast, simple operations (low complexity)",
}

# Agent categories for organization
AGENT_CATEGORIES = [
    ("Document Processing Agents", [
        "agent_separate", "agent_summarize", "agent_sum_disc", "agent_sum_depo",
        "agent_med_rec", "agent_med_chron", "agent_organize", "agent_timeline",
        "agent_contradict"
    ]),
    ("Case Agents", [
        "agent_docket", "agent_complaint",
        "agent_subpoena", "agent_liability", "agent_exposure"
    ]),
    ("UI Functions", [
        "func_chat", "func_email_intelligence", "func_email_compose",
        "func_liability_tab", "func_sent_monitor", "func_attachment_classifier"
    ]),
]


class LLMSettingsDialog(QDialog):
    """Dialog for configuring LLM model preferences per agent/function."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("LLM Settings")
        self.resize(800, 650)
        self.config = LLMConfig()

        # Track widgets for agents and task types
        self.agent_widgets = {}
        self.task_widgets = {}

        self._setup_ui()
        self._load_current_settings()

    def _setup_ui(self):
        layout = QVBoxLayout(self)

        # Provider status header
        status_frame = QFrame()
        status_frame.setFrameShape(QFrame.Shape.StyledPanel)
        status_layout = QHBoxLayout(status_frame)
        status_layout.setContentsMargins(10, 5, 10, 5)

        status_label = QLabel("API Keys:")
        status_label.setStyleSheet("font-weight: bold;")
        status_layout.addWidget(status_label)

        for provider in ["Gemini", "Claude", "OpenAI"]:
            available = self.config.is_provider_available(provider)
            indicator = QLabel(f"{provider}: {'OK' if available else 'Missing'}")
            indicator.setStyleSheet(
                f"color: {'#4CAF50' if available else '#f44336'}; font-weight: bold;"
            )
            status_layout.addWidget(indicator)

        status_layout.addStretch()
        layout.addWidget(status_frame)

        # Main tabs: Agents & Functions | Default Profiles
        self.main_tabs = QTabWidget()

        # Tab 1: Agents & Functions
        agents_tab = self._create_agents_tab()
        self.main_tabs.addTab(agents_tab, "Agents && Functions")

        # Tab 2: Default Profiles (task types)
        defaults_tab = self._create_defaults_tab()
        self.main_tabs.addTab(defaults_tab, "Default Profiles")

        layout.addWidget(self.main_tabs)

        # Buttons
        btn_layout = QHBoxLayout()

        reset_btn = QPushButton("Reset All to Defaults")
        reset_btn.clicked.connect(self._reset_to_defaults)
        btn_layout.addWidget(reset_btn)

        btn_layout.addStretch()

        save_btn = QPushButton("Save")
        save_btn.setStyleSheet("font-weight: bold;")
        save_btn.clicked.connect(self._save_settings)
        btn_layout.addWidget(save_btn)

        cancel_btn = QPushButton("Cancel")
        cancel_btn.clicked.connect(self.reject)
        btn_layout.addWidget(cancel_btn)

        layout.addLayout(btn_layout)

    def _create_agents_tab(self) -> QWidget:
        """Create the agents/functions configuration tab."""
        tab = QWidget()
        layout = QVBoxLayout(tab)

        # Instructions
        info_label = QLabel(
            "Configure the LLM model for each agent/function. "
            "Use 'Default' to inherit from Default Profiles, or select a specific model."
        )
        info_label.setWordWrap(True)
        info_label.setStyleSheet("color: #666; margin-bottom: 10px;")
        layout.addWidget(info_label)

        # Scroll area for agents
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)

        scroll_content = QWidget()
        scroll_layout = QVBoxLayout(scroll_content)
        scroll_layout.setSpacing(15)

        # Create sections for each category
        for category_name, agent_ids in AGENT_CATEGORIES:
            group = QGroupBox(category_name)
            group_layout = QVBoxLayout(group)
            group_layout.setSpacing(8)

            for agent_id in agent_ids:
                row = self._create_agent_row(agent_id)
                group_layout.addLayout(row)

            scroll_layout.addWidget(group)

        scroll_layout.addStretch()
        scroll.setWidget(scroll_content)
        layout.addWidget(scroll)

        return tab

    def _create_agent_row(self, agent_id: str) -> QHBoxLayout:
        """Create a configuration row for a single agent."""
        row = QHBoxLayout()
        row.setSpacing(10)

        info = self.config.get_agent_info(agent_id)

        # Agent name label
        name_label = QLabel(info["name"])
        name_label.setFixedWidth(150)
        name_label.setStyleSheet("font-weight: bold;")
        name_label.setToolTip(info["description"])
        row.addWidget(name_label)

        # Use default checkbox
        use_default_cb = QCheckBox("Use Default")
        use_default_cb.setChecked(True)
        use_default_cb.setFixedWidth(100)
        row.addWidget(use_default_cb)

        # Provider combo
        provider_combo = QComboBox()
        provider_combo.addItems(["Gemini", "Claude", "OpenAI"])
        provider_combo.setFixedWidth(90)
        provider_combo.setEnabled(False)
        row.addWidget(provider_combo)

        # Model combo
        model_combo = QComboBox()
        model_combo.setFixedWidth(180)
        model_combo.setEnabled(False)
        row.addWidget(model_combo)

        # Current setting indicator
        current_label = QLabel("")
        current_label.setStyleSheet("color: #888; font-style: italic;")
        current_label.setFixedWidth(150)
        row.addWidget(current_label)

        row.addStretch()

        # Connect signals
        use_default_cb.toggled.connect(
            lambda checked, pc=provider_combo, mc=model_combo:
            self._toggle_agent_custom(checked, pc, mc)
        )
        provider_combo.currentTextChanged.connect(
            lambda text, mc=model_combo: self._update_model_list(text, mc)
        )

        # Store widgets
        self.agent_widgets[agent_id] = {
            "use_default": use_default_cb,
            "provider": provider_combo,
            "model": model_combo,
            "current_label": current_label,
            "default_task": info["default_task"]
        }

        return row

    def _toggle_agent_custom(self, use_default: bool, provider_combo: QComboBox,
                             model_combo: QComboBox):
        """Toggle between default and custom model selection."""
        provider_combo.setEnabled(not use_default)
        model_combo.setEnabled(not use_default)

        if not use_default and model_combo.count() == 0:
            self._update_model_list(provider_combo.currentText(), model_combo)

    def _create_defaults_tab(self) -> QWidget:
        """Create the default profiles (task types) configuration tab."""
        tab = QWidget()
        layout = QVBoxLayout(tab)

        # Instructions
        info_label = QLabel(
            "Configure default model sequences for each task type. "
            "Agents set to 'Use Default' will inherit these settings based on their task type."
        )
        info_label.setWordWrap(True)
        info_label.setStyleSheet("color: #666; margin-bottom: 10px;")
        layout.addWidget(info_label)

        # Tabs for each task type
        self.task_tabs = QTabWidget()

        for task_type in self.config.get_all_task_types():
            task_tab = self._create_task_tab(task_type)
            display_name = task_type.replace("_", " ").title()
            self.task_tabs.addTab(task_tab, display_name)

        layout.addWidget(self.task_tabs)

        return tab

    def _create_task_tab(self, task_type: str) -> QWidget:
        """Create a tab widget for a specific task type."""
        tab = QWidget()
        layout = QVBoxLayout(tab)

        # Description
        desc = TASK_TYPE_DESCRIPTIONS.get(task_type, "")
        if desc:
            desc_label = QLabel(desc)
            desc_label.setStyleSheet("color: #666; font-style: italic; margin-bottom: 10px;")
            layout.addWidget(desc_label)

        # Model sequence group
        group = QGroupBox("Model Sequence (in order of preference)")
        group_layout = QVBoxLayout(group)

        # Store widgets for this task type
        self.task_widgets[task_type] = {
            "model_rows": [],
            "max_retries": None,
            "timeout": None
        }

        # Create 4 model selection rows
        for i in range(4):
            row_layout = QHBoxLayout()

            priority_label = QLabel(f"{i + 1}.")
            priority_label.setFixedWidth(20)
            row_layout.addWidget(priority_label)

            provider_combo = QComboBox()
            provider_combo.addItem("(None)", None)
            provider_combo.addItems(["Gemini", "Claude", "OpenAI"])
            provider_combo.setFixedWidth(100)
            row_layout.addWidget(provider_combo)

            model_combo = QComboBox()
            model_combo.setFixedWidth(200)
            row_layout.addWidget(model_combo)

            max_tokens_spin = QSpinBox()
            max_tokens_spin.setRange(1024, 128000)
            max_tokens_spin.setSingleStep(1024)
            max_tokens_spin.setValue(8192)
            max_tokens_spin.setPrefix("Tokens: ")
            max_tokens_spin.setFixedWidth(130)
            row_layout.addWidget(max_tokens_spin)

            row_layout.addStretch()

            # Connect provider change to update model list
            provider_combo.currentTextChanged.connect(
                lambda text, mc=model_combo: self._update_model_list(text, mc)
            )

            self.task_widgets[task_type]["model_rows"].append({
                "provider": provider_combo,
                "model": model_combo,
                "max_tokens": max_tokens_spin
            })

            group_layout.addLayout(row_layout)

        layout.addWidget(group)

        # Additional settings
        settings_layout = QHBoxLayout()

        retries_label = QLabel("Max Retries:")
        retries_spin = QSpinBox()
        retries_spin.setRange(1, 10)
        retries_spin.setValue(3)
        self.task_widgets[task_type]["max_retries"] = retries_spin
        settings_layout.addWidget(retries_label)
        settings_layout.addWidget(retries_spin)

        settings_layout.addSpacing(20)

        timeout_label = QLabel("Timeout (seconds):")
        timeout_spin = QSpinBox()
        timeout_spin.setRange(30, 600)
        timeout_spin.setValue(120)
        self.task_widgets[task_type]["timeout"] = timeout_spin
        settings_layout.addWidget(timeout_label)
        settings_layout.addWidget(timeout_spin)

        settings_layout.addStretch()

        layout.addLayout(settings_layout)
        layout.addStretch()

        return tab

    def _update_model_list(self, provider: str, model_combo: QComboBox):
        """Update model combo based on selected provider."""
        model_combo.clear()

        if provider in AVAILABLE_MODELS:
            for model_id, description in AVAILABLE_MODELS[provider]:
                model_combo.addItem(description, model_id)

    def _load_current_settings(self):
        """Load current settings from LLMConfig into the UI."""
        # Load agent settings
        for agent_id, widgets in self.agent_widgets.items():
            agent_config = self.config.get_agent_config(agent_id)

            widgets["use_default"].setChecked(agent_config.use_default)

            if not agent_config.use_default and agent_config.model_sequence:
                # Set custom model
                first_model = agent_config.model_sequence[0]
                idx = widgets["provider"].findText(first_model.provider)
                if idx >= 0:
                    widgets["provider"].setCurrentIndex(idx)

                self._update_model_list(first_model.provider, widgets["model"])
                for j in range(widgets["model"].count()):
                    if widgets["model"].itemData(j) == first_model.model:
                        widgets["model"].setCurrentIndex(j)
                        break

            # Update current label to show what's being used
            self._update_agent_current_label(agent_id)

        # Load task type settings
        for task_type, widgets in self.task_widgets.items():
            task_config = self.config.get_task_config(task_type)

            # Load model sequence
            for i, row in enumerate(widgets["model_rows"]):
                if i < len(task_config.model_sequence):
                    model_spec = task_config.model_sequence[i]

                    # Set provider
                    idx = row["provider"].findText(model_spec.provider)
                    if idx >= 0:
                        row["provider"].setCurrentIndex(idx)

                    # Update model list and select current
                    self._update_model_list(model_spec.provider, row["model"])
                    for j in range(row["model"].count()):
                        if row["model"].itemData(j) == model_spec.model:
                            row["model"].setCurrentIndex(j)
                            break

                    row["max_tokens"].setValue(model_spec.max_tokens)
                else:
                    row["provider"].setCurrentIndex(0)  # (None)

            # Load other settings
            widgets["max_retries"].setValue(task_config.max_retries)
            widgets["timeout"].setValue(task_config.timeout_seconds)

    def _update_agent_current_label(self, agent_id: str):
        """Update the label showing current model for an agent."""
        widgets = self.agent_widgets[agent_id]
        agent_config = self.config.get_agent_config(agent_id)

        if agent_config.use_default:
            # Show what default profile will be used
            default_task = widgets["default_task"]
            task_config = self.config.get_task_config(default_task)
            if task_config.model_sequence:
                first = task_config.model_sequence[0]
                widgets["current_label"].setText(f"-> {first.model}")
            else:
                widgets["current_label"].setText(f"-> {default_task}")
        else:
            widgets["current_label"].setText("")

    def _save_settings(self):
        """Save settings to LLMConfig."""
        try:
            # Save agent settings
            for agent_id, widgets in self.agent_widgets.items():
                use_default = widgets["use_default"].isChecked()

                if use_default:
                    self.config.update_agent_config(agent_id, use_default=True)
                else:
                    provider = widgets["provider"].currentText()
                    model_id = widgets["model"].currentData()

                    if provider and model_id:
                        model_sequence = [ModelSpec(
                            provider=provider,
                            model=model_id,
                            max_tokens=8192
                        )]
                        self.config.update_agent_config(
                            agent_id,
                            model_sequence=model_sequence,
                            use_default=False
                        )

            # Save task type settings
            for task_type, widgets in self.task_widgets.items():
                model_sequence = []

                for row in widgets["model_rows"]:
                    provider = row["provider"].currentText()
                    if provider == "(None)" or not provider:
                        continue

                    model_id = row["model"].currentData()
                    if not model_id:
                        continue

                    model_sequence.append(ModelSpec(
                        provider=provider,
                        model=model_id,
                        max_tokens=row["max_tokens"].value()
                    ))

                if model_sequence:
                    self.config.update_task_config(
                        task_type,
                        model_sequence,
                        max_retries=widgets["max_retries"].value(),
                        timeout_seconds=widgets["timeout"].value()
                    )

            QMessageBox.information(self, "Success", "LLM settings saved successfully.")
            self.accept()

        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to save settings: {e}")

    def _reset_to_defaults(self):
        """Reset all settings to defaults."""
        reply = QMessageBox.question(
            self, "Reset to Defaults",
            "Are you sure you want to reset all LLM settings to defaults?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )

        if reply == QMessageBox.StandardButton.Yes:
            # Remove config file and reload
            from ..llm_config import CONFIG_FILE
            if os.path.exists(CONFIG_FILE):
                try:
                    os.remove(CONFIG_FILE)
                except:
                    pass

            # Reset singleton
            LLMConfig._instance = None
            self.config = LLMConfig()

            # Reload UI
            self._load_current_settings()
            QMessageBox.information(self, "Reset", "Settings reset to defaults.")

