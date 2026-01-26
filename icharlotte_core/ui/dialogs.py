import os
import json
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QFormLayout, QComboBox, QDialogButtonBox,
    QTableWidget, QHeaderView, QPushButton, QHBoxLayout, QMessageBox,
    QTableWidgetItem, QDoubleSpinBox, QSpinBox, QTextEdit, QLabel,
    QGroupBox, QTabWidget, QWidget, QListWidget, QListWidgetItem,
    QGridLayout, QCheckBox, QFrame, QScrollArea
)
from PySide6.QtCore import Qt, QByteArray
from ..config import GEMINI_DATA_DIR
from ..utils import log_event
from ..llm_config import LLMConfig, ModelSpec, TaskConfig

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

class PromptsDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Manage Prompts")
        self.resize(800, 600)
        self.scripts_dir = os.path.join(os.getcwd(), "Scripts")
        
        layout = QVBoxLayout(self)
        
        # File selection
        file_layout = QHBoxLayout()
        self.file_combo = QComboBox()
        self.file_combo.currentTextChanged.connect(self.load_prompt)
        file_layout.addWidget(self.file_combo)
        
        refresh_btn = QPushButton("Refresh")
        refresh_btn.clicked.connect(self.populate_files)
        file_layout.addWidget(refresh_btn)
        
        layout.addLayout(file_layout)
        
        # Editor
        self.editor = QTextEdit()
        layout.addWidget(self.editor)
        
        # Buttons
        btn_layout = QHBoxLayout()
        save_btn = QPushButton("Save")
        save_btn.clicked.connect(self.save_prompt)
        
        close_btn = QPushButton("Close")
        close_btn.clicked.connect(self.accept)
        
        btn_layout.addStretch()
        btn_layout.addWidget(save_btn)
        btn_layout.addWidget(close_btn)
        layout.addLayout(btn_layout)
        
        self.populate_files()
        
    def populate_files(self):
        self.file_combo.blockSignals(True)
        self.file_combo.clear()
        
        try:
            if os.path.exists(self.scripts_dir):
                files = [f for f in os.listdir(self.scripts_dir) if f.endswith("_PROMPT.txt")]
                files.sort()
                self.file_combo.addItems(files)
        except Exception as e:
            log_event(f"Error listing prompt files: {e}", "error")
            
        self.file_combo.blockSignals(False)
        
        if self.file_combo.count() > 0:
            self.load_prompt(self.file_combo.currentText())
            
    def load_prompt(self, filename):
        if not filename:
            self.editor.clear()
            return
            
        path = os.path.join(self.scripts_dir, filename)
        try:
            with open(path, 'r', encoding='utf-8') as f:
                content = f.read()
            self.editor.setPlainText(content)
        except Exception as e:
            self.editor.setPlainText(f"Error loading file: {e}")
            
    def save_prompt(self):
        filename = self.file_combo.currentText()
        if not filename:
            return

        path = os.path.join(self.scripts_dir, filename)
        content = self.editor.toPlainText()

        try:
            with open(path, 'w', encoding='utf-8') as f:
                f.write(content)
            QMessageBox.information(self, "Success", f"Saved {filename}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to save {filename}: {e}")


# =============================================================================
# LLM Settings Dialog
# =============================================================================

# Available models per provider
AVAILABLE_MODELS = {
    "Gemini": [
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
        ("gpt-4o", "GPT-4o"),
        ("gpt-4o-mini", "GPT-4o Mini"),
        ("gpt-4-turbo", "GPT-4 Turbo"),
        ("o1", "O1"),
        ("o1-mini", "O1 Mini"),
    ],
}

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
        "agent_docket", "agent_complaint", "agent_discovery_gen",
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

