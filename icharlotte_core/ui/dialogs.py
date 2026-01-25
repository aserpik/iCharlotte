import os
import json
from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QFormLayout, QComboBox, QDialogButtonBox, 
    QTableWidget, QHeaderView, QPushButton, QHBoxLayout, QMessageBox, 
    QTableWidgetItem, QDoubleSpinBox, QSpinBox, QTextEdit
)
from PyQt6.QtCore import Qt, QByteArray
from ..config import GEMINI_DATA_DIR
from ..utils import log_event

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

