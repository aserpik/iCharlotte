import os
import json
import re
import markdown
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QSplitter, QFrame, QLabel, 
    QComboBox, QPushButton, QListWidget, QListWidgetItem, QTextBrowser, 
    QPlainTextEdit, QMessageBox, QDialog, QFormLayout, QDialogButtonBox, QTextEdit
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QDragEnterEvent, QDropEvent, QTextCursor

from ..config import API_KEYS, GEMINI_DATA_DIR
from ..utils import log_event
from ..llm import LLMWorker, ModelFetcher
from .tabs import ChatTab
from .dialogs import SettingsDialog, SystemPromptDialog

# Constants
ANALYSIS_TYPES = ["Select Analysis...", "Liability Analysis", "Exposure Analysis"]
LIABILITY_TYPES = ["Auto", "Premises Liability", "Wrongful Death", "Nuisance", "Habitability", "Dangerous Condition of Public Property"]
PROMPTS_FILE = os.path.join(GEMINI_DATA_DIR, "liability_prompts.json")

class LiabilityPromptsDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Manage Analysis Prompts")
        self.resize(800, 600)
        self.prompts_data = self.load_prompts()
        
        layout = QVBoxLayout(self)
        
        # Selection Controls
        form_layout = QHBoxLayout()
        
        self.analysis_combo = QComboBox()
        self.analysis_combo.addItems(ANALYSIS_TYPES[1:]) # Skip "Select..."
        self.analysis_combo.currentTextChanged.connect(self.update_liability_visibility)
        form_layout.addWidget(QLabel("Type:"))
        form_layout.addWidget(self.analysis_combo)
        
        self.liability_combo = QComboBox()
        self.liability_combo.addItems(LIABILITY_TYPES)
        self.liability_label = QLabel("Liability:")
        form_layout.addWidget(self.liability_label)
        form_layout.addWidget(self.liability_combo)
        
        # Connect combos to load prompt
        self.analysis_combo.currentTextChanged.connect(self.load_current_prompt)
        self.liability_combo.currentTextChanged.connect(self.load_current_prompt)
        
        layout.addLayout(form_layout)
        
        # Editor
        layout.addWidget(QLabel("Prompt:"))
        self.editor = QTextEdit()
        layout.addWidget(self.editor)
        
        # Buttons
        btn_layout = QHBoxLayout()
        save_btn = QPushButton("Save Current Prompt")
        save_btn.clicked.connect(self.save_current_prompt)
        
        close_btn = QPushButton("Close")
        close_btn.clicked.connect(self.accept)
        
        btn_layout.addStretch()
        btn_layout.addWidget(save_btn)
        btn_layout.addWidget(close_btn)
        layout.addLayout(btn_layout)
        
        # Initial State
        self.update_liability_visibility(self.analysis_combo.currentText())
        self.load_current_prompt()
        
    def load_prompts(self):
        if os.path.exists(PROMPTS_FILE):
            try:
                with open(PROMPTS_FILE, 'r') as f:
                    return json.load(f)
            except Exception as e:
                log_event(f"Error loading liability prompts: {e}", "error")
        return {}

    def get_key(self):
        a_type = self.analysis_combo.currentText()
        if a_type == "Liability Analysis":
            return f"{a_type}|{self.liability_combo.currentText()}"
        else:
            return a_type

    def update_liability_visibility(self, text):
        is_liability = (text == "Liability Analysis")
        self.liability_label.setVisible(is_liability)
        self.liability_combo.setVisible(is_liability)
        self.load_current_prompt()

    def load_current_prompt(self):
        key = self.get_key()
        prompt = self.prompts_data.get(key, "")
        if not prompt:
            # Set default placeholders if empty
            if "Liability" in key:
                prompt = f"Analyze the attached documents for {key.split('|')[1]} liability issues. Identify key arguments for and against liability."
            else:
                prompt = "Analyze the attached documents for potential financial exposure. Estimate range of damages."
        self.editor.setPlainText(prompt)

    def save_current_prompt(self):
        key = self.get_key()
        self.prompts_data[key] = self.editor.toPlainText()
        
        try:
            if not os.path.exists(GEMINI_DATA_DIR):
                os.makedirs(GEMINI_DATA_DIR)
            with open(PROMPTS_FILE, 'w') as f:
                json.dump(self.prompts_data, f, indent=4)
            QMessageBox.information(self, "Success", "Prompt saved.")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to save prompts: {e}")

class LiabilityExposureTab(ChatTab):
    """
    Extends ChatTab but overrides setup_ui to inject specific controls.
    Since ChatTab.setup_ui is monolithic, we mostly re-implement it 
    but reuse the helper methods (update_models, etc.)
    """
    def __init__(self, parent=None):
        # We call super init which calls setup_ui.
        # But we want our own setup_ui.
        # Python resolves methods dynamically, so if we define setup_ui here,
        # super().__init__() will call *our* setup_ui.
        super().__init__(parent)
        self.prompts_data = {} # Will load on demand

    def setup_ui(self):
        main_layout = QHBoxLayout(self)
        splitter = QSplitter(Qt.Orientation.Horizontal)
        main_layout.addWidget(splitter)
        
        # --- Left Panel ---
        left_panel = QFrame()
        left_layout = QVBoxLayout(left_panel)
        left_layout.setAlignment(Qt.AlignmentFlag.AlignTop)
        
        # Provider
        left_layout.addWidget(QLabel("Provider:"))
        self.provider_combo = QComboBox()
        self.provider_combo.addItems(["Gemini", "OpenAI", "Claude"])
        self.provider_combo.currentTextChanged.connect(self.update_models)
        left_layout.addWidget(self.provider_combo)
        
        # Model
        left_layout.addWidget(QLabel("Model:"))
        self.model_combo = QComboBox()
        left_layout.addWidget(self.model_combo)
        self.update_models(self.provider_combo.currentText())
        
        # Settings Buttons
        self.settings_btn = QPushButton("Model Settings")
        self.settings_btn.clicked.connect(self.open_settings)
        left_layout.addWidget(self.settings_btn)
        
        self.sys_prompt_btn = QPushButton("System Instructions")
        self.sys_prompt_btn.clicked.connect(self.open_sys_prompt)
        left_layout.addWidget(self.sys_prompt_btn)
        
        left_layout.addSpacing(10)
        
        # File Selection
        self.select_file_btn = QPushButton("Select File(s)")
        self.select_file_btn.clicked.connect(self.select_files)
        left_layout.addWidget(self.select_file_btn)
        
        self.file_list = QListWidget()
        self.file_list.setFixedHeight(150)
        left_layout.addWidget(self.file_list)
        
        # Clear Buttons
        clear_files_btn = QPushButton("Clear Files")
        clear_files_btn.clicked.connect(self.clear_files)
        left_layout.addWidget(clear_files_btn)

        clear_chat_btn = QPushButton("Clear Chat")
        clear_chat_btn.clicked.connect(self.clear_chat)
        left_layout.addWidget(clear_chat_btn)
        
        left_layout.addSpacing(15)
        
        # --- CUSTOM CONTROLS ---
        left_layout.addWidget(QLabel("<b>Analysis Tools</b>"))
        
        # Analysis Type Dropdown
        self.analysis_type_combo = QComboBox()
        self.analysis_type_combo.addItems(ANALYSIS_TYPES)
        self.analysis_type_combo.currentTextChanged.connect(self.on_analysis_type_changed)
        left_layout.addWidget(self.analysis_type_combo)
        
        # Liability Type Dropdown (Hidden initially)
        self.liability_type_combo = QComboBox()
        self.liability_type_combo.addItems(LIABILITY_TYPES)
        self.liability_type_combo.setVisible(False)
        left_layout.addWidget(self.liability_type_combo)
        
        # Analyze Button (Hidden initially, replaces Send)
        self.analyze_btn = QPushButton("Analyze")
        self.analyze_btn.setStyleSheet("background-color: #673AB7; color: white; font-weight: bold;")
        self.analyze_btn.setVisible(False)
        self.analyze_btn.clicked.connect(self.perform_analysis)
        left_layout.addWidget(self.analyze_btn)
        
        # Stop Analysis Button
        self.stop_analysis_btn = QPushButton("Stop Analysis")
        self.stop_analysis_btn.setStyleSheet("background-color: #e57373; color: white; font-weight: bold;")
        self.stop_analysis_btn.setVisible(False) # Visible only when analyzing
        self.stop_analysis_btn.clicked.connect(self.stop_generation)
        left_layout.addWidget(self.stop_analysis_btn)
        
        left_layout.addStretch()
        splitter.addWidget(left_panel)
        
        # --- Right Panel (Chat) ---
        right_panel = QFrame()
        right_layout = QVBoxLayout(right_panel)
        
        self.chat_history = QTextBrowser()
        self.chat_history.setOpenExternalLinks(True)
        right_layout.addWidget(self.chat_history)
        
        input_layout = QHBoxLayout()
        self.chat_input = QPlainTextEdit()
        self.chat_input.setFixedHeight(80)
        self.chat_input.setPlaceholderText("Type a message... (Enter to send, Shift+Enter for newline)")
        self.chat_input.setAcceptDrops(True)
        self.chat_input.dragEnterEvent = self.dragEnterEvent
        self.chat_input.dropEvent = self.dropEvent
        self.chat_input.keyPressEvent = self.chat_key_press
        
        input_layout.addWidget(self.chat_input)
        
        self.send_btn = QPushButton("Send")
        self.send_btn.setFixedSize(60, 80)
        self.send_btn.clicked.connect(self.send_message)
        input_layout.addWidget(self.send_btn)
        
        self.stop_btn = QPushButton("Stop")
        self.stop_btn.setFixedSize(60, 80)
        self.stop_btn.clicked.connect(self.stop_generation)
        self.stop_btn.setEnabled(False)
        self.stop_btn.setStyleSheet("background-color: #e57373; color: white; font-weight: bold;")
        input_layout.addWidget(self.stop_btn)
        
        right_layout.addLayout(input_layout)
        splitter.addWidget(right_panel)
        splitter.setSizes([250, 950])

    def stop_generation(self):
        super().stop_generation() # Handles worker termination and chat buttons
        
        # Handle Analysis Buttons
        if self.analysis_type_combo.currentText() != "Select Analysis...":
            self.analyze_btn.setVisible(True)
            self.analyze_btn.setEnabled(True)
            self.stop_analysis_btn.setVisible(False)

    def on_analysis_type_changed(self, text):
        if text == "Liability Analysis":
            self.liability_type_combo.setVisible(True)
            self.analyze_btn.setVisible(True)
            self.send_btn.setEnabled(False) # Disable manual send while in analysis mode
            self.chat_input.setEnabled(False)
        elif text == "Exposure Analysis":
            self.liability_type_combo.setVisible(False)
            self.analyze_btn.setVisible(True)
            self.send_btn.setEnabled(False)
            self.chat_input.setEnabled(False)
        else:
            # Reset to normal chat
            self.liability_type_combo.setVisible(False)
            self.analyze_btn.setVisible(False)
            self.send_btn.setEnabled(True)
            self.chat_input.setEnabled(True)

    def perform_analysis(self):
        # 1. Validate inputs
        checked_count = 0
        if hasattr(self, 'file_list'):
            for i in range(self.file_list.count()):
                if self.file_list.item(i).checkState() == Qt.CheckState.Checked:
                    checked_count += 1
        
        if checked_count == 0:
            QMessageBox.warning(self, "No Files Selected", "Please select (check) at least one file to analyze.")
            return

        # 2. Determine Prompt Key
        a_type = self.analysis_type_combo.currentText()
        if a_type == "Liability Analysis":
            l_type = self.liability_type_combo.currentText()
            key = f"{a_type}|{l_type}"
            display_title = f"{l_type} Liability Analysis"
        else:
            key = a_type
            display_title = "Exposure Analysis"

        # 3. Load Prompt
        prompts = {}
        if os.path.exists(PROMPTS_FILE):
            try:
                with open(PROMPTS_FILE, 'r') as f:
                    prompts = json.load(f)
            except: pass
            
        user_prompt = prompts.get(key)
        
        # Fallback if not configured
        if not user_prompt:
             if "Liability" in key:
                 user_prompt = f"Please perform a {display_title}. Identify key arguments for and against liability based on the attached documents."
             else:
                 user_prompt = "Please perform an Exposure Analysis. Estimate the potential damages and financial exposure based on the attached documents."

        # 4. Execute (simulate user sending this prompt)
        # We manually update history to show what's happening
        self.chat_history.append(f"<b>System:</b> Running {display_title}...")
        if checked_count > 0:
            self.chat_history.append(f"<i>(Attached {checked_count} files)</i>")
        
        self.chat_history.append("<b>AI:</b> <i>Thinking...</i>")
        self.analyze_btn.setVisible(False)
        self.stop_analysis_btn.setVisible(True)
        
        # Prepare Content
        file_content = self.read_files_content()
        
        # Filter content for Liability Analysis
        if a_type == "Liability Analysis":
            try:
                # Find "EVALUATION OF LIABILITY" and "EVALUATION OF EXPOSURE" and remove text between
                # Using DOTALL so . matches newlines
                pattern = re.compile(r"(EVALUATION OF LIABILITY)(.*?)(EVALUATION OF EXPOSURE)", re.IGNORECASE | re.DOTALL)
                # Check if it exists before replacing to avoid unneeded modification or log it
                if pattern.search(file_content):
                    log_event("Filtering 'EVALUATION OF LIABILITY' section from content for analysis.")
                    file_content = pattern.sub(r"\1\n[PREVIOUS LIABILITY EVALUATION IGNORED FOR INDEPENDENT ANALYSIS]\n\3", file_content)
            except Exception as e:
                log_event(f"Error filtering liability content: {e}", "error")

        # Start Worker
        self.worker = LLMWorker(
            self.provider_combo.currentText(),
            self.model_combo.currentText(),
            self.system_prompt,
            user_prompt,
            file_content,
            self.settings,
            history=list(self.conversation_history)
        )
        
        # Update history
        full_msg = user_prompt
        if file_content:
            full_msg += "\n\n[ATTACHED FILES]:\n" + file_content
        self.conversation_history.append({'role': 'user', 'content': full_msg})

        self.worker.finished.connect(self.on_analysis_response)
        self.worker.error.connect(self.on_error)
        self.worker.start()

    def on_analysis_response(self, text):
        # Determine Title based on selection (store it in self? or just re-read)
        # Re-reading is safe since UI is blocked
        a_type = self.analysis_type_combo.currentText()
        title = "Analysis Result"
        if a_type == "Liability Analysis":
            title = f"{self.liability_type_combo.currentText()} Liability Analysis"
        elif a_type == "Exposure Analysis":
            title = "Exposure Analysis"

        # Remove "Thinking..."
        cursor = self.chat_history.textCursor()
        cursor.movePosition(QTextCursor.MoveOperation.End)
        cursor.select(QTextCursor.SelectionType.BlockUnderCursor)
        cursor.removeSelectedText()
        cursor.deletePreviousChar()
        
        # Format Output
        try:
            html_text = markdown.markdown(text, extensions=['fenced_code', 'tables'])
        except Exception:
            html_text = text.replace('\n', '<br>')

        self.chat_history.append(f"<b><u>{title}</u></b>")
        self.chat_history.insertHtml(html_text)
        self.chat_history.append("-" * 50)
        self.analyze_btn.setVisible(True)
        self.analyze_btn.setEnabled(True)
        self.stop_analysis_btn.setVisible(False)
        
        # Update history
        self.conversation_history.append({'role': 'assistant', 'content': text})

    def on_error(self, err):
        self.chat_history.append(f"<font color='red'>Error: {err}</font>")
        
        # Reset Analysis UI
        self.analyze_btn.setVisible(True)
        self.analyze_btn.setEnabled(True)
        self.stop_analysis_btn.setVisible(False)
        
        # Reset Chat UI (if applicable)
        self.send_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)
        
        # Ensure correct state based on mode
        if self.analysis_type_combo.currentText() != "Select Analysis...":
            self.send_btn.setEnabled(False) # Should be disabled in analysis mode
