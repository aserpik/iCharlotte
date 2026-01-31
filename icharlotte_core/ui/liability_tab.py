import os
import json
import re
import markdown
import datetime
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QSplitter, QFrame, QLabel,
    QComboBox, QPushButton, QListWidget, QListWidgetItem, QTextBrowser,
    QPlainTextEdit, QMessageBox, QDialog, QFormLayout, QDialogButtonBox, QTextEdit,
    QCheckBox, QScrollArea, QGroupBox
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QDragEnterEvent, QDropEvent, QTextCursor

from ..config import API_KEYS, GEMINI_DATA_DIR
from ..utils import log_event
from ..llm import LLMWorker, ModelFetcher
from .tabs import ChatTab
from .dialogs import SettingsDialog, SystemPromptDialog

# Import document registry and case data manager for report builder
import sys
sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), 'Scripts'))
try:
    from document_registry import DocumentRegistry, get_available_documents
    from case_data_manager import CaseDataManager
except ImportError:
    DocumentRegistry = None
    CaseDataManager = None

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


class ReportBuilderDialog(QDialog):
    """Dialog for selecting documents to include in an investigation report."""

    # Standard intro phrases for different document types
    INTRO_PHRASES = {
        "default": "We have received and reviewed the {doc_name}, the most pertinent portions of which are summarized below:",
        "Deposition": "We have received and reviewed the {doc_name}. The most pertinent portions of the testimony are summarized below:",
        "Medical Records": "We have received and reviewed {doc_name}. The most pertinent portions are summarized below:",
        "Traffic Collision Report": "We have received and reviewed the Traffic Collision Report prepared by {agency}. The most pertinent portions are summarized below:",
        "Police Report": "We have received and reviewed the Police Report. The most pertinent portions are summarized below:",
        "Investigation Report": "We have received and reviewed an investigation report, the most pertinent portions of which are summarized below:",
        "Form Interrogatories": "We have received and reviewed {party}'s Responses to Form Interrogatories. The most pertinent responses are summarized below:",
        "Special Interrogatories": "We have received and reviewed {party}'s Responses to Special Interrogatories. The most pertinent responses are summarized below:",
        "Request for Admissions": "We have received and reviewed {party}'s Responses to Request for Admissions. The most pertinent responses are summarized below:",
        "Request for Production": "We have received and reviewed {party}'s Responses to Request for Production of Documents. The most pertinent responses are summarized below:",
        "Complaint": "We have received and reviewed the Complaint filed in this matter. The most pertinent allegations are summarized below:",
        "Answer": "We have received and reviewed the Answer filed in this matter. The most pertinent portions are summarized below:",
        "Expert Report": "We have received and reviewed an expert report prepared by {expert}. The most pertinent portions are summarized below:",
        "ISO ClaimSearch Report": "We have received and reviewed an ISO ClaimSearch Report. The most pertinent findings are summarized below:",
    }

    def __init__(self, file_number, parent=None):
        super().__init__(parent)
        self.file_number = file_number
        self.selected_documents = []
        self.setWindowTitle(f"Build Investigation Report - {file_number}")
        self.resize(600, 500)
        self.setup_ui()
        self.load_documents()

    def setup_ui(self):
        layout = QVBoxLayout(self)

        # Instructions
        instructions = QLabel(
            "Select the documents to include in the Investigation section of your report.\n"
            "Documents will be added in the order shown, with letter prefixes (A, B, C, etc.)."
        )
        instructions.setWordWrap(True)
        layout.addWidget(instructions)

        # Scrollable area for document checkboxes
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setMinimumHeight(300)

        self.docs_container = QWidget()
        self.docs_layout = QVBoxLayout(self.docs_container)
        self.docs_layout.setAlignment(Qt.AlignmentFlag.AlignTop)

        scroll.setWidget(self.docs_container)
        layout.addWidget(scroll)

        # Select All / Deselect All buttons
        btn_row = QHBoxLayout()
        select_all_btn = QPushButton("Select All")
        select_all_btn.clicked.connect(self.select_all)
        deselect_all_btn = QPushButton("Deselect All")
        deselect_all_btn.clicked.connect(self.deselect_all)
        btn_row.addWidget(select_all_btn)
        btn_row.addWidget(deselect_all_btn)
        btn_row.addStretch()
        layout.addLayout(btn_row)

        # Status label
        self.status_label = QLabel("")
        layout.addWidget(self.status_label)

        # Dialog buttons
        button_box = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)

        self.checkboxes = []

    def load_documents(self):
        """Load available documents from the registry and case data."""
        docs = []

        # Try loading from document registry first
        if DocumentRegistry:
            try:
                registry_docs = get_available_documents(self.file_number)
                if registry_docs:
                    docs.extend(registry_docs)
            except Exception as e:
                log_event(f"Error loading from document registry: {e}", "warning")

        # Also scan case data for summary_* keys as fallback
        if CaseDataManager:
            try:
                data_manager = CaseDataManager()
                all_vars = data_manager.get_all_variables(self.file_number, flatten=False)

                # Find all summary keys not already in docs
                existing_names = {d.get('name', '').lower() for d in docs}

                # Look for various summary key patterns
                summary_prefixes = ['summary_', 'discovery_summary_', 'deposition_summary_',
                                   'medical_summary_', 'complaint_summary_']

                for key, val_obj in all_vars.items():
                    matched_prefix = None
                    for prefix in summary_prefixes:
                        if key.startswith(prefix):
                            matched_prefix = prefix
                            break

                    if matched_prefix:
                        # Extract document name from key
                        doc_name = key[len(matched_prefix):].replace('_', ' ').title()

                        # Determine document type based on prefix
                        if 'discovery' in matched_prefix:
                            doc_type = 'Discovery'
                        elif 'deposition' in matched_prefix:
                            doc_type = 'Deposition'
                        elif 'medical' in matched_prefix:
                            doc_type = 'Medical Records'
                        elif 'complaint' in matched_prefix:
                            doc_type = 'Complaint'
                        else:
                            doc_type = 'Other'

                        if doc_name.lower() not in existing_names:
                            timestamp = val_obj.get('timestamp', '') if isinstance(val_obj, dict) else ''
                            docs.append({
                                'name': doc_name,
                                'document_type': doc_type,
                                'timestamp': timestamp,
                                'summary_key': key  # Store the actual key for retrieval
                            })

            except Exception as e:
                log_event(f"Error scanning case data for summaries: {e}", "warning")

        if not docs:
            self.status_label.setText(
                "No summarized documents found for this case.\n"
                "Run the Summarize agent on documents first."
            )
            return

        # Group by document type
        by_type = {}
        for doc in docs:
            doc_type = doc.get('document_type', 'Other')
            if doc_type not in by_type:
                by_type[doc_type] = []
            by_type[doc_type].append(doc)

        # Create checkboxes grouped by type
        for doc_type in sorted(by_type.keys()):
            group = QGroupBox(doc_type)
            group_layout = QVBoxLayout(group)

            for doc in by_type[doc_type]:
                name = doc.get('name', 'Unknown')
                timestamp = doc.get('timestamp', '')
                if timestamp:
                    try:
                        dt = datetime.datetime.fromisoformat(timestamp)
                        timestamp = dt.strftime('%m/%d/%Y')
                    except:
                        pass

                cb = QCheckBox(f"{name}")
                if timestamp:
                    cb.setToolTip(f"Summarized: {timestamp}")
                cb.setProperty('doc_data', doc)
                self.checkboxes.append(cb)
                group_layout.addWidget(cb)

            self.docs_layout.addWidget(group)

        self.status_label.setText(f"Found {len(docs)} summarized documents")

    def select_all(self):
        for cb in self.checkboxes:
            cb.setChecked(True)

    def deselect_all(self):
        for cb in self.checkboxes:
            cb.setChecked(False)

    def get_selected_documents(self):
        """Return list of selected document data dicts."""
        selected = []
        for cb in self.checkboxes:
            if cb.isChecked():
                doc_data = cb.property('doc_data')
                if doc_data:
                    selected.append(doc_data)
        return selected

    @classmethod
    def get_intro_phrase(cls, doc_type, doc_name):
        """Get the appropriate intro phrase for a document type."""
        # Check for specific matches
        for key in cls.INTRO_PHRASES:
            if key != "default" and key.lower() in doc_type.lower():
                phrase = cls.INTRO_PHRASES[key]
                # Replace placeholders with generic values
                phrase = phrase.replace("{doc_name}", doc_name)
                phrase = phrase.replace("{party}", "the party")
                phrase = phrase.replace("{agency}", "the investigating agency")
                phrase = phrase.replace("{expert}", "the expert")
                return phrase

        # Default phrase
        return cls.INTRO_PHRASES["default"].replace("{doc_name}", doc_name)


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

        left_layout.addSpacing(15)

        # --- REPORT BUILDER ---
        left_layout.addWidget(QLabel("<b>Report Builder</b>"))

        self.build_report_btn = QPushButton("Build Investigation Report")
        self.build_report_btn.setStyleSheet("background-color: #1565C0; color: white; font-weight: bold;")
        self.build_report_btn.clicked.connect(self.open_report_builder)
        left_layout.addWidget(self.build_report_btn)

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

    def get_file_number(self):
        """Get the current file number from the main window."""
        main_win = self.window()
        if main_win and hasattr(main_win, 'file_number'):
            return main_win.file_number
        return None

    def open_report_builder(self):
        """Open the report builder dialog."""
        file_num = self.get_file_number()
        if not file_num:
            QMessageBox.warning(
                self,
                "No Case Loaded",
                "Please open a case first before building a report."
            )
            return

        if not DocumentRegistry or not CaseDataManager:
            QMessageBox.warning(
                self,
                "Module Not Available",
                "The document registry module is not available. Please check your installation."
            )
            return

        dialog = ReportBuilderDialog(file_num, self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            selected_docs = dialog.get_selected_documents()
            if selected_docs:
                self.generate_investigation_report(selected_docs)
            else:
                QMessageBox.information(
                    self,
                    "No Documents Selected",
                    "Please select at least one document to include in the report."
                )

    def generate_investigation_report(self, selected_docs):
        """Generate the investigation report HTML and display it."""
        file_num = self.get_file_number()
        if not file_num:
            return

        try:
            data_manager = CaseDataManager()
        except Exception as e:
            self.chat_history.append(f"<font color='red'>Error initializing case data manager: {e}</font>")
            return

        # Build the report HTML
        html_parts = []

        # Header - centered, bold, underlined
        html_parts.append("<p align='center'><b><u>INVESTIGATION</u></b></p>")
        html_parts.append("<p><br></p>")

        # Intro line
        html_parts.append("<p>Below, please find a summary of our investigation to date.</p>")
        html_parts.append("<p><br></p>")

        # Letter sequence for subsections
        letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"

        for idx, doc in enumerate(selected_docs):
            letter = letters[idx] if idx < len(letters) else f"A{idx - 25}"
            doc_name = doc.get('name', 'Unknown Document')
            doc_type = doc.get('document_type', 'Other')

            # Get the summary from case data
            summary_text = None

            # First, try the explicit summary_key if provided (from case data scan)
            if doc.get('summary_key'):
                summary_text = data_manager.get_value(file_num, doc['summary_key'])

            # Try using the document name
            if not summary_text:
                clean_var_name = re.sub(r"[^a-zA-Z0-9_]", "_", doc_name.lower())
                summary_key = f"summary_{clean_var_name}"
                summary_text = data_manager.get_value(file_num, summary_key)

            # Try alternate key formats using source path
            if not summary_text:
                source_path = doc.get('source_path', '')
                if source_path:
                    base_name = os.path.splitext(os.path.basename(source_path))[0]
                    alt_key = f"summary_{re.sub(r'[^a-zA-Z0-9_]', '_', base_name.lower())}"
                    summary_text = data_manager.get_value(file_num, alt_key)

            if not summary_text:
                summary_text = "[Summary not found in case data]"

            # Format the subsection header - Letter bold only, title bold+underlined
            html_parts.append(f"<p><b>{letter}.</b>&nbsp;&nbsp;&nbsp;<b><u>{doc_name}</u></b></p>")

            # Check if summary already starts with an intro phrase
            summary_lower = summary_text.lower().strip()
            has_intro = (
                summary_lower.startswith('we have received') or
                summary_lower.startswith('we received') or
                summary_lower.startswith('this document') or
                summary_lower.startswith('the document')
            )

            # Add intro phrase only if the summary doesn't already have one
            if not has_intro:
                intro = ReportBuilderDialog.get_intro_phrase(doc_type, doc_name)
                html_parts.append(f"<p>{intro}</p>")

            # Clean and format the summary text for narrative style
            formatted_summary = self._format_narrative_summary(summary_text)
            html_parts.append(formatted_summary)

            # Add blank line between sections
            html_parts.append("<p><br></p>")

        # Combine and display
        full_html = "\n".join(html_parts)

        # Clear and display
        self.chat_history.clear()
        self.chat_history.append(f"<p><i>Generated Investigation Report with {len(selected_docs)} document(s)</i></p>")
        self.chat_history.append("<hr>")
        self.chat_history.insertHtml(full_html)
        self.chat_history.append("<hr>")
        self.chat_history.append("<p><i>Tip: Select and copy (Ctrl+C) the text above to paste into your Word document.</i></p>")

        log_event(f"Generated investigation report with {len(selected_docs)} documents for case {file_num}")

    def _format_narrative_summary(self, text):
        """
        Format summary text into clean narrative HTML.
        - Strip markdown formatting (bold, headers, etc.)
        - First line of each paragraph indented 0.5 inches (using tab)
        - Bullet lists indented 0.5 inches (only when necessary)
        - No bold text, no sub-headers
        """
        if not text:
            return ""

        # Use Unicode non-breaking spaces for indentation
        INDENT = "\u00A0\u00A0\u00A0\u00A0\u00A0\u00A0\u00A0\u00A0"  # 8 non-breaking spaces

        # Strip markdown formatting
        # Remove headers (# ## ### etc.)
        text = re.sub(r'^#{1,6}\s*', '', text, flags=re.MULTILINE)

        # Remove bold (**text** or __text__)
        text = re.sub(r'\*\*([^*]+)\*\*', r'\1', text)
        text = re.sub(r'__([^_]+)__', r'\1', text)

        # Remove italic (*text* or _text_) - be careful not to remove bullet points
        text = re.sub(r'(?<!\*)\*([^*]+)\*(?!\*)', r'\1', text)
        text = re.sub(r'(?<!_)_([^_]+)_(?!_)', r'\1', text)

        # Remove inline code (`text`)
        text = re.sub(r'`([^`]+)`', r'\1', text)

        # Split into paragraphs (double newline separated)
        paragraphs = re.split(r'\n\s*\n', text.strip())

        html_parts = []

        for para in paragraphs:
            para = para.strip()
            if not para:
                continue

            # Check if this paragraph is a bullet list
            lines = para.split('\n')
            is_bullet_list = all(
                re.match(r'^\s*[\*\-\•]\s+', line.strip()) or not line.strip()
                for line in lines if line.strip()
            )

            if is_bullet_list:
                # Format as bullet list with indent
                for line in lines:
                    line = line.strip()
                    if not line:
                        continue
                    # Remove the bullet marker and get content
                    content = re.sub(r'^[\*\-\•]\s+', '', line)
                    html_parts.append(f"<p>{INDENT}•&nbsp;&nbsp;{content}</p>")
            else:
                # Regular paragraph - join lines and indent first line
                clean_lines = []
                for line in lines:
                    line = line.strip()
                    if line:
                        # Check if line starts with a label like "Item 1:" - keep as separate paragraph
                        if re.match(r'^(Item\s+\d+|[A-Z][a-z]+\s+\d+):\s*', line):
                            if clean_lines:
                                # Output accumulated lines first
                                para_text = ' '.join(clean_lines)
                                html_parts.append(f"<p>{INDENT}{para_text}</p>")
                                clean_lines = []
                            # Add labeled item as its own paragraph
                            html_parts.append(f"<p>{INDENT}{line}</p>")
                        else:
                            clean_lines.append(line)

                if clean_lines:
                    para_text = ' '.join(clean_lines)
                    html_parts.append(f"<p>{INDENT}{para_text}</p>")

        return '\n'.join(html_parts)
