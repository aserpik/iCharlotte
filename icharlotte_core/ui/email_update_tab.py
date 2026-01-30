import os
import datetime
import re
import json
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
    QTextEdit, QListWidget, QListWidgetItem, QLabel, QMessageBox, QSplitter,
    QCheckBox, QFileDialog, QProgressBar, QFrame, QAbstractItemView
)
from PySide6.QtCore import Qt, QThread, Signal, QMimeData, QSize, QTimer
from PySide6.QtGui import QDragEnterEvent, QDropEvent

try:
    import win32com.client
    import pythoncom
    import win32gui
except ImportError:
    win32com = None
    pythoncom = None
    win32gui = None

import markdown
from ..llm import LLMWorker
from ..utils import extract_text_from_file, get_case_path, log_event
from ..config import GEMINI_DATA_DIR

# --- Custom Widgets ---

class DragDropListWidget(QListWidget):
    fileDropped = Signal(list)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.setDragDropMode(QAbstractItemView.DragDropMode.DropOnly)
        self.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event: QDropEvent):
        urls = event.mimeData().urls()
        files = [u.toLocalFile() for u in urls if u.isLocalFile()]
        if files:
            self.fileDropped.emit(files)

class EmailSenderWorker(QThread):
    finished = Signal()
    error = Signal(str)

    def __init__(self, recipient, cc, subject, body_html):
        super().__init__()
        self.recipient = recipient
        self.cc = cc
        self.subject = subject
        self.body_html = body_html

    def run(self):
        try:
            log_event("EmailWorker: Initializing COM...")
            if not win32com or not pythoncom:
                raise ImportError("pywin32 module not found.")

            pythoncom.CoInitialize()
            outlook = win32com.client.Dispatch('Outlook.Application')
            mail = outlook.CreateItem(0) # 0 = olMailItem
            
            log_event(f"EmailWorker: Creating mail to {self.recipient}")
            mail.To = self.recipient if self.recipient else ""
            if self.cc:
                mail.CC = self.cc
            mail.Subject = self.subject if self.subject else ""
            
            # Use the provided HTML as the complete body since it already has the signature
            mail.HTMLBody = self.body_html
            
            log_event("EmailWorker: Displaying mail...")
            mail.Display() 
            
            # Force Foreground
            try:
                inspector = mail.GetInspector
                inspector.Activate()
                if win32gui:
                    hwnd = win32gui.FindWindow(None, inspector.Caption)
                    if hwnd:
                        win32gui.ShowWindow(hwnd, 5) # SW_SHOW
                        win32gui.SetForegroundWindow(hwnd)
            except Exception as fe:
                log_event(f"EmailWorker: Foreground failed: {fe}", "warning")

            pythoncom.CoUninitialize()
            self.finished.emit()
            log_event("EmailWorker: Success")
        except Exception as e:
            log_event(f"EmailWorker: ERROR: {e}", "error")
            self.error.emit(str(e))

class EmailUpdateTab(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.current_case_number = None
        self._save_timer = None  # Debounce timer for auto-save
        self._loading_state = False  # Flag to prevent saves during load
        self.load_prompts()
        self.setup_ui()
        self._setup_persistence()

    def load_prompts(self):
        # Load custom prompts from JSON, or use defaults
        self.prompts = {
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
        
        json_path = os.path.join(GEMINI_DATA_DIR, "email_update_prompts.json")
        if os.path.exists(json_path):
            try:
                with open(json_path, 'r', encoding='utf-8') as f:
                    saved = json.load(f)
                    self.prompts.update(saved)
            except Exception as e:
                log_event(f"Error loading prompt config: {e}", "warning")

    def save_prompts(self):
        json_path = os.path.join(GEMINI_DATA_DIR, "email_update_prompts.json")
        try:
            if not os.path.exists(GEMINI_DATA_DIR):
                os.makedirs(GEMINI_DATA_DIR)
            with open(json_path, 'w', encoding='utf-8') as f:
                json.dump(self.prompts, f, indent=4)
        except Exception as e:
            log_event(f"Error saving prompt config: {e}", "error")

    def setup_ui(self):
        main_layout = QVBoxLayout(self)
        
        # Splitter: Left | Middle | Right
        self.main_splitter = QSplitter(Qt.Orientation.Horizontal)
        main_layout.addWidget(self.main_splitter)

        # --- LEFT SECTION (Files) ---
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        
        # Header
        left_layout.addWidget(QLabel("<b>Reference Files</b> (Drag & Drop)"))
        
        # Buttons: Select/Unselect All
        sel_btn_layout = QHBoxLayout()
        self.btn_sel_all = QPushButton("Select All")
        self.btn_sel_all.clicked.connect(self.select_all_files)
        self.btn_unsel_all = QPushButton("Unselect All")
        self.btn_unsel_all.clicked.connect(self.unselect_all_files)
        sel_btn_layout.addWidget(self.btn_sel_all)
        sel_btn_layout.addWidget(self.btn_unsel_all)
        left_layout.addLayout(sel_btn_layout)

        # File List
        self.file_list = DragDropListWidget()
        self.file_list.fileDropped.connect(self.add_files)
        left_layout.addWidget(self.file_list)
        
        # Load Button
        self.btn_load_files = QPushButton("Load Files")
        self.btn_load_files.clicked.connect(self.load_files_dialog)
        left_layout.addWidget(self.btn_load_files)
        
        left_layout.addSpacing(10)
        
        # Generation Buttons
        self.btn_gen_topic = QPushButton("Generate New Topic")
        self.btn_gen_topic.clicked.connect(lambda: self.generate_content("topic"))
        left_layout.addWidget(self.btn_gen_topic)
        
        self.btn_gen_handling = QPushButton("Generate Further Case Handling")
        self.btn_gen_handling.clicked.connect(lambda: self.generate_content("handling"))
        left_layout.addWidget(self.btn_gen_handling)
        
        self.btn_gen_both = QPushButton("Generate Both")
        self.btn_gen_both.clicked.connect(lambda: self.generate_content("both"))
        left_layout.addWidget(self.btn_gen_both)
        
        self.main_splitter.addWidget(left_widget)

        # --- MIDDLE SECTION (Context & Output) ---
        middle_widget = QWidget()
        middle_layout = QVBoxLayout(middle_widget)
        
        middle_splitter = QSplitter(Qt.Orientation.Vertical)
        
        # 1. New Topic Context
        topic_ctx_widget = QWidget()
        t_layout = QVBoxLayout(topic_ctx_widget)
        t_layout.setContentsMargins(0,0,0,0)
        t_layout.addWidget(QLabel("Context for New Topic:"))
        self.ctx_topic = QTextEdit()
        t_layout.addWidget(self.ctx_topic)
        middle_splitter.addWidget(topic_ctx_widget)
        
        # 2. Handling Context
        handling_ctx_widget = QWidget()
        h_layout = QVBoxLayout(handling_ctx_widget)
        h_layout.setContentsMargins(0,0,0,0)
        h_layout.addWidget(QLabel("Context for Further Case Handling:"))
        self.ctx_handling = QTextEdit()
        h_layout.addWidget(self.ctx_handling)
        middle_splitter.addWidget(handling_ctx_widget)
        
        # 3. Topic Output
        topic_out_widget = QWidget()
        to_layout = QVBoxLayout(topic_out_widget)
        to_layout.setContentsMargins(0,0,0,0)
        to_layout.addWidget(QLabel("Generated Topic Output:"))
        self.txt_topic_output = QTextEdit()
        self.txt_topic_output.setReadOnly(False) # Allow user to refine
        to_layout.addWidget(self.txt_topic_output)
        middle_splitter.addWidget(topic_out_widget)

        # 4. Handling Output
        handling_out_widget = QWidget()
        ho_layout = QVBoxLayout(handling_out_widget)
        ho_layout.setContentsMargins(0,0,0,0)
        ho_layout.addWidget(QLabel("Generated Handling Output:"))
        self.txt_handling_output = QTextEdit()
        self.txt_handling_output.setReadOnly(False) # Allow user to refine
        ho_layout.addWidget(self.txt_handling_output)
        middle_splitter.addWidget(handling_out_widget)
        
        middle_layout.addWidget(middle_splitter)
        self.main_splitter.addWidget(middle_widget)

        # --- RIGHT SECTION (Status Reports) ---
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        
        # Status Update Button (Moved from Email Tab)
        self.btn_status_update = QPushButton("Status Update")
        self.btn_status_update.setStyleSheet("""
            QPushButton { background-color: #0078D4; color: white; padding: 10px; font-weight: bold; font-size: 14px; }
            QPushButton:hover { background-color: #106EBE; }
        """)
        self.btn_status_update.clicked.connect(self.create_status_update)
        right_layout.addWidget(self.btn_status_update)
        
        right_layout.addWidget(QLabel("<b>Sent Status Reports</b>"))
        self.sent_list = QListWidget()
        self.sent_list.itemDoubleClicked.connect(self.open_sent_report)
        right_layout.addWidget(self.sent_list)
        
        self.main_splitter.addWidget(right_widget)
        
        # Set Sizes (Equal initially)
        self.main_splitter.setSizes([300, 400, 300])

        # Status Bar
        self.status_bar = QLabel("")
        main_layout.addWidget(self.status_bar)

    # --- Persistence Methods ---

    def _setup_persistence(self):
        """Connect signals for auto-saving state changes."""
        # Create debounce timer
        self._save_timer = QTimer(self)
        self._save_timer.setSingleShot(True)
        self._save_timer.timeout.connect(self._do_save_state)

        # Connect text editors to trigger save on change
        self.ctx_topic.textChanged.connect(self._schedule_save)
        self.ctx_handling.textChanged.connect(self._schedule_save)
        self.txt_topic_output.textChanged.connect(self._schedule_save)
        self.txt_handling_output.textChanged.connect(self._schedule_save)

        # Connect file list changes
        self.file_list.model().rowsInserted.connect(self._schedule_save)
        self.file_list.model().rowsRemoved.connect(self._schedule_save)
        self.file_list.itemChanged.connect(self._schedule_save)

    def _schedule_save(self, *args):
        """Schedule a debounced save (500ms delay)."""
        if self._loading_state:
            return
        if self._save_timer:
            self._save_timer.start(500)  # 500ms debounce

    def _do_save_state(self):
        """Perform the actual save."""
        self.save_state()

    def get_state_path(self):
        """Get the path for the current case's state file."""
        file_num = self.get_file_number()
        if not file_num:
            return None
        return os.path.join(GEMINI_DATA_DIR, f"{file_num}_email_update_state.json")

    def save_state(self):
        """Save all persistent data for the current case."""
        state_path = self.get_state_path()
        if not state_path:
            return

        try:
            # Collect reference files with their checked states
            files_data = []
            for i in range(self.file_list.count()):
                item = self.file_list.item(i)
                files_data.append({
                    'path': item.data(Qt.ItemDataRole.UserRole),
                    'checked': item.checkState() == Qt.CheckState.Checked
                })

            # Collect text editor contents
            state = {
                'reference_files': files_data,
                'ctx_topic': self.ctx_topic.toPlainText(),
                'ctx_handling': self.ctx_handling.toPlainText(),
                'txt_topic_output': self.txt_topic_output.toPlainText(),
                'txt_handling_output': self.txt_handling_output.toPlainText(),
                'sent_reports': []
            }

            # Collect sent reports
            for i in range(self.sent_list.count()):
                item = self.sent_list.item(i)
                state['sent_reports'].append({
                    'text': item.text(),
                    'entry_id': item.data(Qt.ItemDataRole.UserRole)
                })

            # Ensure directory exists
            if not os.path.exists(GEMINI_DATA_DIR):
                os.makedirs(GEMINI_DATA_DIR)

            # Write state file
            with open(state_path, 'w', encoding='utf-8') as f:
                json.dump(state, f, indent=2)

            log_event(f"EmailUpdateTab: Saved state to {state_path}")
        except Exception as e:
            log_event(f"EmailUpdateTab: Error saving state: {e}", "error")

    def load_state(self):
        """Load state for the current case."""
        state_path = self.get_state_path()
        if not state_path or not os.path.exists(state_path):
            return

        try:
            self._loading_state = True  # Prevent save triggers during load

            with open(state_path, 'r', encoding='utf-8') as f:
                state = json.load(f)

            # Clear existing data
            self.file_list.clear()
            self.ctx_topic.clear()
            self.ctx_handling.clear()
            self.txt_topic_output.clear()
            self.txt_handling_output.clear()
            self.sent_list.clear()

            # Restore reference files
            for file_data in state.get('reference_files', []):
                path = file_data.get('path')
                if path and os.path.exists(path):
                    item = QListWidgetItem(os.path.basename(path))
                    item.setData(Qt.ItemDataRole.UserRole, path)
                    item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
                    if file_data.get('checked', True):
                        item.setCheckState(Qt.CheckState.Checked)
                    else:
                        item.setCheckState(Qt.CheckState.Unchecked)
                    self.file_list.addItem(item)

            # Restore text editors
            self.ctx_topic.setPlainText(state.get('ctx_topic', ''))
            self.ctx_handling.setPlainText(state.get('ctx_handling', ''))
            self.txt_topic_output.setPlainText(state.get('txt_topic_output', ''))
            self.txt_handling_output.setPlainText(state.get('txt_handling_output', ''))

            # Restore sent reports
            for report in state.get('sent_reports', []):
                item = QListWidgetItem(report.get('text', ''))
                item.setData(Qt.ItemDataRole.UserRole, report.get('entry_id'))
                self.sent_list.addItem(item)

            log_event(f"EmailUpdateTab: Loaded state from {state_path}")
        except Exception as e:
            log_event(f"EmailUpdateTab: Error loading state: {e}", "error")
        finally:
            self._loading_state = False

    def on_case_changed(self, file_number):
        """Called when the active case changes. Saves old state and loads new."""
        # Save current state before switching
        if self.current_case_number:
            self.save_state()

        # Clear all widgets before loading new state
        self.file_list.clear()
        self.ctx_topic.clear()
        self.ctx_handling.clear()
        self.txt_topic_output.clear()
        self.txt_handling_output.clear()
        self.sent_list.clear()

        # Update current case number
        self.current_case_number = file_number

        # Load state for new case
        if file_number:
            self.load_state()

    # --- File Management ---

    def add_files(self, file_paths):
        for path in file_paths:
            # Check duplicates
            exists = False
            for i in range(self.file_list.count()):
                item = self.file_list.item(i)
                if item.data(Qt.ItemDataRole.UserRole) == path:
                    exists = True
                    break
            
            if not exists:
                item = QListWidgetItem(os.path.basename(path))
                item.setData(Qt.ItemDataRole.UserRole, path)
                item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
                item.setCheckState(Qt.CheckState.Checked)
                self.file_list.addItem(item)

    def load_files_dialog(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, "Select Files", "", 
            "Documents (*.pdf *.docx *.txt *.md *.msg);;All Files (*)"
        )
        if files:
            self.add_files(files)

    def select_all_files(self):
        for i in range(self.file_list.count()):
            self.file_list.item(i).setCheckState(Qt.CheckState.Checked)

    def unselect_all_files(self):
        for i in range(self.file_list.count()):
            self.file_list.item(i).setCheckState(Qt.CheckState.Unchecked)

    def get_selected_files_content(self):
        content = ""
        for i in range(self.file_list.count()):
            item = self.file_list.item(i)
            if item.checkState() == Qt.CheckState.Checked:
                path = item.data(Qt.ItemDataRole.UserRole)
                text = extract_text_from_file(path)
                if text:
                    content += f"\n--- Document: {os.path.basename(path)} ---\n{text}\n" 
        return content

    # --- Generation Logic ---

    def generate_content(self, mode):
        # 1. Gather Content
        file_content = self.get_selected_files_content()
        topic_ctx = self.ctx_topic.toPlainText().strip()
        handling_ctx = self.ctx_handling.toPlainText().strip()
        
        if not file_content and not topic_ctx and not handling_ctx:
            QMessageBox.warning(self, "Warning", "Please provide at least one file or context input.")
            return

        # 2. Read Style Guide
        style_path = os.path.join(os.getcwd(), "Scripts", "status_update_examples.txt")
        if os.path.exists(style_path):
            with open(style_path, 'r', encoding='utf-8') as f:
                style_guide = f.read()
        else:
            style_guide = "No style guide found."

        # 3. Build Prompt using loaded prompts
        system_prompt = self.prompts.get('system_prompt', '')
        topic_inst = self.prompts.get('topic_instruction', '')
        handling_inst = self.prompts.get('handling_instruction', '')
        
        user_prompt = f"STYLE EXAMPLES:\n{style_guide}\n\nCURRENT CASE DATA:\n{file_content}\n"
        
        if mode == "topic" or mode == "both":
            user_prompt += f"\nUSER CONTEXT FOR NEW TOPIC:\n{topic_ctx}\n"
            user_prompt += f"\n{topic_inst}\n"
        
        if mode == "handling" or mode == "both":
            user_prompt += f"\nUSER CONTEXT FOR FURTHER HANDLING:\n{handling_ctx}\n"
            user_prompt += f"\n{handling_inst}\n"
            
        user_prompt += "\nFORMAT OUTPUT AS:\n"
        if mode == "topic" or mode == "both":
            user_prompt += "TOPIC NAME: [Name]\nTOPIC CONTENT: [Paragraph]\n"
        if mode == "handling" or mode == "both":
            user_prompt += "FURTHER HANDLING: [Paragraph]\n"

        user_prompt += "\nSTRICT REQUIREMENT: DO NOT include any salutations, subject lines, or 'Factual Background' unless specifically requested in the context above. Output ONLY the snippets requested."

        self.status_bar.setText("Generating content with Gemini...")
        self.set_buttons_enabled(False)
        
        settings = {'temperature': 1.0, 'max_tokens': -1, 'stream': False}
        self.worker = LLMWorker("Gemini", "gemini-3-flash-preview", system_prompt, user_prompt, "", settings)
        self.worker.finished.connect(self.on_gen_finished)
        self.worker.error.connect(self.on_gen_error)
        self.worker.start()

    def on_gen_finished(self, text):
        log_event("StatusUpdate: Generation complete, parsing for UI...")
        
        # 1. Parse content for Topic
        t_match = re.search(r"TOPIC NAME:\s*(.*?)(?=TOPIC CONTENT:|FURTHER HANDLING:|$)", text, re.DOTALL | re.IGNORECASE)
        c_match = re.search(r"TOPIC CONTENT:\s*(.*?)(?=FURTHER HANDLING:|TOPIC NAME:|$)", text, re.DOTALL | re.IGNORECASE)
        
        if t_match and c_match:
            name = t_match.group(1).strip()
            content = c_match.group(1).strip()
            self.txt_topic_output.setMarkdown(f"**{name}:** {content}")
        elif t_match: # Fallback if only name
            self.txt_topic_output.setMarkdown(f"**{t_match.group(1).strip()}**")
        elif c_match: # Fallback if only content
            self.txt_topic_output.setMarkdown(c_match.group(1).strip())
            
        # 2. Parse content for Handling
        h_match = re.search(r"FURTHER HANDLING:\s*(.*)", text, re.DOTALL | re.IGNORECASE)
        if h_match:
            handling = h_match.group(1).strip()
            self.txt_handling_output.setMarkdown(f"**Further Case Handling:** {handling}")

        self.status_bar.setText("Generation complete.")
        self.set_buttons_enabled(True)

    def on_gen_error(self, err):
        QMessageBox.critical(self, "Error", str(err))
        self.status_bar.setText("Generation failed.")
        self.set_buttons_enabled(True)

    def set_buttons_enabled(self, enabled):
        self.btn_gen_topic.setEnabled(enabled)
        self.btn_gen_handling.setEnabled(enabled)
        self.btn_gen_both.setEnabled(enabled)
        self.btn_status_update.setEnabled(enabled)

    # --- Status Update Email Logic ---

    def create_status_update(self):
        try:
            # 1. Parse Output (Get RAW text from the specific output boxes)
            topic_text = self.txt_topic_output.toPlainText()
            handling_text = self.txt_handling_output.toPlainText()
            
            log_event(f"StatusUpdate: Starting generation for Topic({len(topic_text)}) and Handling({len(handling_text)})")
            
            self.parsed_topic_name = "New Topic"
            self.parsed_topic_content = ""
            self.parsed_handling_content = ""
            
            # Simple parsing for Topic
            # Note: We added the bold markers in on_gen_finished, so we parse those or the raw labels
            t_match = re.search(r"TOPIC NAME:\s*(.*?)(?=TOPIC CONTENT:|$)", topic_text, re.DOTALL | re.IGNORECASE)
            if t_match: 
                self.parsed_topic_name = t_match.group(1).strip()
            else:
                # Fallback: If labels missing, look for our formatted "**Name:** Content"
                parts = topic_text.split(":", 1)
                if len(parts) > 1:
                    self.parsed_topic_name = parts[0].replace("**", "").strip()
                    self.parsed_topic_content = parts[1].strip()

            c_match = re.search(r"TOPIC CONTENT:\s*(.*)", topic_text, re.DOTALL | re.IGNORECASE)
            if c_match: self.parsed_topic_content = c_match.group(1).strip()
            
            # Simple parsing for Handling
            h_match = re.search(r"FURTHER HANDLING:\s*(.*)", handling_text, re.DOTALL | re.IGNORECASE)
            if h_match: 
                self.parsed_handling_content = h_match.group(1).strip()
            else:
                # Fallback for handling
                self.parsed_handling_content = handling_text.replace("**Further Case Handling:**", "").strip()
            
            log_event(f"StatusUpdate: Parsed Topic='{self.parsed_topic_name}', Handling Len={len(self.parsed_handling_content)}")

            # 2. Load Case Vars
            file_num = self.get_file_number()
            log_event(f"StatusUpdate: Identified file number: {file_num}")

            vars = self.load_case_variables()
            if not vars:
                log_event(f"StatusUpdate: No case variables found for {file_num}", "warning")
                QMessageBox.warning(self, "Missing Data", f"No saved variables found for case {file_num}.\n\nPlease run an agent or edit 'Variables' in the Case View tab first.")
                return
            self.vars = vars 
            
            # 3. Check Procedural History for LLM summarization
            proc_hist = str(vars.get("procedural_history", "")).strip()
            if proc_hist and len(proc_hist) > 10:
                log_event("StatusUpdate: Summarizing procedural history...")
                self.status_bar.setText("Summarizing procedural history...")
                self.btn_status_update.setEnabled(False)
                
                today_str = datetime.datetime.now().strftime("%B %d, %Y")
                system_prompt = "You are an expert legal assistant. Extract future hearing dates from a procedural history text and write them as clear, complete sentences."
                user_prompt = (
                    f"Today is {today_str}. Review the following procedural history text. "
                    "Identify any hearing dates that occur AFTER today. "
                    "For each future hearing, generate a complete, professional sentence stating what the hearing is and when it is scheduled. "
                    "Combine these sentences into a single paragraph. "
                    "If there are no future hearings found, return exactly: 'No future hearings are currently scheduled.'\n\n"
                    f"Text to analyze:\n{proc_hist}"
                )
                settings = {'temperature': 1.0, 'max_tokens': -1, 'stream': False}
                
                self.proc_worker = LLMWorker("Gemini", "gemini-3-flash-preview", system_prompt, user_prompt, "", settings)
                self.proc_worker.finished.connect(self.finalize_status_update)
                self.proc_worker.error.connect(lambda e: self.finalize_status_update(f"Error parsing history: {e}"))
                self.proc_worker.start()
            else:
                log_event("StatusUpdate: No procedural history found, proceeding to finalize.")
                self.finalize_status_update("No future hearings are currently scheduled.")
        except Exception as e:
            log_event(f"StatusUpdate: CRASH in create_status_update: {e}", "error")
            QMessageBox.critical(self, "Crash", f"Error initializing report: {e}")

    def _parse_date(self, date_str):
        if not date_str: return None
        formats = ["%Y-%m-%d", "%m/%d/%Y", "%m-%d-%Y", "%d/%m/%Y", "%B %d, %Y", "%b %d, %Y"]
        for fmt in formats:
            try:
                return datetime.datetime.strptime(date_str.strip(), fmt)
            except ValueError:
                continue
        return None

    def _format_date_long(self, date_str):
        dt = self._parse_date(date_str)
        if dt: return dt.strftime("%B %d, %Y")
        return date_str

    def finalize_status_update(self, proc_status_text):
        try:
            self.status_bar.setText("Generating email...")
            self.btn_status_update.setEnabled(True)
            log_event(f"StatusUpdate: Finalizing with proc_text len={len(proc_status_text if proc_status_text else '')}")
            
            if not proc_status_text: proc_status_text = "No future hearings are currently scheduled."
            
            # Gather Data
            adjuster_email = self.vars.get("adjuster_email") or ""
            client_email = self.vars.get("Client_Email") or self.vars.get("client_email") or ""
            
            # Check multiple casing variations for case name
            case_name = self.vars.get("case_name") or self.vars.get("Case_name") or self.vars.get("Case_Name") or ""
            
            claim_num = self.vars.get("claim_number") or self.vars.get("claim_num") or self.vars.get("Claim_Number") or ""
            file_num = self.get_file_number() or self.vars.get("file_number") or ""
            adjuster_name = self.vars.get("adjuster_name") or "Adjuster"
            fact_bg = self.vars.get("factual_background") or ""
            trial_date_raw = self.vars.get("trial_date") or ""
            
            trial_date = self._format_date_long(trial_date_raw) if trial_date_raw else ""
            first_name = adjuster_name.split()[0] if adjuster_name else "Adjuster"

            # 1. Start with a list of HTML segments
            segments = []
            
            # Salutation
            segments.append(f"<p>{first_name},</p>")
            
            # Opening
            if trial_date:
                segments.append(f"<p>Please allow the following to serve as our interim update in the above-referenced matter, which is currently set for a trial to begin on <b>{trial_date}</b>:</p>")
            else:
                segments.append(f"<p>Please allow the following to serve as our interim update in the above-referenced matter, which is currently not set for trial:</p>")
                
            # Factual Background
            if fact_bg:
                segments.append(f"<p><b>Factual Background:</b> {fact_bg}</p>")
                
            # Procedural Status
            segments.append(f"<p><b>Procedural Status:</b> {proc_status_text}</p>")
            
            # 2. Insert Generated Content (Same Line requirement)
            if hasattr(self, 'parsed_topic_name') and self.parsed_topic_content:
                log_event(f"StatusUpdate: Adding Topic section: {self.parsed_topic_name}")
                # Convert markdown to HTML
                content_html = markdown.markdown(self.parsed_topic_content.strip(), extensions=['nl2br'])
                # Force the first paragraph to start with the bold title on the same line
                styled_content = content_html.replace("<p>", f"<p><b>{self.parsed_topic_name}:</b> ", 1)
                # If there was no <p> tag (rare), just prepend
                if styled_content == content_html:
                    styled_content = f"<p><b>{self.parsed_topic_name}:</b> {content_html}</p>"
                segments.append(styled_content)
                
            if hasattr(self, 'parsed_handling_content') and self.parsed_handling_content:
                log_event(f"StatusUpdate: Adding Handling section")
                handling_html = markdown.markdown(self.parsed_handling_content.strip(), extensions=['nl2br'])
                # Same logic for Further Case Handling
                styled_handling = handling_html.replace("<p>", f"<p><b>Further Case Handling:</b> ", 1)
                if styled_handling == handling_html:
                    styled_handling = f"<p><b>Further Case Handling:</b> {handling_html}</p>"
                segments.append(styled_handling)
                    
            # Closing
            segments.append("<p>If you have any questions or conerns, please do not hestitate to call or write.</p>")
            segments.append("<p>Best,<br><br>Andrei</p>")

            # Combine all segments
            full_body_content = "".join(segments)
            
            # Wrap in Outlook-friendly styles
            body_html = f"<div style='font-family: Calibri, sans-serif; font-size: 11pt; color: black;'>{full_body_content}</div>"

            # Signature Block
            signature_html = """
<p class=MsoNormal><a name="_MailAutoSig"><span style='font-family:"Garamond",serif;color:black;'>ANDREI V. SERPIK</span></a></p>
<p class=MsoNormal><i><span style='font-family:"Garamond",serif;color:black;'>Attorney</span></i></p>
<p class=MsoNormal><span style='font-family:"Garamond",serif;'>BORDIN SEMMER LLP</span><br>
<span style='font-family:"Garamond",serif;'>Howard Hughes Center<br></span>
<span style='font-family:"Garamond",serif;color:#045EC2;'>6100 Center Drive, Suite 1100</span><br>
<span style='font-family:"Garamond",serif;'>Los Angeles, California 90045</span><br>
<span style='font-family:"Garamond",serif;color:#0563C1;'>Aserpik@bordinsemmer.com</span><br>
<span style='font-size:7.5pt;font-family:"Garamond",serif;'>TEL</span><span style='font-family:"Garamond",serif;'>&nbsp;&nbsp;</span><span style='font-family:"Garamond",serif;color:#045EC2;'>323.457.2110</span><br>
<span style='font-size:7.5pt;font-family:"Garamond",serif;'>FAX</span><span style='font-family:"Garamond",serif;'>&nbsp;&nbsp;</span><span style='font-family:"Garamond",serif;color:#045EC2;'>323.457.2120</span><br>
<span style='font-family:"Garamond",serif;color:black;'>Los Angeles&nbsp;|&nbsp;San Diego&nbsp;| Bay Area</span></p>
"""
            full_html = body_html + signature_html

            # Subject
            subject = f"{case_name}; {claim_num}; ({file_num})"
            self.pending_subject = subject # Save for monitor
            
            # Send (Open Outlook)
            cc = client_email
            if cc: 
                cc += "; jbordinwosk@bordinsemmer.com"
            else:
                cc = "jbordinwosk@bordinsemmer.com"
                
            log_event(f"StatusUpdate: Opening Outlook draft with subject: {subject}")
            self.email_worker = EmailSenderWorker(adjuster_email, cc, subject, full_html)
            self.email_worker.finished.connect(self.on_email_opened)
            self.email_worker.error.connect(lambda e: QMessageBox.critical(self, "Error", f"Failed to open Outlook: {e}"))
            self.email_worker.start()
        except Exception as e:
            log_event(f"StatusUpdate: CRASH in finalize_status_update: {e}", "error")
            QMessageBox.critical(self, "Crash", f"Error generating email: {e}")

    def on_email_opened(self):
        self.status_bar.setText("Email draft opened in Outlook. Waiting for send to log...")
        # Start monitoring for the send event
        self.start_sent_monitor()

    def start_sent_monitor(self):
        # We'll poll every 10 seconds for 10 minutes
        log_event(f"StatusUpdate: Starting monitor for subject: {getattr(self, 'pending_subject', 'Unknown')}")
        self.monitor_attempts = 0
        self.monitor_timer = QTimer(self)
        self.monitor_timer.timeout.connect(self.check_sent_items)
        self.monitor_timer.start(10000) # 10 seconds

    def check_sent_items(self):
        self.monitor_attempts += 1
        if self.monitor_attempts > 60: # 10 minutes timeout
            self.monitor_timer.stop()
            self.status_bar.setText("Draft closed or send not detected (Monitor Timed Out)")
            return

        expected_subject = getattr(self, 'pending_subject', "").strip()
        if not expected_subject:
            self.monitor_timer.stop()
            return

        try:
            pythoncom.CoInitialize()
            outlook = win32com.client.Dispatch("Outlook.Application")
            mapi = outlook.GetNamespace("MAPI")
            sent_folder = mapi.GetDefaultFolder(5) # 5 = olFolderSentItems
            
            # Sort items by received time to check newest first
            items = sent_folder.Items
            items.Sort("[ReceivedTime]", True) # Descending
            
            # Check the top 20 items (highly likely to find it if just sent)
            found = False
            for i in range(1, min(21, items.Count + 1)):
                item = items.Item(i)
                if expected_subject.lower() in item.Subject.lower():
                    # Found it!
                    found = True
                    self.monitor_timer.stop()
                    
                    # Log to UI
                    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    log_item = QListWidgetItem(f"Status Report Sent - {timestamp}")
                    log_item.setData(Qt.ItemDataRole.UserRole, item.EntryID) 
                    self.sent_list.addItem(log_item)
                    self.sent_list.scrollToBottom()
                    
                    self.status_bar.setText("Email send detected and logged.")
                    log_event(f"StatusUpdate: Successfully logged sent email: {item.Subject}")
                    break
                    
            pythoncom.CoUninitialize()
        except Exception as e:
            # Silent fail for monitor loop to avoid annoying the user
            log_event(f"StatusUpdate: Monitor loop error: {e}", "warning")

    def open_sent_report(self, item):
        entry_id = item.data(Qt.ItemDataRole.UserRole)
        if not entry_id: return

        try:
            pythoncom.CoInitialize()
            outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")
            sent_item = namespace.GetItemFromID(entry_id)
            sent_item.Display()
            pythoncom.CoUninitialize()
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Could not open sent email: {e}")

    def load_case_variables(self):
        file_num = self.get_file_number()
        if not file_num: return {}
        json_path = os.path.join(GEMINI_DATA_DIR, f"{file_num}.json")
        if not os.path.exists(json_path): return {}
        try:
            import json
            with open(json_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
                flat_data = {}
                for k, v in data.items():
                    if isinstance(v, dict) and "value" in v:
                        flat_data[k] = v["value"]
                    else:
                        flat_data[k] = v
                return flat_data
        except: return {}

    def get_file_number(self):
        # Try self.window() first
        win = self.window()
        if hasattr(win, 'file_number') and win.file_number:
            return win.file_number
        
        # Fallback: check parents manually
        p = self.parent()
        while p:
            if hasattr(p, 'file_number') and p.file_number:
                return p.file_number
            p = p.parent()
        return None