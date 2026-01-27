import os
import sys
import json
import subprocess
import markdown
import shutil
import datetime
from functools import partial
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QSplitter, QFrame, QLabel, 
    QComboBox, QPushButton, QListWidget, QListWidgetItem, QTextBrowser, QPlainTextEdit, 
    QFileDialog, QMessageBox, QDialog, QProgressBar, QTabWidget, 
    QTableWidget, QHeaderView, QTableWidgetItem, QCheckBox, QFileIconProvider,
    QLineEdit, QInputDialog, QMenu, QApplication
)
from PySide6.QtCore import Qt, Signal, QThread, QFileInfo
from PySide6.QtGui import QTextCursor, QDragEnterEvent, QDropEvent, QAction

from ..config import API_KEYS, SCRIPTS_DIR, GEMINI_DATA_DIR
from ..utils import log_event, sanitize_filename, format_date_to_mm_dd_yyyy
from ..llm import LLMWorker, ModelFetcher
from .dialogs import SettingsDialog, SystemPromptDialog

try:
    import fitz  # PyMuPDF
except ImportError:
    fitz = None

try:
    import pypdf
except ImportError:
    try:
        import PyPDF2 as pypdf
    except ImportError:
        pypdf = None

class DateTableWidgetItem(QTableWidgetItem):
    def __lt__(self, other):
        try:
            # Parse dates "MM-DD-YYYY"
            d1 = self.text().strip()
            d2 = other.text().strip()
            
            # Handle empty strings
            # If we want empty dates to be last in ascending order: return False if d1 is empty
            if not d1 and d2: return False
            if d1 and not d2: return True
            if not d1 and not d2: return False
            
            # Use datetime for comparison
            dt1 = datetime.datetime.strptime(d1, "%m-%d-%Y")
            dt2 = datetime.datetime.strptime(d2, "%m-%d-%Y")
            return dt1 < dt2
        except ValueError:
            # Fallback to string comparison if parsing fails
            return self.text() < other.text()

# --- OCR Runner ---

class OCRRunner(QThread):
    finished = Signal(bool, str, str) # success, message, final_path
    progress = Signal(int) # percentage

    def __init__(self, script_path, file_path):
        super().__init__()
        self.script_path = script_path
        self.file_path = file_path

    def run(self):
        try:
            process = subprocess.Popen(
                [sys.executable, self.script_path, self.file_path], 
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                creationflags=0x08000000 if os.name == 'nt' else 0,
                bufsize=1
            )
            
            final_path = self.file_path
            
            # Read stdout line by line for progress and results
            while True:
                line = process.stdout.readline()
                if not line and process.poll() is not None:
                    break
                
                if line:
                    line = line.strip()
                    if line.startswith("PROGRESS: "):
                        try:
                            val = int(line.replace("PROGRESS: ", ""))
                            self.progress.emit(val)
                        except: pass
                    elif line.startswith("FINAL_PATH: "):
                        final_path = line.replace("FINAL_PATH: ", "").strip()

            process.wait()
            
            if process.returncode == 0:
                self.finished.emit(True, "Success", final_path)
            else:
                stderr = process.stderr.read()
                self.finished.emit(False, f"OCR process failed: {stderr}", self.file_path)
                
        except Exception as e:
            self.finished.emit(False, str(e), self.file_path)

# --- Chat System ---

class ChatTab(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.settings = {'temperature': 1.0, 'top_p': 0.95, 'max_tokens': -1, 'thinking_level': "None"}
        self.system_prompt = "You are a helpful legal assistant. Do not provide any disclaimers about being an AI or not being an attorney. Provide direct analysis only."
        self.attached_files = []
        self.conversation_history = []
        self.cached_models = {}
        self.fetcher = None
        self.icon_provider = QFileIconProvider()
        
        self.setup_ui()
        
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
        
        # Buttons
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
        
        clear_files_btn = QPushButton("Clear Files")
        clear_files_btn.clicked.connect(self.clear_files)
        left_layout.addWidget(clear_files_btn)

        clear_chat_btn = QPushButton("Clear Chat")
        clear_chat_btn.clicked.connect(self.clear_chat)
        left_layout.addWidget(clear_chat_btn)
        
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
        # Override events for input
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

    def update_models(self, provider):
        self.model_combo.clear()
        
        # Check cache first
        if provider in self.cached_models:
            self.model_combo.addItems(self.cached_models[provider])
            return

        # Not in cache, fetch dynamically
        api_key = API_KEYS.get(provider)
        if not api_key:
            self.model_combo.addItem(f"No API Key for {provider}")
            return
            
        self.model_combo.addItem("Fetching models...")
        self.model_combo.setEnabled(False)
        
        self.fetcher = ModelFetcher(provider, api_key)
        self.fetcher.finished.connect(self.on_models_fetched)
        self.fetcher.error.connect(lambda err: self.on_models_fetched(provider, [f"Error: {err}"]))
        self.fetcher.start()

    def on_models_fetched(self, provider, models):
        self.model_combo.clear()
        self.model_combo.setEnabled(True)
        
        # If error occurred, models might contain error string
        if models and models[0].startswith("Error:"):
            self.model_combo.addItems(models)
            return

        # Update cache
        self.cached_models[provider] = models
        self.model_combo.addItems(models)
        
        # Set default model
        if provider == "Gemini":
            idx = self.model_combo.findText("gemini-3-flash-preview")
            if idx != -1:
                self.model_combo.setCurrentIndex(idx)
            else:
                # If gemini-3 not found, try searching for any gemini-3
                for i in range(self.model_combo.count()):
                    if "gemini-3" in self.model_combo.itemText(i).lower():
                        self.model_combo.setCurrentIndex(i)
                        break

    def chat_key_press(self, event):
        if event.key() in (Qt.Key.Key_Return, Qt.Key.Key_Enter) and not (event.modifiers() & Qt.KeyboardModifier.ShiftModifier):
            self.send_message()
        else:
            QPlainTextEdit.keyPressEvent(self.chat_input, event)

    def open_settings(self):
        dlg = SettingsDialog(self.settings, self)
        if dlg.exec():
            self.settings = dlg.get_settings()

    def open_sys_prompt(self):
        dlg = SystemPromptDialog(self.system_prompt, self)
        if dlg.exec():
            self.system_prompt = dlg.get_prompt()

    def select_files(self):
        files, _ = QFileDialog.getOpenFileNames(self, "Select Files", "", "Documents (*.pdf *.docx *.txt)")
        for f in files:
            self.add_file(f)

    def add_file(self, path):
        # PDF Check Logic
        final_path = path
        if path.lower().endswith(".pdf"):
            res = self.check_pdf_text(path)
            if res is False:
                return # User cancelled or failed OCR
            final_path = res
        
        if final_path not in self.attached_files:
            self.attached_files.append(final_path)
            item = QListWidgetItem(os.path.basename(final_path))
            item.setIcon(self.icon_provider.icon(QFileInfo(final_path)))
            item.setToolTip(final_path)
            item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
            item.setCheckState(Qt.CheckState.Checked)
            item.setData(Qt.ItemDataRole.UserRole, final_path)
            self.file_list.addItem(item)

    def clear_files(self):
        self.attached_files.clear()
        self.file_list.clear()

    def clear_chat(self):
        self.chat_history.clear()
        self.conversation_history = []

    def reset_state(self):
        self.clear_chat()
        self.clear_files()
        self.conversation_history = []

    def check_pdf_text(self, path):
        """Checks if PDF has text. If not, asks to OCR. Returns (possibly updated) path or False."""
        if not fitz:
            return path # Assume OK if we can't check
            
        try:
            doc = fitz.open(path)
            has_text = False
            for page in doc:
                if len(page.get_text().strip()) > 50:
                    has_text = True
                    break
            doc.close()
            
            if not has_text:
                reply = QMessageBox.question(
                    self, "OCR Needed", 
                    f"The file '{os.path.basename(path)}' appears to be an image/scanned PDF.\nDo you want to OCR it before attaching?",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No | QMessageBox.StandardButton.Cancel
                )
                if reply == QMessageBox.StandardButton.Yes:
                    ocr_path = self.run_ocr(path)
                    return ocr_path if ocr_path else False
                elif reply == QMessageBox.StandardButton.Cancel:
                    return False
                else:
                    return path # Attach anyway as is (user said No to OCR)
            return path
        except Exception as e:
            log_event(f"Error checking PDF: {e}", "error")
            return path

    def run_ocr(self, path):
        script_path = os.path.join(SCRIPTS_DIR, "ocr.py")
        if not os.path.exists(script_path):
            QMessageBox.critical(self, "Error", "ocr.py not found.")
            return None

        progress = QDialog(self)
        progress.setWindowTitle("Running OCR...")
        progress.setFixedSize(350, 120)
        layout = QVBoxLayout(progress)
        layout.addWidget(QLabel(f"Processing {os.path.basename(path)}..."))
        
        pbar = QProgressBar()
        pbar.setRange(0, 100)
        pbar.setValue(0)
        layout.addWidget(pbar)
        
        btn_layout = QHBoxLayout()
        cancel_btn = QPushButton("Cancel")
        btn_layout.addStretch()
        btn_layout.addWidget(cancel_btn)
        layout.addLayout(btn_layout)
        
        runner = OCRRunner(script_path, path)
        self.current_ocr_runner = runner # Keep reference
        
        # Connect progress signal
        runner.progress.connect(pbar.setValue)
        
        # State container to capture result from signal
        ocr_result = {"success": False, "final_path": path}

        def on_finished(success, message, final_path):
            ocr_result["success"] = success
            ocr_result["final_path"] = final_path
            if progress.isVisible():
                if success:
                    progress.accept()
                else:
                    if "terminated" not in message.lower():
                        QMessageBox.critical(self, "Error", f"OCR Failed: {message}")
                    progress.reject()
        
        def on_cancel():
            runner.terminate()
            runner.wait()
            progress.reject()
            log_event(f"OCR Cancelled by user: {path}", "warning")

        cancel_btn.clicked.connect(on_cancel)
        runner.finished.connect(on_finished)
        
        runner.start()
        res = progress.exec()
        
        if res == QDialog.DialogCode.Accepted and ocr_result["success"]:
            QMessageBox.information(self, "Success", f"OCR Completed successfully.")
            return ocr_result["final_path"]
        return None

    # Drag and Drop
    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.accept()
        else:
            event.ignore()

    def dropEvent(self, event: QDropEvent):
        for url in event.mimeData().urls():
            path = url.toLocalFile()
            if os.path.isfile(path) and path.lower().endswith((".pdf", ".docx", ".txt")):
                self.add_file(path)

    def read_files_content(self):
        content = ""
        # Iterate over list items to check for check state
        for i in range(self.file_list.count()):
            item = self.file_list.item(i)
            if item.checkState() == Qt.CheckState.Checked:
                path = item.data(Qt.ItemDataRole.UserRole)
                if not path:
                    continue
                
                content += f"\n--- FILE: {os.path.basename(path)} ---\n"
                try:
                    if path.lower().endswith(".txt"):
                        with open(path, 'r', encoding='utf-8', errors='ignore') as f:
                            content += f.read()
                    elif path.lower().endswith(".docx"):
                        from docx import Document
                        doc = Document(path)
                        for p in doc.paragraphs:
                            content += p.text + "\n"
                    elif path.lower().endswith(".pdf"):
                        if fitz:
                            doc = fitz.open(path)
                            for page in doc:
                                content += page.get_text() + "\n"
                            doc.close()
                except Exception as e:
                    content += f"[Error reading file: {e}]\n"
        return content

    def stop_generation(self):
        if hasattr(self, 'worker') and self.worker.isRunning():
            self.worker.terminate()
            self.worker.wait()
            self.chat_history.append("<i>[Generation stopped by user]</i>")
            self.chat_history.append("-" * 50)
            self.send_btn.setEnabled(True)
            self.stop_btn.setEnabled(False)

    def send_message(self):
        user_text = self.chat_input.toPlainText().strip()
        
        checked_count = 0
        for i in range(self.file_list.count()):
            if self.file_list.item(i).checkState() == Qt.CheckState.Checked:
                checked_count += 1
                
        if not user_text and checked_count == 0:
            return
            
        # Display User Message
        self.chat_history.append(f"<b>You:</b> {user_text}")
        if checked_count > 0:
            self.chat_history.append(f"<i>(Attached {checked_count} files)</i>")
        
        self.chat_input.clear()
        self.send_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)
        self.chat_history.append("<b>AI:</b> <i>Thinking...</i>")
        
        # Prepare Content
        file_content = self.read_files_content()
        
        # Start Worker
        self.worker = LLMWorker(
            self.provider_combo.currentText(),
            self.model_combo.currentText(),
            self.system_prompt,
            user_text,
            file_content,
            self.settings,
            history=list(self.conversation_history)
        )
        # Update history
        # Note: We attach files to the current message in the history too, so the model knows what was sent.
        # However, to save tokens, we might NOT want to store the full file content in history if we are re-reading it?
        # But if we use history, we MUST rely on the history content.
        # For now, let's append the full content (text + files) to the user history item?
        # No, LLMHandler.generate appends files to 'user_prompt'.
        # We should just append 'user_text' to history here, and rely on LLMHandler to combine them for the CURRENT request.
        # BUT for FUTURE requests, if we only stored 'user_text', the model won't see the files in history.
        # So we SHOULD append files to the history item.
        
        full_msg = user_text
        if file_content:
            full_msg += "\n\n[ATTACHED FILES]:\n" + file_content
            
        self.conversation_history.append({'role': 'user', 'content': full_msg})

        self.worker.finished.connect(self.on_response)
        self.worker.error.connect(self.on_error)
        self.worker.start()

    def on_response(self, text):
        # Remove "Thinking..."
        cursor = self.chat_history.textCursor()
        cursor.movePosition(QTextCursor.MoveOperation.End)
        cursor.select(QTextCursor.SelectionType.BlockUnderCursor)
        cursor.removeSelectedText()
        cursor.deletePreviousChar() # Newline
        
        # Convert Markdown to HTML
        try:
            # extensions=['fenced_code', 'codehilite'] can be added for better code block support
            # assuming standard markdown for now
            html_text = markdown.markdown(text, extensions=['fenced_code', 'tables'])
        except Exception as e:
            log_event(f"Markdown conversion failed: {e}", "error")
            html_text = text.replace('\n', '<br>')

        self.chat_history.append("<b>AI:</b>")
        self.chat_history.insertHtml(html_text)
        self.chat_history.append("-" * 50)
        self.send_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)
        
        # Update history
        self.conversation_history.append({'role': 'assistant', 'content': text})

    def on_error(self, err):
        self.chat_history.append(f"<font color='red'>Error: {err}</font>")
        self.send_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)

class IndexTab(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.file_number = None
        self.index_data = {}  # {pdf_path: [docs]}
        self.suggestions = {} # Grouped by original folder: {folder_path: [suggestions]}
        self.current_org_folder = None
        self.icon_provider = QFileIconProvider()
        self.setup_ui()

    def setup_ui(self):
        self.main_layout = QVBoxLayout(self)
        self.main_layout.setContentsMargins(0, 0, 0, 0)
        
        self.sub_tabs = QTabWidget()
        self.main_layout.addWidget(self.sub_tabs)
        
        # --- Sub-Tab 1: Document Index ---
        self.doc_index_widget = QWidget()
        self.sub_tabs.addTab(self.doc_index_widget, "Document Index")
        doc_layout = QHBoxLayout(self.doc_index_widget)
        
        splitter = QSplitter(Qt.Orientation.Horizontal)
        doc_layout.addWidget(splitter)

        # Left: PDF List
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        left_layout.addWidget(QLabel("Processed PDFs:"))
        self.pdf_list = QListWidget()
        self.pdf_list.currentItemChanged.connect(self.on_pdf_selected)
        left_layout.addWidget(self.pdf_list)
        splitter.addWidget(left_widget)

        # Right: Index Table
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)

        # Search Bar
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Search by Date or Title...")
        self.search_input.textChanged.connect(self.filter_documents)
        right_layout.addWidget(self.search_input)

        self.doc_table = QTableWidget()
        self.doc_table.setColumnCount(6)
        self.doc_table.setHorizontalHeaderLabels(["Sep.", "Merge Group", "ID", "Date", "Pages", "Title"])
        header = self.doc_table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents) # Sep
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.Interactive)      # Merge Group
        self.doc_table.setColumnWidth(1, 100)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents) # ID
        header.setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents) # Date
        header.setSectionResizeMode(4, QHeaderView.ResizeMode.ResizeToContents) # Pages
        header.setSectionResizeMode(5, QHeaderView.ResizeMode.Stretch)          # Title
        self.doc_table.setSortingEnabled(True)
        
        # Context Menu for Batch Edit
        self.doc_table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.doc_table.customContextMenuRequested.connect(self.show_context_menu)
        
        right_layout.addWidget(self.doc_table)

        btn_layout = QHBoxLayout()

        self.add_doc_btn = QPushButton("Add Document")
        self.add_doc_btn.setToolTip("Add a document that was not identified by the agent (e.g., a missed page range)")
        self.add_doc_btn.setStyleSheet("background-color: #FF9800; color: white; font-weight: bold; padding: 10px;")
        self.add_doc_btn.clicked.connect(self.add_document_row)
        btn_layout.addWidget(self.add_doc_btn)

        self.delete_doc_btn = QPushButton("Delete Selected")
        self.delete_doc_btn.setToolTip("Delete the selected document rows")
        self.delete_doc_btn.setStyleSheet("background-color: #f44336; color: white; font-weight: bold; padding: 10px;")
        self.delete_doc_btn.clicked.connect(self.delete_selected_rows)
        btn_layout.addWidget(self.delete_doc_btn)

        btn_layout.addStretch()

        self.process_btn = QPushButton("Process Documents")
        self.process_btn.setToolTip("Extracts checked 'Sep.' items and builds merged PDFs for 'Merge Group' items.")
        self.process_btn.setStyleSheet("background-color: #4CAF50; color: white; font-weight: bold; padding: 10px;")
        self.process_btn.clicked.connect(self.process_documents)
        btn_layout.addWidget(self.process_btn)
        right_layout.addLayout(btn_layout)
        
        splitter.addWidget(right_widget)
        splitter.setSizes([300, 900])

        # --- Sub-Tab 2: File Organization ---
        self.file_org_widget = QWidget()
        self.sub_tabs.addTab(self.file_org_widget, "File Organization")
        org_main_layout = QHBoxLayout(self.file_org_widget)
        
        org_splitter = QSplitter(Qt.Orientation.Horizontal)
        org_main_layout.addWidget(org_splitter)
        
        # Left: Processed Folders
        org_left_widget = QWidget()
        org_left_layout = QVBoxLayout(org_left_widget)
        org_left_layout.addWidget(QLabel("Processed Folders:"))
        self.org_folder_list = QListWidget()
        self.org_folder_list.currentItemChanged.connect(self.on_org_folder_selected)
        org_left_layout.addWidget(self.org_folder_list)
        org_splitter.addWidget(org_left_widget)
        
        # Right: Suggestions Table
        org_right_widget = QWidget()
        org_right_layout = QVBoxLayout(org_right_widget)
        
        # Selection Toggles
        toggle_layout = QHBoxLayout()
        btn_select_all = QPushButton("Select All")
        btn_select_all.clicked.connect(partial(self.set_org_checkboxes, Qt.CheckState.Checked))
        btn_unselect_all = QPushButton("Unselect All")
        btn_unselect_all.clicked.connect(partial(self.set_org_checkboxes, Qt.CheckState.Unchecked))
        toggle_layout.addWidget(btn_select_all)
        toggle_layout.addWidget(btn_unselect_all)
        toggle_layout.addStretch()
        org_right_layout.addLayout(toggle_layout)
        
        self.org_table = QTableWidget()
        self.org_table.setColumnCount(7)
        self.org_table.setHorizontalHeaderLabels(["", "Original File", "Apply Name", "Suggested Name", "Apply Folder", "Suggested Folder", "Reasoning"])
        self.org_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        self.org_table.setColumnWidth(0, 30)
        self.org_table.setColumnWidth(1, 150)
        self.org_table.setColumnWidth(2, 80)
        self.org_table.setColumnWidth(3, 200)
        self.org_table.setColumnWidth(4, 80)
        self.org_table.setColumnWidth(5, 150)
        self.org_table.horizontalHeader().setSectionResizeMode(6, QHeaderView.ResizeMode.Stretch)
        org_right_layout.addWidget(self.org_table)
        
        org_btn_layout = QHBoxLayout()
        self.apply_org_btn = QPushButton("Apply Selected Changes")
        self.apply_org_btn.setStyleSheet("background-color: #2196F3; color: white; font-weight: bold; padding: 10px;")
        self.apply_org_btn.clicked.connect(self.apply_organization_changes)
        org_btn_layout.addWidget(self.apply_org_btn)
        org_right_layout.addLayout(org_btn_layout)
        
        org_splitter.addWidget(org_right_widget)
        org_splitter.setSizes([300, 900])

    def set_org_checkboxes(self, state):
        for i in range(self.org_table.rowCount()):
            chk = self.org_table.item(i, 0)
            if chk:
                chk.setCheckState(state)

    def on_org_folder_selected(self, current, previous):
        if not current:
            self.org_table.setRowCount(0)
            return
            
        folder_path = current.data(Qt.ItemDataRole.UserRole)
        self.current_org_folder = folder_path
        self.populate_org_table(folder_path)

    def populate_org_table(self, folder_path):
        suggestions = self.suggestions.get(folder_path, [])
        self.org_table.setRowCount(len(suggestions))
        
        for i, s in enumerate(suggestions):
            orig_path = s.get("original_path", "")
            orig_name = os.path.basename(orig_path)
            
            # Col 0: Select Checkbox
            chk_all = QTableWidgetItem()
            chk_all.setFlags(Qt.ItemFlag.ItemIsUserCheckable | Qt.ItemFlag.ItemIsEnabled)
            chk_all.setCheckState(Qt.CheckState.Checked)
            self.org_table.setItem(i, 0, chk_all)
            
            # Col 1: Original File
            item_orig = QTableWidgetItem(orig_name)
            item_orig.setFlags(item_orig.flags() & ~Qt.ItemFlag.ItemIsEditable)
            item_orig.setData(Qt.ItemDataRole.UserRole, orig_path)
            item_orig.setToolTip(orig_path)
            self.org_table.setItem(i, 1, item_orig)
            
            # Col 2: Apply Name Checkbox
            chk_name = QTableWidgetItem()
            chk_name.setFlags(Qt.ItemFlag.ItemIsUserCheckable | Qt.ItemFlag.ItemIsEnabled)
            chk_name.setCheckState(Qt.CheckState.Checked)
            self.org_table.setItem(i, 2, chk_name)
            
            # Col 3: Suggested Name
            self.org_table.setItem(i, 3, QTableWidgetItem(s.get("suggested_name", "")))
            
            # Col 4: Apply Folder Checkbox
            chk_folder = QTableWidgetItem()
            chk_folder.setFlags(Qt.ItemFlag.ItemIsUserCheckable | Qt.ItemFlag.ItemIsEnabled)
            chk_folder.setCheckState(Qt.CheckState.Checked)
            self.org_table.setItem(i, 4, chk_folder)
            
            # Col 5: Suggested Folder
            self.org_table.setItem(i, 5, QTableWidgetItem(s.get("suggested_folder", "")))
            
            # Col 6: Reasoning
            reasoning = s.get("reasoning", "")
            if s.get("error"):
                reasoning = f"ERROR: {reasoning}"
            item_reason = QTableWidgetItem(reasoning)
            item_reason.setFlags(item_reason.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.org_table.setItem(i, 6, item_reason)

    def add_organization_suggestions(self, suggestions):
        # Group suggestions by original folder
        new_grouped = self.suggestions if self.suggestions else {}
        for s in suggestions:
            orig_path = s.get("original_path", "")
            folder = os.path.dirname(orig_path)
            if folder not in new_grouped:
                new_grouped[folder] = []
            
            # Update existing or add new
            existing_idx = next((i for i, item in enumerate(new_grouped[folder]) if item.get("original_path") == orig_path), -1)
            if existing_idx != -1:
                new_grouped[folder][existing_idx] = s
            else:
                new_grouped[folder].append(s)
            
        self.suggestions = new_grouped
        self.save_org_data()
        self.refresh_org_folder_list()
        self.sub_tabs.setCurrentIndex(1) # Switch to Organization tab

    def apply_organization_changes(self):
        if not self.current_org_folder:
            return
            
        suggestions_for_folder = self.suggestions.get(self.current_org_folder, [])
        approved = []
        
        for i in range(self.org_table.rowCount()):
            if self.org_table.item(i, 0).checkState() == Qt.CheckState.Checked:
                orig_path = self.org_table.item(i, 1).data(Qt.ItemDataRole.UserRole)
                
                # Find matching suggestion metadata (like doc_type)
                s_meta = next((s for s in suggestions_for_folder if s.get("original_path") == orig_path), {})
                
                # Apply Name
                if self.org_table.item(i, 2).checkState() == Qt.CheckState.Checked:
                    suggested_name = self.org_table.item(i, 3).text()
                else:
                    suggested_name = os.path.basename(orig_path)
                    
                # Apply Folder
                if self.org_table.item(i, 4).checkState() == Qt.CheckState.Checked:
                    suggested_folder = self.org_table.item(i, 5).text()
                    # Use main window's case path if available, else local folder
                    target_dir = self.window().case_path if hasattr(self.window(), 'case_path') else os.path.dirname(orig_path)
                else:
                    suggested_folder = "" # Stay in original folder
                    target_dir = os.path.dirname(orig_path)

                approved.append({
                    "original_path": orig_path,
                    "doc_type": s_meta.get("doc_type"),
                    "suggested_name": suggested_name,
                    "suggested_folder": suggested_folder,
                    "target_dir": target_dir
                })
        
        if not approved:
            QMessageBox.information(self, "Info", "No changes selected.")
            return
            
        # We need to reach the MainWindow to call apply_organization
        main_win = self.window()
        if hasattr(main_win, 'apply_organization'):
            main_win.apply_organization(approved)
            
            # Remove the processed items from suggestions
            processed_paths = [a['original_path'] for a in approved]
            self.suggestions[self.current_org_folder] = [
                s for s in self.suggestions[self.current_org_folder] 
                if s['original_path'] not in processed_paths
            ]
            
            # If folder is now empty, remove it
            if not self.suggestions[self.current_org_folder]:
                del self.suggestions[self.current_org_folder]
                
            self.save_org_data()
            self.refresh_org_folder_list()
        else:
            QMessageBox.critical(self, "Error", "Main window does not support apply_organization.")

    def refresh_org_folder_list(self):
        self.org_folder_list.clear()
        for folder in sorted(self.suggestions.keys()):
            from PySide6.QtWidgets import QListWidgetItem
            display_name = os.path.basename(folder) if folder else "Root"
            item = QListWidgetItem(display_name)
            item.setData(Qt.ItemDataRole.UserRole, folder)
            item.setToolTip(folder)
            self.org_folder_list.addItem(item)
            
        if self.org_folder_list.count() > 0:
            self.org_folder_list.setCurrentRow(0)
        else:
            self.org_table.setRowCount(0)
            self.current_org_folder = None

    def load_data(self, file_number):
        self.file_number = file_number
        self.index_data = {}
        self.suggestions = {}
        self.pdf_list.clear()
        self.org_folder_list.clear()
        self.doc_table.setRowCount(0)
        self.org_table.setRowCount(0)
        
        # Load Index Data
        idx_path = os.path.join(GEMINI_DATA_DIR, f"{file_number}_index.json")
        if os.path.exists(idx_path):
            try:
                with open(idx_path, 'r') as f:
                    self.index_data = json.load(f)
            except Exception as e:
                log_event(f"Error loading index: {e}", "error")
        
        for path in self.index_data:
            item = QListWidgetItem(path)
            item.setIcon(self.icon_provider.icon(QFileInfo(path)))
            self.pdf_list.addItem(item)
            
        # Load Organization Data
        org_path = os.path.join(GEMINI_DATA_DIR, f"{file_number}_org.json")
        if os.path.exists(org_path):
            try:
                with open(org_path, 'r') as f:
                    self.suggestions = json.load(f)
            except Exception as e:
                log_event(f"Error loading organization data: {e}", "error")
        
        self.refresh_org_folder_list()

    def save_data(self):
        if not self.file_number: return
        if not os.path.exists(GEMINI_DATA_DIR):
            os.makedirs(GEMINI_DATA_DIR)
        
        json_path = os.path.join(GEMINI_DATA_DIR, f"{self.file_number}_index.json")
        try:
            with open(json_path, 'w') as f:
                json.dump(self.index_data, f, indent=4)
        except Exception as e:
            log_event(f"Error saving index: {e}", "error")

    def save_org_data(self):
        if not self.file_number: return
        if not os.path.exists(GEMINI_DATA_DIR):
            os.makedirs(GEMINI_DATA_DIR)
            
        json_path = os.path.join(GEMINI_DATA_DIR, f"{self.file_number}_org.json")
        try:
            with open(json_path, 'w') as f:
                json.dump(self.suggestions, f, indent=4)
        except Exception as e:
            log_event(f"Error saving organization data: {e}", "error")

    def add_pdf(self, path, docs):
        self.index_data[path] = docs
        self.save_data()
        
        # Add to list if not present
        items = self.pdf_list.findItems(path, Qt.MatchFlag.MatchExactly)
        if not items:
            item = QListWidgetItem(path)
            item.setIcon(self.icon_provider.icon(QFileInfo(path)))
            self.pdf_list.addItem(item)
            self.pdf_list.setCurrentRow(self.pdf_list.count() - 1)
        else:
            self.pdf_list.setCurrentItem(items[0])
            self.on_pdf_selected(items[0], None)

    def filter_documents(self, text):
        text = text.lower()
        for row in range(self.doc_table.rowCount()):
            date_item = self.doc_table.item(row, 3)
            title_widget = self.doc_table.cellWidget(row, 5)  # Title is now a QLineEdit

            date_text = date_item.text().lower() if date_item else ""
            title_text = title_widget.text().lower() if isinstance(title_widget, QLineEdit) else ""

            if text in date_text or text in title_text:
                self.doc_table.setRowHidden(row, False)
            else:
                self.doc_table.setRowHidden(row, True)

    def on_pdf_selected(self, current, previous):
        if not current: return
        path = current.text()
        docs = self.index_data.get(path, [])

        self.doc_table.setSortingEnabled(False)
        self.doc_table.setRowCount(0)
        for doc in docs:
            self._add_doc_to_table(doc)

        self.doc_table.setSortingEnabled(True)

        # Apply filter
        if hasattr(self, 'search_input'):
            self.filter_documents(self.search_input.text())

    def _add_doc_to_table(self, doc, check_sep=False):
        """Helper to add a document row to the table. Used both for loading and adding new docs."""
        row = self.doc_table.rowCount()
        self.doc_table.insertRow(row)

        # Col 0: Sep. Checkbox
        chk_item = QTableWidgetItem()
        chk_item.setFlags(Qt.ItemFlag.ItemIsUserCheckable | Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable)
        chk_item.setCheckState(Qt.CheckState.Checked if check_sep else Qt.CheckState.Unchecked)
        self.doc_table.setItem(row, 0, chk_item)

        # Col 1: Merge Group (LineEdit)
        merge_edit = QLineEdit()
        merge_edit.setPlaceholderText("")
        self.doc_table.setCellWidget(row, 1, merge_edit)

        # Col 2: ID (read-only)
        id_item = QTableWidgetItem(str(doc.get('id', '')))
        id_item.setFlags(id_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
        self.doc_table.setItem(row, 2, id_item)

        # Col 3: Date (read-only display, but stored)
        date_val = doc.get('date', '')
        formatted_date = format_date_to_mm_dd_yyyy(date_val)
        date_item = DateTableWidgetItem(formatted_date)
        date_item.setFlags(date_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
        self.doc_table.setItem(row, 3, date_item)

        # Col 4: Pages (EDITABLE - use LineEdit for better UX)
        start = doc.get('start', '')
        end = doc.get('end', '')
        pages = f"{start}-{end}" if start != end else str(start)
        pages_edit = QLineEdit(pages)
        pages_edit.setPlaceholderText("e.g., 5-7")
        pages_edit.setToolTip("Edit page range (format: start-end or single page)")
        self.doc_table.setCellWidget(row, 4, pages_edit)

        # Col 5: Title (EDITABLE - use LineEdit)
        title_edit = QLineEdit(str(doc.get('title', '')))
        title_edit.setPlaceholderText("Document title")
        title_edit.setToolTip("Edit document title")
        self.doc_table.setCellWidget(row, 5, title_edit)

        return row

    def add_document_row(self):
        """Add a new document row that the user can fill in manually."""
        current_item = self.pdf_list.currentItem()
        if not current_item:
            QMessageBox.warning(self, "Warning", "Please select a PDF first.")
            return

        # Find the next available ID
        max_id = 0
        for row in range(self.doc_table.rowCount()):
            id_item = self.doc_table.item(row, 2)
            if id_item:
                try:
                    max_id = max(max_id, int(id_item.text()))
                except ValueError:
                    pass

        new_id = max_id + 1

        # Create a new doc dict for the new row
        new_doc = {
            'id': str(new_id),
            'title': 'New Document',
            'date': '',
            'start': '',
            'end': ''
        }

        # Disable sorting temporarily to add at the end
        self.doc_table.setSortingEnabled(False)
        row = self._add_doc_to_table(new_doc, check_sep=True)
        self.doc_table.setSortingEnabled(True)

        # Select and scroll to the new row
        self.doc_table.selectRow(row)
        self.doc_table.scrollToItem(self.doc_table.item(row, 0))

        # Focus on the Pages field for immediate editing
        pages_widget = self.doc_table.cellWidget(row, 4)
        if pages_widget:
            pages_widget.setFocus()
            pages_widget.selectAll()

    def delete_selected_rows(self):
        """Delete the selected document rows."""
        selected_rows = set()
        for range_ in self.doc_table.selectedRanges():
            for r in range(range_.topRow(), range_.bottomRow() + 1):
                selected_rows.add(r)

        if not selected_rows:
            QMessageBox.information(self, "Info", "No rows selected.")
            return

        reply = QMessageBox.question(
            self, "Confirm Delete",
            f"Delete {len(selected_rows)} selected document(s)?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )

        if reply == QMessageBox.StandardButton.Yes:
            # Delete in reverse order to maintain correct indices
            for row in sorted(selected_rows, reverse=True):
                self.doc_table.removeRow(row)

    def save_table_to_index(self):
        """Save the current table state (including edits and additions) back to index_data and file."""
        current_item = self.pdf_list.currentItem()
        if not current_item:
            QMessageBox.warning(self, "Warning", "No PDF selected.")
            return

        pdf_path = current_item.text()
        new_docs = []

        for row in range(self.doc_table.rowCount()):
            doc_obj = self._get_doc_from_row(row)

            # Validate page range
            if doc_obj['start'] is None or doc_obj['end'] is None:
                QMessageBox.warning(
                    self, "Validation Error",
                    f"Row {row + 1} (ID {doc_obj['id']}): Invalid page range. Cannot save."
                )
                return

            new_docs.append({
                'id': doc_obj['id'],
                'title': doc_obj['title'],
                'date': doc_obj['date'],
                'start': doc_obj['start'],
                'end': doc_obj['end']
            })

        # Update in-memory data
        self.index_data[pdf_path] = new_docs
        self.save_data()

        QMessageBox.information(self, "Success", f"Saved {len(new_docs)} document(s) to index.")

    def show_context_menu(self, position):
        menu = QMenu()

        add_doc_action = QAction("Add New Document", self)
        add_doc_action.triggered.connect(self.add_document_row)
        menu.addAction(add_doc_action)

        delete_action = QAction("Delete Selected Row(s)", self)
        delete_action.triggered.connect(self.delete_selected_rows)
        menu.addAction(delete_action)

        menu.addSeparator()

        set_group_action = QAction("Set Merge Group for Selected", self)
        set_group_action.triggered.connect(self.set_merge_group_batch)
        menu.addAction(set_group_action)

        clear_group_action = QAction("Clear Merge Group for Selected", self)
        clear_group_action.triggered.connect(self.clear_merge_group_batch)
        menu.addAction(clear_group_action)

        menu.addSeparator()

        save_changes_action = QAction("Save Changes to Index", self)
        save_changes_action.triggered.connect(self.save_table_to_index)
        menu.addAction(save_changes_action)

        menu.exec(self.doc_table.viewport().mapToGlobal(position))

    def set_merge_group_batch(self):
        selected_rows = set()
        for range_ in self.doc_table.selectedRanges():
            for r in range(range_.topRow(), range_.bottomRow() + 1):
                if not self.doc_table.isRowHidden(r):
                    selected_rows.add(r)
        
        if not selected_rows:
            return

        group_name, ok = QInputDialog.getText(self, "Set Merge Group", "Enter Merge Group Name:")
        if ok:
            for row in selected_rows:
                widget = self.doc_table.cellWidget(row, 1)
                if isinstance(widget, QLineEdit):
                    widget.setText(group_name)

    def clear_merge_group_batch(self):
        selected_rows = set()
        for range_ in self.doc_table.selectedRanges():
            for r in range(range_.topRow(), range_.bottomRow() + 1):
                if not self.doc_table.isRowHidden(r):
                    selected_rows.add(r)
        
        for row in selected_rows:
            widget = self.doc_table.cellWidget(row, 1)
            if isinstance(widget, QLineEdit):
                widget.setText("")

    def _parse_pages(self, pages_str):
        """Parse a pages string like '5-7' or '8' into (start, end) tuple."""
        pages_str = pages_str.strip()
        if '-' in pages_str:
            parts = pages_str.split('-')
            try:
                start = int(parts[0].strip())
                end = int(parts[1].strip())
                return start, end
            except (ValueError, IndexError):
                return None, None
        else:
            try:
                page = int(pages_str)
                return page, page
            except ValueError:
                return None, None

    def _get_doc_from_row(self, row):
        """Extract document data from a table row, reading from the editable widgets."""
        doc_id = self.doc_table.item(row, 2).text() if self.doc_table.item(row, 2) else ""

        # Get pages from widget
        pages_widget = self.doc_table.cellWidget(row, 4)
        pages_str = pages_widget.text() if isinstance(pages_widget, QLineEdit) else ""
        start, end = self._parse_pages(pages_str)

        # Get title from widget
        title_widget = self.doc_table.cellWidget(row, 5)
        title = title_widget.text() if isinstance(title_widget, QLineEdit) else ""

        # Get date from item
        date_item = self.doc_table.item(row, 3)
        date = date_item.text() if date_item else ""

        return {
            'id': doc_id,
            'title': title,
            'date': date,
            'start': start,
            'end': end
        }

    def process_documents(self):
        current_item = self.pdf_list.currentItem()
        if not current_item: return

        pdf_path = current_item.text()
        if not os.path.exists(pdf_path):
            QMessageBox.warning(self, "Error", f"Source PDF not found: {pdf_path}")
            return

        # Check if table has any rows
        if self.doc_table.rowCount() == 0:
            QMessageBox.information(self, "Info", "No documents in the table.")
            return

        # 1. Collect Tasks - read directly from the table, not from self.index_data
        separate_tasks = []  # List of doc dicts built from table data
        merge_groups = {}    # { group_name: [doc_dicts] }

        rows = self.doc_table.rowCount()
        validation_errors = []

        for row in range(rows):
            if self.doc_table.isRowHidden(row): continue

            # Build doc_obj from current table values (user edits)
            doc_obj = self._get_doc_from_row(row)

            # Check "Sep."
            is_sep_checked = self.doc_table.item(row, 0).checkState() == Qt.CheckState.Checked

            # Check "Merge Group"
            merge_widget = self.doc_table.cellWidget(row, 1)
            group_name = merge_widget.text().strip() if isinstance(merge_widget, QLineEdit) else ""

            # Skip if neither sep nor merge group
            if not is_sep_checked and not group_name:
                continue

            # Validate page range
            if doc_obj['start'] is None or doc_obj['end'] is None:
                validation_errors.append(f"Row {row + 1} (ID {doc_obj['id']}): Invalid page range")
                continue

            if not doc_obj['title'].strip():
                validation_errors.append(f"Row {row + 1} (ID {doc_obj['id']}): Title is empty")
                continue

            if is_sep_checked:
                separate_tasks.append(doc_obj)

            if group_name:
                if group_name not in merge_groups:
                    merge_groups[group_name] = []
                merge_groups[group_name].append(doc_obj)

        # Show validation errors if any
        if validation_errors:
            error_msg = "The following rows have issues:\n\n" + "\n".join(validation_errors)
            QMessageBox.warning(self, "Validation Errors", error_msg)

        if not separate_tasks and not merge_groups:
            QMessageBox.information(self, "Info", "No actions selected (Check 'Sep.' or enter a 'Merge Group').")
            return

        # 2. Setup Output
        base_dir = os.path.dirname(pdf_path)
        source_name = os.path.splitext(os.path.basename(pdf_path))[0]
        output_folder = os.path.join(base_dir, f"PULLED-{source_name}")
        
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)

        if not pypdf:
             QMessageBox.critical(self, "Error", "pypdf library not found.")
             return

        created_files = []
        errors = []

        try:
            reader = pypdf.PdfReader(pdf_path)
            
            # 3. Process Individual Extractions
            for doc in separate_tasks:
                try:
                    writer = pypdf.PdfWriter()
                    start = int(doc['start']) - 1
                    end = int(doc['end'])
                    
                    if start < 0: start = 0
                    if end > len(reader.pages): end = len(reader.pages)
                    
                    for i in range(start, end):
                        writer.add_page(reader.pages[i])
                    
                    safe_title = sanitize_filename(doc['title'])
                    if len(safe_title) > 50: safe_title = safe_title[:50]
                    
                    out_name = f"{source_name} - {doc['id']} - {safe_title}.pdf"
                    out_path = os.path.join(output_folder, out_name)
                    
                    with open(out_path, "wb") as f:
                        writer.write(f)
                    created_files.append(os.path.basename(out_path))
                except Exception as e:
                    errors.append(f"Failed to separate '{doc['title']}': {e}")

            # 4. Process Merged Groups
            for group_name, group_docs in merge_groups.items():
                try:
                    writer = pypdf.PdfWriter()
                    
                    # Sort by ID or original order? Usually standard list order is fine.
                    # group_docs is appended in row order, so it matches table order.
                    
                    for doc in group_docs:
                        start = int(doc['start']) - 1
                        end = int(doc['end'])
                        
                        if start < 0: start = 0
                        if end > len(reader.pages): end = len(reader.pages)
                        
                        for i in range(start, end):
                            writer.add_page(reader.pages[i])
                    
                    safe_title = sanitize_filename(group_name)
                    if not safe_title.lower().endswith('.pdf'):
                        safe_title += ".pdf"
                        
                    out_path = os.path.join(output_folder, safe_title)
                    
                    with open(out_path, "wb") as f:
                        writer.write(f)
                    created_files.append(f"{os.path.basename(out_path)} ({len(group_docs)} docs)")
                except Exception as e:
                    errors.append(f"Failed to merge group '{group_name}': {e}")

            # 5. Report Results
            msg = f"Processed {len(separate_tasks) + len(merge_groups)} tasks.\n\nFiles Created:\n" + "\n".join(created_files[:10])
            if len(created_files) > 10:
                msg += f"\n...and {len(created_files) - 10} more."
            
            if errors:
                msg += "\n\nErrors:\n" + "\n".join(errors)
                QMessageBox.warning(self, "Result with Errors", msg)
            else:
                QMessageBox.information(self, "Success", msg)

        except Exception as e:
            QMessageBox.critical(self, "Critical Error", f"Processing failed: {e}")
