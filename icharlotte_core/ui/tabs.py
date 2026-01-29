import os
import sys
import json
import subprocess
import markdown
import shutil
import datetime
import base64
import time
from functools import partial
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QSplitter, QFrame, QLabel,
    QComboBox, QPushButton, QListWidget, QListWidgetItem, QTextBrowser, QPlainTextEdit,
    QFileDialog, QMessageBox, QDialog, QProgressBar, QTabWidget,
    QTableWidget, QHeaderView, QTableWidgetItem, QCheckBox, QFileIconProvider,
    QLineEdit, QInputDialog, QMenu, QApplication, QToolButton, QScrollArea
)
from PySide6.QtCore import Qt, Signal, QThread, QFileInfo, QTimer
from PySide6.QtGui import QTextCursor, QDragEnterEvent, QDropEvent, QAction, QPixmap

from ..config import API_KEYS, SCRIPTS_DIR, GEMINI_DATA_DIR
from ..utils import log_event, sanitize_filename, format_date_to_mm_dd_yyyy
from ..llm import LLMWorker, ModelFetcher
from .dialogs import SettingsDialog, SystemPromptDialog
from .pdf_viewer_widget import PdfViewerWidget
from .chat_widgets import (
    ConversationSidebar, ResizableInputArea, ContextIndicator,
    MessageWidget, SearchResultsWidget, get_theme, THEMES
)
from ..chat import ChatPersistence, TokenCounter, Message, Conversation, BUILTIN_PROMPTS

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
    """Enhanced chat tab with conversation management, streaming, and persistence."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)

        # Core state
        self.settings = {'temperature': 1.0, 'top_p': 0.95, 'max_tokens': -1, 'thinking_level': "None"}
        self.system_prompt = "You are a helpful legal assistant. Do not provide any disclaimers about being an AI or not being an attorney. Provide direct analysis only."
        self.attached_files = []
        self.conversation_history = []  # Legacy: for backward compatibility
        self.cached_models = {}
        self.fetcher = None
        self.icon_provider = QFileIconProvider()

        # New persistence state
        self.file_number = None
        self.persistence = None
        self.current_conversation_id = None
        self.current_conversation = None

        # Streaming state
        self.stream_text = ""
        self.stream_start_pos = 0
        self.stream_start_time = None
        self.worker = None

        # Theme
        self.theme = 'light'

        self.setup_ui()

    def setup_ui(self):
        main_layout = QHBoxLayout(self)
        main_layout.setContentsMargins(0, 0, 0, 0)

        # Main splitter for three panels
        self.main_splitter = QSplitter(Qt.Orientation.Horizontal)
        main_layout.addWidget(self.main_splitter)

        # --- Conversation Sidebar (Left) - Collapsible ---
        self.conv_sidebar = ConversationSidebar(theme=self.theme)
        self.conv_sidebar.setFixedWidth(200)
        self.conv_sidebar.conversation_selected.connect(self.on_conversation_selected)
        self.conv_sidebar.save_conversation_requested.connect(self.on_save_conversation)
        self.conv_sidebar.conversation_renamed.connect(self.on_conversation_renamed)
        self.conv_sidebar.conversation_deleted.connect(self.on_conversation_deleted)
        self.main_splitter.addWidget(self.conv_sidebar)

        # Collapse sidebar by default
        self.sidebar_collapsed = True
        self.conv_sidebar.setVisible(False)

        # --- Settings Panel (Middle-Left) ---
        settings_panel = QFrame()
        settings_layout = QVBoxLayout(settings_panel)
        settings_layout.setAlignment(Qt.AlignmentFlag.AlignTop)
        settings_layout.setContentsMargins(8, 8, 8, 8)

        # Toggle conversations sidebar button
        self.toggle_sidebar_btn = QPushButton("Show Conversations")
        self.toggle_sidebar_btn.setToolTip("Show/hide the conversations sidebar")
        self.toggle_sidebar_btn.clicked.connect(self.toggle_sidebar)
        settings_layout.addWidget(self.toggle_sidebar_btn)

        # Provider
        settings_layout.addWidget(QLabel("Provider:"))
        self.provider_combo = QComboBox()
        self.provider_combo.addItems(["Gemini", "OpenAI", "Claude"])
        self.provider_combo.currentTextChanged.connect(self.update_models)
        settings_layout.addWidget(self.provider_combo)

        # Model
        settings_layout.addWidget(QLabel("Model:"))
        self.model_combo = QComboBox()
        settings_layout.addWidget(self.model_combo)
        self.update_models(self.provider_combo.currentText())

        # Buttons
        self.settings_btn = QPushButton("Model Settings")
        self.settings_btn.clicked.connect(self.open_settings)
        settings_layout.addWidget(self.settings_btn)

        self.sys_prompt_btn = QPushButton("System Instructions")
        self.sys_prompt_btn.clicked.connect(self.open_sys_prompt)
        settings_layout.addWidget(self.sys_prompt_btn)

        settings_layout.addSpacing(10)

        # Theme toggle
        theme_layout = QHBoxLayout()
        theme_layout.addWidget(QLabel("Theme:"))
        self.theme_combo = QComboBox()
        self.theme_combo.addItems(["Light", "Dark"])
        self.theme_combo.currentTextChanged.connect(self.on_theme_changed)
        theme_layout.addWidget(self.theme_combo)
        settings_layout.addLayout(theme_layout)

        settings_layout.addSpacing(10)

        # File Selection
        self.select_file_btn = QPushButton("Select File(s)")
        self.select_file_btn.clicked.connect(self.select_files)
        settings_layout.addWidget(self.select_file_btn)

        self.file_list = QListWidget()
        self.file_list.setFixedHeight(120)
        self.file_list.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.file_list.customContextMenuRequested.connect(self.show_file_context_menu)
        settings_layout.addWidget(self.file_list)

        clear_files_btn = QPushButton("Clear Files")
        clear_files_btn.clicked.connect(self.clear_files)
        settings_layout.addWidget(clear_files_btn)

        clear_chat_btn = QPushButton("Clear Chat")
        clear_chat_btn.setToolTip("Clear the current conversation and start a new one")
        clear_chat_btn.clicked.connect(self.clear_current_chat)
        settings_layout.addWidget(clear_chat_btn)

        settings_layout.addStretch()

        settings_panel.setFixedWidth(200)
        self.main_splitter.addWidget(settings_panel)

        # --- Chat Panel (Right) ---
        chat_panel = QFrame()
        chat_layout = QVBoxLayout(chat_panel)
        chat_layout.setContentsMargins(8, 8, 8, 8)
        chat_layout.setSpacing(8)

        # Search results overlay (hidden by default)
        self.search_results = SearchResultsWidget(theme=self.theme)
        self.search_results.hide()
        self.search_results.result_selected.connect(self.on_search_result_selected)

        # Chat history display
        self.chat_history = QTextBrowser()
        self.chat_history.setOpenExternalLinks(True)
        self.chat_history.setAcceptDrops(True)
        chat_layout.addWidget(self.chat_history, 1)

        # Context indicator
        self.context_indicator = ContextIndicator(theme=self.theme)
        self.context_indicator.clicked.connect(self.show_context_details)
        chat_layout.addWidget(self.context_indicator)

        # Input area with toolbar
        input_container = QWidget()
        input_layout = QVBoxLayout(input_container)
        input_layout.setContentsMargins(0, 0, 0, 0)
        input_layout.setSpacing(4)

        # Toolbar row
        toolbar_layout = QHBoxLayout()

        # Quick prompts dropdown
        self.template_btn = QToolButton()
        self.template_btn.setText("Templates")
        self.template_btn.setPopupMode(QToolButton.ToolButtonPopupMode.InstantPopup)
        self.update_template_menu()
        toolbar_layout.addWidget(self.template_btn)

        toolbar_layout.addStretch()

        # Attachment indicator
        self.attachment_label = QLabel("")
        self.attachment_label.setStyleSheet("color: #666; font-size: 11px;")
        toolbar_layout.addWidget(self.attachment_label)

        input_layout.addLayout(toolbar_layout)

        # Input row with text and buttons
        input_row = QHBoxLayout()

        self.chat_input = QPlainTextEdit()
        self.chat_input.setMinimumHeight(60)
        self.chat_input.setMaximumHeight(150)
        self.chat_input.setPlaceholderText("Type a message... (Enter to send, Shift+Enter for newline)")
        self.chat_input.setAcceptDrops(True)
        self.chat_input.dragEnterEvent = self.dragEnterEvent
        self.chat_input.dropEvent = self.dropEvent
        self.chat_input.keyPressEvent = self.chat_key_press
        self.chat_input.textChanged.connect(self.on_input_changed)
        input_row.addWidget(self.chat_input)

        # Button column
        btn_layout = QVBoxLayout()
        btn_layout.setSpacing(4)

        self.send_btn = QPushButton("Send")
        self.send_btn.setFixedSize(60, 40)
        self.send_btn.setStyleSheet("background-color: #4CAF50; color: white; font-weight: bold;")
        self.send_btn.clicked.connect(self.send_message)
        btn_layout.addWidget(self.send_btn)

        self.stop_btn = QPushButton("Stop")
        self.stop_btn.setFixedSize(60, 40)
        self.stop_btn.clicked.connect(self.stop_generation)
        self.stop_btn.setEnabled(False)
        self.stop_btn.setStyleSheet("background-color: #e57373; color: white; font-weight: bold;")
        btn_layout.addWidget(self.stop_btn)

        input_row.addLayout(btn_layout)
        input_layout.addLayout(input_row)

        chat_layout.addWidget(input_container)
        self.main_splitter.addWidget(chat_panel)

        # Set splitter sizes
        self.main_splitter.setSizes([200, 200, 800])

    # --- Persistence Methods ---

    def load_case(self, file_number: str):
        """Load conversations for a case. Called when case switches."""
        # Stop any running threads to prevent "QThread destroyed while running" errors
        if self.worker is not None and self.worker.isRunning():
            self.worker.request_stop()
            self.worker.wait(1000)

        if self.fetcher is not None and self.fetcher.isRunning():
            self.fetcher.wait(1000)

        # Clear attached files from previous case
        self.clear_files()

        self.file_number = file_number
        self.persistence = ChatPersistence(file_number)

        # Refresh conversation list
        self.refresh_conversation_list()

        # Load most recent conversation or create new
        recent_id = self.persistence.get_most_recent_conversation_id()
        if recent_id:
            self.on_conversation_selected(recent_id)
        else:
            self.on_new_conversation()

        # Load theme preference
        settings = self.persistence.get_settings()
        theme = settings.get('theme', 'light')
        self.theme_combo.setCurrentText(theme.capitalize())

    def refresh_conversation_list(self):
        """Refresh the conversation sidebar."""
        if not self.persistence:
            return
        data = self.persistence.load()
        conversations = data.get('conversations', [])
        self.conv_sidebar.set_conversations(conversations)

    def save_current_state(self):
        """Save current conversation state."""
        if not self.persistence or not self.current_conversation_id:
            return

        # Update conversation with current settings
        self.persistence.update_conversation(
            self.current_conversation_id,
            provider=self.provider_combo.currentText(),
            model=self.model_combo.currentText(),
            system_prompt=self.system_prompt,
            settings=self.settings
        )

    # --- Conversation Management ---

    def on_conversation_selected(self, conv_id: str):
        """Load selected conversation."""
        print(f"[DEBUG] on_conversation_selected called with conv_id={conv_id}")
        if not self.persistence:
            print("[DEBUG] No persistence, returning early")
            return

        self.save_current_state()
        self.current_conversation_id = conv_id
        self.current_conversation = self.persistence.get_conversation(conv_id)
        print(f"[DEBUG] Got conversation: {self.current_conversation is not None}")
        self.conv_sidebar.set_current_conversation(conv_id)

        if self.current_conversation:
            print(f"[DEBUG] Loading conversation with {len(self.current_conversation.messages)} messages")
            # Restore settings
            self.provider_combo.setCurrentText(self.current_conversation.provider)
            # Model will be set after provider change triggers model fetch
            QTimer.singleShot(500, lambda: self.model_combo.setCurrentText(self.current_conversation.model))
            self.system_prompt = self.current_conversation.system_prompt or self.system_prompt
            if self.current_conversation.settings:
                self.settings.update(self.current_conversation.settings)

            # Load messages
            self.load_conversation_messages()

            # Update legacy history for compatibility
            self.conversation_history = self.current_conversation.get_history_for_llm()

        self.update_context_indicator()

    def on_new_conversation(self):
        """Create a new conversation."""
        if not self.persistence:
            # No case loaded, just clear chat
            self.clear_chat()
            return

        self.save_current_state()

        # Create new conversation
        conv_id = self.persistence.create_conversation(
            provider=self.provider_combo.currentText(),
            model=self.model_combo.currentText(),
            system_prompt=self.system_prompt
        )

        self.current_conversation_id = conv_id
        self.current_conversation = self.persistence.get_conversation(conv_id)
        self.conversation_history = []

        self.refresh_conversation_list()
        self.conv_sidebar.set_current_conversation(conv_id)
        self.clear_chat_display()
        self.update_context_indicator()

    def on_conversation_renamed(self, conv_id: str, new_name: str):
        """Handle conversation rename."""
        if self.persistence:
            self.persistence.rename_conversation(conv_id, new_name)
            self.refresh_conversation_list()

    def on_conversation_deleted(self, conv_id: str):
        """Handle conversation deletion."""
        if self.persistence:
            self.persistence.delete_conversation(conv_id)
            self.refresh_conversation_list()

            # If deleted current conversation, load another or create new
            if conv_id == self.current_conversation_id:
                recent_id = self.persistence.get_most_recent_conversation_id()
                if recent_id:
                    self.on_conversation_selected(recent_id)
                else:
                    self.on_new_conversation()

    def on_save_conversation(self):
        """Save the current conversation and optionally rename it."""
        if not self.persistence or not self.current_conversation_id:
            return

        # Save current state
        self.save_current_state()

        # Check if conversation has messages and might need a name
        conv = self.persistence.get_conversation(self.current_conversation_id)
        if conv and conv.messages:
            # If name is still default (starts with "Chat "), offer to rename
            if conv.name.startswith("Chat "):
                # Generate a suggested name from first message
                first_msg = conv.messages[0].content[:50] if conv.messages else ""
                suggested = first_msg.split('\n')[0].strip()
                if len(suggested) > 40:
                    suggested = suggested[:40] + "..."

                new_name, ok = QInputDialog.getText(
                    self, "Save Conversation",
                    "Enter a name for this conversation:",
                    text=suggested if suggested else conv.name
                )
                if ok and new_name.strip():
                    self.persistence.rename_conversation(self.current_conversation_id, new_name.strip())

        # Refresh the sidebar to show the saved conversation
        self.refresh_conversation_list()
        self.conv_sidebar.set_current_conversation(self.current_conversation_id)

        # Show confirmation
        QMessageBox.information(self, "Saved", "Conversation saved successfully.")

    def clear_current_chat(self):
        """Clear the current chat and start a new conversation."""
        # Save current conversation first if it has messages
        if self.persistence and self.current_conversation_id:
            conv = self.persistence.get_conversation(self.current_conversation_id)
            if conv and conv.messages:
                self.save_current_state()

        # Create a new conversation
        self.on_new_conversation()

    def load_conversation_messages(self):
        """Load and display messages from current conversation."""
        self.clear_chat_display()

        if not self.current_conversation:
            return

        for msg in self.current_conversation.messages:
            self.display_message(msg.role, msg.content, msg.attachments, msg.pinned)

    def clear_chat_display(self):
        """Clear the chat display (not the data)."""
        self.chat_history.clear()

    # --- Message Display ---

    def display_message(self, role: str, content: str, attachments=None, pinned=False):
        """Display a message in the chat history."""
        colors = get_theme(self.theme)

        if role == 'user':
            prefix = "<b>You:</b>"
            bg_color = colors['user_bubble']
        else:
            prefix = "<b>AI:</b>"
            bg_color = colors['assistant_bubble']

        # Pin indicator
        pin_html = " <span style='color: #FF9800; font-size: 10px;'>[pinned]</span>" if pinned else ""

        # Attachment indicator
        att_html = ""
        if attachments:
            att_count = len(attachments) if isinstance(attachments, list) else 0
            if att_count > 0:
                att_html = f"<br><i style='color: #666; font-size: 11px;'>({att_count} file(s) attached)</i>"

        self.chat_history.append(f"{prefix}{pin_html}")

        if att_html:
            self.chat_history.insertHtml(att_html)

        # Convert markdown to HTML for assistant messages
        if role == 'assistant':
            try:
                html_text = markdown.markdown(content, extensions=['fenced_code', 'tables'])
                self.chat_history.insertHtml(html_text)
            except Exception as e:
                log_event(f"Markdown conversion failed: {e}", "error")
                self.chat_history.append(content)
        else:
            self.chat_history.append(content)

        self.chat_history.append("-" * 50)

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

        # Clean up any previous fetcher to prevent "QThread destroyed while running" error
        if self.fetcher is not None:
            try:
                self.fetcher.finished.disconnect()
                self.fetcher.error.disconnect()
            except (TypeError, RuntimeError):
                pass  # Signals may already be disconnected
            if self.fetcher.isRunning():
                self.fetcher.wait(1000)  # Wait up to 1 second for it to finish

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
        files, _ = QFileDialog.getOpenFileNames(
            self, "Select Files", "",
            "All Supported (*.pdf *.docx *.txt *.png *.jpg *.jpeg *.gif *.webp);;Documents (*.pdf *.docx *.txt);;Images (*.png *.jpg *.jpeg *.gif *.webp)"
        )
        for f in files:
            self.add_file(f)

    def add_file(self, path):
        """Add a file to the attachment list."""
        final_path = path
        ext = os.path.splitext(path)[1].lower()

        # PDF Check Logic
        if ext == ".pdf":
            res = self.check_pdf_text(path)
            if res is False:
                return  # User cancelled or failed OCR
            final_path = res

        if final_path not in self.attached_files:
            self.attached_files.append(final_path)
            item = QListWidgetItem(os.path.basename(final_path))

            # Use custom icon for images, system icon for others
            if ext in ['.png', '.jpg', '.jpeg', '.gif', '.webp']:
                # Try to create a thumbnail for images
                try:
                    pixmap = QPixmap(final_path)
                    if not pixmap.isNull():
                        pixmap = pixmap.scaled(32, 32, Qt.AspectRatioMode.KeepAspectRatio)
                        item.setIcon(pixmap)
                    else:
                        item.setIcon(self.icon_provider.icon(QFileInfo(final_path)))
                except:
                    item.setIcon(self.icon_provider.icon(QFileInfo(final_path)))
            else:
                item.setIcon(self.icon_provider.icon(QFileInfo(final_path)))

            item.setToolTip(final_path)
            item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
            item.setCheckState(Qt.CheckState.Checked)
            item.setData(Qt.ItemDataRole.UserRole, final_path)
            self.file_list.addItem(item)

    def clear_files(self):
        if self.attached_files is None:
            self.attached_files = []
        else:
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
        """Handle file drops - supports documents and images.

        Files are collected immediately but processed asynchronously to avoid
        blocking Windows Explorer during OCR or other heavy operations.
        """
        supported_extensions = (".pdf", ".docx", ".txt", ".png", ".jpg", ".jpeg", ".gif", ".webp")
        files_to_add = []
        for url in event.mimeData().urls():
            path = url.toLocalFile()
            if os.path.isfile(path) and path.lower().endswith(supported_extensions):
                files_to_add.append(path)

        # Accept the event immediately to release Windows Explorer
        event.accept()

        # Defer file processing to after the drop handler returns
        if files_to_add:
            QTimer.singleShot(0, lambda: self._process_dropped_files(files_to_add))

    def _process_dropped_files(self, file_paths):
        """Process dropped files asynchronously after drop event completes."""
        for path in file_paths:
            self.add_file(path)

    def read_files_content(self):
        """Read content from attached files, including image base64 encoding."""
        content = ""
        image_extensions = ['.png', '.jpg', '.jpeg', '.gif', '.webp']

        for i in range(self.file_list.count()):
            item = self.file_list.item(i)
            if item.checkState() == Qt.CheckState.Checked:
                path = item.data(Qt.ItemDataRole.UserRole)
                if not path:
                    continue

                ext = os.path.splitext(path)[1].lower()
                content += f"\n--- FILE: {os.path.basename(path)} ---\n"

                try:
                    if ext in image_extensions:
                        # For images, include base64 encoded data
                        with open(path, 'rb') as f:
                            image_data = base64.b64encode(f.read()).decode('utf-8')
                        mime_type = {
                            '.png': 'image/png',
                            '.jpg': 'image/jpeg',
                            '.jpeg': 'image/jpeg',
                            '.gif': 'image/gif',
                            '.webp': 'image/webp'
                        }.get(ext, 'image/png')
                        content += f"[Image: {mime_type}, base64 encoded, {len(image_data)} chars]\n"
                        content += f"data:{mime_type};base64,{image_data[:100]}...[truncated for display]\n"
                        # Note: Full image data is passed through the file_contents parameter
                    elif ext == ".txt":
                        with open(path, 'r', encoding='utf-8', errors='ignore') as f:
                            content += f.read()
                    elif ext == ".docx":
                        from docx import Document
                        doc = Document(path)
                        for p in doc.paragraphs:
                            content += p.text + "\n"
                    elif ext == ".pdf":
                        if fitz:
                            doc = fitz.open(path)
                            for page in doc:
                                content += page.get_text() + "\n"
                            doc.close()
                except Exception as e:
                    content += f"[Error reading file: {e}]\n"
        return content

    def stop_generation(self):
        """Stop the current generation."""
        if hasattr(self, 'worker') and self.worker and self.worker.isRunning():
            self.worker.request_stop()
            self.worker.wait(2000)  # Wait up to 2 seconds
            if self.worker.isRunning():
                self.worker.terminate()
                self.worker.wait()

            # If we have streamed text, save it
            if self.stream_text:
                self.chat_history.append("<br><i>[Generation stopped by user]</i>")
                self.finalize_response(self.stream_text)
            else:
                self.chat_history.append("<i>[Generation stopped by user]</i>")
                self.chat_history.append("-" * 50)

            self.send_btn.setEnabled(True)
            self.stop_btn.setEnabled(False)

    def send_message(self):
        """Send a message with streaming support."""
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

        # Prepare Content
        file_content = self.read_files_content()
        attachments = self.get_attachment_info()

        # Build user message for history
        full_msg = user_text
        if file_content:
            full_msg += "\n\n[ATTACHED FILES]:\n" + file_content

        # Save user message to persistence
        if self.persistence and self.current_conversation_id:
            token_count = TokenCounter.estimate_tokens(full_msg, self.provider_combo.currentText())
            user_message = Message(
                role='user',
                content=full_msg,
                attachments=attachments,
                token_count=token_count
            )
            self.persistence.add_message(self.current_conversation_id, user_message)

        # Update legacy history
        self.conversation_history.append({'role': 'user', 'content': full_msg})

        # Enable streaming
        settings = {**self.settings, 'stream': True}

        # Initialize streaming state
        self.stream_text = ""
        self.stream_start_time = time.time()

        # Show initial "thinking" indicator that will be replaced
        self.chat_history.append("<b>AI:</b> ")
        cursor = self.chat_history.textCursor()
        cursor.movePosition(QTextCursor.MoveOperation.End)
        self.stream_start_pos = cursor.position()

        # Start Worker
        self.worker = LLMWorker(
            self.provider_combo.currentText(),
            self.model_combo.currentText(),
            self.system_prompt,
            user_text,
            file_content,
            settings,
            history=list(self.conversation_history[:-1])  # Exclude the message we just added
        )

        self.worker.new_token.connect(self.on_streaming_token)
        self.worker.finished.connect(self.on_stream_complete)
        self.worker.error.connect(self.on_error)
        self.worker.start()

        self.update_context_indicator()

    def on_streaming_token(self, token: str):
        """Handle real-time token display during streaming."""
        self.stream_text += token

        # Append token to chat display
        cursor = self.chat_history.textCursor()
        cursor.movePosition(QTextCursor.MoveOperation.End)
        cursor.insertText(token)
        self.chat_history.setTextCursor(cursor)
        self.chat_history.ensureCursorVisible()

    def on_stream_complete(self, full_text: str):
        """Handle completion of streaming response."""
        # Replace streamed plain text with rendered markdown
        self.finalize_response(full_text)

        self.send_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)

    def finalize_response(self, text: str):
        """Finalize the response with markdown rendering and save to persistence."""
        # Calculate response time
        response_time = int((time.time() - self.stream_start_time) * 1000) if self.stream_start_time else 0

        # Remove the plain text we streamed and replace with formatted HTML
        cursor = self.chat_history.textCursor()
        cursor.setPosition(self.stream_start_pos)
        cursor.movePosition(QTextCursor.MoveOperation.End, QTextCursor.MoveMode.KeepAnchor)
        cursor.removeSelectedText()

        # Convert markdown to HTML
        try:
            html_text = markdown.markdown(text, extensions=['fenced_code', 'tables'])
        except Exception as e:
            log_event(f"Markdown conversion failed: {e}", "error")
            html_text = text.replace('\n', '<br>')

        cursor.insertHtml(html_text)
        self.chat_history.append("")  # New line
        self.chat_history.append("-" * 50)

        # Save assistant message to persistence
        if self.persistence and self.current_conversation_id:
            token_count = TokenCounter.estimate_tokens(text, self.provider_combo.currentText())
            assistant_message = Message(
                role='assistant',
                content=text,
                token_count=token_count,
                model_used=self.model_combo.currentText(),
                response_time_ms=response_time
            )
            self.persistence.add_message(self.current_conversation_id, assistant_message)

        # Update legacy history
        self.conversation_history.append({'role': 'assistant', 'content': text})

        self.update_context_indicator()

    def on_error(self, err: str):
        """Handle generation error."""
        self.chat_history.append(f"<font color='red'>Error: {err}</font>")
        self.chat_history.append("-" * 50)
        self.send_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)

    # --- Helper Methods ---

    def toggle_sidebar(self):
        """Toggle the conversations sidebar visibility."""
        self.sidebar_collapsed = not self.sidebar_collapsed
        self.conv_sidebar.setVisible(not self.sidebar_collapsed)

        if self.sidebar_collapsed:
            self.toggle_sidebar_btn.setText("Show Conversations")
        else:
            self.toggle_sidebar_btn.setText("Hide Conversations")

    def get_attachment_info(self) -> list:
        """Get attachment information for the current message."""
        attachments = []
        for i in range(self.file_list.count()):
            item = self.file_list.item(i)
            if item.checkState() == Qt.CheckState.Checked:
                path = item.data(Qt.ItemDataRole.UserRole)
                if path:
                    name = os.path.basename(path)
                    ext = os.path.splitext(path)[1].lower()
                    file_type = 'image' if ext in ['.png', '.jpg', '.jpeg', '.gif', '.webp'] else 'file'
                    attachments.append({
                        'name': name,
                        'path': path,
                        'type': file_type
                    })
        return attachments

    def update_context_indicator(self):
        """Update the context usage indicator."""
        provider = self.provider_combo.currentText()
        model = self.model_combo.currentText()

        # Calculate token usage
        total_tokens = TokenCounter.calculate_context_usage(
            self.conversation_history,
            self.system_prompt,
            '',  # Current file content not counted here
            provider
        )['total_tokens']

        context_limit = TokenCounter.get_context_limit(model, provider)
        self.context_indicator.update_usage(total_tokens, context_limit)

    def show_context_details(self):
        """Show detailed context usage breakdown."""
        provider = self.provider_combo.currentText()
        model = self.model_combo.currentText()

        usage = TokenCounter.calculate_context_usage(
            self.conversation_history,
            self.system_prompt,
            self.read_files_content(),
            provider
        )

        limit = TokenCounter.get_context_limit(model, provider)

        details = f"""Context Usage Details:

System Prompt: ~{TokenCounter.format_token_count(usage['system_tokens'])} tokens
Messages: ~{TokenCounter.format_token_count(usage['message_tokens'])} tokens
Attached Files: ~{TokenCounter.format_token_count(usage['file_tokens'])} tokens
---
Total: ~{TokenCounter.format_token_count(usage['total_tokens'])} tokens
Limit: {TokenCounter.format_token_count(limit)} tokens
Usage: {TokenCounter.get_usage_percentage(usage['total_tokens'], model, provider):.1f}%
"""
        QMessageBox.information(self, "Context Usage", details)

    def on_input_changed(self):
        """Handle input text changes for auto-resize."""
        doc = self.chat_input.document()
        height = int(doc.size().height()) + 20
        height = max(60, min(150, height))
        self.chat_input.setFixedHeight(height)

        # Update attachment indicator
        checked_count = sum(1 for i in range(self.file_list.count())
                          if self.file_list.item(i).checkState() == Qt.CheckState.Checked)
        if checked_count > 0:
            self.attachment_label.setText(f"{checked_count} file(s)")
            self.attachment_label.setVisible(True)
        else:
            self.attachment_label.setVisible(False)

    def on_theme_changed(self, theme_text: str):
        """Handle theme change."""
        self.theme = theme_text.lower()
        self.apply_theme()

        # Save preference
        if self.persistence:
            self.persistence.update_settings(theme=self.theme)

    def apply_theme(self):
        """Apply the current theme to all widgets."""
        colors = get_theme(self.theme)

        # Update sidebar
        self.conv_sidebar.theme = self.theme
        self.conv_sidebar.apply_theme()

        # Update context indicator
        self.context_indicator.theme = self.theme
        self.context_indicator.apply_theme()

        # Update chat display
        self.chat_history.setStyleSheet(f"""
            QTextBrowser {{
                background-color: {colors['bg']};
                color: {colors['text']};
                border: 1px solid {colors['border']};
                border-radius: 8px;
                padding: 8px;
            }}
        """)

        # Update input
        self.chat_input.setStyleSheet(f"""
            QPlainTextEdit {{
                background-color: {colors['bg']};
                color: {colors['text']};
                border: 1px solid {colors['border']};
                border-radius: 8px;
                padding: 8px;
            }}
        """)

    def update_template_menu(self):
        """Update the quick prompts template menu."""
        menu = QMenu(self)

        # Built-in prompts
        builtin_menu = menu.addMenu("Built-in")
        for prompt in BUILTIN_PROMPTS:
            action = QAction(prompt.name, self)
            action.triggered.connect(lambda checked, p=prompt: self.insert_template(p.prompt))
            builtin_menu.addAction(action)

        # Custom prompts (if persistence available)
        if self.persistence:
            prompts = self.persistence.get_quick_prompts()
            custom_prompts = [p for p in prompts if not p.is_builtin]
            if custom_prompts:
                menu.addSeparator()
                custom_menu = menu.addMenu("Custom")
                for prompt in custom_prompts:
                    action = QAction(prompt.name, self)
                    action.triggered.connect(lambda checked, p=prompt: self.insert_template(p.prompt))
                    custom_menu.addAction(action)

        menu.addSeparator()
        manage_action = QAction("Manage Templates...", self)
        manage_action.triggered.connect(self.open_template_manager)
        menu.addAction(manage_action)

        self.template_btn.setMenu(menu)

    def insert_template(self, prompt: str):
        """Insert a template prompt into the input."""
        current = self.chat_input.toPlainText()
        if current:
            self.chat_input.setPlainText(current + "\n\n" + prompt)
        else:
            self.chat_input.setPlainText(prompt)

    def open_template_manager(self):
        """Open the template management dialog."""
        from .chat_dialogs import PromptTemplateDialog
        dlg = PromptTemplateDialog(self.persistence, self)
        if dlg.exec():
            self.update_template_menu()

    def show_file_context_menu(self, pos):
        """Show context menu for file list."""
        item = self.file_list.itemAt(pos)
        if not item:
            return

        menu = QMenu(self)

        remove_action = QAction("Remove", self)
        remove_action.triggered.connect(lambda: self.remove_file(item))
        menu.addAction(remove_action)

        menu.exec(self.file_list.mapToGlobal(pos))

    def remove_file(self, item):
        """Remove a file from the list."""
        path = item.data(Qt.ItemDataRole.UserRole)
        if path in self.attached_files:
            self.attached_files.remove(path)
        row = self.file_list.row(item)
        self.file_list.takeItem(row)

    def on_search_result_selected(self, conv_id: str, msg_id: str):
        """Handle search result selection."""
        self.search_results.hide()
        self.on_conversation_selected(conv_id)
        # TODO: Scroll to specific message if msg_id provided

    def search_conversations(self, query: str):
        """Search across all conversations."""
        if not self.persistence or not query:
            self.search_results.hide()
            return

        results = self.persistence.search_conversations(query)
        if results:
            self.search_results.set_results(results)
        else:
            self.search_results.hide()

    # --- Legacy Compatibility ---

    def clear_chat(self):
        """Clear chat and create new conversation."""
        self.chat_history.clear()
        self.conversation_history = []
        if self.persistence:
            self.on_new_conversation()

    def reset_state(self):
        """Reset all state (called on case switch)."""
        # Stop any running threads to prevent "QThread destroyed while running" errors
        if self.worker is not None and self.worker.isRunning():
            self.worker.request_stop()
            self.worker.wait(1000)

        if self.fetcher is not None and self.fetcher.isRunning():
            self.fetcher.wait(1000)

        self.clear_chat()
        self.clear_files()
        self.conversation_history = []
        self.current_conversation_id = None
        self.current_conversation = None

class IndexTab(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.file_number = None
        self.index_data = {}  # {pdf_path: [docs]}
        self.suggestions = {} # Grouped by original folder: {folder_path: [suggestions]}
        self.current_org_folder = None
        self.current_pdf_path = None
        self.marked_start_page = None
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

        self.doc_splitter = QSplitter(Qt.Orientation.Horizontal)
        doc_layout.addWidget(self.doc_splitter)

        # Left: PDF List (collapsible)
        self.left_widget = QWidget()
        left_layout = QVBoxLayout(self.left_widget)
        left_layout.setContentsMargins(0, 0, 0, 0)

        # Header with collapse toggle
        left_header = QHBoxLayout()
        left_header.setContentsMargins(4, 4, 4, 4)
        self.pdf_list_label = QLabel("Processed PDFs:")
        left_header.addWidget(self.pdf_list_label)
        left_header.addStretch()

        self.collapse_pdf_btn = QToolButton()
        self.collapse_pdf_btn.setText("◀")
        self.collapse_pdf_btn.setToolTip("Collapse panel")
        self.collapse_pdf_btn.setFixedSize(24, 24)
        self.collapse_pdf_btn.setStyleSheet("QToolButton { border: none; font-size: 12px; }")
        self.collapse_pdf_btn.clicked.connect(self.toggle_pdf_list_collapse)
        left_header.addWidget(self.collapse_pdf_btn)
        left_layout.addLayout(left_header)

        self.pdf_list = QListWidget()
        self.pdf_list.currentItemChanged.connect(self.on_pdf_selected)
        left_layout.addWidget(self.pdf_list)
        self.doc_splitter.addWidget(self.left_widget)

        # Store original size for restore
        self._pdf_list_width = 200
        self._pdf_list_collapsed = False

        # Middle: Index Table + Controls
        middle_widget = QWidget()
        middle_layout = QVBoxLayout(middle_widget)
        middle_layout.setContentsMargins(0, 0, 0, 0)

        # Header row with expand button (hidden by default) and search bar
        middle_header = QHBoxLayout()
        middle_header.setContentsMargins(4, 4, 4, 4)

        self.expand_pdf_btn = QToolButton()
        self.expand_pdf_btn.setText("▶")
        self.expand_pdf_btn.setToolTip("Show PDF list")
        self.expand_pdf_btn.setFixedSize(24, 24)
        self.expand_pdf_btn.setStyleSheet("QToolButton { border: 1px solid #ccc; border-radius: 4px; font-size: 12px; background: #f0f0f0; }")
        self.expand_pdf_btn.clicked.connect(self.toggle_pdf_list_collapse)
        self.expand_pdf_btn.setVisible(False)  # Hidden until collapsed
        middle_header.addWidget(self.expand_pdf_btn)

        # Search Bar
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Search by Date or Title...")
        self.search_input.textChanged.connect(self.filter_documents)
        middle_header.addWidget(self.search_input)

        middle_layout.addLayout(middle_header)

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

        # Connect click signals for PDF navigation
        self.doc_table.cellClicked.connect(self.on_doc_clicked)
        self.doc_table.cellDoubleClicked.connect(self.on_doc_double_clicked)

        # Context Menu for Batch Edit
        self.doc_table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.doc_table.customContextMenuRequested.connect(self.show_context_menu)

        middle_layout.addWidget(self.doc_table)

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
        middle_layout.addLayout(btn_layout)

        self.doc_splitter.addWidget(middle_widget)

        # Right: PDF Preview
        preview_widget = QWidget()
        preview_layout = QVBoxLayout(preview_widget)
        preview_layout.setContentsMargins(0, 0, 0, 0)

        # PDF Viewer with page tracking
        self.pdf_viewer = PdfViewerWidget()
        preview_layout.addWidget(self.pdf_viewer)

        # Mark Range controls
        mark_layout = QHBoxLayout()
        self.mark_status = QLabel("Range: Not set")
        self.mark_status.setStyleSheet("font-weight: bold; padding: 5px;")
        mark_layout.addWidget(self.mark_status)

        self.mark_start_btn = QPushButton("Mark Start")
        self.mark_start_btn.setToolTip("Mark current page as start of new document")
        self.mark_start_btn.setStyleSheet("background-color: #2196F3; color: white; padding: 8px;")
        self.mark_start_btn.clicked.connect(self.mark_start_page)
        mark_layout.addWidget(self.mark_start_btn)

        self.mark_end_btn = QPushButton("Mark End && Add")
        self.mark_end_btn.setToolTip("Mark current page as end and add new document")
        self.mark_end_btn.setStyleSheet("background-color: #4CAF50; color: white; padding: 8px;")
        self.mark_end_btn.clicked.connect(self.mark_end_and_add)
        mark_layout.addWidget(self.mark_end_btn)

        self.clear_mark_btn = QPushButton("Clear")
        self.clear_mark_btn.setToolTip("Clear the marked page range")
        self.clear_mark_btn.clicked.connect(self.clear_marked_range)
        mark_layout.addWidget(self.clear_mark_btn)

        preview_layout.addLayout(mark_layout)
        self.doc_splitter.addWidget(preview_widget)

        self.doc_splitter.setSizes([200, 400, 500])
        self.doc_splitter.setCollapsible(0, True)

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

    def toggle_pdf_list_collapse(self):
        """Toggle collapse/expand of the processed PDFs panel."""
        sizes = self.doc_splitter.sizes()
        if self._pdf_list_collapsed:
            # Expand
            self.doc_splitter.setSizes([self._pdf_list_width, sizes[1] - self._pdf_list_width, sizes[2]])
            self.collapse_pdf_btn.setText("◀")
            self.collapse_pdf_btn.setToolTip("Collapse panel")
            self.expand_pdf_btn.setVisible(False)
            self._pdf_list_collapsed = False
        else:
            # Collapse - store current width first
            if sizes[0] > 0:
                self._pdf_list_width = sizes[0]
            self.doc_splitter.setSizes([0, sizes[1] + sizes[0], sizes[2]])
            self.collapse_pdf_btn.setText("▶")
            self.collapse_pdf_btn.setToolTip("Expand panel")
            self.expand_pdf_btn.setVisible(True)
            self._pdf_list_collapsed = True

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
        self.current_pdf_path = path
        docs = self.index_data.get(path, [])

        # Load PDF in preview
        if hasattr(self, 'pdf_viewer') and os.path.exists(path):
            self.pdf_viewer.load_pdf(path)

        # Clear any marked range when switching PDFs
        if hasattr(self, 'clear_marked_range'):
            self.clear_marked_range()

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
        title_edit.setCursorPosition(0)  # Show beginning of text, not end
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

    def on_doc_clicked(self, row, column):
        """Single click: Navigate to first page of document."""
        if not hasattr(self, 'pdf_viewer'):
            return
        pages_widget = self.doc_table.cellWidget(row, 4)
        if pages_widget:
            start, _ = self._parse_pages(pages_widget.text())
            if start:
                self.pdf_viewer.go_to_page(start)

    def on_doc_double_clicked(self, row, column):
        """Double click: Navigate to last page of document."""
        if not hasattr(self, 'pdf_viewer'):
            return
        pages_widget = self.doc_table.cellWidget(row, 4)
        if pages_widget:
            _, end = self._parse_pages(pages_widget.text())
            if end:
                self.pdf_viewer.go_to_page(end)

    def mark_start_page(self):
        """Mark current page as start of new document range."""
        if not hasattr(self, 'pdf_viewer'):
            return
        page = self.pdf_viewer.get_current_page()
        if page:
            self.marked_start_page = page
            self.mark_status.setText(f"Range: {page} - ?")
            self.mark_status.setStyleSheet("font-weight: bold; padding: 5px; background-color: #FFF3E0;")

    def mark_end_and_add(self):
        """Mark current page as end and add new document."""
        if not hasattr(self, 'pdf_viewer'):
            return

        if not self.marked_start_page:
            QMessageBox.warning(self, "Warning", "Please mark a start page first.")
            return

        end_page = self.pdf_viewer.get_current_page()
        if not end_page:
            QMessageBox.warning(self, "Warning", "Could not get current page.")
            return

        # Swap if end is before start
        start_page = self.marked_start_page
        if end_page < start_page:
            start_page, end_page = end_page, start_page

        # Create new document with marked range
        new_doc = {
            'id': str(self._get_next_doc_id()),
            'title': 'New Document',
            'date': '',
            'start': start_page,
            'end': end_page
        }

        self.doc_table.setSortingEnabled(False)
        row = self._add_doc_to_table(new_doc, check_sep=True)
        self.doc_table.setSortingEnabled(True)
        self.doc_table.selectRow(row)
        self.doc_table.scrollToItem(self.doc_table.item(row, 0))

        # Focus on title for immediate editing
        title_widget = self.doc_table.cellWidget(row, 5)
        if title_widget:
            title_widget.setFocus()
            title_widget.selectAll()

        # Reset marked range
        self.clear_marked_range()

    def clear_marked_range(self):
        """Clear the marked page range."""
        self.marked_start_page = None
        if hasattr(self, 'mark_status'):
            self.mark_status.setText("Range: Not set")
            self.mark_status.setStyleSheet("font-weight: bold; padding: 5px;")

    def _get_next_doc_id(self):
        """Get next available document ID."""
        max_id = 0
        for row in range(self.doc_table.rowCount()):
            id_item = self.doc_table.item(row, 2)
            if id_item:
                try:
                    max_id = max(max_id, int(id_item.text()))
                except ValueError:
                    pass
        return max_id + 1

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
