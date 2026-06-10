import os
import re
import sys
import json
import subprocess
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
    QLineEdit, QInputDialog, QMenu, QApplication, QToolButton, QScrollArea,
    QSizePolicy, QSlider
)
from PySide6.QtCore import Qt, Signal, QThread, QFileInfo, QTimer, QSettings, QEvent
from PySide6.QtGui import QTextCursor, QTextDocument, QDragEnterEvent, QDropEvent, QAction, QActionGroup, QPixmap, QBrush

from ..config import API_KEYS, SCRIPTS_DIR, GEMINI_DATA_DIR
from ..utils import log_event
from ..llm import LLMWorker, ModelFetcher
from .dialogs import SettingsDialog, SystemPromptDialog
from .chat_widgets import (
    ConversationSidebar, ResizableInputArea, ContextIndicator,
    MessageWidget, SearchResultsWidget, get_theme, THEMES
)
from ..chat import (
    ChatPersistence,
    TokenCounter,
    Message,
    Conversation,
    BUILTIN_PROMPTS,
    TRANSCRIBE_PROMPT,
    ChatResearchError,
    ChatResearchSettings,
    CourtListenerMode,
)
from ..chat.legal_research import ChatLegalResearchService
from ..chat.markdown_render import render_markdown, CHAT_MARKDOWN_CSS
from ..mediation_brief import (
    MediationBriefGenerator, MediationBriefWorker, RefinementWorker,
    RoutingWorker, SECTION_HEADINGS,
)

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

# Matches carrier report filenames: carrier001..carrier015 with optional
# trailing text (e.g. " (FSR)", "(lit plan)", " - Final"), .doc or .docx.
# Anchored at start to reject prefixed variants like "[draft]carrier001.docx".
CARRIER_REPORT_RE = re.compile(
    r'^carrier0(0[1-9]|1[0-5])(?![0-9]).*\.docx?$',
    re.IGNORECASE,
)

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
                        except Exception: pass
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


class ChatLegalResearchWorker(QThread):
    status_update = Signal(str)
    debug_update = Signal(str, str, str, object)  # phase, message, level, details
    research_finished = Signal(object)
    research_failed = Signal(str, str)  # kind, message

    def __init__(
        self,
        *,
        user_text,
        file_content,
        provider,
        model,
        settings,
        research_settings,
        parent=None,
    ):
        super().__init__(parent)
        self.user_text = user_text or ""
        self.file_content = file_content or ""
        self.provider = provider
        self.model = model
        self.settings = dict(settings or {})
        self.research_settings = research_settings

    def request_stop(self):
        self.requestInterruption()

    def _raise_if_cancelled(self):
        if self.isInterruptionRequested():
            raise ChatResearchError("Legal research cancelled.")

    def run(self):
        from icharlotte_core.llm import LLMHandler

        def llm_for_research(system_prompt, user_prompt):
            self._raise_if_cancelled()
            result = LLMHandler.generate(
                provider=self.provider,
                model=self.model,
                system_prompt=system_prompt,
                user_prompt=user_prompt,
                file_contents="",
                settings={**self.settings, "stream": False, "temperature": 0.2},
            )
            self._raise_if_cancelled()
            return result

        def status(message):
            self._raise_if_cancelled()
            self.status_update.emit(str(message))

        def debug_event(*, phase, message, level="info", details=None):
            if self.isInterruptionRequested():
                return
            self.debug_update.emit(
                str(phase),
                str(message),
                str(level),
                dict(details or {}),
            )

        try:
            service = ChatLegalResearchService.from_environment(
                llm_callback=llm_for_research
            )
            self._raise_if_cancelled()
            packet = service.research(
                user_text=self.user_text,
                context_text=self.file_content[:100000] if self.file_content else "",
                settings=self.research_settings,
                status_callback=status,
                debug_callback=debug_event,
            )
            self._raise_if_cancelled()
            self.research_finished.emit(packet)
        except ChatResearchError as exc:
            self.research_failed.emit("stopped", str(exc))
        except Exception as exc:
            self.research_failed.emit("error", str(exc))


# --- Chat System ---


class ResizableListWidget(QListWidget):
    """QListWidget that can be resized by dragging its bottom edge."""

    HANDLE_HEIGHT = 8  # pixels from bottom edge that trigger resize

    def __init__(self, parent=None):
        super().__init__(parent)
        self._resizing = False
        self._start_y = 0
        self._start_height = 0
        self.setMouseTracking(True)
        self.viewport().setMouseTracking(True)
        self.viewport().installEventFilter(self)
        self.setFixedHeight(120)

    def _in_handle_zone(self, widget_pos_y):
        """Check if y position (in widget coordinates) is in the resize handle zone."""
        return widget_pos_y >= self.height() - self.HANDLE_HEIGHT

    def _viewport_to_widget_y(self, viewport_y):
        """Convert viewport-local y to widget-local y."""
        return viewport_y + self.viewport().y()

    def eventFilter(self, obj, event):
        """Intercept viewport mouse events for resize handling."""
        if obj is not self.viewport():
            return super().eventFilter(obj, event)

        if event.type() == QEvent.Type.MouseMove:
            if self._resizing:
                delta = int(event.globalPosition().y() - self._start_y)
                new_h = max(self.minimumHeight(), min(self.maximumHeight(), self._start_height + delta))
                self.setFixedHeight(new_h)
                return True
            widget_y = self._viewport_to_widget_y(int(event.position().y()))
            if self._in_handle_zone(widget_y):
                self.setCursor(Qt.CursorShape.SizeVerCursor)
            else:
                self.unsetCursor()
            return False

        elif event.type() == QEvent.Type.MouseButtonPress:
            if event.button() == Qt.MouseButton.LeftButton:
                widget_y = self._viewport_to_widget_y(int(event.position().y()))
                if self._in_handle_zone(widget_y):
                    self._resizing = True
                    self._start_y = event.globalPosition().y()
                    self._start_height = self.height()
                    return True
            return False

        elif event.type() == QEvent.Type.MouseButtonRelease:
            if self._resizing:
                self._resizing = False
                return True
            return False

        return False

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton and self._in_handle_zone(int(event.position().y())):
            self._resizing = True
            self._start_y = event.globalPosition().y()
            self._start_height = self.height()
            event.accept()
        else:
            super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        if self._resizing:
            delta = int(event.globalPosition().y() - self._start_y)
            new_h = max(self.minimumHeight(), min(self.maximumHeight(), self._start_height + delta))
            self.setFixedHeight(new_h)
            event.accept()
        else:
            if self._in_handle_zone(int(event.position().y())):
                self.setCursor(Qt.CursorShape.SizeVerCursor)
            else:
                self.unsetCursor()
            super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event):
        if self._resizing:
            self._resizing = False
            event.accept()
        else:
            super().mouseReleaseEvent(event)


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
        self.chat_research_worker = None

        # Mediation brief state
        self.med_brief_generator = None
        self.med_brief_worker = None

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
        self.conv_sidebar.setMinimumWidth(150)
        self.conv_sidebar.setMaximumWidth(400)
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

        # Seed the default provider/model from the Workbench (agent_chat). A loaded
        # conversation or a live pick still overrides this; on_models_fetched
        # consumes self._seed_model once when the model list first populates.
        self._seed_model = None
        try:
            from ..llm_config import get_primary_model_for_agent
            seed_provider, seed_model = get_primary_model_for_agent("agent_chat", default=(None, None))
            if seed_provider:
                self._seed_model = seed_model
                self.provider_combo.blockSignals(True)
                self.provider_combo.setCurrentText(seed_provider)
                self.provider_combo.blockSignals(False)
        except Exception:
            pass

        self.update_models(self.provider_combo.currentText())

        # Buttons
        self.settings_btn = QPushButton("Model Settings")
        self.settings_btn.clicked.connect(self.open_settings)
        settings_layout.addWidget(self.settings_btn)

        self.sys_prompt_btn = QPushButton("System Instructions")
        self.sys_prompt_btn.clicked.connect(self.open_sys_prompt)
        settings_layout.addWidget(self.sys_prompt_btn)

        # Templates dropdown (under System Instructions)
        self.template_btn = QToolButton()
        self.template_btn.setText("Templates ▾")
        self.template_btn.setPopupMode(QToolButton.ToolButtonPopupMode.InstantPopup)
        self.template_btn.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self.update_template_menu()
        settings_layout.addWidget(self.template_btn)

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

        self.file_list = ResizableListWidget()
        self.file_list.setMinimumHeight(60)
        self.file_list.setMaximumHeight(400)
        self.file_list.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.file_list.customContextMenuRequested.connect(self.show_file_context_menu)
        self.file_list.itemDoubleClicked.connect(self.open_file_from_list)
        settings_layout.addWidget(self.file_list)

        # Select All / Deselect All / Clear Files / Import Reports buttons
        file_btn_layout = QHBoxLayout()
        select_all_btn = QPushButton("All")
        select_all_btn.setToolTip("Select all files")
        select_all_btn.clicked.connect(lambda: self._set_all_file_checks(Qt.CheckState.Checked))
        deselect_all_btn = QPushButton("None")
        deselect_all_btn.setToolTip("Deselect all files")
        deselect_all_btn.clicked.connect(lambda: self._set_all_file_checks(Qt.CheckState.Unchecked))
        clear_files_btn = QPushButton("Clear")
        clear_files_btn.setToolTip("Remove all files")
        clear_files_btn.clicked.connect(self.clear_files)
        import_reports_btn = QPushButton("Import Reports")
        import_reports_btn.setToolTip("Import carrier reports from the case's STATUS folder")
        import_reports_btn.clicked.connect(self.import_carrier_reports)
        file_btn_layout.addWidget(select_all_btn)
        file_btn_layout.addWidget(deselect_all_btn)
        file_btn_layout.addWidget(clear_files_btn)
        file_btn_layout.addWidget(import_reports_btn)
        settings_layout.addLayout(file_btn_layout)

        # --- Saved Documents (document text library) ---
        from PySide6.QtWidgets import QTreeWidget, QTreeWidgetItem
        self._QTreeWidgetItem = QTreeWidgetItem
        settings_layout.addWidget(QLabel("Saved Documents:"))
        self.library_tree = QTreeWidget()
        self.library_tree.setHeaderHidden(True)
        self.library_tree.setMinimumHeight(60)
        self.library_tree.setMaximumHeight(300)
        settings_layout.addWidget(self.library_tree)
        self.library_tree.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.library_tree.customContextMenuRequested.connect(self._show_library_context_menu)
        # Connect ONCE here (not in _refresh_library_tree, which runs repeatedly).
        self.library_tree.itemChanged.connect(self._on_library_item_changed)

        lib_btn_layout = QHBoxLayout()
        add_lib_btn = QPushButton("Add to Library…")
        add_lib_btn.clicked.connect(self.add_to_library)
        lib_all_btn = QPushButton("All")
        lib_all_btn.clicked.connect(lambda: self._set_all_library_checks(True))
        lib_none_btn = QPushButton("None")
        lib_none_btn.clicked.connect(lambda: self._set_all_library_checks(False))
        lib_refresh_btn = QPushButton("Refresh")
        lib_refresh_btn.clicked.connect(self._refresh_library_tree)
        for b in (add_lib_btn, lib_all_btn, lib_none_btn, lib_refresh_btn):
            lib_btn_layout.addWidget(b)
        settings_layout.addLayout(lib_btn_layout)
        self.library_selected_label = QLabel("Selected: 0 docs · ~0 tokens")
        settings_layout.addWidget(self.library_selected_label)

        clear_chat_btn = QPushButton("Clear Chat")
        clear_chat_btn.setToolTip("Clear the current conversation and start a new one")
        clear_chat_btn.clicked.connect(self.clear_current_chat)
        settings_layout.addWidget(clear_chat_btn)

        settings_layout.addStretch()

        settings_panel.setMinimumWidth(150)
        settings_panel.setMaximumWidth(350)
        self.settings_panel = settings_panel
        self.main_splitter.addWidget(settings_panel)

        # --- Chat Panel (Right) ---
        chat_panel = QFrame()
        chat_layout = QVBoxLayout(chat_panel)
        chat_layout.setContentsMargins(8, 8, 8, 8)
        chat_layout.setSpacing(0)

        # Search results overlay (hidden by default)
        self.search_results = SearchResultsWidget(theme=self.theme)
        self.search_results.hide()
        self.search_results.result_selected.connect(self.on_search_result_selected)

        # Vertical splitter for chat output and input
        self.chat_splitter = QSplitter(Qt.Orientation.Vertical)
        self.chat_splitter.setHandleWidth(6)  # Make handle easier to grab
        self.chat_splitter.setStyleSheet("""
            QSplitter::handle:vertical {
                background-color: #d0d0d0;
                height: 6px;
                margin: 2px 40px;
                border-radius: 3px;
            }
            QSplitter::handle:vertical:hover {
                background-color: #4CAF50;
            }
        """)
        chat_layout.addWidget(self.chat_splitter)

        # Chat history display (upper pane)
        self.chat_history = QTextBrowser()
        self.chat_history.setOpenExternalLinks(True)
        self.chat_history.setAcceptDrops(True)
        self.chat_history.setMinimumHeight(100)
        # Style tables/code/blockquote/etc. emitted by render_markdown — Qt's
        # QTextDocument renders bare <table> tags with no visible gridlines
        # until a default stylesheet is attached.
        self.chat_history.document().setDefaultStyleSheet(CHAT_MARKDOWN_CSS)
        self.chat_splitter.addWidget(self.chat_history)

        # Input area container (lower pane)
        input_container = QWidget()
        input_layout = QVBoxLayout(input_container)
        input_layout.setContentsMargins(0, 4, 0, 0)
        input_layout.setSpacing(4)

        # Context indicator
        self.context_indicator = ContextIndicator(theme=self.theme)
        self.context_indicator.clicked.connect(self.show_context_details)
        input_layout.addWidget(self.context_indicator)

        # Toolbar row
        toolbar_layout = QHBoxLayout()

        # Legal Research checkbox
        self.legal_research_check = QCheckBox("Legal Research")
        self.legal_research_check.setStyleSheet("font-size: 11px;")
        self.legal_research_check.setToolTip(
            "Search CA case law and statutes, inject verified citations into response"
        )
        toolbar_layout.addWidget(self.legal_research_check)

        self.research_sources_btn = QToolButton()
        self.research_sources_btn.setPopupMode(QToolButton.ToolButtonPopupMode.InstantPopup)
        self.research_sources_btn.setToolTip("Choose legal research sources")
        self._build_research_sources_menu()
        toolbar_layout.addWidget(self.research_sources_btn)

        toolbar_layout.addStretch()

        # Attachment indicator
        self.attachment_label = QLabel("")
        self.attachment_label.setStyleSheet("color: #666; font-size: 11px;")
        toolbar_layout.addWidget(self.attachment_label)

        input_layout.addLayout(toolbar_layout)

        # Input row with text and buttons
        input_row = QHBoxLayout()

        self.chat_input = QPlainTextEdit()
        self.chat_input.setMinimumHeight(40)
        self.chat_input.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
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
        input_layout.addLayout(input_row, 1)  # stretch=1 so input fills available space

        input_container.setMinimumHeight(100)
        self.input_container = input_container
        self.chat_splitter.addWidget(input_container)

        # Set chat splitter default sizes (output large, input smaller)
        self.chat_splitter.setSizes([500, 150])

        self.main_splitter.addWidget(chat_panel)

        # Set main splitter default sizes
        self.main_splitter.setSizes([200, 200, 800])

        # Connect splitter signals for persistence
        self.main_splitter.splitterMoved.connect(self._on_main_splitter_moved)
        self.chat_splitter.splitterMoved.connect(self._on_chat_splitter_moved)

        # Load saved splitter sizes
        self._load_splitter_sizes()

    # --- Persistence Methods ---

    def load_case(self, file_number: str):
        """Load conversations for a case. Called when case switches."""
        # Stop any running threads to prevent "QThread destroyed while running" errors
        if self.worker is not None and self.worker.isRunning():
            self.worker.request_stop()
            self.worker.wait(1000)

        if self.chat_research_worker is not None and self.chat_research_worker.isRunning():
            self.chat_research_worker.request_stop()
            self.chat_research_worker.wait(1000)

        if self.fetcher is not None and self.fetcher.isRunning():
            self.fetcher.wait(1000)

        # Clear mediation brief state
        self.med_brief_generator = None
        if self.med_brief_worker and self.med_brief_worker.isRunning():
            self.med_brief_worker.request_stop()
            self.med_brief_worker.wait(1000)
        self.med_brief_worker = None
        if hasattr(self, 'add_quotes_btn') and self.add_quotes_btn:
            self.add_quotes_btn.setVisible(False)

        # Clear attached files from previous case (without persisting)
        self._clear_files_no_persist()

        self.file_number = file_number
        self.persistence = ChatPersistence(file_number)

        # Refresh template menu now that persistence is available
        self.update_template_menu()

        # Restore persisted attached files for this case
        self._restore_attached_files()

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

        # Repopulate the Saved Documents library tree for the new case
        self._refresh_library_tree()
        if getattr(self, "persistence", None):
            try:
                saved = self.persistence.get_setting("library_selected_ids", [])
            except Exception:
                saved = []
            self._restore_checked_entry_ids(saved)

    # ------------------------------------------------------------------
    # Saved Documents (document text library)
    # ------------------------------------------------------------------
    def _pick_files_for_library(self):
        from PySide6.QtWidgets import QFileDialog
        paths, _ = QFileDialog.getOpenFileNames(
            self, "Add documents to library", "",
            "Documents (*.pdf *.docx *.doc *.txt *.msg);;All files (*.*)")
        return list(paths)

    def add_to_library(self):
        lib = self._library()
        if lib is None:
            QMessageBox.information(self, "No Case Loaded",
                                    "Load a case before adding to its library.")
            return
        paths = self._pick_files_for_library()
        if not paths:
            return
        prior = getattr(self, "_lib_thread", None)
        if prior is not None:
            try:
                if prior.isRunning():
                    return
            except RuntimeError:
                pass  # prior C++ object already deleted; safe to proceed
        from PySide6.QtCore import QThread, Signal, QObject

        class _Worker(QObject):
            done = Signal()

            def __init__(self, lib, paths):
                super().__init__()
                self._lib, self._paths = lib, paths

            def run(self):
                try:
                    self._lib.add_entry("manual", self._paths, {})
                finally:
                    self.done.emit()

        self._lib_thread = QThread()
        self._lib_worker = _Worker(lib, paths)
        self._lib_worker.moveToThread(self._lib_thread)
        self._lib_thread.started.connect(self._lib_worker.run)
        self._lib_worker.done.connect(self._lib_thread.quit)
        self._lib_worker.done.connect(self._refresh_library_tree)
        self._lib_thread.finished.connect(self._lib_worker.deleteLater)
        self._lib_thread.finished.connect(self._lib_thread.deleteLater)
        self._lib_thread.start()

    def _library(self):
        """Return a DocumentLibrary for the current case, or None."""
        from ..doc_library.library import DocumentLibrary
        root = getattr(self, "_case_root_for_library", None) \
            or getattr(self.window(), "case_path", None)
        return DocumentLibrary(root) if root else None

    def _show_library_context_menu(self, pos):
        from PySide6.QtCore import Qt
        from PySide6.QtWidgets import QMenu
        item = self.library_tree.itemAt(pos)
        if item is None:
            return
        data = item.data(0, Qt.ItemDataRole.UserRole) or {}
        if data.get("kind") != "entry":
            item = item.parent()
            data = item.data(0, Qt.ItemDataRole.UserRole) or {} if item else {}
        if data.get("kind") != "entry":
            return
        entry_id = data.get("id")
        menu = QMenu(self)
        rename_act = menu.addAction("Rename")
        reset_act = menu.addAction("Reset to auto name")
        menu.addSeparator()
        remove_act = menu.addAction("Remove from library")
        chosen = menu.exec(self.library_tree.viewport().mapToGlobal(pos))
        if chosen is None:
            return
        if chosen == rename_act:
            self.library_tree.editItem(item, 0)
        elif chosen == reset_act:
            self._reset_library_entry(entry_id)
        elif chosen == remove_act:
            self._delete_library_entry(entry_id, confirm=True)

    def _reset_library_entry(self, entry_id):
        lib = self._library()
        if lib is None:
            return
        try:
            lib.reset_label(entry_id)
        except Exception:
            pass
        self._refresh_library_tree()

    def _delete_library_entry(self, entry_id, confirm=True):
        lib = self._library()
        if lib is None:
            return
        if confirm:
            reply = QMessageBox.question(
                self, "Remove from library",
                "Remove this document from the library? The saved text will be deleted.",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No)
            if reply != QMessageBox.StandardButton.Yes:
                return
        try:
            lib.delete_entry(entry_id)
        except Exception:
            pass
        self._refresh_library_tree()

    @staticmethod
    def refresh_open_library_trees(tabs_widget):
        """Refresh the Saved Documents tree on every ChatTab in a QTabWidget.

        Called when a background task-completion capture finishes, so an
        already-open Chat tab shows the newly-saved document without the user
        having to click Refresh. Best-effort and signal-safe (runs on GUI thread).
        """
        if tabs_widget is None:
            return
        for i in range(tabs_widget.count()):
            w = tabs_widget.widget(i)
            if isinstance(w, ChatTab):
                try:
                    w._refresh_library_tree()
                except Exception:
                    pass

    def _refresh_library_tree(self):
        from PySide6.QtCore import Qt
        try:
            _preserve_ids = self._collect_checked_entry_ids()
        except Exception:
            _preserve_ids = []
        self.library_tree.blockSignals(True)
        try:
            self.library_tree.clear()
            lib = self._library()
            entries = []
            if lib is not None:
                try:
                    entries = lib.list_entries()
                except Exception:
                    entries = []
            Item = self._QTreeWidgetItem
            for e in entries:
                top = Item([e.label])
                top.setData(0, Qt.ItemDataRole.UserRole, {"kind": "entry", "id": e.id, "label": e.label})
                top.setFlags(top.flags() | Qt.ItemFlag.ItemIsUserCheckable
                             | Qt.ItemFlag.ItemIsEditable
                             | Qt.ItemFlag.ItemIsAutoTristate)
                top.setCheckState(0, Qt.CheckState.Unchecked)
                for m in e.members:
                    label = m.source_name + (" [extract failed]" if m.error else "")
                    child = Item([label])
                    child.setData(0, Qt.ItemDataRole.UserRole,
                                  {"kind": "member", "blob": m.blob,
                                   "name": m.source_name, "tokens": m.est_tokens})
                    child.setFlags(child.flags() | Qt.ItemFlag.ItemIsUserCheckable)
                    child.setCheckState(0, Qt.CheckState.Unchecked)
                    top.addChild(child)
                self.library_tree.addTopLevelItem(top)
            # Default to collapsed: show entries, hide member files until the
            # user expands an entry.
            self.library_tree.collapseAll()
        finally:
            self.library_tree.blockSignals(False)
        # Re-apply the checked selection captured before the rebuild.
        # _restore_checked_entry_ids handles blockSignals and calls
        # _update_library_selected_label() once at the end.
        self._restore_checked_entry_ids(_preserve_ids)

    def _set_all_library_checks(self, checked: bool):
        from PySide6.QtCore import Qt
        state = Qt.CheckState.Checked if checked else Qt.CheckState.Unchecked
        for i in range(self.library_tree.topLevelItemCount()):
            self.library_tree.topLevelItem(i).setCheckState(0, state)

    def _on_library_item_changed(self, item, column):
        from PySide6.QtCore import Qt
        data = item.data(0, Qt.ItemDataRole.UserRole) or {}
        # Inline rename of an entry (top-level item, text actually changed).
        if data.get("kind") == "entry" and column == 0:
            new_label = item.text(0).strip()
            if new_label and new_label != data.get("label"):
                lib = self._library()
                if lib:
                    try:
                        lib.rename_entry(data["id"], new_label)
                        data["label"] = new_label
                        item.setData(0, Qt.ItemDataRole.UserRole, data)
                    except Exception:
                        pass
        self._update_library_selected_label()
        self._persist_library_selection()

    def _collect_checked_entry_ids(self):
        from PySide6.QtCore import Qt
        ids = []
        for i in range(self.library_tree.topLevelItemCount()):
            top = self.library_tree.topLevelItem(i)
            if top.checkState(0) in (Qt.CheckState.Checked,
                                     Qt.CheckState.PartiallyChecked):
                data = top.data(0, Qt.ItemDataRole.UserRole) or {}
                if data.get("id"):
                    ids.append(data["id"])
        return ids

    def _restore_checked_entry_ids(self, ids):
        from PySide6.QtCore import Qt
        wanted = set(ids or [])
        self.library_tree.blockSignals(True)
        try:
            for i in range(self.library_tree.topLevelItemCount()):
                top = self.library_tree.topLevelItem(i)
                data = top.data(0, Qt.ItemDataRole.UserRole) or {}
                if data.get("id") in wanted:
                    top.setCheckState(0, Qt.CheckState.Checked)
        finally:
            self.library_tree.blockSignals(False)
        self._update_library_selected_label()

    def _persist_library_selection(self):
        if not getattr(self, "persistence", None):
            return
        try:
            self.persistence.set_setting(
                "library_selected_ids", self._collect_checked_entry_ids())
        except Exception:
            pass

    def _iter_checked_library_members(self):
        from PySide6.QtCore import Qt
        for i in range(self.library_tree.topLevelItemCount()):
            top = self.library_tree.topLevelItem(i)
            for j in range(top.childCount()):
                child = top.child(j)
                if child.checkState(0) == Qt.CheckState.Checked:
                    yield child.data(0, Qt.ItemDataRole.UserRole)

    def read_library_content(self) -> str:
        lib = self._library()
        if lib is None:
            return ""
        content = ""
        for m in self._iter_checked_library_members():
            blob = m.get("blob")
            if not blob:
                continue
            text = lib.get_member_text(blob)
            if not text:
                content += f"\n--- FILE: {m.get('name', 'document')} ---\n[saved text unavailable]\n"
                continue
            content += f"\n--- FILE: {m.get('name', 'document')} ---\n{text}\n"
        return content

    def _context_limit(self) -> int:
        forced = getattr(self, "_context_limit_for_test", None)
        if forced is not None:
            return forced
        from icharlotte_core.chat.token_counter import TokenCounter
        return TokenCounter.get_context_limit(
            self.model_combo.currentText(), self.provider_combo.currentText())

    def _library_budget_warning(self, file_content: str, history_tokens: int) -> str:
        from icharlotte_core.chat.token_counter import TokenCounter
        reserve = 16384
        used = (TokenCounter.estimate_tokens(file_content) + history_tokens + reserve)
        limit = self._context_limit()
        if used <= limit:
            return None
        return (f"The selected documents (~{TokenCounter.format_token_count(used)} "
                f"tokens) exceed this model's context window "
                f"(~{TokenCounter.format_token_count(limit)}). The request may be "
                f"truncated or rejected. Deselect some Saved Documents, or proceed "
                f"anyway?")

    def _update_library_selected_label(self):
        members = list(self._iter_checked_library_members())
        toks = sum(int(m.get("tokens", 0)) for m in members)
        self.library_selected_label.setText(
            f"Selected: {len(members)} docs · ~{toks} tokens")

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
        log_event(f"Selecting conversation: {conv_id}")
        if not self.persistence:
            log_event("Cannot select conversation - no persistence", "warning")
            return

        self.save_current_state()
        self.current_conversation_id = conv_id
        self.current_conversation = self.persistence.get_conversation(conv_id)

        if self.current_conversation:
            log_event(f"Loaded conversation '{self.current_conversation.name}' with {len(self.current_conversation.messages)} messages")
        else:
            log_event(f"ERROR: Could not find conversation {conv_id}", "error")

        self.conv_sidebar.set_current_conversation(conv_id)

        if self.current_conversation:
            # Restore settings
            self.provider_combo.setCurrentText(self.current_conversation.provider)
            # Model will be set after provider change triggers model fetch
            model_to_restore = self.current_conversation.model
            def restore_model_selection(model=model_to_restore):
                try:
                    self.model_combo.setCurrentText(model)
                except RuntimeError:
                    # The tab may have been closed before the delayed model
                    # selection fires.
                    pass
            QTimer.singleShot(500, restore_model_selection)
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
            log_event("Cannot save - no persistence or conversation ID", "warning")
            QMessageBox.warning(self, "Cannot Save", "No active conversation to save.")
            return

        # Save current state
        self.save_current_state()

        # Check if conversation has messages and might need a name
        conv = self.persistence.get_conversation(self.current_conversation_id)
        if not conv:
            log_event(f"Cannot save - conversation {self.current_conversation_id} not found", "error")
            QMessageBox.warning(self, "Error", "Conversation not found.")
            return

        log_event(f"Saving conversation '{conv.name}' with {len(conv.messages)} messages")

        if conv.messages:
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
                    log_event(f"Renamed conversation to '{new_name.strip()}'")
        else:
            # Conversation has no messages - still allow naming
            new_name, ok = QInputDialog.getText(
                self, "Name Conversation",
                "This conversation has no messages yet.\nEnter a name for this conversation:",
                text=conv.name
            )
            if ok and new_name.strip():
                self.persistence.rename_conversation(self.current_conversation_id, new_name.strip())
                log_event(f"Named empty conversation '{new_name.strip()}'")

        # Refresh the sidebar to show the saved conversation
        self.refresh_conversation_list()
        self.conv_sidebar.set_current_conversation(self.current_conversation_id)

        # Show confirmation
        msg_count = len(conv.messages)
        QMessageBox.information(self, "Saved", f"Conversation saved with {msg_count} message(s).")

    def clear_current_chat(self):
        """Clear the chat display without saving or creating new conversations."""
        # Just clear the display
        self.clear_chat_display()
        self.conversation_history = []

        # Delete messages from current conversation in persistence (if any)
        if self.persistence and self.current_conversation_id:
            # Clear messages from the conversation data
            data = self.persistence.load()
            for conv_data in data.get('conversations', []):
                if conv_data.get('id') == self.current_conversation_id:
                    conv_data['messages'] = []
                    conv_data['total_tokens_used'] = 0
                    break
            self.persistence.save()
            # Reload the conversation object
            self.current_conversation = self.persistence.get_conversation(self.current_conversation_id)

    def load_conversation_messages(self):
        """Load and display messages from current conversation."""
        self.clear_chat_display()

        if not self.current_conversation:
            self.chat_history.append("<i style='color: #888;'>No conversation selected.</i>")
            return

        if not self.current_conversation.messages:
            # Show a helpful message for empty conversations
            conv_name = self.current_conversation.name
            self.chat_history.append(f"<i style='color: #888;'>Conversation '{conv_name}' loaded. Start chatting!</i>")
            self.chat_history.append("-" * 50)
            log_event(f"Loaded empty conversation: {self.current_conversation_id}")
            return

        log_event(f"Loading {len(self.current_conversation.messages)} messages from conversation {self.current_conversation_id}")
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
            if "\n\n[ATTACHED FILES]:\n" in content:
                content = content.split("\n\n[ATTACHED FILES]:\n", 1)[0]
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
                html_text = render_markdown(content)
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

        # Not in cache, fetch dynamically. The fetcher falls back to a curated
        # local list when a provider key is absent or model listing fails.
        api_key = API_KEYS.get(provider)
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

        # One-time seed of the Workbench-configured default model (agent_chat).
        seed = getattr(self, "_seed_model", None)
        if seed:
            self._seed_model = None
            idx = self.model_combo.findText(seed)
            if idx != -1:
                self.model_combo.setCurrentIndex(idx)
                return

        # Set default model
        if provider == "Gemini":
            idx = self.model_combo.findText("gemini-3.1-flash-lite")
            if idx == -1:
                idx = self.model_combo.findText("gemini-3.5-flash")
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
        dlg = SettingsDialog(
            self.settings,
            self,
            selected_model=self.model_combo.currentText()
        )
        if dlg.exec():
            self.settings = dlg.get_settings()

    def open_sys_prompt(self):
        dlg = SystemPromptDialog(self.system_prompt, self)
        if dlg.exec():
            self.system_prompt = dlg.get_prompt()

    def select_files(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, "Select Files", "",
            "All Supported (*.pdf *.docx *.doc *.txt *.msg *.png *.jpg *.jpeg *.gif *.webp *.mp3 *.mp4 *.m4a *.wav *.ogg *.flac *.aac *.wma *.avi *.mkv *.mov *.webm);;"
            "Documents (*.pdf *.docx *.doc *.txt *.msg);;"
            "Images (*.png *.jpg *.jpeg *.gif *.webp);;"
            "Audio/Video (*.mp3 *.mp4 *.m4a *.wav *.ogg *.flac *.aac *.wma *.avi *.mkv *.mov *.webm)"
        )
        for f in files:
            self.add_file(f)

    def add_file(self, path):
        """Add a file to the attachment list."""
        final_path = path
        ext = os.path.splitext(path)[1].lower()

        # PDF Check Logic — non-blocking OCR
        needs_ocr = False
        if ext == ".pdf":
            needs_ocr = self._check_pdf_needs_ocr(path)
            if needs_ocr is None:
                return  # User cancelled

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
                except Exception:
                    item.setIcon(self.icon_provider.icon(QFileInfo(final_path)))
            else:
                item.setIcon(self.icon_provider.icon(QFileInfo(final_path)))

            item.setToolTip(final_path)
            item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
            item.setCheckState(Qt.CheckState.Checked)
            item.setData(Qt.ItemDataRole.UserRole, final_path)
            self.file_list.addItem(item)

            # Persist the added file
            if self.persistence:
                self.persistence.add_attached_file(final_path)

            # Launch background OCR if needed (after file is attached)
            if needs_ocr:
                self._start_background_ocr(path, item)

    def _set_all_file_checks(self, state):
        """Set all file list items to the given check state."""
        for i in range(self.file_list.count()):
            self.file_list.item(i).setCheckState(state)

    def clear_files(self):
        """Clear all attached files and persist the change."""
        if self.attached_files is None:
            self.attached_files = []
        else:
            self.attached_files.clear()
        self.file_list.clear()
        # Persist the cleared state
        if self.persistence:
            self.persistence.clear_attached_files()

    def import_carrier_reports(self):
        """Scan {case_path}/STATUS/ for carrier00X.doc(x) files and attach them.

        Matches carrier001..carrier015 with optional trailing text after the
        number. Rejects filenames with any prefix before "carrier". Adds matches
        to the existing attached-files list (deduped), leaving current files
        intact. Shows a popup summarizing the result.
        """
        main_win = self.window()
        case_path = getattr(main_win, 'case_path', None)
        if not case_path:
            QMessageBox.warning(
                self,
                "Import Reports",
                "No case is currently selected.",
            )
            return

        status_dir = os.path.join(case_path, "STATUS")
        if not os.path.isdir(status_dir):
            QMessageBox.warning(
                self,
                "Import Reports",
                f"No STATUS folder found at:\n{status_dir}",
            )
            return

        try:
            entries = os.listdir(status_dir)
        except (PermissionError, OSError) as e:
            QMessageBox.warning(
                self,
                "Import Reports",
                f"Could not read STATUS folder:\n{e}",
            )
            return

        imported = 0
        already_attached = 0
        failed = []
        attached_normcase = {os.path.normcase(p) for p in self.attached_files}
        for name in entries:
            if not CARRIER_REPORT_RE.match(name):
                continue
            full_path = os.path.join(status_dir, name)
            if os.path.normcase(full_path) in attached_normcase:
                already_attached += 1
                continue
            try:
                self.add_file(full_path)
            except OSError as e:
                failed.append(f"{name}: {e}")
                continue
            attached_normcase.add(os.path.normcase(full_path))
            imported += 1

        if imported == 0 and already_attached == 0 and not failed:
            QMessageBox.information(
                self,
                "Import Reports",
                "No carrier reports (carrier001\u2013carrier015) found in STATUS.",
            )
        elif imported == 0 and already_attached > 0 and not failed:
            QMessageBox.information(
                self,
                "Import Reports",
                f"All {already_attached} matching report(s) were already attached.",
            )
        else:
            msg = f"Imported {imported} carrier report(s) from STATUS."
            if already_attached > 0:
                msg += f"\n({already_attached} already attached, skipped.)"
            if failed:
                msg += f"\n\nFailed to import {len(failed)} file(s):\n" + "\n".join(failed)
                QMessageBox.warning(self, "Import Reports", msg)
            else:
                QMessageBox.information(self, "Import Reports", msg)

    def _clear_files_no_persist(self):
        """Clear attached files from UI without persisting (used when switching cases)."""
        if self.attached_files is None:
            self.attached_files = []
        else:
            self.attached_files.clear()
        self.file_list.clear()

    def _restore_attached_files(self):
        """Restore attached files from persistence."""
        if not self.persistence:
            return
        persisted_files = self.persistence.get_attached_files()
        for file_path in persisted_files:
            # Only add files that still exist
            if os.path.exists(file_path):
                self._add_file_to_ui(file_path)
            else:
                # Remove non-existent files from persistence
                self.persistence.remove_attached_file(file_path)

    def _add_file_to_ui(self, final_path):
        """Add a file to the UI without OCR checks (for restoring persisted files)."""
        if final_path not in self.attached_files:
            self.attached_files.append(final_path)
            item = QListWidgetItem(os.path.basename(final_path))

            ext = os.path.splitext(final_path)[1].lower()
            # Use custom icon for images, system icon for others
            if ext in ['.png', '.jpg', '.jpeg', '.gif', '.webp']:
                try:
                    pixmap = QPixmap(final_path)
                    if not pixmap.isNull():
                        pixmap = pixmap.scaled(32, 32, Qt.AspectRatioMode.KeepAspectRatio)
                        item.setIcon(pixmap)
                    else:
                        item.setIcon(self.icon_provider.icon(QFileInfo(final_path)))
                except Exception:
                    item.setIcon(self.icon_provider.icon(QFileInfo(final_path)))
            else:
                item.setIcon(self.icon_provider.icon(QFileInfo(final_path)))

            item.setToolTip(final_path)
            item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
            item.setCheckState(Qt.CheckState.Checked)
            item.setData(Qt.ItemDataRole.UserRole, final_path)
            self.file_list.addItem(item)

    def _check_pdf_needs_ocr(self, path):
        """Check if PDF needs OCR. Returns True (needs OCR), False (has text), or None (user cancelled)."""
        if not fitz:
            return False

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
                    f"The file '{os.path.basename(path)}' appears to be an image/scanned PDF.\n"
                    f"Do you want to OCR it in the background?",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No | QMessageBox.StandardButton.Cancel
                )
                if reply == QMessageBox.StandardButton.Yes:
                    return True
                elif reply == QMessageBox.StandardButton.Cancel:
                    return None
                else:
                    return False  # Attach as-is
            return False
        except Exception as e:
            log_event(f"Error checking PDF: {e}", "error")
            return False

    def _start_background_ocr(self, path, list_item):
        """Launch OCR in background with status in the Status tab."""
        from .widgets import StatusWidget

        script_path = os.path.join(SCRIPTS_DIR, "ocr.py")
        if not os.path.exists(script_path):
            log_event("ocr.py not found — skipping background OCR", "error")
            return

        # Mark the list item as pending OCR
        basename = os.path.basename(path)
        list_item.setText(f"{basename} (OCR in progress...)")
        list_item.setForeground(Qt.GlobalColor.gray)

        # Create status widget in the Status tab
        status_widget = StatusWidget("OCR", f"Processing {basename}")

        main_window = self.window()
        if hasattr(main_window, 'status_list_layout'):
            main_window.status_list_layout.insertWidget(0, status_widget)

        # Create and start the OCR runner
        runner = OCRRunner(script_path, path)

        # Keep references so they don't get GC'd
        if not hasattr(self, '_ocr_runners'):
            self._ocr_runners = []
        self._ocr_runners.append(runner)

        # Connect signals
        runner.progress.connect(lambda pct: status_widget.update_progress(pct, f"OCR {pct}%"))
        runner.finished.connect(lambda success, message, final_path:
            self._on_background_ocr_finished(success, message, final_path, path, list_item, status_widget, runner))

        # Connect cancel
        status_widget.cancel_requested.connect(lambda: self._cancel_background_ocr(runner, list_item, path, status_widget))

        runner.start()
        log_event(f"Background OCR started: {path}", "info")

    def _on_background_ocr_finished(self, success, message, final_path, original_path, list_item, status_widget, runner):
        """Handle background OCR completion — swap the attachment path."""
        basename = os.path.basename(original_path)

        if success and final_path and final_path != original_path:
            # Update the attachment list entry
            list_item.setText(os.path.basename(final_path))
            list_item.setForeground(QBrush())  # Reset to theme default
            list_item.setToolTip(final_path)
            list_item.setData(Qt.ItemDataRole.UserRole, final_path)

            # Update internal tracking
            if original_path in self.attached_files:
                idx = self.attached_files.index(original_path)
                self.attached_files[idx] = final_path

            # Update persistence
            if self.persistence:
                self.persistence.remove_attached_file(original_path)
                self.persistence.add_attached_file(final_path)

            status_widget.set_output_file(final_path)
            status_widget.set_finished(True)
            status_widget.append_log(f"OCR completed: {final_path}\n")
            log_event(f"Background OCR completed: {original_path} → {final_path}", "info")
        elif success:
            # OCR succeeded but path didn't change (text was added in-place)
            list_item.setText(basename)
            list_item.setForeground(QBrush())  # Reset to theme default
            status_widget.set_finished(True)
            status_widget.append_log(f"OCR completed: {original_path}\n")
            log_event(f"Background OCR completed (in-place): {original_path}", "info")
        else:
            # OCR failed — leave the original file attached
            list_item.setText(f"{basename} (OCR failed)")
            list_item.setForeground(Qt.GlobalColor.red)
            status_widget.set_finished(False)
            status_widget.append_log(f"OCR failed: {message}\n")
            log_event(f"Background OCR failed: {original_path} — {message}", "error")

        # Clean up runner reference
        if hasattr(self, '_ocr_runners') and runner in self._ocr_runners:
            self._ocr_runners.remove(runner)

    def _cancel_background_ocr(self, runner, list_item, original_path, status_widget):
        """Cancel a running background OCR."""
        runner.terminate()
        runner.wait()
        basename = os.path.basename(original_path)
        list_item.setText(basename)
        list_item.setForeground(QBrush())  # Reset to theme default
        status_widget.set_finished(False)
        status_widget.append_log("OCR cancelled by user.\n")
        log_event(f"Background OCR cancelled: {original_path}", "warning")
        if hasattr(self, '_ocr_runners') and runner in self._ocr_runners:
            self._ocr_runners.remove(runner)

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
        supported_extensions = (".pdf", ".docx", ".doc", ".txt", ".msg", ".png", ".jpg", ".jpeg", ".gif", ".webp",
                                ".mp3", ".mp4", ".m4a", ".wav", ".ogg", ".flac", ".aac", ".wma", ".avi", ".mkv", ".mov", ".webm")
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
                        from ..document_processor import extract_docx_text
                        content += extract_docx_text(path) + "\n"
                    elif ext == ".doc":
                        content += self._extract_doc_text(path)
                    elif ext == ".pdf":
                        if fitz:
                            doc = fitz.open(path)
                            for page in doc:
                                content += page.get_text() + "\n"
                            doc.close()
                    elif ext == ".msg":
                        try:
                            import pythoncom
                            import win32com.client
                            pythoncom.CoInitialize()
                            outlook = win32com.client.Dispatch("Outlook.Application")
                            namespace = outlook.GetNamespace("MAPI")
                            item = namespace.OpenSharedItem(os.path.abspath(path))
                            subject = item.Subject or ""
                            sender = item.SenderName or ""
                            body = item.Body or ""
                            if not body.strip() and item.HTMLBody:
                                import re
                                body = re.sub(r'<[^>]+>', '', item.HTMLBody).strip()
                            item.Close(0)
                            pythoncom.CoUninitialize()
                            content += f"From: {sender}\nSubject: {subject}\n\n{body}\n"
                        except Exception as msg_err:
                            content += f"[Error reading .msg file: {msg_err}]\n"
                    elif ext in self.AUDIO_VIDEO_EXTENSIONS:
                        file_size = os.path.getsize(path)
                        size_mb = file_size / (1024 * 1024)
                        content += f"[Audio/Video file: {os.path.basename(path)}, {size_mb:.1f} MB]\n"
                        content += "[WARNING: This model cannot process audio/video content]\n"
                except Exception as e:
                    content += f"[Error reading file: {e}]\n"
        return content

    AUDIO_VIDEO_EXTENSIONS = {'.mp3', '.mp4', '.m4a', '.wav', '.ogg', '.flac', '.aac', '.wma',
                              '.avi', '.mkv', '.mov', '.webm'}

    @staticmethod
    def _extract_doc_text(path: str) -> str:
        """Extract text from a legacy .doc file via Word COM automation.

        Deliberately never calls ``word.Quit()`` and never touches
        ``word.Visible`` — the user may have Word open with other documents,
        and the global safety rule is that we must not close or hide it. We
        open the .doc read-only, pull its text, and close only the document
        we opened.
        """
        try:
            import pythoncom
            import win32com.client
        except ImportError as e:
            return f"[Error reading .doc file: win32com not available: {e}]\n"

        pythoncom.CoInitialize()
        try:
            try:
                word = win32com.client.Dispatch("Word.Application")
            except Exception as e:
                return f"[Error reading .doc file: could not launch Word: {e}]\n"

            doc = None
            try:
                abs_path = os.path.abspath(path)
                doc = word.Documents.Open(
                    FileName=abs_path,
                    ConfirmConversions=False,
                    ReadOnly=True,
                    AddToRecentFiles=False,
                )
                text = doc.Content.Text or ""
                return text if text.endswith("\n") else text + "\n"
            except Exception as e:
                return f"[Error reading .doc file: {e}]\n"
            finally:
                if doc is not None:
                    try:
                        doc.Close(SaveChanges=0)
                    except Exception:
                        pass
                # Intentionally NOT calling word.Quit() — never close the
                # user's Word session (global safety rule).
        finally:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    def stop_generation(self):
        """Stop the current generation."""
        if (
            self.chat_research_worker is not None
            and self.chat_research_worker.isRunning()
        ):
            self.chat_research_worker.request_stop()
            self.chat_history.append("<i>[Legal research stop requested]</i>")
            self.stop_btn.setEnabled(False)
            return

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
        # --- Mediation brief refinement mode ---
        if (self.med_brief_generator and self.med_brief_generator.is_active
                and (not self.med_brief_worker or not self.med_brief_worker.isRunning())):
            user_text = self.chat_input.toPlainText().strip()
            if user_text:
                self._route_brief_refinement(user_text)
                return
        # --- End mediation brief check ---

        # Warn if no case is loaded (messages won't persist)
        if not self.persistence or not self.current_conversation_id:
            reply = QMessageBox.warning(
                self, "No Case Loaded",
                "No case is loaded. Your chat will not be saved.\n\nDo you want to continue anyway?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No
            )
            if reply != QMessageBox.StandardButton.Yes:
                return

        user_text = self.chat_input.toPlainText().strip()

        checked_count = 0
        for i in range(self.file_list.count()):
            if self.file_list.item(i).checkState() == Qt.CheckState.Checked:
                checked_count += 1

        lib_checked = len(list(self._iter_checked_library_members()))
        if not user_text and checked_count == 0 and lib_checked == 0:
            return

        # Display User Message
        self.chat_history.append(f"<b>You:</b> {user_text}")
        if checked_count > 0:
            self.chat_history.append(f"<i>(Attached {checked_count} files)</i>")

        self.chat_input.clear()
        self.send_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)

        # Prepare Content
        file_content = self.read_files_content() + self.read_library_content()
        warn = self._library_budget_warning(file_content, history_tokens=0)
        if warn:
            reply = QMessageBox.warning(
                self, "Context Budget", warn,
                QMessageBox.StandardButton.Ok | QMessageBox.StandardButton.Cancel,
                QMessageBox.StandardButton.Cancel)
            if reply == QMessageBox.StandardButton.Cancel:
                self.send_btn.setEnabled(True)
                self.stop_btn.setEnabled(False)
                return
        attachments = self.get_attachment_info()

        # Detect audio/video attachments
        audio_files = self._get_checked_audio_files()
        current_provider = self.provider_combo.currentText()
        current_model = self.model_combo.currentText()
        current_settings = dict(self.settings)

        # Safety: warn for non-Gemini providers (they can't process media)
        if audio_files and current_provider != "Gemini":
            file_names = ", ".join(os.path.basename(f) for f in audio_files)
            reply = QMessageBox.warning(
                self, "Audio/Video Files Detected",
                f"You have audio/video files attached:\n{file_names}\n\n"
                f"{current_provider} models cannot process audio/video and will "
                f"likely hallucinate a fake transcript.\n\n"
                f"Switch to Gemini for native audio/video support, or click "
                f"No to send anyway.",
                QMessageBox.StandardButton.No | QMessageBox.StandardButton.Cancel,
                QMessageBox.StandardButton.Cancel
            )
            if reply == QMessageBox.StandardButton.Cancel:
                self.send_btn.setEnabled(True)
                self.stop_btn.setEnabled(False)
                return
            # User chose No — clear audio_files so we don't try to upload
            audio_files = []

        # Legal Research: if checked, run selected research before LLM call.
        if self.legal_research_check.isChecked():
            self._start_chat_legal_research(
                user_text,
                file_content,
                provider=current_provider,
                model=current_model,
                settings=current_settings,
                attachments=attachments,
                audio_files=audio_files,
            )
            return

        self._continue_send_message_after_research(
            user_text,
            file_content,
            attachments=attachments,
            audio_files=audio_files,
            provider=current_provider,
            model=current_model,
            settings=current_settings,
            research_packet=None,
        )

    def _continue_send_message_after_research(
        self,
        user_text,
        file_content,
        *,
        attachments,
        audio_files,
        provider,
        model,
        settings,
        research_packet,
    ):
        self._pending_research = research_packet

        # Build user message for history
        full_msg = user_text
        if file_content:
            full_msg += "\n\n[ATTACHED FILES]:\n" + file_content

        # Save user message to persistence
        if self.persistence and self.current_conversation_id:
            token_count = TokenCounter.estimate_tokens(full_msg, provider)
            user_message = Message(
                role='user',
                content=full_msg,
                attachments=attachments,
                token_count=token_count
            )
            log_event(f"Saving user message to conversation {self.current_conversation_id}")
            self.persistence.add_message(self.current_conversation_id, user_message)
        else:
            log_event(f"WARNING: Cannot save user message - persistence={self.persistence is not None}, conv_id={self.current_conversation_id}", "warning")

        # Update legacy history
        self.conversation_history.append({'role': 'user', 'content': full_msg})

        # Enable streaming
        worker_settings = {**settings, 'stream': True}

        # Initialize streaming state
        self.stream_text = ""
        self.stream_start_time = time.time()

        # Show initial "thinking" indicator that will be replaced
        self.chat_history.append("<b>AI:</b> ")
        cursor = self.chat_history.textCursor()
        cursor.movePosition(QTextCursor.MoveOperation.End)
        self.stream_start_pos = cursor.position()

        # Sending jumps to the latest, so the response follows from the start
        # (until the user deliberately scrolls up).
        self._chat_scroll_to_bottom()

        # Build system prompt (augment with legal authority if research was done)
        effective_system_prompt = self.system_prompt
        if research_packet:
            effective_system_prompt = research_packet.build_augmented_system_prompt(
                self.system_prompt
            )

        # Start Worker (pass media files for Gemini native upload)
        media_for_upload = audio_files if (audio_files and provider == "Gemini") else None
        if media_for_upload:
            self.chat_history.append("<i>Uploading media to Gemini...</i>")
            QApplication.processEvents()

        self.worker = LLMWorker(
            provider,
            model,
            effective_system_prompt,
            user_text,
            file_content,
            worker_settings,
            history=list(self.conversation_history[:-1]),  # Exclude the message we just added
            media_files=media_for_upload
        )

        self.worker.new_token.connect(self.on_streaming_token)
        self.worker.finished.connect(self.on_stream_complete)
        self.worker.error.connect(self.on_error)
        self.worker.start()

        self.update_context_indicator()

    def _start_chat_legal_research(
        self,
        user_text,
        file_content,
        *,
        provider,
        model,
        settings,
        attachments,
        audio_files,
    ):
        from icharlotte_core import task_debug

        research_settings = self._current_chat_research_settings()
        debug_run_id = task_debug.start_run(
            task_id="chat_legal_research",
            task_title="Chat Legal Research",
            source="chat.legal_research",
            details={
                "provider": provider,
                "model": model,
                "settings": dict(settings),
                "research_settings": {
                    "firm_authority": research_settings.firm_authority,
                    "local_corpus": research_settings.local_corpus,
                    "courtlistener_mode": research_settings.courtlistener_mode.value,
                },
                "user_text_length": len(user_text or ""),
                "context_text_length": len(file_content or ""),
            },
        )
        self.chat_history.append("<i>Researching selected legal authority</i>")
        payload = {
            "user_text": user_text,
            "file_content": file_content,
            "attachments": attachments,
            "audio_files": audio_files,
            "provider": provider,
            "model": model,
            "settings": dict(settings),
        }
        worker = ChatLegalResearchWorker(
            user_text=user_text,
            file_content=file_content,
            provider=provider,
            model=model,
            settings=settings,
            research_settings=research_settings,
        )
        self.chat_research_worker = worker
        worker.status_update.connect(self._on_chat_legal_research_status)
        worker.debug_update.connect(
            lambda phase, message, level, details, run_id=debug_run_id: (
                self._emit_chat_legal_research_debug(
                    run_id,
                    phase,
                    message,
                    level,
                    details,
                )
            )
        )
        worker.research_finished.connect(
            lambda packet, w=worker, run_id=debug_run_id, p=payload: (
                self._on_chat_legal_research_finished(w, run_id, p, packet)
            )
        )
        worker.research_failed.connect(
            lambda kind, message, w=worker, run_id=debug_run_id: (
                self._on_chat_legal_research_failed(w, run_id, kind, message)
            )
        )
        worker.finished.connect(lambda w=worker: self._on_chat_legal_research_thread_finished(w))
        worker.finished.connect(worker.deleteLater)
        worker.start()

    def _on_chat_legal_research_status(self, message):
        self.chat_history.append(f"<i>  {message}</i>")

    def _emit_chat_legal_research_debug(
        self,
        run_id,
        phase,
        message,
        level="info",
        details=None,
    ):
        from icharlotte_core import task_debug

        task_debug.emit_event(
            run_id,
            "chat_legal_research",
            "Chat Legal Research",
            phase=phase,
            message=message,
            level=level,
            source="chat.legal_research",
            details=details or {},
        )

    def _on_chat_legal_research_finished(self, worker, run_id, payload, packet):
        if worker is not self.chat_research_worker:
            return
        from icharlotte_core import task_debug

        task_debug.finish_run(
            run_id,
            status="success",
            message="Task complete",
            details={
                "selected_authority_count": len(
                    getattr(packet, "selected_authorities", []) or []
                ),
                "warnings": list(getattr(packet, "warnings", []) or []),
                "searches": list(getattr(packet, "searches", []) or []),
            },
        )
        self._continue_send_message_after_research(
            payload["user_text"],
            payload["file_content"],
            attachments=payload["attachments"],
            audio_files=payload["audio_files"],
            provider=payload["provider"],
            model=payload["model"],
            settings=payload["settings"],
            research_packet=packet,
        )

    def _on_chat_legal_research_failed(self, worker, run_id, kind, message):
        if worker is not self.chat_research_worker:
            return
        from icharlotte_core import task_debug

        prefix = "Legal research stopped" if kind == "stopped" else "Legal research error"
        task_debug.finish_run(
            run_id,
            status="error",
            message=f"{prefix}: {message}",
            details={"error": str(message)},
        )
        self._pending_research = None
        self.chat_history.append(f"<font color='orange'>{prefix}: {message}</font>")
        self.send_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)

    def _on_chat_legal_research_thread_finished(self, worker):
        if worker is self.chat_research_worker:
            self.chat_research_worker = None

    def _run_chat_legal_research(
        self,
        user_text,
        file_content,
        *,
        provider=None,
        model=None,
        settings=None,
        research_settings=None,
    ):
        from icharlotte_core.llm import LLMHandler
        from icharlotte_core import task_debug

        provider = self.provider_combo.currentText() if provider is None else provider
        model = self.model_combo.currentText() if model is None else model
        settings = dict(settings) if settings is not None else dict(self.settings)
        if research_settings is None:
            research_settings = self._current_chat_research_settings()
        debug_run_id = task_debug.start_run(
            task_id="chat_legal_research",
            task_title="Chat Legal Research",
            source="chat.legal_research",
            details={
                "provider": provider,
                "model": model,
                "settings": dict(settings),
                "research_settings": {
                    "firm_authority": research_settings.firm_authority,
                    "local_corpus": research_settings.local_corpus,
                    "courtlistener_mode": research_settings.courtlistener_mode.value,
                },
                "user_text_length": len(user_text or ""),
                "context_text_length": len(file_content or ""),
            },
        )

        def llm_for_research(system_prompt, user_prompt):
            return LLMHandler.generate(
                provider=provider,
                model=model,
                system_prompt=system_prompt,
                user_prompt=user_prompt,
                file_contents="",
                settings={**settings, "stream": False, "temperature": 0.2},
            )

        def status(message):
            self.chat_history.append(f"<i>  {message}</i>")
            QApplication.processEvents()

        def debug_event(*, phase, message, level="info", details=None):
            task_debug.emit_event(
                debug_run_id,
                "chat_legal_research",
                "Chat Legal Research",
                phase=phase,
                message=message,
                level=level,
                source="chat.legal_research",
                details=details or {},
            )

        self.chat_history.append("<i>Researching selected legal authority</i>")
        QApplication.processEvents()
        try:
            service = ChatLegalResearchService.from_environment(
                llm_callback=llm_for_research
            )
            packet = service.research(
                user_text=user_text,
                context_text=file_content[:100000] if file_content else "",
                settings=research_settings,
                status_callback=status,
                debug_callback=debug_event,
            )
            task_debug.finish_run(
                debug_run_id,
                status="success",
                message="Task complete",
                details={
                    "selected_authority_count": len(
                        getattr(packet, "selected_authorities", []) or []
                    ),
                    "warnings": list(getattr(packet, "warnings", []) or []),
                    "searches": list(getattr(packet, "searches", []) or []),
                },
            )
            return packet
        except ChatResearchError as exc:
            task_debug.finish_run(
                debug_run_id,
                status="error",
                message=f"Legal research stopped: {exc}",
                details={"error": str(exc)},
            )
            self._pending_research = None
            self.chat_history.append(
                f"<font color='orange'>Legal research stopped: {exc}</font>"
            )
            self.send_btn.setEnabled(True)
            self.stop_btn.setEnabled(False)
            return None
        except Exception as exc:
            task_debug.finish_run(
                debug_run_id,
                status="error",
                message=f"Legal research error: {exc}",
                details={"error": str(exc)},
            )
            self._pending_research = None
            self.chat_history.append(
                f"<font color='orange'>Legal research error: {exc}</font>"
            )
            self.send_btn.setEnabled(True)
            self.stop_btn.setEnabled(False)
            return None

    def _get_checked_audio_files(self):
        """Return list of checked audio/video file paths."""
        audio_files = []
        for i in range(self.file_list.count()):
            item = self.file_list.item(i)
            if item.checkState() == Qt.CheckState.Checked:
                path = item.data(Qt.ItemDataRole.UserRole)
                if path:
                    ext = os.path.splitext(path)[1].lower()
                    if ext in self.AUDIO_VIDEO_EXTENSIONS:
                        audio_files.append(path)
        return audio_files



    # Pixels of slack when deciding whether the chat view is "parked" at the
    # bottom. Smaller than a line of text, so scrolling up even one notch
    # disengages auto-follow.
    _SCROLL_BOTTOM_TOLERANCE_PX = 4

    def _chat_is_at_bottom(self) -> bool:
        """True when the chat view is parked at (or within a hair of) the bottom."""
        bar = self.chat_history.verticalScrollBar()
        return bar.value() >= bar.maximum() - self._SCROLL_BOTTOM_TOLERANCE_PX

    def _chat_scroll_to_bottom(self):
        """Pin the chat view to the bottom."""
        bar = self.chat_history.verticalScrollBar()
        bar.setValue(bar.maximum())

    def on_streaming_token(self, token: str):
        """Handle real-time token display during streaming."""
        self.stream_text += token

        # Only auto-follow the stream when the user is already parked at the
        # bottom. Capture this BEFORE inserting, because the insert grows the
        # scroll range. If the user has scrolled up to read, leave their
        # viewport untouched (scrolling back to the bottom re-arms following).
        follow = self._chat_is_at_bottom()

        # Append the token via a local cursor. We deliberately do NOT call
        # setTextCursor()/ensureCursorVisible() here: forcing the cursor into
        # view on every token is what yanked a scrolled-up reader back down.
        cursor = self.chat_history.textCursor()
        cursor.movePosition(QTextCursor.MoveOperation.End)
        cursor.insertText(token)

        if follow:
            self._chat_scroll_to_bottom()

    def on_stream_complete(self, full_text: str):
        """Handle completion of streaming response."""
        # Replace streamed plain text with rendered markdown
        self.finalize_response(full_text)

        self.send_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)

    @staticmethod
    def _research_basis_html_block(lines):
        return "".join(f"<div>{line}</div>" for line in lines)

    @staticmethod
    def _research_basis_plain_text(html_block):
        doc = QTextDocument()
        doc.setHtml(html_block)
        return doc.toPlainText().strip()

    def finalize_response(self, text: str):
        """Finalize the response with markdown rendering and save to persistence."""
        # Remember whether the user is following the bottom BEFORE we mutate the
        # document, so a scrolled-up reader isn't yanked down when the final
        # formatted text and separators are appended.
        follow = self._chat_is_at_bottom()
        pending_research = getattr(self, '_pending_research', None)

        # Calculate response time
        response_time = int((time.time() - self.stream_start_time) * 1000) if self.stream_start_time else 0

        # Remove the plain text we streamed and replace with formatted HTML
        cursor = self.chat_history.textCursor()
        cursor.setPosition(self.stream_start_pos)
        cursor.movePosition(QTextCursor.MoveOperation.End, QTextCursor.MoveMode.KeepAnchor)
        cursor.removeSelectedText()

        # Apply deterministic citation cross-check if legal research was done
        if pending_research:
            try:
                known_names = pending_research.get_known_case_names()
                if known_names:
                    from icharlotte_core.legal_research.engine import LegalResearchEngine
                    text = LegalResearchEngine._deterministic_citation_check(
                        text, known_names
                    )
            except Exception as e:
                print(f"[ChatTab] Deterministic citation check failed: {e}")

        # Convert markdown to HTML (tables, fenced code, GFM line breaks, lists)
        try:
            html_text = render_markdown(text)
        except Exception as e:
            log_event(f"Markdown conversion failed: {e}", "error")
            html_text = text.replace('\n', '<br>')

        research_basis_html = ""
        research_basis_plain = ""
        if pending_research:
            try:
                research_basis_html = self._research_basis_html_block(
                    pending_research.format_research_basis_html()
                )
                research_basis_plain = self._research_basis_plain_text(research_basis_html)
            except Exception as e:
                print(f"[ChatTab] Research basis display failed: {e}")
            self._pending_research = None

        cursor.insertHtml(html_text)
        if research_basis_html:
            cursor.insertBlock()
            cursor.insertHtml(research_basis_html)
        self.chat_history.append("")  # New line
        self.chat_history.append("-" * 50)

        # Save assistant message to persistence
        saved_text = text
        if research_basis_plain:
            saved_text = f"{text.rstrip()}\n\n{research_basis_plain}"
        if self.persistence and self.current_conversation_id:
            token_count = TokenCounter.estimate_tokens(saved_text, self.provider_combo.currentText())
            assistant_message = Message(
                role='assistant',
                content=saved_text,
                token_count=token_count,
                model_used=self.model_combo.currentText(),
                response_time_ms=response_time
            )
            log_event(f"Saving assistant message to conversation {self.current_conversation_id}")
            self.persistence.add_message(self.current_conversation_id, assistant_message)
        else:
            log_event(f"WARNING: Cannot save assistant message - persistence={self.persistence is not None}, conv_id={self.current_conversation_id}", "warning")

        # Update legacy history
        self.conversation_history.append({'role': 'assistant', 'content': saved_text})

        if follow:
            self._chat_scroll_to_bottom()

        self.update_context_indicator()

    def on_error(self, err: str):
        """Handle generation error."""
        self._pending_research = None
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

        # Save sidebar visibility state
        self._save_sidebar_state()

    # --- Chat Legal Research Source Persistence ---

    def _load_chat_research_settings(self):
        settings = QSettings("iCharlotte", "iCharlotte")
        return ChatResearchSettings.from_values(
            firm_authority=settings.value("chat_tab/legal_research_firm_authority", True),
            local_corpus=settings.value("chat_tab/legal_research_local_corpus", True),
            courtlistener_mode=settings.value(
                "chat_tab/legal_research_courtlistener_mode",
                CourtListenerMode.FALLBACK_CURRENT_LAW.value,
            ),
        )

    def _save_chat_research_settings(self, research_settings):
        settings = QSettings("iCharlotte", "iCharlotte")
        settings.setValue(
            "chat_tab/legal_research_firm_authority",
            research_settings.firm_authority,
        )
        settings.setValue(
            "chat_tab/legal_research_local_corpus",
            research_settings.local_corpus,
        )
        settings.setValue(
            "chat_tab/legal_research_courtlistener_mode",
            research_settings.courtlistener_mode.value,
        )

    def _build_research_sources_menu(self):
        menu = QMenu(self.research_sources_btn)
        current = self._load_chat_research_settings()

        self.firm_authority_action = QAction("Firm/sample-motion authority", menu)
        self.firm_authority_action.setCheckable(True)
        self.firm_authority_action.setChecked(current.firm_authority)

        self.local_corpus_action = QAction("Local California corpus", menu)
        self.local_corpus_action.setCheckable(True)
        self.local_corpus_action.setChecked(current.local_corpus)

        menu.addAction(self.firm_authority_action)
        menu.addAction(self.local_corpus_action)
        menu.addSeparator()

        mode_group = QActionGroup(menu)
        mode_group.setExclusive(True)
        self.courtlistener_off_action = QAction("CourtListener API: Off", menu)
        self.courtlistener_fallback_action = QAction(
            "CourtListener API: Fallback/current-law",
            menu,
        )
        self.courtlistener_always_action = QAction("CourtListener API: Always search", menu)
        for action in (
            self.courtlistener_off_action,
            self.courtlistener_fallback_action,
            self.courtlistener_always_action,
        ):
            action.setCheckable(True)
            mode_group.addAction(action)
            menu.addAction(action)

        if current.courtlistener_mode == CourtListenerMode.OFF:
            self.courtlistener_off_action.setChecked(True)
        elif current.courtlistener_mode == CourtListenerMode.ALWAYS_SEARCH:
            self.courtlistener_always_action.setChecked(True)
        else:
            self.courtlistener_fallback_action.setChecked(True)

        for action in (
            self.firm_authority_action,
            self.local_corpus_action,
            self.courtlistener_off_action,
            self.courtlistener_fallback_action,
            self.courtlistener_always_action,
        ):
            action.triggered.connect(self._on_research_source_changed)

        self.research_sources_btn.setMenu(menu)
        self._refresh_research_sources_label()

    def _current_chat_research_settings(self):
        if self.courtlistener_off_action.isChecked():
            mode = CourtListenerMode.OFF
        elif self.courtlistener_always_action.isChecked():
            mode = CourtListenerMode.ALWAYS_SEARCH
        else:
            mode = CourtListenerMode.FALLBACK_CURRENT_LAW
        return ChatResearchSettings.from_values(
            firm_authority=self.firm_authority_action.isChecked(),
            local_corpus=self.local_corpus_action.isChecked(),
            courtlistener_mode=mode.value,
        )

    def _on_research_source_changed(self):
        current = self._current_chat_research_settings()
        if current.courtlistener_mode == CourtListenerMode.ALWAYS_SEARCH:
            self.courtlistener_always_action.setChecked(True)
        elif current.courtlistener_mode == CourtListenerMode.OFF:
            self.courtlistener_off_action.setChecked(True)
        else:
            self.courtlistener_fallback_action.setChecked(True)
        self._save_chat_research_settings(current)
        self._refresh_research_sources_label()

    def _refresh_research_sources_label(self):
        current = self._current_chat_research_settings()
        parts = []
        if current.firm_authority:
            parts.append("Firm")
        if current.local_corpus:
            parts.append("Local")
        if current.courtlistener_mode == CourtListenerMode.FALLBACK_CURRENT_LAW:
            parts.append("CL Fallback")
        elif current.courtlistener_mode == CourtListenerMode.ALWAYS_SEARCH:
            parts.append("CL Always")
        else:
            parts.append("CL Off")
        self.research_sources_btn.setText("Sources: " + " + ".join(parts))

    # --- Splitter Persistence ---

    def _on_main_splitter_moved(self, pos, index):
        """Save main splitter sizes when moved."""
        settings = QSettings("iCharlotte", "iCharlotte")
        sizes = self.main_splitter.sizes()
        settings.setValue("chat_tab/main_splitter_sizes", sizes)

    def _on_chat_splitter_moved(self, pos, index):
        """Save chat splitter sizes when moved."""
        settings = QSettings("iCharlotte", "iCharlotte")
        sizes = self.chat_splitter.sizes()
        settings.setValue("chat_tab/chat_splitter_sizes", sizes)

    def _save_sidebar_state(self):
        """Save sidebar collapsed state."""
        settings = QSettings("iCharlotte", "iCharlotte")
        settings.setValue("chat_tab/sidebar_collapsed", self.sidebar_collapsed)

    def _load_splitter_sizes(self):
        """Load saved splitter sizes from settings."""
        settings = QSettings("iCharlotte", "iCharlotte")

        # Load main splitter sizes
        main_sizes = settings.value("chat_tab/main_splitter_sizes")
        if main_sizes:
            try:
                # Convert to list of ints if needed
                if isinstance(main_sizes, list):
                    main_sizes = [int(s) for s in main_sizes]
                    self.main_splitter.setSizes(main_sizes)
            except (ValueError, TypeError):
                pass  # Use default sizes

        # Load chat splitter sizes
        chat_sizes = settings.value("chat_tab/chat_splitter_sizes")
        if chat_sizes:
            try:
                if isinstance(chat_sizes, list):
                    chat_sizes = [int(s) for s in chat_sizes]
                    self.chat_splitter.setSizes(chat_sizes)
            except (ValueError, TypeError):
                pass  # Use default sizes

        # Load sidebar state
        sidebar_collapsed = settings.value("chat_tab/sidebar_collapsed")
        if sidebar_collapsed is not None:
            # QSettings may return string 'true'/'false' or bool
            if isinstance(sidebar_collapsed, str):
                self.sidebar_collapsed = sidebar_collapsed.lower() == 'true'
            else:
                self.sidebar_collapsed = bool(sidebar_collapsed)

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
        # Skip if indicator not present (e.g., a subclass overrides setup_ui)
        if not hasattr(self, 'context_indicator'):
            return

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
        """Handle input text changes."""
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
        """(Re)populate the quick prompts template menu from built-ins + the global store."""
        # Reuse one persistent menu so we can refresh it on aboutToShow without re-wiring.
        menu = self.template_btn.menu()
        if menu is None:
            menu = QMenu(self.template_btn)
            self.template_btn.setMenu(menu)
            # Rebuild every time the menu opens so newly added/edited/deleted global
            # templates appear in all chat tabs without needing a restart.
            menu.aboutToShow.connect(lambda s=self: ChatTab.update_template_menu(s))
        menu.clear()

        # Built-in prompts
        for prompt in BUILTIN_PROMPTS:
            if prompt.id == 'builtin_mediation_brief':
                continue  # Handled separately below
            action = QAction(prompt.name, menu)
            action.triggered.connect(lambda checked, p=prompt: self.insert_template(p.prompt))
            menu.addAction(action)

        # Mediation Brief (special — triggers generation, not text insert)
        menu.addSeparator()
        med_brief_action = QAction("Mediation Brief", menu)
        med_brief_action.triggered.connect(self._on_mediation_brief_selected)
        menu.addAction(med_brief_action)

        # Custom prompts from the global store (shared across all cases/sessions)
        try:
            from ..chat.global_prompts import get_global_quick_prompt_store
            custom_prompts = get_global_quick_prompt_store().get_quick_prompts()
        except Exception as e:
            custom_prompts = []
            log_event(f"[ChatTab] Could not load global quick prompts: {e}", "error")
        if custom_prompts:
            menu.addSeparator()
            for prompt in custom_prompts:
                action = QAction(prompt.name, menu)
                action.triggered.connect(lambda checked, p=prompt: self.insert_template(p.prompt))
                menu.addAction(action)

        menu.addSeparator()
        manage_action = QAction("Manage Templates...", menu)
        manage_action.triggered.connect(self.open_template_manager)
        menu.addAction(manage_action)

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
        # PromptTemplateDialog persists templates to the global store; the persistence
        # arg is accepted only for backward compatibility and is unused for templates.
        dlg = PromptTemplateDialog(self.persistence, self)
        if dlg.exec():
            self.update_template_menu()

    def open_file_from_list(self, item):
        """Open a file from the attachment list with the system default application."""
        path = item.data(Qt.ItemDataRole.UserRole)
        if path and os.path.exists(path):
            os.startfile(path)

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
        """Remove a file from the list and persist the change."""
        path = item.data(Qt.ItemDataRole.UserRole)
        if path in self.attached_files:
            self.attached_files.remove(path)
        row = self.file_list.row(item)
        self.file_list.takeItem(row)
        # Persist the removal
        if self.persistence:
            self.persistence.remove_attached_file(path)

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
        """Clear chat - only creates new conversation if current one has messages."""
        # Delegate to clear_current_chat which has the proper logic
        self.clear_current_chat()

    def reset_state(self):
        """Reset all state (called on case switch)."""
        # Stop any running threads to prevent "QThread destroyed while running" errors
        if self.worker is not None and self.worker.isRunning():
            self.worker.request_stop()
            self.worker.wait(1000)

        if self.chat_research_worker is not None and self.chat_research_worker.isRunning():
            self.chat_research_worker.request_stop()
            self.chat_research_worker.wait(1000)

        if self.fetcher is not None and self.fetcher.isRunning():
            self.fetcher.wait(1000)

        self.clear_chat()
        self.clear_files()
        self.conversation_history = []
        self.current_conversation_id = None
        self.current_conversation = None

    # --- Mediation Brief Integration ---

    def _on_mediation_brief_selected(self):
        """Handle Mediation Brief template selection — show confirmation dialog."""
        checked_files = []
        for i in range(self.file_list.count()):
            item = self.file_list.item(i)
            if item.checkState() == Qt.CheckState.Checked:
                path = item.data(Qt.ItemDataRole.UserRole)
                if path:
                    checked_files.append(os.path.basename(path))

        if not checked_files:
            QMessageBox.warning(
                self, "No Documents",
                "Please select documents in the file list before generating a mediation brief."
            )
            return

        case_name = self.file_number or "Unknown Case"
        file_list_text = "\n".join(f"  - {f}" for f in checked_files[:15])
        if len(checked_files) > 15:
            file_list_text += f"\n  ... and {len(checked_files) - 15} more"

        msg = (
            f"Generate a Mediation Brief for case {case_name}?\n\n"
            f"Documents ({len(checked_files)}):\n{file_list_text}\n\n"
            "This will generate a comprehensive brief section-by-section. "
            "The process may take several minutes."
        )

        reply = QMessageBox.question(
            self, "Generate Mediation Brief", msg,
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No
        )

        if reply == QMessageBox.StandardButton.Yes:
            self._start_mediation_brief_generation()

    def _start_mediation_brief_generation(self):
        """Start the mediation brief generation pipeline."""
        main_win = self.window()
        case_path = getattr(main_win, 'case_path', None)

        generator = MediationBriefGenerator()

        caption_path = None
        if case_path:
            caption_path = generator.find_caption_template(case_path)

        if not caption_path:
            search_loc = case_path or ""
            caption_path, _ = QFileDialog.getOpenFileName(
                self,
                f"No caption template found in {search_loc}. Select one:",
                search_loc,
                "Word Documents (*.docx)"
            )
            if not caption_path:
                return

        generator.caption_template_path = caption_path
        generator.document_content = self.read_files_content()

        if not generator.document_content.strip():
            QMessageBox.warning(
                self, "No Content",
                "Could not read any content from the selected documents."
            )
            return

        generator.get_style_excerpts()
        self.med_brief_generator = generator

        self.chat_history.append("<b>Mediation Brief Generator</b>")
        self.chat_history.append("<i>Starting generation...</i>")
        self.chat_history.append("")

        self.send_btn.setEnabled(False)

        self.med_brief_worker = MediationBriefWorker(generator, parent=self)
        self.med_brief_worker.section_started.connect(self._on_brief_section_started)
        self.med_brief_worker.section_complete.connect(self._on_brief_section_complete)
        self.med_brief_worker.all_complete.connect(self._on_brief_all_complete)
        self.med_brief_worker.error.connect(self._on_brief_error)
        self.med_brief_worker.start()

    def _on_brief_section_started(self, section_name: str, index: int, total: int):
        """Display progress when a section starts generating."""
        if section_name == "planning":
            self.chat_history.append("<i>Analyzing documents...</i>")
        else:
            heading = SECTION_HEADINGS.get(section_name, ("", section_name.upper()))
            display = f"{heading[0]}. {heading[1]}" if heading[0] else section_name
            self.chat_history.append(f"<i>Generating {display} ({index} of {total - 1})...</i>")

    def _on_brief_section_complete(self, section_name: str, text: str):
        """Display completed section text in chat."""
        heading = SECTION_HEADINGS.get(section_name)
        if heading:
            self.chat_history.append(
                f"<br><b>{heading[0]}.&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;{heading[1]}</b>"
            )

        try:
            html = render_markdown(text)
        except Exception:
            html = text.replace('\n', '<br>')
        self.chat_history.append(html)
        self.chat_history.append("<hr>")
        self.chat_history.ensureCursorVisible()

    def _on_brief_all_complete(self, sections: dict):
        """Handle completion of all sections — assemble document and save."""
        self.send_btn.setEnabled(True)
        self.chat_history.append("<b>All sections generated. Assembling document...</b>")

        gen = self.med_brief_generator
        main_win = self.window()
        case_path = getattr(main_win, 'case_path', None)
        default_dir = os.path.dirname(case_path) if case_path else ""
        default_name = os.path.join(default_dir, "Defendant's Confidential Mediation Brief.docx")

        import tempfile
        with tempfile.TemporaryDirectory() as tmpdir:
            temp_output = os.path.join(tmpdir, "mediation_brief.docx")
            try:
                gen.assemble_document(gen.caption_template_path, temp_output)
            except Exception as e:
                self.chat_history.append(f"<b style='color:red'>Assembly error: {e}</b>")
                return

            save_path, _ = QFileDialog.getSaveFileName(
                self,
                "Save Mediation Brief",
                default_name,
                "Word Documents (*.docx)"
            )
            if save_path:
                shutil.copy2(temp_output, save_path)
                gen.saved_path = save_path
                self.chat_history.append(f"<b>Mediation brief saved to:</b> {save_path}")
            else:
                self.chat_history.append(
                    "<i>Save cancelled. You can refine sections and save later.</i>"
                )

        self.chat_history.append("")
        self.chat_history.append(
            "<i>You can now refine the brief by typing instructions "
            "(e.g., 'make the Damages section more aggressive'), "
            "or use the Add Quotes button to insert deposition testimony. "
            "Send a normal message to exit brief mode.</i>"
        )

        # Show "Add Quotes" button
        if not hasattr(self, 'add_quotes_btn') or self.add_quotes_btn is None:
            self.add_quotes_btn = QPushButton("Add Quotes")
            self.add_quotes_btn.clicked.connect(self._on_add_quotes_clicked)
            send_layout = self.send_btn.parent().layout()
            if send_layout:
                send_layout.insertWidget(send_layout.indexOf(self.send_btn), self.add_quotes_btn)
        self.add_quotes_btn.setVisible(True)
        self.add_quotes_btn.setEnabled(True)

    def _on_brief_error(self, error_msg: str):
        """Handle generation error."""
        self.send_btn.setEnabled(True)
        self.chat_history.append(f"<b style='color:red'>Generation error: {error_msg}</b>")

    def _route_brief_refinement(self, user_text: str):
        """Route a refinement message asynchronously (off UI thread)."""
        self.chat_history.append(f"<b>You:</b> {user_text}")
        self.chat_input.clear()
        self.send_btn.setEnabled(False)
        self.chat_history.append("<i>Routing...</i>")

        self._pending_refinement_text = user_text
        worker = RoutingWorker(self.med_brief_generator, user_text, parent=self)
        worker.result.connect(self._on_routing_result)
        worker.error.connect(self._on_brief_error)
        self.med_brief_worker = worker
        worker.start()

    def _on_routing_result(self, sections: list):
        """Handle routing worker result — start refinement or pass to normal chat."""
        if sections:
            self._start_brief_refinement(self._pending_refinement_text, sections)
        else:
            # Not a brief refinement — re-enable send and let user know
            self.send_btn.setEnabled(True)
            self.chat_history.append(
                "<i>Message not related to the brief. Send again to chat normally, "
                "or type a brief refinement instruction.</i>"
            )

    def _start_brief_refinement(self, instruction: str, section_names: list):
        """Regenerate specified sections with user's instruction."""
        section_display = ", ".join(
            SECTION_HEADINGS.get(s, ("", s))[1] for s in section_names
        )
        self.chat_history.append(f"<i>Regenerating: {section_display}...</i>")
        self.send_btn.setEnabled(False)

        gen = self.med_brief_generator
        worker = RefinementWorker(gen, section_names, instruction, parent=self)
        worker.section_complete.connect(self._on_brief_section_complete)
        worker.all_complete.connect(
            lambda regenerated: self._on_brief_all_complete(gen.sections)
        )
        worker.error.connect(self._on_brief_error)
        self.med_brief_worker = worker
        worker.start()

    def _on_add_quotes_clicked(self):
        """Open the Quote Insertion dialog."""
        from .quote_dialog import QuoteInsertionDialog

        if not self.med_brief_generator or not self.med_brief_generator.is_active:
            return

        dlg = QuoteInsertionDialog(self.med_brief_generator, parent=self)
        dlg.quotes_to_insert.connect(self._on_quotes_confirmed)
        dlg.exec()

    def _on_quotes_confirmed(self, quotes: list, section_name: str,
                              subsection: str, mode: str):
        """Handle confirmed quote insertion from the dialog."""
        gen = self.med_brief_generator
        count = len(quotes)

        self.chat_history.append(
            f"<b>Inserting {count} quote(s) into "
            f"{SECTION_HEADINGS.get(section_name, ('', section_name))[1]}...</b>"
        )

        if mode == "quick":
            sub_title = subsection if subsection else None
            gen.insert_quotes_quick(quotes, section_name, sub_title)
            self.chat_history.append(f"<i>{count} quote(s) inserted.</i>")
            self._reassemble_and_save()
        else:
            # Weave In — regenerate section with LLM
            self.send_btn.setEnabled(False)
            if hasattr(self, 'add_quotes_btn') and self.add_quotes_btn:
                self.add_quotes_btn.setEnabled(False)

            quote_text_parts = []
            for q in quotes:
                quote_text_parts.append(
                    f"DEPO_QUOTE_START\n{q['qa_text']}\nDEPO_QUOTE_END\n"
                    f"({q['deponent']} Depo Trns., at p. {q['page_line']}.)"
                )
            all_quotes = "\n\n".join(quote_text_parts)

            instruction = (
                f"Incorporate the following deposition testimony into this section "
                f"at the most appropriate location. Weave it into the argument "
                f"naturally with proper context and transitions. Include the "
                f"testimony verbatim — do not change any wording.\n\n{all_quotes}"
            )

            worker = RefinementWorker(
                gen, [section_name], instruction, parent=self
            )
            worker.section_complete.connect(self._on_brief_section_complete)
            worker.all_complete.connect(
                lambda regenerated: self._on_quote_weave_complete()
            )
            worker.error.connect(self._on_brief_error)
            self.med_brief_worker = worker
            worker.start()

    def _reassemble_and_save(self):
        """Reassemble the Word document and save (overwrite or Save As)."""
        import tempfile
        gen = self.med_brief_generator

        self.chat_history.append("<i>Assembling document...</i>")

        with tempfile.TemporaryDirectory() as tmpdir:
            temp_output = os.path.join(tmpdir, "mediation_brief.docx")
            try:
                gen.assemble_document(gen.caption_template_path, temp_output)
            except Exception as e:
                self.chat_history.append(f"<b style='color:red'>Assembly error: {e}</b>")
                return

            if gen.saved_path and os.path.exists(os.path.dirname(gen.saved_path)):
                shutil.copy2(temp_output, gen.saved_path)
                self.chat_history.append(f"<b>Document updated:</b> {gen.saved_path}")
            else:
                main_win = self.window()
                case_path = getattr(main_win, 'case_path', None)
                default_dir = os.path.dirname(case_path) if case_path else ""
                default_name = os.path.join(
                    default_dir, "Defendant's Confidential Mediation Brief.docx"
                )
                save_path, _ = QFileDialog.getSaveFileName(
                    self, "Save Mediation Brief", default_name,
                    "Word Documents (*.docx)"
                )
                if save_path:
                    shutil.copy2(temp_output, save_path)
                    gen.saved_path = save_path
                    self.chat_history.append(f"<b>Document saved:</b> {save_path}")
                else:
                    self.chat_history.append("<i>Save cancelled.</i>")

    def _on_quote_weave_complete(self):
        """Handle completion of Weave In mode — reassemble and save."""
        self.send_btn.setEnabled(True)
        if hasattr(self, 'add_quotes_btn') and self.add_quotes_btn:
            self.add_quotes_btn.setEnabled(True)
        self.chat_history.append("<i>Quotes woven into section.</i>")
        self._reassemble_and_save()

class IndexTab(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.file_number = None
        self.index_data = {}  # {pdf_path: [docs]}
        self.current_pdf_path = None
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

        # Right of the PDF list: the shared workbench.
        from .separator_workbench import SeparatorWorkbench
        self.workbench = SeparatorWorkbench()
        self.workbench.reanalyze_requested.connect(self._on_workbench_reanalyze)
        self.workbench.processing_complete.connect(self._on_workbench_processing_complete)

        # Host-level affordance: persist manual table edits back into the index
        # cache (the workbench is storage-agnostic by design).
        save_row = QHBoxLayout()
        save_row.addStretch()
        self.save_index_btn = QPushButton("Save Edits to Index")
        self.save_index_btn.setToolTip("Save the current table edits (titles, page ranges) to the cached index for this case.")
        self.save_index_btn.clicked.connect(self.save_table_to_index)
        save_row.addWidget(self.save_index_btn)

        right_host = QWidget()
        right_host_layout = QVBoxLayout(right_host)
        right_host_layout.setContentsMargins(0, 0, 0, 0)
        right_host_layout.addWidget(self.workbench, 1)
        right_host_layout.addLayout(save_row)
        self.doc_splitter.addWidget(right_host)

        self.doc_splitter.setSizes([200, 900])
        self.doc_splitter.setCollapsible(0, True)

    def toggle_pdf_list_collapse(self):
        """Toggle collapse/expand of the processed PDFs panel."""
        sizes = self.doc_splitter.sizes()
        if self._pdf_list_collapsed:
            # Expand
            self.doc_splitter.setSizes([self._pdf_list_width, sizes[1] - self._pdf_list_width])
            self.collapse_pdf_btn.setText("◀")
            self.collapse_pdf_btn.setToolTip("Collapse panel")
            self._pdf_list_collapsed = False
        else:
            # Collapse - store current width first
            if sizes[0] > 0:
                self._pdf_list_width = sizes[0]
            self.doc_splitter.setSizes([0, sizes[1] + sizes[0]])
            self.collapse_pdf_btn.setText("▶")
            self.collapse_pdf_btn.setToolTip("Expand panel")
            self._pdf_list_collapsed = True

    def load_data(self, file_number):
        self.file_number = file_number
        self.index_data = {}
        self.pdf_list.clear()
        if hasattr(self, 'workbench'):
            self.workbench.load_docs("", [])

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

        if hasattr(self, 'workbench'):
            self.workbench.set_busy(False)

    def _on_workbench_reanalyze(self, sensitivity):
        if not self.current_pdf_path:
            return
        main_window = self.window()
        if hasattr(main_window, 'run_separator_path'):
            main_window.run_separator_path(self.current_pdf_path, sensitivity=sensitivity)

    def _on_workbench_processing_complete(self, summary):
        created = summary.get("created", [])
        errors = summary.get("errors", [])
        msg = f"Processed {len(created)} item(s).\n\nFiles Created:\n" + "\n".join(created[:10])
        if len(created) > 10:
            msg += f"\n...and {len(created) - 10} more."
        if errors:
            msg += "\n\nErrors:\n" + "\n".join(errors)
            QMessageBox.warning(self, "Result with Errors", msg)
        else:
            QMessageBox.information(self, "Success", msg)

    def on_pdf_selected(self, current, previous):
        if not current:
            return
        path = current.text()
        self.current_pdf_path = path
        docs = self.index_data.get(path, [])
        self.workbench.load_docs(path, docs)

    def save_table_to_index(self):
        if not self.current_pdf_path:
            QMessageBox.warning(self, "Warning", "No PDF selected.")
            return
        wb = self.workbench
        new_docs = []
        for row in range(wb.doc_table.rowCount()):
            doc_obj = wb._get_doc_from_row(row)
            if doc_obj['start'] is None or doc_obj['end'] is None:
                QMessageBox.warning(self, "Validation Error",
                    f"Row {row + 1} (ID {doc_obj['id']}): Invalid page range. Cannot save.")
                return
            new_docs.append(doc_obj)
        self.index_data[self.current_pdf_path] = new_docs
        self.save_data()
        QMessageBox.information(self, "Success", f"Saved {len(new_docs)} document(s) to index.")
