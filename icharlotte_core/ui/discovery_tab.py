"""
Discovery Tab — Propound / Respond sub-tabs for generating written discovery.

Implements the full left-pane controls (mode, type, party, documents, LLM,
prompt, generate button) and right-pane editor shell (toolbar, document
sub-tabs, empty state).
"""

import logging
import os
from typing import Dict, List, Optional

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QTabWidget, QSplitter, QLabel,
    QComboBox, QRadioButton, QButtonGroup, QCheckBox, QTextEdit,
    QPushButton, QListWidgetItem, QMenu, QScrollArea, QGroupBox,
    QDialog, QLineEdit, QFormLayout, QDialogButtonBox, QSizePolicy,
    QPlainTextEdit,
)
from PySide6.QtCore import Qt, QTimer, Signal, QFileInfo, Slot
from PySide6.QtGui import QDragEnterEvent, QDropEvent, QAction

from icharlotte_core.ui.tabs import ResizableListWidget
from icharlotte_core.llm import LLMWorker, ModelFetcher
from icharlotte_core.config import API_KEYS
from icharlotte_core.discovery.models import (
    Party, PartyRole, DiscoveryMode, DiscoveryType, CustomStyle,
    DiscoverySet, generate_abbreviation,
)
from icharlotte_core.discovery.engine import DiscoveryEngine
from icharlotte_core.discovery.assembler import DiscoveryAssembler
from icharlotte_core.discovery.templates import extract_requests_from_text

logger = logging.getLogger(__name__)

try:
    import fitz  # PyMuPDF
except ImportError:
    fitz = None


# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

_SUPPORTED_DROP_EXTENSIONS = (
    ".pdf", ".docx", ".txt", ".msg",
    ".png", ".jpg", ".jpeg", ".gif", ".webp",
)

_CUSTOM_STYLE_PLACEHOLDERS = {
    CustomStyle.CUSTOM_ONLY: "Describe what discovery requests to generate...",
    CustomStyle.STANDARD_PLUS_CUSTOM: (
        "Describe additional requests to generate beyond the standard set..."
    ),
    CustomStyle.MODIFIED_STANDARD: (
        "Describe how to modify the standard requests..."
    ),
}


# ---------------------------------------------------------------------------
# PartyEditDialog
# ---------------------------------------------------------------------------

class PartyEditDialog(QDialog):
    """Dialog for adding or editing a party."""

    def __init__(self, parent=None, party: Optional[Party] = None):
        super().__init__(parent)
        self.setWindowTitle("Edit Party" if party else "Add Party")
        self.setMinimumWidth(350)

        layout = QFormLayout(self)

        self.name_edit = QLineEdit()
        self.name_edit.setPlaceholderText("Party name")
        layout.addRow("Name:", self.name_edit)

        self.role_combo = QComboBox()
        self.role_combo.addItem("Plaintiff", PartyRole.PLAINTIFF)
        self.role_combo.addItem("Defendant", PartyRole.DEFENDANT)
        self.role_combo.addItem("Cross-Defendant", PartyRole.CROSS_DEFENDANT)
        self.role_combo.addItem("Cross-Complainant", PartyRole.CROSS_COMPLAINANT)
        layout.addRow("Role:", self.role_combo)

        self.our_client_cb = QCheckBox("Our client")
        layout.addRow("", self.our_client_cb)

        self.buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        self.buttons.accepted.connect(self.accept)
        self.buttons.rejected.connect(self.reject)
        layout.addRow(self.buttons)

        # Populate from existing party
        if party:
            self.name_edit.setText(party.name)
            idx = self.role_combo.findData(party.role)
            if idx >= 0:
                self.role_combo.setCurrentIndex(idx)
            self.our_client_cb.setChecked(party.is_our_client)

    def get_party(self) -> Party:
        """Return a Party constructed from the dialog fields."""
        return Party(
            name=self.name_edit.text().strip(),
            role=self.role_combo.currentData(),
            is_our_client=self.our_client_cb.isChecked(),
        )


# ---------------------------------------------------------------------------
# PropoundTab
# ---------------------------------------------------------------------------

class PropoundTab(QWidget):
    """Left + right pane for propounding discovery."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)

        # State
        self.file_number: Optional[str] = None
        self.parties: List[Party] = []
        self.cached_models: Dict[str, list] = {}
        self.fetcher: Optional[ModelFetcher] = None
        self.generated_sets: List[DiscoverySet] = []

        # Debounce timer for prompt auto-save
        self._prompt_save_timer = QTimer(self)
        self._prompt_save_timer.setSingleShot(True)
        self._prompt_save_timer.setInterval(800)
        self._prompt_save_timer.timeout.connect(self._save_prompt)

        # Debounce timer for output auto-save
        self._output_save_timer = QTimer(self)
        self._output_save_timer.setSingleShot(True)
        self._output_save_timer.setInterval(800)
        self._output_save_timer.timeout.connect(self._save_output)

        self._build_ui()
        self._connect_signals()
        # Apply initial visibility rules
        self._on_mode_changed()

    # ------------------------------------------------------------------
    # UI construction
    # ------------------------------------------------------------------

    def _build_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)

        self.splitter = QSplitter(Qt.Orientation.Horizontal)
        root.addWidget(self.splitter)

        # ---- Left pane (scroll area) ----
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setMinimumWidth(280)
        scroll.setMaximumWidth(420)

        left_widget = QWidget()
        self.left_layout = QVBoxLayout(left_widget)
        self.left_layout.setContentsMargins(8, 8, 8, 8)
        self.left_layout.setSpacing(10)
        scroll.setWidget(left_widget)
        self.splitter.addWidget(scroll)

        self._build_mode_group()
        self._build_standard_type()
        self._build_custom_style()
        self._build_discovery_types()
        self._build_directed_to()
        self._build_document_box()
        self._build_llm_section()
        self._build_prompt_section()
        self._build_generate_button()

        self.left_layout.addStretch()

        # ---- Right pane ----
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        right_layout.setContentsMargins(4, 4, 4, 4)

        self._build_right_toolbar(right_layout)
        self._build_right_content(right_layout)

        self.splitter.addWidget(right_widget)
        self.splitter.setStretchFactor(0, 0)
        self.splitter.setStretchFactor(1, 1)

    # --- Left-pane sections ---

    def _build_mode_group(self):
        group = QGroupBox("Discovery Mode")
        layout = QVBoxLayout(group)

        self.mode_group = QButtonGroup(self)
        self.rb_standard = QRadioButton("Initial \u2014 Standard")
        self.rb_custom = QRadioButton("Initial \u2014 Custom")
        self.rb_additional = QRadioButton("Additional Discovery")

        self.mode_group.addButton(self.rb_standard, 0)
        self.mode_group.addButton(self.rb_custom, 1)
        self.mode_group.addButton(self.rb_additional, 2)

        self.rb_standard.setChecked(True)

        layout.addWidget(self.rb_standard)
        layout.addWidget(self.rb_custom)
        layout.addWidget(self.rb_additional)

        self.left_layout.addWidget(group)

    def _build_standard_type(self):
        self.standard_type_group = QGroupBox("Standard Type")
        layout = QVBoxLayout(self.standard_type_group)

        self.standard_type_combo = QComboBox()
        self.standard_type_combo.addItem("Standard Negligence")
        self.standard_type_combo.addItem("Standard Wrongful Death (coming soon)")
        # Disable the coming-soon item
        model = self.standard_type_combo.model()
        item = model.item(1)
        item.setEnabled(False)

        layout.addWidget(self.standard_type_combo)
        self.left_layout.addWidget(self.standard_type_group)

    def _build_custom_style(self):
        self.custom_style_group = QGroupBox("Custom Style")
        layout = QVBoxLayout(self.custom_style_group)

        self.custom_style_btn_group = QButtonGroup(self)
        self.rb_custom_only = QRadioButton("Custom Only")
        self.rb_standard_plus = QRadioButton("Standard + Custom")
        self.rb_modified = QRadioButton("Modified Standard")

        self.custom_style_btn_group.addButton(self.rb_custom_only, 0)
        self.custom_style_btn_group.addButton(self.rb_standard_plus, 1)
        self.custom_style_btn_group.addButton(self.rb_modified, 2)

        self.rb_custom_only.setChecked(True)

        layout.addWidget(self.rb_custom_only)
        layout.addWidget(self.rb_standard_plus)
        layout.addWidget(self.rb_modified)

        self.left_layout.addWidget(self.custom_style_group)

    def _build_discovery_types(self):
        group = QGroupBox("Discovery Types")
        layout = QHBoxLayout(group)

        self.cb_si = QCheckBox("SI")
        self.cb_rpd = QCheckBox("RPD")
        self.cb_rfa = QCheckBox("RFA")

        self.cb_si.setChecked(True)
        self.cb_rpd.setChecked(True)
        self.cb_rfa.setChecked(False)

        layout.addWidget(self.cb_si)
        layout.addWidget(self.cb_rpd)
        layout.addWidget(self.cb_rfa)

        self.left_layout.addWidget(group)

    def _build_directed_to(self):
        group = QGroupBox("Directed To")
        layout = QHBoxLayout(group)

        self.party_combo = QComboBox()
        self.party_combo.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        layout.addWidget(self.party_combo)

        self.add_party_btn = QPushButton("+")
        self.add_party_btn.setFixedWidth(30)
        self.add_party_btn.setToolTip("Add a new party")
        layout.addWidget(self.add_party_btn)

        self.left_layout.addWidget(group)

        # Enable right-click context menu on the combo
        self.party_combo.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)

    def _build_document_box(self):
        group = QGroupBox("Context Documents")
        layout = QVBoxLayout(group)

        self.doc_list = ResizableListWidget()
        self.doc_list.setAcceptDrops(True)
        self.doc_list.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.doc_list.setToolTip("Drag and drop files here")
        layout.addWidget(self.doc_list)

        self.left_layout.addWidget(group)

    def _build_llm_section(self):
        self.llm_group = QGroupBox("LLM Provider / Model")
        layout = QVBoxLayout(self.llm_group)

        self.provider_combo = QComboBox()
        self.provider_combo.addItems(["Gemini", "OpenAI", "Claude"])
        layout.addWidget(self.provider_combo)

        self.model_combo = QComboBox()
        layout.addWidget(self.model_combo)

        self.left_layout.addWidget(self.llm_group)

    def _build_prompt_section(self):
        self.prompt_group = QGroupBox("Instructions / Prompt")
        layout = QVBoxLayout(self.prompt_group)

        self.prompt_edit = QTextEdit()
        self.prompt_edit.setPlaceholderText(
            "Describe what discovery requests to generate..."
        )
        self.prompt_edit.setMinimumHeight(80)
        self.prompt_edit.setMaximumHeight(200)
        layout.addWidget(self.prompt_edit)

        self.left_layout.addWidget(self.prompt_group)

    def _build_generate_button(self):
        self.generate_btn = QPushButton("Generate Discovery")
        self.generate_btn.setStyleSheet(
            "QPushButton { background-color: #2563eb; color: white; "
            "padding: 8px 16px; border-radius: 4px; font-weight: bold; }"
            "QPushButton:hover { background-color: #1d4ed8; }"
            "QPushButton:disabled { background-color: #93c5fd; }"
        )
        self.left_layout.addWidget(self.generate_btn)

    # --- Right-pane sections ---

    def _build_right_toolbar(self, parent_layout):
        toolbar = QHBoxLayout()

        self.save_btn = QPushButton("Save as .docx")
        self.save_all_btn = QPushButton("Save All")
        self.clear_btn = QPushButton("Clear")
        self.status_label = QLabel("")
        self.status_label.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)

        toolbar.addWidget(self.save_btn)
        toolbar.addWidget(self.save_all_btn)
        toolbar.addWidget(self.clear_btn)
        toolbar.addWidget(self.status_label)

        parent_layout.addLayout(toolbar)

    def _build_right_content(self, parent_layout):
        self.doc_tabs = QTabWidget()
        parent_layout.addWidget(self.doc_tabs)

        # Empty-state label (shown when no documents generated)
        self.empty_label = QLabel("Configure settings and click Generate...")
        self.empty_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.empty_label.setStyleSheet("color: #909090; font-size: 14px;")
        parent_layout.addWidget(self.empty_label)

        # Start with empty state visible, tabs hidden
        self.doc_tabs.hide()
        self.empty_label.show()

    # ------------------------------------------------------------------
    # Signal wiring
    # ------------------------------------------------------------------

    def _connect_signals(self):
        # Mode radios
        self.mode_group.buttonClicked.connect(lambda _: self._on_mode_changed())

        # Custom style radios -> update prompt placeholder
        self.custom_style_btn_group.buttonClicked.connect(
            lambda _: self._on_custom_style_changed()
        )

        # Party management
        self.add_party_btn.clicked.connect(self._on_add_party)
        self.party_combo.customContextMenuRequested.connect(self._on_party_context_menu)

        # Document box context menu
        self.doc_list.customContextMenuRequested.connect(self._on_doc_context_menu)

        # LLM provider change
        self.provider_combo.currentTextChanged.connect(self._update_models)

        # Generate / Save stubs
        self.generate_btn.clicked.connect(self._on_generate)
        self.save_btn.clicked.connect(self._save_current)
        self.save_all_btn.clicked.connect(self._save_all)
        self.clear_btn.clicked.connect(self._clear_editor)

        # Auto-save prompt on edits (debounced)
        self.prompt_edit.textChanged.connect(self._prompt_save_timer.start)

        # Kick off initial model fetch
        self._update_models(self.provider_combo.currentText())

    # ------------------------------------------------------------------
    # Conditional visibility
    # ------------------------------------------------------------------

    def _on_mode_changed(self):
        """Show/hide controls based on the selected discovery mode."""
        mode = self._current_mode()

        is_standard = mode == DiscoveryMode.INITIAL_STANDARD
        is_custom = mode == DiscoveryMode.INITIAL_CUSTOM
        is_additional = mode == DiscoveryMode.ADDITIONAL

        self.standard_type_group.setVisible(is_standard)
        self.custom_style_group.setVisible(is_custom)
        self.llm_group.setVisible(is_custom or is_additional)
        self.prompt_group.setVisible(is_custom or is_additional)

        # Button label
        if is_additional:
            self.generate_btn.setText("Generate Additional Discovery")
        else:
            self.generate_btn.setText("Generate Discovery")

        # Update prompt placeholder if custom
        if is_custom:
            self._on_custom_style_changed()

    def _on_custom_style_changed(self):
        """Update the prompt placeholder based on custom style selection."""
        style = self._current_custom_style()
        placeholder = _CUSTOM_STYLE_PLACEHOLDERS.get(
            style,
            "Describe what discovery requests to generate...",
        )
        self.prompt_edit.setPlaceholderText(placeholder)

    # ------------------------------------------------------------------
    # Mode / style helpers
    # ------------------------------------------------------------------

    def _current_mode(self) -> DiscoveryMode:
        btn_id = self.mode_group.checkedId()
        return {
            0: DiscoveryMode.INITIAL_STANDARD,
            1: DiscoveryMode.INITIAL_CUSTOM,
            2: DiscoveryMode.ADDITIONAL,
        }.get(btn_id, DiscoveryMode.INITIAL_STANDARD)

    def _current_custom_style(self) -> CustomStyle:
        btn_id = self.custom_style_btn_group.checkedId()
        return {
            0: CustomStyle.CUSTOM_ONLY,
            1: CustomStyle.STANDARD_PLUS_CUSTOM,
            2: CustomStyle.MODIFIED_STANDARD,
        }.get(btn_id, CustomStyle.CUSTOM_ONLY)

    def selected_discovery_types(self) -> List[DiscoveryType]:
        """Return list of checked discovery types."""
        types = []
        if self.cb_si.isChecked():
            types.append(DiscoveryType.SI)
        if self.cb_rpd.isChecked():
            types.append(DiscoveryType.RPD)
        if self.cb_rfa.isChecked():
            types.append(DiscoveryType.RFA)
        return types

    # ------------------------------------------------------------------
    # Party management
    # ------------------------------------------------------------------

    def _on_add_party(self):
        dlg = PartyEditDialog(self)
        if dlg.exec() == QDialog.DialogCode.Accepted:
            party = dlg.get_party()
            if not party.name:
                return
            self.parties.append(party)
            self._regenerate_abbreviations()
            self._refresh_party_combo()
            self._save_party_roster()

    def _on_party_context_menu(self, pos):
        idx = self.party_combo.currentIndex()
        if idx < 0 or idx >= len(self.parties):
            return

        menu = QMenu(self)
        edit_action = QAction("Edit", self)
        remove_action = QAction("Remove", self)
        menu.addAction(edit_action)
        menu.addAction(remove_action)

        action = menu.exec(self.party_combo.mapToGlobal(pos))
        if action == edit_action:
            self._edit_party(idx)
        elif action == remove_action:
            self._remove_party(idx)

    def _edit_party(self, idx: int):
        party = self.parties[idx]
        dlg = PartyEditDialog(self, party)
        if dlg.exec() == QDialog.DialogCode.Accepted:
            updated = dlg.get_party()
            if not updated.name:
                return
            self.parties[idx] = updated
            self._regenerate_abbreviations()
            self._refresh_party_combo()
            self._save_party_roster()

    def _remove_party(self, idx: int):
        del self.parties[idx]
        self._regenerate_abbreviations()
        self._refresh_party_combo()
        self._save_party_roster()

    def _regenerate_abbreviations(self):
        """Recalculate abbreviations for all parties."""
        for party in self.parties:
            party.abbreviation = generate_abbreviation(party, self.parties)

    def _refresh_party_combo(self):
        """Rebuild the party combo from self.parties."""
        self.party_combo.blockSignals(True)
        current_text = self.party_combo.currentText()
        self.party_combo.clear()
        for p in self.parties:
            label = f"{p.name} ({p.role_label})"
            self.party_combo.addItem(label)
        # Try to restore previous selection
        idx = self.party_combo.findText(current_text)
        if idx >= 0:
            self.party_combo.setCurrentIndex(idx)
        self.party_combo.blockSignals(False)

    def _save_party_roster(self):
        """Persist party roster back to CaseDataManager."""
        if not self.file_number:
            return
        try:
            from Scripts.case_data_manager import CaseDataManager
            manager = CaseDataManager()

            roster = [p.to_dict() for p in self.parties]
            manager.save_variable(
                self.file_number, "discovery_party_roster", roster,
                source="discovery_tab", auto_tag=False,
            )

            # Also sync plaintiffs / defendants variables
            plaintiffs = [p.name for p in self.parties if p.role == PartyRole.PLAINTIFF]
            defendants = [p.name for p in self.parties if p.role == PartyRole.DEFENDANT]
            if plaintiffs:
                manager.save_variable(
                    self.file_number, "plaintiffs", plaintiffs,
                    source="discovery_tab", auto_tag=False,
                )
            if defendants:
                manager.save_variable(
                    self.file_number, "defendants", defendants,
                    source="discovery_tab", auto_tag=False,
                )
        except Exception as e:
            print(f"[DiscoveryTab] Failed to save party roster: {e}")

    def _load_parties_from_case(self):
        """Load party roster from CaseDataManager, seeding from plaintiffs/defendants if needed."""
        if not self.file_number:
            return
        try:
            from Scripts.case_data_manager import CaseDataManager
            manager = CaseDataManager()

            # Try saved roster first
            roster_data = manager.get_value(self.file_number, "discovery_party_roster")
            if roster_data and isinstance(roster_data, list):
                self.parties = [Party.from_dict(d) for d in roster_data]
                self._regenerate_abbreviations()
                self._refresh_party_combo()
                return

            # Seed from plaintiffs / defendants
            self.parties = []
            client_name = manager.get_value(self.file_number, "client_name") or ""
            if isinstance(client_name, list):
                client_name = client_name[0] if client_name else ""
            client_name = str(client_name).strip().lower()

            plaintiffs = manager.get_value(self.file_number, "plaintiffs")
            if isinstance(plaintiffs, list):
                for name in plaintiffs:
                    n = str(name).strip()
                    is_client = bool(client_name and client_name in n.lower())
                    self.parties.append(Party(
                        name=n, role=PartyRole.PLAINTIFF, is_our_client=is_client,
                    ))
            elif isinstance(plaintiffs, str) and plaintiffs.strip():
                n = plaintiffs.strip()
                is_client = bool(client_name and client_name in n.lower())
                self.parties.append(Party(
                    name=n, role=PartyRole.PLAINTIFF, is_our_client=is_client,
                ))

            defendants = manager.get_value(self.file_number, "defendants")
            if isinstance(defendants, list):
                for name in defendants:
                    n = str(name).strip()
                    is_client = bool(client_name and client_name in n.lower())
                    self.parties.append(Party(
                        name=n, role=PartyRole.DEFENDANT, is_our_client=is_client,
                    ))
            elif isinstance(defendants, str) and defendants.strip():
                n = defendants.strip()
                is_client = bool(client_name and client_name in n.lower())
                self.parties.append(Party(
                    name=n, role=PartyRole.DEFENDANT, is_our_client=is_client,
                ))

            self._regenerate_abbreviations()
            self._refresh_party_combo()
        except Exception as e:
            print(f"[DiscoveryTab] Failed to load parties: {e}")

    # ------------------------------------------------------------------
    # Document box (context documents)
    # ------------------------------------------------------------------

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.accept()
        else:
            event.ignore()

    def dropEvent(self, event: QDropEvent):
        """Handle file drops with deferred processing."""
        files_to_add = []
        for url in event.mimeData().urls():
            path = url.toLocalFile()
            if os.path.isfile(path) and path.lower().endswith(_SUPPORTED_DROP_EXTENSIONS):
                files_to_add.append(path)

        event.accept()

        if files_to_add:
            QTimer.singleShot(0, lambda: self._process_dropped_files(files_to_add))

    def _process_dropped_files(self, file_paths: List[str]):
        """Add dropped files to the document list."""
        for path in file_paths:
            self._add_document(path)

    def _add_document(self, path: str):
        """Add a single document to the list (skip duplicates)."""
        # Check for duplicates
        for i in range(self.doc_list.count()):
            existing = self.doc_list.item(i).data(Qt.ItemDataRole.UserRole)
            if existing == path:
                return

        item = QListWidgetItem(os.path.basename(path))
        item.setData(Qt.ItemDataRole.UserRole, path)
        item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
        item.setCheckState(Qt.CheckState.Checked)
        item.setToolTip(path)
        self.doc_list.addItem(item)
        self._save_documents()

    def _on_doc_context_menu(self, pos):
        item = self.doc_list.itemAt(pos)
        if not item:
            return
        menu = QMenu(self)
        remove_action = QAction("Remove", self)
        menu.addAction(remove_action)

        action = menu.exec(self.doc_list.mapToGlobal(pos))
        if action == remove_action:
            row = self.doc_list.row(item)
            self.doc_list.takeItem(row)
            self._save_documents()

    def read_files_content(self) -> str:
        """Read text from all checked documents in the list.

        Extracts text from .txt, .docx (via python-docx), and .pdf (via
        PyMuPDF).  Other formats are noted but not extracted.
        """
        content = ""
        for i in range(self.doc_list.count()):
            item = self.doc_list.item(i)
            if item.checkState() != Qt.CheckState.Checked:
                continue
            path = item.data(Qt.ItemDataRole.UserRole)
            if not path or not os.path.isfile(path):
                continue

            ext = os.path.splitext(path)[1].lower()
            content += f"\n--- FILE: {os.path.basename(path)} ---\n"

            try:
                if ext == ".txt":
                    with open(path, "r", encoding="utf-8", errors="ignore") as f:
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
                    else:
                        content += "[PyMuPDF not available]\n"
                elif ext == ".msg":
                    try:
                        import pythoncom
                        import win32com.client
                        pythoncom.CoInitialize()
                        outlook = win32com.client.Dispatch("Outlook.Application")
                        namespace = outlook.GetNamespace("MAPI")
                        msg = namespace.OpenSharedItem(os.path.abspath(path))
                        subject = msg.Subject or ""
                        sender = msg.SenderName or ""
                        body = msg.Body or ""
                        msg.Close(0)
                        pythoncom.CoUninitialize()
                        content += f"From: {sender}\nSubject: {subject}\n\n{body}\n"
                    except Exception as msg_err:
                        content += f"[Error reading .msg: {msg_err}]\n"
                else:
                    content += f"[Unsupported format for text extraction: {ext}]\n"
            except Exception as e:
                content += f"[Error reading file: {e}]\n"

        return content

    # ------------------------------------------------------------------
    # LLM provider / model
    # ------------------------------------------------------------------

    def _update_models(self, provider: str):
        """Fetch and populate the model combo for the selected provider."""
        self.model_combo.clear()

        # Check cache first
        if provider in self.cached_models:
            self.model_combo.addItems(self.cached_models[provider])
            return

        api_key = API_KEYS.get(provider)
        self.model_combo.addItem("Fetching models...")
        self.model_combo.setEnabled(False)

        # Clean up previous fetcher
        if self.fetcher is not None:
            try:
                self.fetcher.finished.disconnect()
                self.fetcher.error.disconnect()
            except (TypeError, RuntimeError):
                pass
            if self.fetcher.isRunning():
                self.fetcher.wait(1000)

        self.fetcher = ModelFetcher(provider, api_key)
        self.fetcher.finished.connect(self._on_models_fetched)
        self.fetcher.error.connect(
            lambda err: self._on_models_fetched(provider, [f"Error: {err}"])
        )
        self.fetcher.start()

    def _on_models_fetched(self, provider: str, models: list):
        self.model_combo.clear()
        self.model_combo.setEnabled(True)

        if models and models[0].startswith("Error:"):
            self.model_combo.addItems(models)
            return

        self.cached_models[provider] = models
        self.model_combo.addItems(models)

        # Default selection
        if provider == "Gemini":
            idx = self.model_combo.findText("gemini-3.1-flash-lite")
            if idx == -1:
                idx = self.model_combo.findText("gemini-3.5-flash")
            if idx != -1:
                self.model_combo.setCurrentIndex(idx)
            else:
                for i in range(self.model_combo.count()):
                    if "gemini-3" in self.model_combo.itemText(i).lower():
                        self.model_combo.setCurrentIndex(i)
                        break

    # ------------------------------------------------------------------
    # Case loading
    # ------------------------------------------------------------------

    def load_case(self, file_number: str):
        """Load case data — called when the active case changes."""
        # Save current case state before switching
        if self.file_number and self.file_number != file_number:
            self._save_state()

        # Clear all UI state
        self._clear_state()

        self.file_number = file_number
        self._load_parties_from_case()
        self._load_state()

    def _clear_state(self):
        """Clear all UI state (documents, prompt, output) for a case switch."""
        self.doc_list.clear()
        self.prompt_edit.blockSignals(True)
        self.prompt_edit.clear()
        self.prompt_edit.blockSignals(False)
        self.generated_sets = []
        self.doc_tabs.clear()
        self.doc_tabs.hide()
        self.empty_label.show()
        self.status_label.setText("")

    # ------------------------------------------------------------------
    # Session persistence (per-case)
    # ------------------------------------------------------------------

    def _get_manager(self):
        """Return a CaseDataManager instance, or None on error."""
        try:
            from Scripts.case_data_manager import CaseDataManager
            return CaseDataManager()
        except Exception as e:
            logger.warning("CaseDataManager unavailable: %s", e)
            return None

    def _save_state(self):
        """Persist documents, prompt, and output for the current case."""
        self._save_documents()
        self._save_prompt()
        self._save_output()

    def _load_state(self):
        """Restore documents, prompt, and output for the current case."""
        self._load_documents()
        self._load_prompt()
        self._load_output()

    def _save_documents(self):
        """Persist the document list to CaseDataManager."""
        if not self.file_number:
            return
        manager = self._get_manager()
        if not manager:
            return
        docs = []
        for i in range(self.doc_list.count()):
            item = self.doc_list.item(i)
            path = item.data(Qt.ItemDataRole.UserRole)
            checked = item.checkState() == Qt.CheckState.Checked
            if path:
                docs.append({"path": path, "checked": checked})
        try:
            manager.save_variable(
                self.file_number, "discovery_documents", docs,
                source="discovery_tab", auto_tag=False,
            )
        except Exception as e:
            logger.warning("Failed to save discovery documents: %s", e)

    def _load_documents(self):
        """Restore the document list from CaseDataManager."""
        if not self.file_number:
            return
        manager = self._get_manager()
        if not manager:
            return
        try:
            docs = manager.get_value(self.file_number, "discovery_documents")
            if not docs or not isinstance(docs, list):
                return
            for entry in docs:
                path = entry.get("path", "") if isinstance(entry, dict) else str(entry)
                if not path or not os.path.isfile(path):
                    continue
                item = QListWidgetItem(os.path.basename(path))
                item.setData(Qt.ItemDataRole.UserRole, path)
                item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
                checked = entry.get("checked", True) if isinstance(entry, dict) else True
                item.setCheckState(
                    Qt.CheckState.Checked if checked else Qt.CheckState.Unchecked
                )
                item.setToolTip(path)
                self.doc_list.addItem(item)
        except Exception as e:
            logger.warning("Failed to load discovery documents: %s", e)

    @Slot()
    def _save_prompt(self):
        """Persist the prompt text to CaseDataManager."""
        if not self.file_number:
            return
        manager = self._get_manager()
        if not manager:
            return
        try:
            manager.save_variable(
                self.file_number, "discovery_prompt",
                self.prompt_edit.toPlainText(),
                source="discovery_tab", auto_tag=False,
            )
        except Exception as e:
            logger.warning("Failed to save discovery prompt: %s", e)

    def _load_prompt(self):
        """Restore the prompt text from CaseDataManager."""
        if not self.file_number:
            return
        manager = self._get_manager()
        if not manager:
            return
        try:
            text = manager.get_value(self.file_number, "discovery_prompt")
            if text and isinstance(text, str):
                self.prompt_edit.blockSignals(True)
                self.prompt_edit.setPlainText(text)
                self.prompt_edit.blockSignals(False)
        except Exception as e:
            logger.warning("Failed to load discovery prompt: %s", e)

    def _save_output(self):
        """Persist the editor output tabs to CaseDataManager."""
        if not self.file_number:
            return
        manager = self._get_manager()
        if not manager:
            return
        tabs_data = []
        for i in range(self.doc_tabs.count()):
            editor = self.doc_tabs.widget(i)
            label = self.doc_tabs.tabText(i)
            text = ""
            if isinstance(editor, QPlainTextEdit):
                text = editor.toPlainText()
            tabs_data.append({"label": label, "text": text})
        try:
            manager.save_variable(
                self.file_number, "discovery_output", tabs_data,
                source="discovery_tab", auto_tag=False,
            )
        except Exception as e:
            logger.warning("Failed to save discovery output: %s", e)

    def _load_output(self):
        """Restore the editor output tabs from CaseDataManager."""
        if not self.file_number:
            return
        manager = self._get_manager()
        if not manager:
            return
        try:
            tabs_data = manager.get_value(self.file_number, "discovery_output")
            if not tabs_data or not isinstance(tabs_data, list):
                return
            has_content = any(
                t.get("text", "").strip() for t in tabs_data if isinstance(t, dict)
            )
            if not has_content:
                return
            self.empty_label.hide()
            self.doc_tabs.show()
            for entry in tabs_data:
                if not isinstance(entry, dict):
                    continue
                label = entry.get("label", "")
                text = entry.get("text", "")
                editor = QPlainTextEdit()
                editor.setPlainText(text)
                editor.setLineWrapMode(QPlainTextEdit.LineWrapMode.WidgetWidth)
                editor.textChanged.connect(self._output_save_timer.start)
                self.doc_tabs.addTab(editor, label)
        except Exception as e:
            logger.warning("Failed to load discovery output: %s", e)

    # ------------------------------------------------------------------
    # Case path resolution
    # ------------------------------------------------------------------

    @property
    def case_path(self) -> Optional[str]:
        """Resolve the physical case folder for the current file_number."""
        if not self.file_number:
            return None
        try:
            from icharlotte_core.utils import get_case_path
            return get_case_path(self.file_number)
        except Exception:
            return None

    # ------------------------------------------------------------------
    # Generate, LLM, display, and save
    # ------------------------------------------------------------------

    def _on_generate(self):
        """Generate discovery based on current mode and settings."""
        # 1. Gather inputs
        mode = self._current_mode()
        disc_types = self.selected_discovery_types()
        directed_to = self._get_directed_to_party()
        our_client = self._get_our_client()

        # 2. Validate
        if not disc_types:
            self.status_label.setText("Select at least one discovery type.")
            return
        if directed_to is None:
            self.status_label.setText("Select the party discovery is directed to.")
            return
        if our_client is None:
            self.status_label.setText(
                "Mark one party as 'Our client' (right-click party to edit)."
            )
            return

        # 3. Create engine
        discovery_dir = os.path.join(os.getcwd(), "discovery")
        engine = DiscoveryEngine(discovery_dir)

        # 4. Build variables
        # For standard negligence, defendant is typically the responding party
        defendant = None
        for p in self.parties:
            if p.role == PartyRole.DEFENDANT:
                defendant = p
                break

        variables = {
            "responding_party_name": directed_to.name,
            "propounding_party_name": our_client.name,
            "defendant_name": defendant.name if defendant else directed_to.name,
        }

        # 5. Disable generate button, show status
        self.generate_btn.setEnabled(False)
        self.status_label.setText("Generating...")

        # 6. Branch by mode
        if mode == DiscoveryMode.INITIAL_STANDARD:
            try:
                results = engine.generate_standard(
                    standard_type="negligence",
                    discovery_types=disc_types,
                    directed_to=directed_to,
                    propounding_party=our_client,
                    variables=variables,
                )
                self._display_results(results)
            except Exception as e:
                logger.exception("Standard generation failed")
                self.status_label.setText(f"Error: {e}")
            finally:
                self.generate_btn.setEnabled(True)

        elif mode == DiscoveryMode.INITIAL_CUSTOM:
            custom_style = self._current_custom_style()
            user_prompt = self.prompt_edit.toPlainText().strip()

            if not user_prompt:
                self.status_label.setText("Enter instructions in the prompt box.")
                self.generate_btn.setEnabled(True)
                return

            context_text = self.read_files_content()

            try:
                results = engine.generate_custom(
                    custom_style=custom_style,
                    standard_type="negligence",
                    discovery_types=disc_types,
                    directed_to=directed_to,
                    propounding_party=our_client,
                    user_instructions=user_prompt,
                    context_text=context_text,
                    variables=variables,
                )
                self.generated_sets = results

                # For each DiscoverySet, build LLM prompt and run in background
                for ds in results:
                    # Determine start number for LLM-generated requests
                    start_number = len(ds.requests) + 1 + ds.previous_count

                    # Build the standard requests text for reference
                    standard_text = ds.plain_text() if ds.requests else ""

                    prompt = engine.build_llm_prompt(
                        discovery_type=ds.discovery_type,
                        user_instructions=user_prompt,
                        custom_style=custom_style,
                        context_text=context_text,
                        start_number=start_number,
                        standard_requests_text=standard_text,
                    )
                    self._run_llm(prompt, ds, custom_style)
            except Exception as e:
                logger.exception("Custom generation failed")
                self.status_label.setText(f"Error: {e}")
                self.generate_btn.setEnabled(True)

        elif mode == DiscoveryMode.ADDITIONAL:
            user_prompt = self.prompt_edit.toPlainText().strip()
            if not user_prompt:
                self.status_label.setText("Enter instructions in the prompt box.")
                self.generate_btn.setEnabled(True)
                return

            cp = self.case_path
            if not cp or not os.path.isdir(cp):
                self.status_label.setText("Could not resolve case folder for this file number.")
                self.generate_btn.setEnabled(True)
                return

            context_text = self.read_files_content()

            try:
                results = engine.prepare_additional(
                    case_path=cp,
                    discovery_types=disc_types,
                    directed_to=directed_to,
                    propounding_party=our_client,
                )
                self.generated_sets = results

                for ds in results:
                    start_number = ds.previous_count + 1

                    prompt = engine.build_llm_prompt(
                        discovery_type=ds.discovery_type,
                        user_instructions=user_prompt,
                        custom_style=CustomStyle.CUSTOM_ONLY,
                        context_text=context_text,
                        start_number=start_number,
                    )
                    self._run_llm(prompt, ds)
            except Exception as e:
                logger.exception("Additional generation failed")
                self.status_label.setText(f"Error: {e}")
                self.generate_btn.setEnabled(True)

    def _get_directed_to_party(self) -> Optional[Party]:
        """Return the selected 'directed to' party from the combo box."""
        idx = self.party_combo.currentIndex()
        if 0 <= idx < len(self.parties):
            return self.parties[idx]
        return None

    def _get_our_client(self) -> Optional[Party]:
        """Return the party marked as 'our client', or None."""
        for p in self.parties:
            if p.is_our_client:
                return p
        return None

    # ------------------------------------------------------------------
    # LLM worker
    # ------------------------------------------------------------------

    def _run_llm(self, prompt: str, discovery_set: DiscoverySet,
                 custom_style: Optional[CustomStyle] = None):
        """Kick off an LLM worker thread for a single DiscoverySet."""
        provider = self.provider_combo.currentText()
        model = self.model_combo.currentText()

        worker = LLMWorker(
            provider=provider,
            model=model,
            system="You are a California litigation attorney.",
            user=prompt,
            files=[],
            settings={"stream": False},
        )

        # Store custom_style on the worker so the finished handler can check it
        worker._ds = discovery_set
        worker._custom_style = custom_style

        worker.finished.connect(
            lambda text, _ds=discovery_set, _cs=custom_style: self._on_llm_finished(text, _ds, _cs)
        )
        worker.error.connect(self._on_llm_error)

        # Store a reference so the worker isn't garbage collected
        if not hasattr(self, "_llm_workers"):
            self._llm_workers = []
        self._llm_workers.append(worker)

        worker.start()

    def _on_llm_finished(self, response_text: str, discovery_set: DiscoverySet,
                         custom_style: Optional[CustomStyle] = None):
        """Handle LLM completion for a DiscoverySet."""
        parsed = extract_requests_from_text(response_text, discovery_set.discovery_type)

        if custom_style == CustomStyle.STANDARD_PLUS_CUSTOM and discovery_set.requests:
            # Extend: renumber the new requests to continue from the existing ones
            offset = len(discovery_set.requests) + discovery_set.previous_count
            for req in parsed:
                req.number = offset + parsed.index(req) + 1
            discovery_set.requests.extend(parsed)
        else:
            # Replace: use parsed requests directly
            discovery_set.requests = parsed

        self._display_results(self.generated_sets)
        self.generate_btn.setEnabled(True)
        self.status_label.setText(
            f"Generated {len(discovery_set.requests)} requests for "
            f"{discovery_set.discovery_type.abbreviation}, "
            f"Set {discovery_set.set_word}."
        )

    def _on_llm_error(self, error_msg: str):
        """Handle LLM error."""
        self.generate_btn.setEnabled(True)
        self.status_label.setText(f"LLM Error: {error_msg}")
        logger.error("LLM error during discovery generation: %s", error_msg)

    # ------------------------------------------------------------------
    # Display results
    # ------------------------------------------------------------------

    def _display_results(self, results: List[DiscoverySet]):
        """Display generated discovery sets in the right-pane tabs."""
        self.generated_sets = results

        # Clear existing tabs
        self.doc_tabs.clear()

        if not results:
            self.doc_tabs.hide()
            self.empty_label.show()
            return

        # Hide empty state, show tabs + save buttons
        self.empty_label.hide()
        self.doc_tabs.show()

        for ds in results:
            editor = QPlainTextEdit()
            editor.setPlainText(ds.plain_text())
            editor.setLineWrapMode(QPlainTextEdit.LineWrapMode.WidgetWidth)
            editor.textChanged.connect(self._output_save_timer.start)

            # Build tab label: e.g. "SI(1) tPltf"
            directed_abbr = ds.directed_to.abbreviation or ds.directed_to.name.split()[0]
            tab_label = f"{ds.discovery_type.abbreviation}({ds.set_number}) t{directed_abbr}"
            self.doc_tabs.addTab(editor, tab_label)

        # Update status
        total_requests = sum(len(ds.requests) for ds in results)
        set_count = len(results)
        self.status_label.setText(
            f"{total_requests} requests across {set_count} "
            f"{'set' if set_count == 1 else 'sets'}."
        )
        self._save_output()

    def _clear_editor(self):
        """Clear the right-pane editor, resetting to empty state."""
        self.generated_sets = []
        self.doc_tabs.clear()
        self.doc_tabs.hide()
        self.empty_label.show()
        self.status_label.setText("")
        self._save_output()

    # ------------------------------------------------------------------
    # Save
    # ------------------------------------------------------------------

    def _save_current(self):
        """Save the currently visible discovery set tab as a .docx."""
        idx = self.doc_tabs.currentIndex()
        if idx < 0 or idx >= len(self.generated_sets):
            self.status_label.setText("No document to save.")
            return

        ds = self.generated_sets[idx]
        editor = self.doc_tabs.widget(idx)
        if isinstance(editor, QPlainTextEdit):
            plain_text = editor.toPlainText()
        else:
            plain_text = ds.plain_text()

        self._save_discovery_set(ds, plain_text)

    def _save_all(self):
        """Save all generated discovery sets as .docx files."""
        if not self.generated_sets:
            self.status_label.setText("No documents to save.")
            return

        saved = []
        for idx, ds in enumerate(self.generated_sets):
            editor = self.doc_tabs.widget(idx)
            if isinstance(editor, QPlainTextEdit):
                plain_text = editor.toPlainText()
            else:
                plain_text = ds.plain_text()
            result = self._save_discovery_set(ds, plain_text)
            if result:
                saved.append(result)

        if saved:
            self.status_label.setText(
                f"Saved {len(saved)} file{'s' if len(saved) != 1 else ''}: "
                f"{', '.join(os.path.basename(s) for s in saved)}"
            )

    def _save_discovery_set(self, ds: DiscoverySet, plain_text: str) -> Optional[str]:
        """Save a single DiscoverySet to a .docx file. Returns the output path or None."""
        cp = self.case_path
        if not cp or not os.path.isdir(cp):
            self.status_label.setText("Could not resolve case folder.")
            return None

        caption_path = DiscoveryAssembler.find_caption_page(cp)
        if not caption_path:
            self.status_label.setText("No caption page found in case folder.")
            return None

        output_dir = os.path.join(cp, "NOTES", "AI OUTPUT", "DISCOVERY REQUESTS")
        output_path = os.path.join(output_dir, ds.filename)

        try:
            assembler = DiscoveryAssembler(caption_path)
            assembler.assemble_from_plain_text(
                plain_text=plain_text,
                discovery_set=ds,
                output_path=output_path,
            )
            self.status_label.setText(f"Saved: {ds.filename}")
            logger.info("Saved discovery document: %s", output_path)
            return output_path
        except Exception as e:
            logger.exception("Failed to save discovery set")
            self.status_label.setText(f"Save error: {e}")
            return None


# ---------------------------------------------------------------------------
# DiscoveryTab — Main tab with Propound / Respond sub-tabs
# ---------------------------------------------------------------------------

class DiscoveryTab(QWidget):
    """Top-level Discovery tab containing Propound and Respond sub-tabs."""

    def __init__(self, parent=None):
        super().__init__(parent)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)

        self.tabs = QTabWidget()
        layout.addWidget(self.tabs)

        # Propound sub-tab (full implementation)
        self.propound_tab = PropoundTab()
        self.tabs.addTab(self.propound_tab, "Propound")

        # Respond sub-tab (lazy import to avoid circular dependency)
        from icharlotte_core.ui.respond_tab import RespondTab
        self.respond_tab = RespondTab()
        self.tabs.addTab(self.respond_tab, "Respond")

    def load_case(self, file_number: str):
        """Delegate to both sub-tabs."""
        self.propound_tab.load_case(file_number)
        self.respond_tab.load_case(file_number)

    def save_state(self):
        """Save state for both sub-tabs (called on app close)."""
        if self.propound_tab.file_number:
            self.propound_tab._save_state()
        if self.respond_tab.file_number:
            self.respond_tab._save_state()
