"""
Discovery Tab — Propound / Respond sub-tabs for generating written discovery.

Implements the full left-pane controls (mode, type, party, documents, LLM,
prompt, generate button) and right-pane editor shell (toolbar, document
sub-tabs, empty state).
"""

import os
from typing import Dict, List, Optional

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QTabWidget, QSplitter, QLabel,
    QComboBox, QRadioButton, QButtonGroup, QCheckBox, QTextEdit,
    QPushButton, QListWidgetItem, QMenu, QScrollArea, QGroupBox,
    QDialog, QLineEdit, QFormLayout, QDialogButtonBox, QSizePolicy,
)
from PySide6.QtCore import Qt, QTimer, Signal, QFileInfo
from PySide6.QtGui import QDragEnterEvent, QDropEvent, QAction

from icharlotte_core.ui.tabs import ResizableListWidget
from icharlotte_core.llm import ModelFetcher
from icharlotte_core.config import API_KEYS
from icharlotte_core.discovery.models import (
    Party, PartyRole, DiscoveryMode, DiscoveryType, CustomStyle,
    generate_abbreviation,
)

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
        self.status_label = QLabel("")
        self.status_label.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)

        toolbar.addWidget(self.save_btn)
        toolbar.addWidget(self.save_all_btn)
        toolbar.addWidget(self.status_label)

        parent_layout.addLayout(toolbar)

    def _build_right_content(self, parent_layout):
        self.doc_tabs = QTabWidget()
        parent_layout.addWidget(self.doc_tabs)

        # Empty-state label (shown when no documents generated)
        self.empty_label = QLabel("Configure settings and click Generate...")
        self.empty_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.empty_label.setStyleSheet("color: #888; font-size: 14px;")
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
            plaintiffs = manager.get_value(self.file_number, "plaintiffs")
            if isinstance(plaintiffs, list):
                for name in plaintiffs:
                    self.parties.append(Party(
                        name=str(name), role=PartyRole.PLAINTIFF, is_our_client=False,
                    ))
            elif isinstance(plaintiffs, str) and plaintiffs.strip():
                self.parties.append(Party(
                    name=plaintiffs.strip(), role=PartyRole.PLAINTIFF, is_our_client=False,
                ))

            defendants = manager.get_value(self.file_number, "defendants")
            if isinstance(defendants, list):
                for name in defendants:
                    self.parties.append(Party(
                        name=str(name), role=PartyRole.DEFENDANT, is_our_client=False,
                    ))
            elif isinstance(defendants, str) and defendants.strip():
                self.parties.append(Party(
                    name=defendants.strip(), role=PartyRole.DEFENDANT, is_our_client=False,
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
        if not api_key:
            self.model_combo.addItem(f"No API Key for {provider}")
            return

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
            idx = self.model_combo.findText("gemini-3.1-flash-lite-preview")
            if idx == -1:
                idx = self.model_combo.findText("gemini-3-flash-preview")
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
        self.file_number = file_number
        self._load_parties_from_case()

    # ------------------------------------------------------------------
    # Stubs (wired in Task 8)
    # ------------------------------------------------------------------

    def _on_generate(self):
        pass

    def _save_current(self):
        pass

    def _save_all(self):
        pass


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

        # Respond sub-tab (placeholder)
        respond_placeholder = QLabel("Respond tab \u2014 coming soon")
        respond_placeholder.setAlignment(Qt.AlignmentFlag.AlignCenter)
        respond_placeholder.setStyleSheet("color: #888; font-size: 14px;")
        self.tabs.addTab(respond_placeholder, "Respond")

    def load_case(self, file_number: str):
        """Delegate to PropoundTab."""
        self.propound_tab.load_case(file_number)
