"""
Respond Tab — UI widget for generating discovery responses.

Mirrors PropoundTab's left/right pane split, party management, LLM integration,
and state persistence.  Orchestrates the three-phase response pipeline:
  Phase 1: Parse incoming discovery (response_parser)
  Phase 2: Select objections (objection_selector)
  Phase 3: Draft substantive responses (response_drafter)
"""

import logging
import os
from typing import Dict, List, Optional

from PySide6.QtCore import Qt, QTimer, Slot
from PySide6.QtGui import QAction, QDragEnterEvent, QDropEvent
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QSplitter, QLabel,
    QPushButton, QComboBox, QTabWidget, QPlainTextEdit,
    QGroupBox, QScrollArea, QSizePolicy, QMenu,
    QListWidgetItem, QDialog,
)

from icharlotte_core.llm import LLMWorker, ModelFetcher
from icharlotte_core.config import API_KEYS
from icharlotte_core.discovery.response_rules import ResponseRules
from icharlotte_core.discovery.response_parser import (
    ParsedDiscovery, build_parse_prompt, parse_llm_response,
    detect_discovery_type,
)
from icharlotte_core.discovery.objection_selector import (
    ObjectionMenu, select_fi_objections, rule_based_preselect,
    build_objection_prompt, parse_objection_response, merge_objections,
    format_objections,
)
from icharlotte_core.discovery.response_drafter import (
    get_fi_fixed_response, build_si_prompt, build_rfa_prompt,
    build_rpd_prompt, build_fi_17_1_prompt, assemble_plain_text,
    detect_inapplicable_fi,
)
from icharlotte_core.discovery.response_assembler import (
    ResponseAssembler, build_response_title, build_response_filename,
)
from icharlotte_core.discovery.assembler import DiscoveryAssembler
from icharlotte_core.ui.response_rules_dialog import ResponseRulesDialog
from icharlotte_core.ui.tabs import ResizableListWidget

logger = logging.getLogger(__name__)

try:
    import fitz  # PyMuPDF
except ImportError:
    fitz = None

try:
    from Scripts.case_data_manager import CaseDataManager
except ImportError:
    CaseDataManager = None


# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

_DISCOVERY_EXTENSIONS = (".pdf",)
_CONTEXT_EXTENSIONS = (".pdf", ".docx", ".txt")

_SUPPORTED_DROP_EXTENSIONS = _DISCOVERY_EXTENSIONS + (".docx", ".txt")

# Gemini 1M token limit ≈ 3-4M chars for English text.  Leave headroom
# for the response and internal overhead by capping user prompts at ~2.5M chars.
_MAX_PROMPT_CHARS = 2_500_000


def _truncate_prompt(text: str, limit: int = _MAX_PROMPT_CHARS) -> str:
    """Truncate a prompt string to stay within token limits."""
    if len(text) <= limit:
        return text
    return text[:limit] + "\n\n[... content truncated to fit model context window ...]"


# ---------------------------------------------------------------------------
# RespondTab
# ---------------------------------------------------------------------------

class RespondTab(QWidget):
    """Left + right pane for generating discovery responses."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)

        # State
        self.file_number: Optional[str] = None
        self.rules = ResponseRules()
        self.cached_models: Dict[str, list] = {}
        self.fetcher: Optional[ModelFetcher] = None
        self._llm_workers: list = []
        self._parsed_discoveries: Dict[str, ParsedDiscovery] = {}
        self._rfa_responses_cache: Optional[str] = None
        self._objection_menu = ObjectionMenu.load_defaults()

        # Debounce timer for output auto-save
        self._output_save_timer = QTimer(self)
        self._output_save_timer.setSingleShot(True)
        self._output_save_timer.setInterval(800)
        self._output_save_timer.timeout.connect(self._save_output)

        self._build_ui()
        self._connect_signals()

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

        self._build_discovery_input()
        self._build_context_documents()
        self._build_llm_section()
        self._build_buttons()

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

    def _build_discovery_input(self):
        group = QGroupBox("Discovery to Respond To")
        layout = QVBoxLayout(group)

        header = QLabel("Drop discovery PDFs here (Form Interrogatories, "
                        "Special Interrogatories, RFAs, RPDs)")
        header.setWordWrap(True)
        header.setStyleSheet("color: #2563eb; font-size: 11px;")
        layout.addWidget(header)

        self.discovery_list = ResizableListWidget()
        self.discovery_list.setAcceptDrops(True)
        self.discovery_list.setContextMenuPolicy(
            Qt.ContextMenuPolicy.CustomContextMenu
        )
        self.discovery_list.setToolTip("Drag and drop discovery PDFs here")
        layout.addWidget(self.discovery_list)

        self.left_layout.addWidget(group)

    def _build_context_documents(self):
        group = QGroupBox("Context Documents")
        layout = QVBoxLayout(group)

        header = QLabel("Case documents for the LLM to reference when "
                        "drafting responses")
        header.setWordWrap(True)
        header.setStyleSheet("color: #b8860b; font-size: 11px;")
        layout.addWidget(header)

        self.context_list = ResizableListWidget()
        self.context_list.setAcceptDrops(True)
        self.context_list.setContextMenuPolicy(
            Qt.ContextMenuPolicy.CustomContextMenu
        )
        self.context_list.setToolTip("Drag and drop context documents here")
        layout.addWidget(self.context_list)

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

    def _build_buttons(self):
        self.rules_btn = QPushButton("Response Rules...")
        self.left_layout.addWidget(self.rules_btn)

        self.generate_btn = QPushButton("Generate Responses")
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

        self.refresh_17_1_btn = QPushButton("Refresh FI 17.1")
        self.refresh_17_1_btn.setEnabled(False)
        self.refresh_17_1_btn.setToolTip(
            "Re-generate FI 17.1 response using current RFA answers"
        )
        self.refresh_17_1_btn.setStyleSheet(
            "QPushButton { background-color: #e11d48; color: white; "
            "padding: 4px 10px; border-radius: 3px; }"
            "QPushButton:hover { background-color: #be123c; }"
            "QPushButton:disabled { background-color: #fda4af; color: #fff; }"
        )

        self.status_label = QLabel("")
        self.status_label.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed
        )

        toolbar.addWidget(self.save_btn)
        toolbar.addWidget(self.save_all_btn)
        toolbar.addWidget(self.clear_btn)
        toolbar.addWidget(self.refresh_17_1_btn)
        toolbar.addWidget(self.status_label)

        parent_layout.addLayout(toolbar)

    def _build_right_content(self, parent_layout):
        self.output_tabs = QTabWidget()
        parent_layout.addWidget(self.output_tabs)

        # Empty-state label (shown when no documents generated)
        self.empty_label = QLabel("Drop discovery PDFs and click Generate")
        self.empty_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.empty_label.setStyleSheet("color: #888; font-size: 14px;")
        parent_layout.addWidget(self.empty_label)

        # Start with empty state visible, tabs hidden
        self.output_tabs.hide()
        self.empty_label.show()

    # ------------------------------------------------------------------
    # Signal wiring
    # ------------------------------------------------------------------

    def _connect_signals(self):
        # Document list context menus
        self.discovery_list.customContextMenuRequested.connect(
            self._on_discovery_context_menu
        )
        self.context_list.customContextMenuRequested.connect(
            self._on_context_context_menu
        )

        # LLM provider change
        self.provider_combo.currentTextChanged.connect(self._update_models)

        # Buttons
        self.rules_btn.clicked.connect(self._on_rules_clicked)
        self.generate_btn.clicked.connect(self._on_generate)
        self.save_btn.clicked.connect(self._save_current)
        self.save_all_btn.clicked.connect(self._save_all)
        self.clear_btn.clicked.connect(self._clear_editor)
        self.refresh_17_1_btn.clicked.connect(self._on_refresh_17_1)

        # Kick off initial model fetch
        self._update_models(self.provider_combo.currentText())

    # ------------------------------------------------------------------
    # Document list management
    # ------------------------------------------------------------------

    def _on_discovery_context_menu(self, pos):
        item = self.discovery_list.itemAt(pos)
        if not item:
            return
        menu = QMenu(self)
        remove_action = QAction("Remove", self)
        menu.addAction(remove_action)

        action = menu.exec(self.discovery_list.mapToGlobal(pos))
        if action == remove_action:
            row = self.discovery_list.row(item)
            self.discovery_list.takeItem(row)
            self._save_discovery_files()

    def _on_context_context_menu(self, pos):
        item = self.context_list.itemAt(pos)
        if not item:
            return
        menu = QMenu(self)
        remove_action = QAction("Remove", self)
        menu.addAction(remove_action)

        action = menu.exec(self.context_list.mapToGlobal(pos))
        if action == remove_action:
            row = self.context_list.row(item)
            self.context_list.takeItem(row)
            self._save_context_files()

    def _add_discovery_file(self, path: str):
        """Add a discovery PDF to the list (skip duplicates)."""
        if not path.lower().endswith(_DISCOVERY_EXTENSIONS):
            return
        for i in range(self.discovery_list.count()):
            existing = self.discovery_list.item(i).data(Qt.ItemDataRole.UserRole)
            if existing == path:
                return

        item = QListWidgetItem(os.path.basename(path))
        item.setData(Qt.ItemDataRole.UserRole, path)
        item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
        item.setCheckState(Qt.CheckState.Checked)
        item.setToolTip(path)
        self.discovery_list.addItem(item)
        self._save_discovery_files()

    def _add_context_file(self, path: str):
        """Add a context document to the list (skip duplicates)."""
        if not path.lower().endswith(_CONTEXT_EXTENSIONS):
            return
        for i in range(self.context_list.count()):
            existing = self.context_list.item(i).data(Qt.ItemDataRole.UserRole)
            if existing == path:
                return

        item = QListWidgetItem(os.path.basename(path))
        item.setData(Qt.ItemDataRole.UserRole, path)
        item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
        item.setCheckState(Qt.CheckState.Checked)
        item.setToolTip(path)
        self.context_list.addItem(item)
        self._save_context_files()

    # ------------------------------------------------------------------
    # Drag and drop
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
            if os.path.isfile(path) and path.lower().endswith(
                _SUPPORTED_DROP_EXTENSIONS
            ):
                files_to_add.append(path)

        event.accept()

        if files_to_add:
            QTimer.singleShot(
                0, lambda: self._process_dropped_files(files_to_add)
            )

    def _process_dropped_files(self, file_paths: List[str]):
        """Route dropped files to the appropriate list."""
        for path in file_paths:
            if path.lower().endswith(".pdf"):
                self._add_discovery_file(path)
            else:
                self._add_context_file(path)

    # ------------------------------------------------------------------
    # File reading
    # ------------------------------------------------------------------

    def _read_files_from_list(self, list_widget) -> str:
        """Read text content from checked files in a list widget."""
        texts = []
        for i in range(list_widget.count()):
            item = list_widget.item(i)
            if item.checkState() != Qt.CheckState.Checked:
                continue
            path = item.data(Qt.ItemDataRole.UserRole)
            if not path or not os.path.isfile(path):
                continue
            try:
                if path.lower().endswith(".pdf"):
                    if fitz:
                        doc = fitz.open(path)
                        text = "\n".join(page.get_text() for page in doc)
                        doc.close()
                        texts.append(text)
                    else:
                        texts.append("[PyMuPDF not available]")
                elif path.lower().endswith(".docx"):
                    from docx import Document
                    doc = Document(path)
                    text = "\n".join(p.text for p in doc.paragraphs)
                    texts.append(text)
                elif path.lower().endswith(".txt"):
                    with open(path, "r", encoding="utf-8", errors="replace") as f:
                        texts.append(f.read())
            except Exception as e:
                logger.warning("Failed to read %s: %s", path, e)
        return "\n\n---\n\n".join(texts)

    def read_discovery_content(self) -> str:
        """Extract text from checked discovery PDFs."""
        return self._read_files_from_list(self.discovery_list)

    def read_context_content(self) -> str:
        """Extract text from checked context documents."""
        return self._read_files_from_list(self.context_list)

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
            idx = self.model_combo.findText("gemini-3.1-pro-preview")
            if idx == -1:
                idx = self.model_combo.findText("gemini-3-pro-preview")
            if idx != -1:
                self.model_combo.setCurrentIndex(idx)
            else:
                for i in range(self.model_combo.count()):
                    if "gemini-3" in self.model_combo.itemText(i).lower():
                        self.model_combo.setCurrentIndex(i)
                        break

    # ------------------------------------------------------------------
    # Rules dialog
    # ------------------------------------------------------------------

    def _on_rules_clicked(self):
        dlg = ResponseRulesDialog(self.rules, parent=self)
        if dlg.exec() == QDialog.DialogCode.Accepted:
            self.rules = dlg.get_rules()
            self._save_rules()

    # ------------------------------------------------------------------
    # Generate responses — Phase 1, 2, 3
    # ------------------------------------------------------------------

    def _on_generate(self):
        """Validate inputs, read discovery PDFs, launch Phase 1 parse workers."""
        # 1. Validate: at least one checked discovery PDF
        checked_paths = []
        for i in range(self.discovery_list.count()):
            item = self.discovery_list.item(i)
            if item.checkState() == Qt.CheckState.Checked:
                path = item.data(Qt.ItemDataRole.UserRole)
                if path and os.path.isfile(path):
                    checked_paths.append(
                        (path, os.path.basename(path))
                    )

        if not checked_paths:
            self.status_label.setText(
                "Add and check at least one discovery PDF."
            )
            return

        # 2. Disable generate button, show status
        self.generate_btn.setEnabled(False)
        self.status_label.setText("Parsing discovery documents...")
        self._parsed_discoveries.clear()
        self._rfa_responses_cache = None

        # 3. Read context once
        context_text = self.read_context_content()
        logger.info("Context text length: %d chars (~%dk tokens)",
                    len(context_text), len(context_text) // 4000)

        # 5. For each checked discovery PDF, launch Phase 1
        for path, label in checked_paths:
            try:
                text = self._read_single_file(path)
                if not text.strip():
                    self.status_label.setText(f"Empty file: {label}")
                    continue
                prompt = build_parse_prompt(text)
                self._run_phase1_llm(prompt, label, context_text)
            except Exception as e:
                logger.exception("Failed to read %s", path)
                self.status_label.setText(f"Error reading {label}: {e}")

    def _read_single_file(self, path: str) -> str:
        """Read text from a single file."""
        if path.lower().endswith(".pdf") and fitz:
            doc = fitz.open(path)
            text = "\n".join(page.get_text() for page in doc)
            doc.close()
            return text
        return ""

    def _run_phase1_llm(self, prompt: str, file_label: str, context_text: str):
        """Launch LLMWorker for Phase 1 (parse)."""
        provider = self.provider_combo.currentText()
        model = self.model_combo.currentText()

        worker = LLMWorker(
            provider=provider,
            model=model,
            system="You are a legal document parser. Return only valid JSON.",
            user=_truncate_prompt(prompt),
            files=[],
            settings={"stream": False},
        )

        worker.finished.connect(
            lambda text, fl=file_label, ct=context_text: self._on_parse_finished(
                text, fl, ct
            )
        )
        worker.error.connect(self._on_llm_error)

        self._llm_workers.append(worker)
        worker.start()

    def _on_parse_finished(self, response_text: str, file_label: str,
                           context_text: str):
        """Handle Phase 1 completion: parse the LLM response, run Phase 2+3."""
        try:
            parsed = parse_llm_response(response_text)
            self._parsed_discoveries[file_label] = parsed
            self.status_label.setText(
                f"Parsed {file_label}: {len(parsed.requests)} requests "
                f"({parsed.discovery_type})"
            )
            self._run_phase_2_3(parsed, file_label, context_text)
        except Exception as e:
            logger.exception("Failed to parse discovery %s", file_label)
            self.status_label.setText(f"Parse error ({file_label}): {e}")
            self.generate_btn.setEnabled(True)

    def _run_phase_2_3(self, parsed: ParsedDiscovery, file_label: str,
                       context_text: str):
        """
        Run Phase 2 (objection selection) and Phase 3 (response drafting)
        for a parsed discovery document.

        Uses a simplified approach: builds a single comprehensive prompt per
        discovery type that asks the LLM to draft the complete response.
        For FI, fixed responses are used where applicable.
        """
        disc_type = parsed.discovery_type.upper()

        # Phase 2: Build objections map
        objections_map: Dict[str, str] = {}
        responses_map: Dict[str, str] = {}

        if disc_type == "FI":
            # FI: fixed objections from rules for all requests
            fi_obj_text = select_fi_objections(self.rules)
            for req in parsed.requests:
                objections_map[req.number] = fi_obj_text

                # Phase 3: Fixed responses for certain FI numbers
                fixed = get_fi_fixed_response(req.number, self.rules)
                if fixed is not None:
                    responses_map[req.number] = fixed

            # Find FIs that need LLM-drafted responses
            unfixed = [
                req for req in parsed.requests
                if req.number not in responses_map
                and not detect_inapplicable_fi(req.number)
            ]

            # Handle inapplicable FIs
            for req in parsed.requests:
                if detect_inapplicable_fi(req.number):
                    responses_map[req.number] = (
                        "This interrogatory is not applicable to the present action."
                    )

            if unfixed:
                # Build a combined prompt for all unfixed FIs
                self._draft_remaining_fi(
                    parsed, unfixed, objections_map, responses_map,
                    file_label, context_text
                )
            else:
                # All FI responses are fixed — assemble now
                self._finalize_response(
                    parsed, objections_map, responses_map, file_label
                )
        else:
            # SI, RFA, RPD: rule-based pre-select + LLM for objections and responses
            for req in parsed.requests:
                rule_ids = rule_based_preselect(
                    req, self.rules, self._objection_menu
                )
                # Use first defined term for placeholder substitution
                term = req.defined_terms_used[0] if req.defined_terms_used else None
                obj_text = format_objections(rule_ids, self._objection_menu, term=term)
                objections_map[req.number] = obj_text

            # Build a single comprehensive LLM prompt for all requests
            self._draft_responses_combined(
                parsed, objections_map, responses_map,
                file_label, context_text
            )

    def _draft_remaining_fi(
        self,
        parsed: ParsedDiscovery,
        unfixed: list,
        objections_map: Dict[str, str],
        responses_map: Dict[str, str],
        file_label: str,
        context_text: str,
    ):
        """Launch LLM for FI requests that need drafted responses."""
        # Build a combined prompt for all unfixed FIs
        requests_text = "\n\n".join(
            f"FORM INTERROGATORY NO. {req.number}:\n{req.text}"
            for req in unfixed
        )
        prompt = (
            "You are a California civil litigation defense attorney drafting "
            "responses to Form Interrogatories.\n\n"
            "Draft substantive responses for each of the following Form "
            "Interrogatories. For each one, write ONLY the factual response "
            "(objections are handled separately).\n\n"
            "Format your output as:\n"
            "FI {number}: [your response]\n\n"
            f"INTERROGATORIES:\n{requests_text}\n\n"
        )
        if context_text:
            max_ctx = 500_000
            ctx = context_text[:max_ctx]
            if len(context_text) > max_ctx:
                ctx += "\n\n[Context truncated due to length]"
            prompt += f"CASE CONTEXT:\n{ctx}\n\n"
        if self.rules.custom_instructions:
            prompt += f"ADDITIONAL INSTRUCTIONS:\n{self.rules.custom_instructions}\n"

        provider = self.provider_combo.currentText()
        model = self.model_combo.currentText()

        worker = LLMWorker(
            provider=provider,
            model=model,
            system="You are a California litigation defense attorney.",
            user=_truncate_prompt(prompt),
            files=[],
            settings={"stream": False},
        )

        worker.finished.connect(
            lambda text: self._on_fi_draft_finished(
                text, parsed, unfixed, objections_map, responses_map,
                file_label
            )
        )
        worker.error.connect(self._on_llm_error)

        self._llm_workers.append(worker)
        worker.start()

    def _on_fi_draft_finished(
        self,
        response_text: str,
        parsed: ParsedDiscovery,
        unfixed: list,
        objections_map: Dict[str, str],
        responses_map: Dict[str, str],
        file_label: str,
    ):
        """Parse LLM-drafted FI responses and finalize."""
        import re
        # Parse "FI {number}: response" blocks
        pattern = re.compile(r"FI\s+([\d.]+)\s*:\s*", re.IGNORECASE)
        parts = pattern.split(response_text)
        # parts = [preamble, num1, text1, num2, text2, ...]
        if len(parts) >= 3:
            for i in range(1, len(parts) - 1, 2):
                num = parts[i].strip()
                text = parts[i + 1].strip()
                responses_map[num] = text
        else:
            # Fallback: assign the entire response to the first unfixed FI
            if unfixed:
                responses_map[unfixed[0].number] = response_text.strip()

        self._finalize_response(parsed, objections_map, responses_map, file_label)

    def _draft_responses_combined(
        self,
        parsed: ParsedDiscovery,
        objections_map: Dict[str, str],
        responses_map: Dict[str, str],
        file_label: str,
        context_text: str,
    ):
        """Launch a single LLM call for SI/RFA/RPD response drafting."""
        disc_type = parsed.discovery_type.upper()

        # Truncate context to avoid exceeding token limits.
        # ~4 chars per token; leave room for instructions + requests.
        max_context_chars = 500_000
        truncated_context = context_text[:max_context_chars]
        if len(context_text) > max_context_chars:
            truncated_context += "\n\n[Context truncated due to length]"

        # Build type-specific drafting instructions (included ONCE)
        if disc_type == "SI":
            from icharlotte_core.discovery.response_drafter import _SI_STYLE_INSTRUCTIONS
            style = _SI_STYLE_INSTRUCTIONS.get(
                self.rules.si_response_style, _SI_STYLE_INSTRUCTIONS["moderate"]
            )
            type_instruction = (
                f"DRAFTING INSTRUCTION: {style}\n"
                "Draft substantive responses only — objections are handled separately.\n"
                "Write only the factual, responsive answer for each interrogatory."
            )
        elif disc_type == "RFA":
            from icharlotte_core.discovery.response_drafter import _RFA_POSTURE_INSTRUCTIONS
            posture = _RFA_POSTURE_INSTRUCTIONS.get(
                self.rules.rfa_default_posture, _RFA_POSTURE_INSTRUCTIONS["cautious"]
            )
            type_instruction = (
                "For each Request for Admission, respond with EXACTLY one of:\n"
                '- "Admit"\n'
                '- "Deny"\n'
                '- "After a reasonable inquiry concerning the matter in this request, '
                "the information known or readily obtainable to Responding Party is "
                'insufficient to enable Responding Party to admit the matter."\n\n'
                f"POSTURE: {posture}"
            )
        elif disc_type == "RPD":
            from icharlotte_core.discovery.response_drafter import _RPD_POSTURE_INSTRUCTIONS
            posture = _RPD_POSTURE_INSTRUCTIONS.get(
                self.rules.rpd_default_posture, _RPD_POSTURE_INSTRUCTIONS["context_dependent"]
            )
            type_instruction = (
                "For each Request for Production, respond with EXACTLY one of:\n"
                '- "Responding Party will comply with this request and produce all '
                "non-privileged documents in Responding Party's possession, custody "
                "and control that Responding Party understands to be responsive to this "
                "Request. Responding Party identifies and refers to the documents "
                'produced concurrently herewith."\n'
                '- "Upon a diligent search and reasonable inquiry made in an effort to '
                "locate the item(s) requested, Responding Party is unable to comply "
                "with this request at this time because the documents responsive to "
                "this request, if they exist, are not in the possession, custody or "
                'control of Responding Party."\n\n'
                f"POSTURE: {posture}"
            )
        else:
            type_instruction = "Draft substantive responses."

        # Build combined prompt: instructions + context (ONCE) + requests (listed)
        requests_listing = "\n\n".join(
            f"REQUEST NO. {req.number}:\n{req.text}"
            for req in parsed.requests
        )

        combined = (
            "You are a California civil litigation defense attorney.\n\n"
            f"{type_instruction}\n\n"
            f"Format your output as:\n"
            f"RESPONSE {parsed.requests[0].number if parsed.requests else '1'}: "
            f"[your response]\n"
            f"RESPONSE {parsed.requests[1].number if len(parsed.requests) > 1 else '2'}: "
            f"[your response]\n"
            f"... and so on for all {len(parsed.requests)} requests.\n\n"
            f"CASE CONTEXT:\n{truncated_context}\n\n"
            f"REQUESTS TO RESPOND TO:\n{requests_listing}\n"
        )

        if self.rules.custom_instructions:
            combined += f"\nADDITIONAL INSTRUCTIONS:\n{self.rules.custom_instructions}\n"

        provider = self.provider_combo.currentText()
        model = self.model_combo.currentText()

        logger.info("Combined prompt length: %d chars (~%dk tokens)",
                    len(combined), len(combined) // 4000)

        worker = LLMWorker(
            provider=provider,
            model=model,
            system="You are a California litigation defense attorney.",
            user=_truncate_prompt(combined),
            files=[],
            settings={"stream": False},
        )

        worker.finished.connect(
            lambda text: self._on_combined_draft_finished(
                text, parsed, objections_map, responses_map, file_label,
                disc_type
            )
        )
        worker.error.connect(self._on_llm_error)

        self._llm_workers.append(worker)
        worker.start()

    def _on_combined_draft_finished(
        self,
        response_text: str,
        parsed: ParsedDiscovery,
        objections_map: Dict[str, str],
        responses_map: Dict[str, str],
        file_label: str,
        disc_type: str,
    ):
        """Parse combined LLM response for SI/RFA/RPD and finalize."""
        import re

        # Try to split on "RESPONSE {number}:" or "Response {number}:" patterns
        pattern = re.compile(
            r"(?:RESPONSE|Response)\s+([\d.]+)\s*[:\-]\s*", re.IGNORECASE
        )
        parts = pattern.split(response_text)

        if len(parts) >= 3:
            for i in range(1, len(parts) - 1, 2):
                num = parts[i].strip()
                text = parts[i + 1].strip()
                responses_map[num] = text
        else:
            # Fallback: try splitting by request numbers
            for i, req in enumerate(parsed.requests):
                # Simple fallback: just assign chunks
                responses_map[req.number] = response_text.strip()

        self._finalize_response(parsed, objections_map, responses_map, file_label)

    def _finalize_response(
        self,
        parsed: ParsedDiscovery,
        objections_map: Dict[str, str],
        responses_map: Dict[str, str],
        file_label: str,
    ):
        """Assemble plain text and display the result."""
        try:
            plain_text = assemble_plain_text(
                parsed, self.rules, objections_map, responses_map
            )

            # Cache RFA responses for FI 17.1
            if parsed.discovery_type.upper() == "RFA":
                self._rfa_responses_cache = plain_text

            self._display_result(file_label, plain_text)
            self.generate_btn.setEnabled(True)
            self.status_label.setText(
                f"Generated responses for {file_label} "
                f"({len(parsed.requests)} requests)"
            )
            self._update_refresh_17_1_state()
        except Exception as e:
            logger.exception("Failed to assemble response for %s", file_label)
            self.status_label.setText(f"Assembly error ({file_label}): {e}")
            self.generate_btn.setEnabled(True)

    def _on_llm_error(self, error_msg: str):
        """Handle LLM error."""
        self.generate_btn.setEnabled(True)
        self.status_label.setText(f"LLM Error: {error_msg}")
        logger.error("LLM error during response generation: %s", error_msg)

    # ------------------------------------------------------------------
    # FI 17.1 refresh
    # ------------------------------------------------------------------

    def _update_refresh_17_1_state(self):
        """Enable the Refresh 17.1 button only when both FI and RFA tabs exist."""
        has_fi = False
        has_rfa = False
        for i in range(self.output_tabs.count()):
            label = self.output_tabs.tabText(i).upper()
            if "FI" in label or "FORM" in label:
                has_fi = True
            if "RFA" in label or "ADMISSION" in label:
                has_rfa = True
        self.refresh_17_1_btn.setEnabled(has_fi and has_rfa)

    def _on_refresh_17_1(self):
        """Read RFA tab text, launch LLM for FI 17.1 response."""
        # Find the RFA tab text
        rfa_text = None
        for i in range(self.output_tabs.count()):
            label = self.output_tabs.tabText(i).upper()
            if "RFA" in label or "ADMISSION" in label:
                editor = self.output_tabs.widget(i)
                if isinstance(editor, QPlainTextEdit):
                    rfa_text = editor.toPlainText()
                break

        if not rfa_text:
            self.status_label.setText("No RFA responses found to use for 17.1.")
            return

        self.refresh_17_1_btn.setEnabled(False)
        self.status_label.setText("Generating FI 17.1 response...")

        prompt = build_fi_17_1_prompt(rfa_text, self.rules)
        provider = self.provider_combo.currentText()
        model = self.model_combo.currentText()

        worker = LLMWorker(
            provider=provider,
            model=model,
            system="You are a California litigation defense attorney.",
            user=_truncate_prompt(prompt),
            files=[],
            settings={"stream": False},
        )

        worker.finished.connect(self._on_17_1_finished)
        worker.error.connect(self._on_llm_error)

        self._llm_workers.append(worker)
        worker.start()

    def _on_17_1_finished(self, response_text: str):
        """Update the FI 17.1 response in the FI tab."""
        # Find the FI tab and update 17.1 in it
        for i in range(self.output_tabs.count()):
            label = self.output_tabs.tabText(i).upper()
            if "FI" in label or "FORM" in label:
                editor = self.output_tabs.widget(i)
                if isinstance(editor, QPlainTextEdit):
                    current = editor.toPlainText()
                    # Replace the placeholder
                    updated = current.replace(
                        "[PENDING \u2014 complete after RFA responses are finalized]",
                        response_text.strip(),
                    )
                    editor.setPlainText(updated)
                break

        self.refresh_17_1_btn.setEnabled(True)
        self.status_label.setText("FI 17.1 response updated.")
        self._output_save_timer.start()

    # ------------------------------------------------------------------
    # Display results
    # ------------------------------------------------------------------

    def _display_result(self, label: str, text: str):
        """Add a tab with QPlainTextEdit containing text."""
        self.empty_label.hide()
        self.output_tabs.show()

        editor = QPlainTextEdit()
        editor.setPlainText(text)
        editor.setLineWrapMode(QPlainTextEdit.LineWrapMode.WidgetWidth)
        editor.textChanged.connect(self._output_save_timer.start)
        self.output_tabs.addTab(editor, label)

        # Switch to the new tab
        self.output_tabs.setCurrentIndex(self.output_tabs.count() - 1)
        self._save_output()

    def _clear_editor(self):
        """Clear the right-pane editor, resetting to empty state."""
        self._parsed_discoveries.clear()
        self._rfa_responses_cache = None
        self.output_tabs.clear()
        self.output_tabs.hide()
        self.empty_label.show()
        self.status_label.setText("")
        self.refresh_17_1_btn.setEnabled(False)
        self._save_output()

    # ------------------------------------------------------------------
    # Save to .docx
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

    def _save_current(self):
        """Save the currently visible tab as a .docx."""
        idx = self.output_tabs.currentIndex()
        if idx < 0:
            self.status_label.setText("No document to save.")
            return

        label = self.output_tabs.tabText(idx)
        editor = self.output_tabs.widget(idx)
        if not isinstance(editor, QPlainTextEdit):
            return

        plain_text = editor.toPlainText()
        parsed = self._parsed_discoveries.get(label)

        if not parsed:
            self.status_label.setText("No parsed discovery data for this tab.")
            return

        result = self._save_response_doc(parsed, plain_text, label)
        if result:
            self.status_label.setText(f"Saved: {os.path.basename(result)}")

    def _save_all(self):
        """Save all tabs as .docx files."""
        if self.output_tabs.count() == 0:
            self.status_label.setText("No documents to save.")
            return

        saved = []
        for i in range(self.output_tabs.count()):
            label = self.output_tabs.tabText(i)
            editor = self.output_tabs.widget(i)
            if not isinstance(editor, QPlainTextEdit):
                continue

            plain_text = editor.toPlainText()
            parsed = self._parsed_discoveries.get(label)
            if not parsed:
                continue

            result = self._save_response_doc(parsed, plain_text, label)
            if result:
                saved.append(result)

        if saved:
            self.status_label.setText(
                f"Saved {len(saved)} file{'s' if len(saved) != 1 else ''}: "
                f"{', '.join(os.path.basename(s) for s in saved)}"
            )

    def _save_response_doc(
        self, parsed: ParsedDiscovery, plain_text: str, label: str
    ) -> Optional[str]:
        """Save a single response to a .docx file. Returns the output path or None."""
        cp = self.case_path
        if not cp or not os.path.isdir(cp):
            self.status_label.setText("Could not resolve case folder.")
            return None

        caption_path = DiscoveryAssembler.find_caption_page(cp)
        if not caption_path:
            self.status_label.setText("No caption page found in case folder.")
            return None

        # Build filename — derive abbreviation from responding party name
        responding_name = parsed.responding_party
        # Extract a short abbreviation: use the first distinctive word
        words = responding_name.replace(",", "").split()
        skip = {"defendant", "plaintiff", "cross-defendant", "cross-complainant"}
        abbreviation = next(
            (w for w in words if w.lower() not in skip), "Def"
        )
        filename = build_response_filename(
            abbreviation=abbreviation,
            disc_type=parsed.discovery_type,
            set_number=parsed.set_number,
        )

        output_dir = os.path.join(cp, "NOTES", "AI OUTPUT", "DISCOVERY RESPONSES")
        output_path = os.path.join(output_dir, filename)

        try:
            assembler = ResponseAssembler(caption_path)
            assembler.assemble(
                parsed=parsed,
                response_text=plain_text,
                rules=self.rules,
                output_path=output_path,
            )
            logger.info("Saved response document: %s", output_path)
            return output_path
        except Exception as e:
            logger.exception("Failed to save response document")
            self.status_label.setText(f"Save error: {e}")
            return None

    # ------------------------------------------------------------------
    # Case loading
    # ------------------------------------------------------------------

    def load_case(self, file_number: str):
        """Load case data -- called when the active case changes."""
        # Save current case state before switching
        if self.file_number and self.file_number != file_number:
            self._save_state()

        # Clear all UI state
        self._clear_state()

        self.file_number = file_number
        self._load_state()

    def _clear_state(self):
        """Clear all UI state for a case switch."""
        self.discovery_list.clear()
        self.context_list.clear()
        self._parsed_discoveries.clear()
        self._rfa_responses_cache = None
        self.output_tabs.clear()
        self.output_tabs.hide()
        self.empty_label.show()
        self.status_label.setText("")
        self.refresh_17_1_btn.setEnabled(False)

    # ------------------------------------------------------------------
    # Session persistence (per-case)
    # ------------------------------------------------------------------

    def _get_manager(self):
        """Return a CaseDataManager instance, or None on error."""
        try:
            from Scripts.case_data_manager import CaseDataManager as CDM
            return CDM()
        except Exception as e:
            logger.warning("CaseDataManager unavailable: %s", e)
            return None

    def _save_state(self):
        """Persist all state for the current case."""
        self._save_discovery_files()
        self._save_context_files()
        self._save_output()
        self._save_rules()

    def _load_state(self):
        """Restore all state for the current case."""
        self._load_discovery_files()
        self._load_context_files()
        self._load_output()
        self._load_rules()

    def _save_discovery_files(self):
        """Persist the discovery file list to CaseDataManager."""
        if not self.file_number:
            return
        manager = self._get_manager()
        if not manager:
            return
        docs = []
        for i in range(self.discovery_list.count()):
            item = self.discovery_list.item(i)
            path = item.data(Qt.ItemDataRole.UserRole)
            checked = item.checkState() == Qt.CheckState.Checked
            if path:
                docs.append({"path": path, "checked": checked})
        try:
            manager.save_variable(
                self.file_number, "respond_discovery_files", docs,
                source="respond_tab", auto_tag=False,
            )
        except Exception as e:
            logger.warning("Failed to save discovery files: %s", e)

    def _load_discovery_files(self):
        """Restore the discovery file list from CaseDataManager."""
        if not self.file_number:
            return
        manager = self._get_manager()
        if not manager:
            return
        try:
            docs = manager.get_value(self.file_number, "respond_discovery_files")
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
                self.discovery_list.addItem(item)
        except Exception as e:
            logger.warning("Failed to load discovery files: %s", e)

    def _save_context_files(self):
        """Persist the context file list to CaseDataManager."""
        if not self.file_number:
            return
        manager = self._get_manager()
        if not manager:
            return
        docs = []
        for i in range(self.context_list.count()):
            item = self.context_list.item(i)
            path = item.data(Qt.ItemDataRole.UserRole)
            checked = item.checkState() == Qt.CheckState.Checked
            if path:
                docs.append({"path": path, "checked": checked})
        try:
            manager.save_variable(
                self.file_number, "respond_context_files", docs,
                source="respond_tab", auto_tag=False,
            )
        except Exception as e:
            logger.warning("Failed to save context files: %s", e)

    def _load_context_files(self):
        """Restore the context file list from CaseDataManager."""
        if not self.file_number:
            return
        manager = self._get_manager()
        if not manager:
            return
        try:
            docs = manager.get_value(self.file_number, "respond_context_files")
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
                self.context_list.addItem(item)
        except Exception as e:
            logger.warning("Failed to load context files: %s", e)

    @Slot()
    def _save_output(self):
        """Persist the editor output tabs to CaseDataManager."""
        if not self.file_number:
            return
        manager = self._get_manager()
        if not manager:
            return
        tabs_data = []
        for i in range(self.output_tabs.count()):
            editor = self.output_tabs.widget(i)
            label = self.output_tabs.tabText(i)
            text = ""
            if isinstance(editor, QPlainTextEdit):
                text = editor.toPlainText()
            tabs_data.append({"label": label, "text": text})
        try:
            manager.save_variable(
                self.file_number, "respond_output", tabs_data,
                source="respond_tab", auto_tag=False,
            )
        except Exception as e:
            logger.warning("Failed to save response output: %s", e)

    def _load_output(self):
        """Restore the editor output tabs from CaseDataManager."""
        if not self.file_number:
            return
        manager = self._get_manager()
        if not manager:
            return
        try:
            tabs_data = manager.get_value(self.file_number, "respond_output")
            if not tabs_data or not isinstance(tabs_data, list):
                return
            has_content = any(
                t.get("text", "").strip() for t in tabs_data if isinstance(t, dict)
            )
            if not has_content:
                return
            self.empty_label.hide()
            self.output_tabs.show()
            for entry in tabs_data:
                if not isinstance(entry, dict):
                    continue
                label = entry.get("label", "")
                text = entry.get("text", "")
                editor = QPlainTextEdit()
                editor.setPlainText(text)
                editor.setLineWrapMode(QPlainTextEdit.LineWrapMode.WidgetWidth)
                editor.textChanged.connect(self._output_save_timer.start)
                self.output_tabs.addTab(editor, label)
            self._update_refresh_17_1_state()
        except Exception as e:
            logger.warning("Failed to load response output: %s", e)

    def _save_rules(self):
        """Persist response rules to CaseDataManager."""
        if not self.file_number:
            return
        manager = self._get_manager()
        if not manager:
            return
        try:
            from dataclasses import asdict
            manager.save_variable(
                self.file_number, "respond_rules", asdict(self.rules),
                source="respond_tab", auto_tag=False,
            )
        except Exception as e:
            logger.warning("Failed to save response rules: %s", e)

    def _load_rules(self):
        """Restore response rules from CaseDataManager."""
        if not self.file_number:
            return
        manager = self._get_manager()
        if not manager:
            return
        try:
            rules_data = manager.get_value(self.file_number, "respond_rules")
            if rules_data and isinstance(rules_data, dict):
                self.rules = ResponseRules(**{
                    k: v for k, v in rules_data.items()
                    if k in {f.name for f in __import__('dataclasses').fields(ResponseRules)}
                })
        except Exception as e:
            logger.warning("Failed to load response rules: %s", e)
