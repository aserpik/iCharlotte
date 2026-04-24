"""
QuoteInsertionDialog — search deposition transcripts and select Q&A passages
to insert into a mediation brief section.

Workflow:
  1. User adds transcript PDFs/DOCXs via "Add Transcript(s)" button.
  2. User types a description of the testimony they want to find.
  3. User picks the target section (and optional subsection) and insert mode.
  4. "Search" button fires QuoteSearchWorker in a background thread.
  5. Results populate as QuoteResultWidget cards (checked by default).
  6. "Insert Selected" collects checked quotes and emits quotes_to_insert signal.
"""

import logging
import os
import re
from typing import Dict, List, Optional

from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QPlainTextEdit, QComboBox, QRadioButton, QButtonGroup,
    QListWidget, QListWidgetItem, QScrollArea, QWidget,
    QCheckBox, QFileDialog, QGroupBox, QProgressBar,
    QMessageBox, QFrame,
)
from PySide6.QtCore import Qt, Signal, QThread

from ..mediation_brief import MediationBriefGenerator, SECTION_ORDER, SECTION_HEADINGS

logger = logging.getLogger(__name__)

# Sections that make sense for quote insertion (narrative sections only)
_QUOTE_SECTIONS = [
    s for s in SECTION_ORDER
    if s not in ("introduction", "procedural_status", "settlement_position", "conclusion")
]


class QuoteSearchWorker(QThread):
    """Background thread that calls generator.search_quotes().

    ``brief_sections`` is optional context describing the current text of
    the mediation brief — passed through to ``search_quotes`` so the LLM
    can orient on which arguments actually appear in each section.
    Callers without brief context (or outside a brief workflow) should
    pass ``None``.
    """

    results_ready = Signal(list)
    error = Signal(str)

    def __init__(self, generator: MediationBriefGenerator,
                 transcript_paths: List[str], description: str,
                 brief_sections: Optional[Dict[str, str]] = None,
                 parent=None):
        super().__init__(parent)
        self._generator = generator
        self._transcript_paths = transcript_paths
        self._description = description
        self._brief_sections = brief_sections

    def run(self):
        try:
            results = self._generator.search_quotes(
                self._transcript_paths,
                self._description,
                brief_sections=self._brief_sections,
            )
            self.results_ready.emit(results)
        except Exception as exc:
            logger.exception("QuoteSearchWorker failed")
            self.error.emit(str(exc))


class QuoteResultWidget(QFrame):
    """Card widget displaying a single quote result with edit support."""

    def __init__(self, quote_data: Dict, parent=None):
        super().__init__(parent)
        self.quote_data = dict(quote_data)  # local copy — edits land here

        self.setFrameShape(QFrame.Shape.StyledPanel)
        self.setFrameShadow(QFrame.Shadow.Raised)
        self.setStyleSheet(
            "QuoteResultWidget { border: 1px solid #ccc; border-radius: 4px;"
            " background: #fafafa; margin: 2px; }"
        )

        root = QVBoxLayout(self)
        root.setContentsMargins(8, 6, 8, 6)
        root.setSpacing(4)

        # ── Row 1: checkbox + header label ──────────────────────────────────
        header_row = QHBoxLayout()

        self.checkbox = QCheckBox()
        self.checkbox.setChecked(True)
        header_row.addWidget(self.checkbox)

        deponent = quote_data.get("deponent", "Unknown")
        source   = quote_data.get("source", "")
        page_line = quote_data.get("page_line", "")
        header_text = deponent
        if source:
            header_text += f"  —  {source}"
        if page_line:
            header_text += f"  (p. {page_line})"

        header_label = QLabel(header_text)
        header_label.setStyleSheet("font-weight: bold; color: #000;")
        header_row.addWidget(header_label, 1)

        self.edit_btn = QPushButton("Edit")
        self.edit_btn.setFixedWidth(55)
        self.edit_btn.setCheckable(True)
        self.edit_btn.clicked.connect(self._toggle_edit)
        header_row.addWidget(self.edit_btn)

        root.addLayout(header_row)

        # ── Row 2: relevance note ────────────────────────────────────────────
        relevance = quote_data.get("relevance", "")
        if relevance:
            rel_label = QLabel(relevance)
            rel_label.setStyleSheet("font-style: italic; color: #222;")
            rel_label.setWordWrap(True)
            root.addWidget(rel_label)

        # ── Row 3: Q&A display (label) ───────────────────────────────────────
        qa_text = quote_data.get("qa_text", "")

        self.qa_display = QLabel(qa_text)
        self.qa_display.setWordWrap(True)
        self.qa_display.setTextInteractionFlags(
            Qt.TextInteractionFlag.TextSelectableByMouse
        )
        self.qa_display.setStyleSheet(
            "background: #f0f0f0; color: #000; padding: 6px;"
            " border-radius: 3px; font-family: monospace; font-size: 12px;"
        )
        root.addWidget(self.qa_display)

        # ── Row 4: Q&A editor (hidden by default) ───────────────────────────
        self.qa_editor = QPlainTextEdit()
        self.qa_editor.setPlainText(qa_text)
        self.qa_editor.setMinimumHeight(80)
        self.qa_editor.setVisible(False)
        self.qa_editor.setStyleSheet(
            "background: #fff9e6; color: #000; font-family: monospace;"
            " font-size: 12px;"
        )
        root.addWidget(self.qa_editor)

    # ── helpers ─────────────────────────────────────────────────────────────

    def _toggle_edit(self, checked: bool):
        if checked:
            # Switch to editor
            self.qa_editor.setPlainText(self.qa_display.text())
            self.qa_display.setVisible(False)
            self.qa_editor.setVisible(True)
            self.edit_btn.setText("Done")
        else:
            # Save and switch back to display
            edited = self.qa_editor.toPlainText()
            self.quote_data["qa_text"] = edited
            self.qa_display.setText(edited)
            self.qa_editor.setVisible(False)
            self.qa_display.setVisible(True)
            self.edit_btn.setText("Edit")

    def is_selected(self) -> bool:
        return self.checkbox.isChecked()

    def get_quote_data(self) -> Dict:
        # If currently in edit mode, flush the editor first
        if self.edit_btn.isChecked():
            self.quote_data["qa_text"] = self.qa_editor.toPlainText()
        return dict(self.quote_data)


class QuoteInsertionDialog(QDialog):
    """
    Modal dialog for searching deposition transcripts and selecting quotes to
    insert into a mediation brief section.

    Emits ``quotes_to_insert(quotes, section_name, subsection_or_empty, mode)``
    when the user clicks "Insert Selected".
      - quotes: list of quote dicts (deponent, source, page_line, relevance, qa_text)
      - section_name: canonical section key (e.g. "statement_of_facts")
      - subsection_or_empty: subsection title string, or "" for "Auto"
      - mode: "quick" or "weave"
    """

    quotes_to_insert = Signal(list, str, str, str)

    def __init__(self, generator: MediationBriefGenerator, parent=None):
        super().__init__(parent)
        self._generator = generator
        self._worker: Optional[QuoteSearchWorker] = None
        self._result_widgets: List[QuoteResultWidget] = []

        self.setWindowTitle("Insert Deposition Quotes")
        self.setMinimumWidth(640)
        self.setMinimumHeight(700)

        root = QVBoxLayout(self)
        root.setSpacing(8)

        # ── 1. Transcript upload ─────────────────────────────────────────────
        transcript_group = QGroupBox("Transcripts")
        tg_layout = QVBoxLayout(transcript_group)

        self.transcript_list = QListWidget()
        self.transcript_list.setMaximumHeight(100)
        self.transcript_list.setSelectionMode(
            QListWidget.SelectionMode.ExtendedSelection
        )
        tg_layout.addWidget(self.transcript_list)

        t_btn_row = QHBoxLayout()
        add_transcript_btn = QPushButton("Add Transcript(s)")
        add_transcript_btn.clicked.connect(self._add_transcripts)
        t_btn_row.addWidget(add_transcript_btn)

        remove_transcript_btn = QPushButton("Remove Selected")
        remove_transcript_btn.clicked.connect(self._remove_selected_transcripts)
        t_btn_row.addWidget(remove_transcript_btn)
        t_btn_row.addStretch()
        tg_layout.addLayout(t_btn_row)

        root.addWidget(transcript_group)

        # ── 2. Search description ────────────────────────────────────────────
        desc_group = QGroupBox("Search Description")
        desc_layout = QVBoxLayout(desc_group)

        self.desc_edit = QPlainTextEdit()
        self.desc_edit.setPlaceholderText(
            "Describe the testimony you are looking for, e.g.:\n"
            "'plaintiff admits she did not seek medical treatment for six months'"
        )
        self.desc_edit.setMaximumHeight(80)
        self.desc_edit.textChanged.connect(self._update_search_button)
        desc_layout.addWidget(self.desc_edit)

        root.addWidget(desc_group)

        # ── 3. Placement controls ────────────────────────────────────────────
        placement_group = QGroupBox("Placement")
        pg_layout = QHBoxLayout(placement_group)

        pg_layout.addWidget(QLabel("Section:"))
        self.section_combo = QComboBox()
        for s in _QUOTE_SECTIONS:
            _, heading = SECTION_HEADINGS[s]
            self.section_combo.addItem(heading, s)
        self.section_combo.currentIndexChanged.connect(self._on_section_changed)
        pg_layout.addWidget(self.section_combo)

        pg_layout.addSpacing(12)
        pg_layout.addWidget(QLabel("Subsection:"))
        self.subsection_combo = QComboBox()
        self.subsection_combo.addItem("Auto (LLM decides)", "")
        pg_layout.addWidget(self.subsection_combo)

        pg_layout.addSpacing(16)
        self._mode_group = QButtonGroup(self)
        self.quick_radio = QRadioButton("Quick Insert")
        self.weave_radio = QRadioButton("Weave In")
        self.quick_radio.setChecked(True)
        self._mode_group.addButton(self.quick_radio)
        self._mode_group.addButton(self.weave_radio)
        pg_layout.addWidget(self.quick_radio)
        pg_layout.addWidget(self.weave_radio)
        pg_layout.addStretch()

        root.addWidget(placement_group)

        # ── 4. Search button + progress ──────────────────────────────────────
        search_row = QHBoxLayout()

        self.search_btn = QPushButton("Search")
        self.search_btn.setEnabled(False)
        self.search_btn.setStyleSheet(
            "QPushButton { background-color: #2196F3; color: white;"
            " padding: 6px 18px; border-radius: 3px; }"
            "QPushButton:disabled { background-color: #90CAF9; }"
        )
        self.search_btn.clicked.connect(self._start_search)
        search_row.addWidget(self.search_btn)

        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 0)   # indeterminate
        self.progress_bar.setVisible(False)
        self.progress_bar.setMaximumHeight(18)
        search_row.addWidget(self.progress_bar, 1)

        root.addLayout(search_row)

        # ── 5. Results panel ─────────────────────────────────────────────────
        results_group = QGroupBox("Results")
        rg_layout = QVBoxLayout(results_group)

        self.status_label = QLabel("Add transcripts and enter a search description to begin.")
        self.status_label.setStyleSheet("color: #666; font-style: italic;")
        rg_layout.addWidget(self.status_label)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)

        self._results_container = QWidget()
        self._results_layout = QVBoxLayout(self._results_container)
        self._results_layout.setAlignment(Qt.AlignmentFlag.AlignTop)
        self._results_layout.setSpacing(6)
        scroll.setWidget(self._results_container)

        rg_layout.addWidget(scroll, 1)
        root.addWidget(results_group, 1)

        # ── 6. Bottom buttons ────────────────────────────────────────────────
        btn_row = QHBoxLayout()
        btn_row.addStretch()

        self.insert_btn = QPushButton("Insert Selected")
        self.insert_btn.setEnabled(False)
        self.insert_btn.setStyleSheet(
            "QPushButton { background-color: #4CAF50; color: white;"
            " padding: 6px 18px; border-radius: 3px; }"
            "QPushButton:disabled { background-color: #A5D6A7; }"
        )
        self.insert_btn.clicked.connect(self._insert_selected)
        btn_row.addWidget(self.insert_btn)

        cancel_btn = QPushButton("Cancel")
        cancel_btn.clicked.connect(self.reject)
        btn_row.addWidget(cancel_btn)

        root.addLayout(btn_row)

        # Populate subsections for default section
        self._on_section_changed(0)

    # ── Transcript management ────────────────────────────────────────────────

    def _add_transcripts(self):
        paths, _ = QFileDialog.getOpenFileNames(
            self,
            "Select Deposition Transcripts",
            "",
            "Transcripts (*.pdf *.docx);;PDF Files (*.pdf);;Word Documents (*.docx)",
        )
        for path in paths:
            # Avoid duplicates
            existing = [
                self.transcript_list.item(i).data(Qt.ItemDataRole.UserRole)
                for i in range(self.transcript_list.count())
            ]
            if path not in existing:
                item = QListWidgetItem(os.path.basename(path))
                item.setData(Qt.ItemDataRole.UserRole, path)
                self.transcript_list.addItem(item)
        self._update_search_button()

    def _remove_selected_transcripts(self):
        for item in self.transcript_list.selectedItems():
            self.transcript_list.takeItem(self.transcript_list.row(item))
        self._update_search_button()

    def _transcript_paths(self) -> List[str]:
        return [
            self.transcript_list.item(i).data(Qt.ItemDataRole.UserRole)
            for i in range(self.transcript_list.count())
        ]

    # ── Placement: subsection population ────────────────────────────────────

    def _on_section_changed(self, _index: int):
        section_name = self.section_combo.currentData()
        self.subsection_combo.clear()
        self.subsection_combo.addItem("Auto (LLM decides)", "")
        if section_name and hasattr(self._generator, "sections"):
            section_text = self._generator.sections.get(section_name, "")
            subsections = self._parse_subsections(section_text)
            for title in subsections:
                self.subsection_combo.addItem(title, title)

    @staticmethod
    def _parse_subsections(section_text: str) -> List[str]:
        """Extract SUBSECTION: marker titles from section text."""
        return re.findall(r"^SUBSECTION:\s*(.+)$", section_text, re.MULTILINE)

    # ── Search button state ──────────────────────────────────────────────────

    def _update_search_button(self):
        has_transcripts = self.transcript_list.count() > 0
        has_description = bool(self.desc_edit.toPlainText().strip())
        self.search_btn.setEnabled(has_transcripts and has_description)

    # ── Search execution ─────────────────────────────────────────────────────

    def _start_search(self):
        paths = self._transcript_paths()
        description = self.desc_edit.toPlainText().strip()

        if not paths:
            QMessageBox.warning(self, "Missing Transcripts",
                                "Please add at least one transcript file.")
            return
        if not description:
            QMessageBox.warning(self, "Missing Description",
                                "Please enter a search description.")
            return

        # Clear previous results
        self._clear_results()
        self.status_label.setText("Searching…")
        self.status_label.setVisible(True)
        self.progress_bar.setVisible(True)
        self.search_btn.setEnabled(False)
        self.insert_btn.setEnabled(False)

        # Pass the generator's current in-memory sections so the LLM sees
        # the actual brief text as orientation context.
        brief_sections = dict(getattr(self._generator, "sections", {}) or {})
        self._worker = QuoteSearchWorker(
            self._generator, paths, description,
            brief_sections=brief_sections, parent=self,
        )
        self._worker.results_ready.connect(self._on_results_ready)
        self._worker.error.connect(self._on_search_error)
        self._worker.finished.connect(self._on_worker_finished)
        self._worker.start()

    def _on_results_ready(self, results: list):
        self._clear_results()
        if not results:
            self.status_label.setText(
                "No matching testimony found. Try broadening your description."
            )
            return
        self.status_label.setVisible(False)
        for quote_data in results:
            card = QuoteResultWidget(quote_data, self._results_container)
            card.checkbox.stateChanged.connect(self._update_insert_button)
            self._results_layout.addWidget(card)
            self._result_widgets.append(card)
        self._update_insert_button()

    def _on_search_error(self, msg: str):
        self.status_label.setText(f"Search failed: {msg}")
        logger.error("Quote search error: %s", msg)

    def _on_worker_finished(self):
        self.progress_bar.setVisible(False)
        self._update_search_button()
        self._worker = None

    def _clear_results(self):
        for w in self._result_widgets:
            self._results_layout.removeWidget(w)
            w.deleteLater()
        self._result_widgets.clear()

    # ── Insert ───────────────────────────────────────────────────────────────

    def _update_insert_button(self):
        any_checked = any(w.is_selected() for w in self._result_widgets)
        self.insert_btn.setEnabled(bool(self._result_widgets) and any_checked)

    def _insert_selected(self):
        selected = [w.get_quote_data() for w in self._result_widgets if w.is_selected()]
        if not selected:
            QMessageBox.information(self, "Nothing Selected",
                                    "Please check at least one quote to insert.")
            return

        section_name = self.section_combo.currentData() or ""
        subsection   = self.subsection_combo.currentData() or ""
        mode         = "quick" if self.quick_radio.isChecked() else "weave"

        self.quotes_to_insert.emit(selected, section_name, subsection, mode)
        self.accept()
