"""WordQuoteInsertionDialog — Word-cursor variant of QuoteInsertionDialog.

Used by the Word AI assistant popup's "Mediation Brief: Add Quotes" flow.
Provides the same transcript upload + search description + result selection
UI as :class:`QuoteInsertionDialog`, but without the section/subsection
combos or the Quick/Weave mode toggle — quotes are inserted at the Word
cursor, not into a named section.

Reuses :class:`QuoteSearchWorker` and :class:`QuoteResultWidget` from the
existing dialog module so search behaviour and per-result rendering stay in
a single place.
"""

from __future__ import annotations

import os
from typing import Dict, List, Mapping, Optional

from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QDragEnterEvent, QDropEvent
from PySide6.QtWidgets import (
    QDialog,
    QFileDialog,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QListWidget,
    QListWidgetItem,
    QPlainTextEdit,
    QProgressBar,
    QPushButton,
    QScrollArea,
    QVBoxLayout,
    QWidget,
)

from icharlotte_core.mediation_brief import MediationBriefGenerator
from icharlotte_core.ui.quote_dialog import QuoteResultWidget, QuoteSearchWorker


class WordQuoteInsertionDialog(QDialog):
    """Modal dialog for searching deposition transcripts and picking quotes to
    insert at the current Word cursor.

    Emits ``quotes_to_insert(quotes: List[Dict])`` when the user clicks
    "Insert Selected".  The caller is responsible for splicing the quotes
    into the live Word document.
    """

    quotes_to_insert = Signal(list)

    def __init__(
        self,
        parent=None,
        brief_sections: Optional[Mapping[str, str]] = None,
    ):
        """Construct the dialog.

        Args:
            parent: Qt parent widget.
            brief_sections: Optional mapping of canonical section name →
                current body text, parsed from the live Word document by
                the caller via
                :func:`mediation_brief_live.parse_brief_from_word_doc`.
                When provided, the search LLM receives the brief's current
                text as orientation context so it can pick quotes that
                support the actual defense arguments. Pass ``None`` when
                the active document is not a recognised mediation brief.
        """
        super().__init__(parent)
        self._worker: Optional[QuoteSearchWorker] = None
        self._result_widgets: List[QuoteResultWidget] = []
        self._brief_sections: Optional[Dict[str, str]] = (
            dict(brief_sections) if brief_sections else None
        )

        self.setWindowTitle("Insert Deposition Quotes")
        self.setMinimumWidth(640)
        self.setMinimumHeight(640)
        self.setAcceptDrops(True)

        root = QVBoxLayout(self)
        root.setSpacing(8)

        # ── 1. Transcript upload ─────────────────────────────────────────
        transcript_group = QGroupBox("Transcripts")
        tg_layout = QVBoxLayout(transcript_group)

        self.transcript_list = QListWidget()
        self.transcript_list.setMaximumHeight(100)
        self.transcript_list.setSelectionMode(
            QListWidget.SelectionMode.ExtendedSelection
        )
        tg_layout.addWidget(self.transcript_list)

        t_btn_row = QHBoxLayout()
        add_btn = QPushButton("Add Transcript(s)")
        add_btn.clicked.connect(self._add_transcripts)
        t_btn_row.addWidget(add_btn)

        remove_btn = QPushButton("Remove Selected")
        remove_btn.clicked.connect(self._remove_selected_transcripts)
        t_btn_row.addWidget(remove_btn)
        t_btn_row.addStretch()
        tg_layout.addLayout(t_btn_row)

        root.addWidget(transcript_group)

        # ── 2. Search description ────────────────────────────────────────
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

        # ── 3. Search button + progress ──────────────────────────────────
        search_row = QHBoxLayout()
        self.search_btn = QPushButton("Search")
        self.search_btn.setEnabled(False)
        self.search_btn.clicked.connect(self._start_search)
        search_row.addWidget(self.search_btn)

        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 0)  # indeterminate
        self.progress_bar.setVisible(False)
        self.progress_bar.setMaximumHeight(18)
        search_row.addWidget(self.progress_bar, 1)
        root.addLayout(search_row)

        # ── 4. Results panel ─────────────────────────────────────────────
        results_group = QGroupBox("Results")
        rg_layout = QVBoxLayout(results_group)

        self.status_label = QLabel(
            "Add transcripts and enter a search description to begin."
        )
        self.status_label.setStyleSheet("color: #666; font-style: italic;")
        rg_layout.addWidget(self.status_label)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self._results_container = QWidget()
        self._results_layout = QVBoxLayout(self._results_container)
        self._results_layout.setAlignment(Qt.AlignmentFlag.AlignTop)
        scroll.setWidget(self._results_container)
        rg_layout.addWidget(scroll, 1)

        root.addWidget(results_group, 1)

        # ── 5. Action buttons ────────────────────────────────────────────
        btn_row = QHBoxLayout()
        btn_row.addStretch()
        cancel_btn = QPushButton("Cancel")
        cancel_btn.clicked.connect(self.reject)
        btn_row.addWidget(cancel_btn)

        self.insert_btn = QPushButton("Insert Selected")
        self.insert_btn.setEnabled(False)
        self.insert_btn.clicked.connect(self._on_insert_clicked)
        btn_row.addWidget(self.insert_btn)

        root.addLayout(btn_row)

    # ------------------------------------------------------------------
    # Transcript management
    # ------------------------------------------------------------------

    def _add_transcripts(self):
        paths, _ = QFileDialog.getOpenFileNames(
            self,
            "Select deposition transcript(s)",
            "",
            "Transcripts (*.pdf *.docx);;All Files (*)",
        )
        self._append_transcript_paths(paths)

    def _append_transcript_paths(self, paths) -> int:
        """Add transcript file paths to the list, skipping duplicates and
        unsupported extensions. Returns the number of new items added.
        """
        added = 0
        existing = {
            self.transcript_list.item(i).data(Qt.ItemDataRole.UserRole)
            for i in range(self.transcript_list.count())
        }
        for path in paths:
            if not path or path in existing:
                continue
            ext = os.path.splitext(path)[1].lower()
            if ext not in (".pdf", ".docx"):
                continue
            item = QListWidgetItem(os.path.basename(path))
            item.setData(Qt.ItemDataRole.UserRole, path)
            self.transcript_list.addItem(item)
            existing.add(path)
            added += 1
        if added:
            self._update_search_button()
        return added

    def dragEnterEvent(self, event: QDragEnterEvent):
        mime = event.mimeData()
        if mime.hasUrls() and any(
            u.isLocalFile()
            and os.path.splitext(u.toLocalFile())[1].lower() in (".pdf", ".docx")
            for u in mime.urls()
        ):
            event.acceptProposedAction()
        else:
            super().dragEnterEvent(event)

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            super().dragMoveEvent(event)

    def dropEvent(self, event: QDropEvent):
        mime = event.mimeData()
        if not mime.hasUrls():
            super().dropEvent(event)
            return
        paths = [u.toLocalFile() for u in mime.urls() if u.isLocalFile()]
        if self._append_transcript_paths(paths):
            event.acceptProposedAction()
        else:
            event.ignore()

    def _remove_selected_transcripts(self):
        for item in self.transcript_list.selectedItems():
            self.transcript_list.takeItem(self.transcript_list.row(item))
        self._update_search_button()

    def _current_transcript_paths(self) -> List[str]:
        return [
            self.transcript_list.item(i).data(Qt.ItemDataRole.UserRole)
            for i in range(self.transcript_list.count())
        ]

    # ------------------------------------------------------------------
    # Search
    # ------------------------------------------------------------------

    def _update_search_button(self):
        has_transcripts = self.transcript_list.count() > 0
        has_description = bool(self.desc_edit.toPlainText().strip())
        self.search_btn.setEnabled(has_transcripts and has_description)

    def _start_search(self):
        paths = self._current_transcript_paths()
        description = self.desc_edit.toPlainText().strip()
        if not paths or not description:
            return

        self._clear_results()
        self.progress_bar.setVisible(True)
        self.search_btn.setEnabled(False)
        self.status_label.setText("Searching transcripts…")

        generator = MediationBriefGenerator()
        self._worker = QuoteSearchWorker(
            generator, paths, description,
            brief_sections=self._brief_sections, parent=self,
        )
        self._worker.results_ready.connect(self._on_search_done)
        self._worker.error.connect(self._on_search_error)
        self._worker.start()

    def _clear_results(self):
        while self._results_layout.count():
            item = self._results_layout.takeAt(0)
            w = item.widget()
            if w is not None:
                w.setParent(None)
                w.deleteLater()
        self._result_widgets = []
        self.insert_btn.setEnabled(False)

    def _on_search_done(self, results: List[Dict]):
        self.progress_bar.setVisible(False)
        self.search_btn.setEnabled(True)
        if not results:
            self.status_label.setText("No matches found. Try revising your description.")
            return
        self.status_label.setText(f"Found {len(results)} result(s). Select quotes to insert.")
        for r in results:
            w = QuoteResultWidget(r, parent=self._results_container)
            w.checkbox.toggled.connect(self._update_insert_button)
            self._results_layout.addWidget(w)
            self._result_widgets.append(w)
        self._update_insert_button()

    def _on_search_error(self, msg: str):
        self.progress_bar.setVisible(False)
        self.search_btn.setEnabled(True)
        self.status_label.setText(f"Search failed: {msg}")

    def _update_insert_button(self):
        any_checked = any(w.is_selected() for w in self._result_widgets)
        self.insert_btn.setEnabled(any_checked)

    def _on_insert_clicked(self):
        selected = [w.get_quote_data() for w in self._result_widgets if w.is_selected()]
        if not selected:
            return
        self.quotes_to_insert.emit(selected)
        self.accept()
