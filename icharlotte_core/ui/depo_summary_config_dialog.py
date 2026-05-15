"""Modal dialog the user fills out after phase 1 of the deposition agent."""

import os
import re
from pathlib import Path

from PySide6.QtCore import Qt, QTimer
from PySide6.QtWidgets import (
    QCheckBox, QComboBox, QDialog, QDialogButtonBox, QFileDialog, QHBoxLayout,
    QLabel, QLineEdit, QListWidget, QListWidgetItem, QMessageBox, QPlainTextEdit,
    QPushButton, QSpinBox, QVBoxLayout, QWidget,
)

from icharlotte_core.deposition import session_manager


class DepoSummaryConfigDialog(QDialog):
    def __init__(self, session_path, parent=None):
        super().__init__(parent)
        self.session_path = Path(session_path)
        self._session = session_manager.read_session(self.session_path)

        self.setWindowTitle("Configure Deposition Summary")
        self.setModal(True)
        self.setWindowModality(Qt.ApplicationModal)
        self.resize(800, 750)

        root = QVBoxLayout(self)

        header_text = (
            f"Configure summary for <b>{self._session.get('deponent_name', '')}</b> "
            f"({self._session.get('deponent_type', '')}, "
            f"{self._session.get('deposition_date', 'date unknown')})"
        )
        root.addWidget(QLabel(header_text))

        # Topics list with native drag-reorder, checkboxes, and double-click rename.
        # No setItemWidget — custom widgets absorb mouse events and break drag.
        root.addWidget(QLabel("Topics (drag to reorder, uncheck to omit, double-click to rename):"))
        self.topics_list = QListWidget()
        self.topics_list.setDragDropMode(QListWidget.InternalMove)
        self.topics_list.setSelectionMode(QListWidget.SingleSelection)
        self.topics_list.setDefaultDropAction(Qt.MoveAction)
        for t in self._session.get("topics", []):
            item = QListWidgetItem(t.get("title", ""))
            item.setFlags(
                item.flags()
                | Qt.ItemIsUserCheckable
                | Qt.ItemIsEditable
                | Qt.ItemIsDragEnabled
            )
            item.setCheckState(Qt.Checked)
            self.topics_list.addItem(item)
        root.addWidget(self.topics_list, 1)

        # Summary bias row
        bias_row = QHBoxLayout()
        bias_row.addWidget(QLabel("Summary bias:"))
        self.bias_combo = QComboBox()
        for label, value in (
            ("Neutral", "neutral"),
            ("Most favorable to plaintiff", "pro_plaintiff"),
            ("Most favorable to defense", "pro_defense"),
            ("Custom…", "custom"),
        ):
            self.bias_combo.addItem(label, value)
        bias_row.addWidget(self.bias_combo)
        self.bias_custom_edit = QLineEdit()
        self.bias_custom_edit.setPlaceholderText(
            "Describe the editorial lens (e.g., 'Highlight any inconsistencies in injury testimony')"
        )
        self.bias_custom_edit.setVisible(False)
        bias_row.addWidget(self.bias_custom_edit, 1)
        root.addLayout(bias_row)

        self.bias_combo.currentIndexChanged.connect(self._on_bias_combo_changed)

        # Context documents drop zone
        ctx_header = QHBoxLayout()
        ctx_header.addWidget(QLabel("Context documents (drop .pdf, .doc, .docx here):"))
        ctx_header.addStretch(1)
        self._ctx_add_btn = QPushButton("Add files…")
        self._ctx_add_btn.clicked.connect(self._on_add_context_files)
        ctx_header.addWidget(self._ctx_add_btn)
        root.addLayout(ctx_header)

        self.context_docs_list = QListWidget()
        self.context_docs_list.setFixedHeight(80)
        self.context_docs_list.setAcceptDrops(True)
        self.context_docs_list.dragEnterEvent = self._context_drag_enter
        self.context_docs_list.dragMoveEvent = self._context_drag_enter
        self.context_docs_list.dropEvent = self._context_drop
        root.addWidget(self.context_docs_list)

        self._ctx_status_label = QLabel("")
        self._ctx_status_label.setStyleSheet("color: #c62828; font-style: italic; font-size: 11px;")
        root.addWidget(self._ctx_status_label)

        self._context_doc_paths: list[Path] = []
        self._ctx_status_clear_timer = QTimer(self)
        self._ctx_status_clear_timer.setSingleShot(True)
        self._ctx_status_clear_timer.timeout.connect(lambda: self._ctx_status_label.setText(""))

        # Additional topics
        root.addWidget(QLabel("Additional topics (one per line):"))
        self.added_topics_edit = QPlainTextEdit()
        self.added_topics_edit.setPlaceholderText(
            "One topic per line. These are appended after the checked topics above, in order."
        )
        self.added_topics_edit.setFixedHeight(70)
        root.addWidget(self.added_topics_edit)

        # Settings row
        settings_row = QHBoxLayout()
        settings_row.addWidget(QLabel("Bullets per topic:"))
        self.bullets_spinbox = QSpinBox()
        self.bullets_spinbox.setRange(1, 15)
        self.bullets_spinbox.setValue(5)
        settings_row.addWidget(self.bullets_spinbox)

        settings_row.addSpacing(20)
        settings_row.addWidget(QLabel("Deponent label:"))
        self.deponent_label_edit = QLineEdit("Plaintiff")
        settings_row.addWidget(self.deponent_label_edit, 1)

        settings_row.addSpacing(20)
        self.cross_check_checkbox = QCheckBox("Run cross-check pass")
        self.cross_check_checkbox.setChecked(True)
        settings_row.addWidget(self.cross_check_checkbox)
        root.addLayout(settings_row)

        # Custom rules
        root.addWidget(QLabel("Custom rules:"))
        self.custom_rules_edit = QPlainTextEdit()
        self.custom_rules_edit.setPlaceholderText(
            "Any extra instructions for the summary (tense, citation style, things to avoid, etc.)."
        )
        self.custom_rules_edit.setFixedHeight(90)
        root.addWidget(self.custom_rules_edit)

        # Buttons
        buttons = QDialogButtonBox()
        buttons.addButton("Cancel", QDialogButtonBox.RejectRole)
        generate_btn = buttons.addButton("Generate Summary", QDialogButtonBox.AcceptRole)
        generate_btn.setDefault(True)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        root.addWidget(buttons)

    def topic_rows_in_order(self):
        """Yield each topic QListWidgetItem in current visual order.

        Items expose .text() for the title and .checkState() for the checkbox state.
        """
        for i in range(self.topics_list.count()):
            yield self.topics_list.item(i)

    def _on_bias_combo_changed(self, _index):
        is_custom = self.bias_combo.currentData() == "custom"
        self.bias_custom_edit.setVisible(is_custom)
        if not is_custom:
            self.bias_custom_edit.clear()

    _CONTEXT_DOC_EXTS = (".pdf", ".doc", ".docx")

    def _context_drag_enter(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            event.ignore()

    def _context_drop(self, event):
        urls = event.mimeData().urls()
        QTimer.singleShot(0, lambda: self._add_context_paths_from_urls(urls))
        event.acceptProposedAction()

    def _add_context_paths_from_urls(self, urls):
        accepted, rejected = [], []
        for url in urls:
            if not url.isLocalFile():
                rejected.append(url.toString())
                continue
            p = Path(url.toLocalFile())
            if p.suffix.lower() in self._CONTEXT_DOC_EXTS and p.exists():
                accepted.append(p)
            else:
                rejected.append(p.name)
        for p in accepted:
            self._append_context_path(p)
        if rejected:
            self._show_ctx_status(
                f"Unsupported file type — skipped: {', '.join(rejected[:3])}"
                + (" …" if len(rejected) > 3 else "")
            )

    def _append_context_path(self, path: Path):
        if path in self._context_doc_paths:
            return  # de-dupe
        self._context_doc_paths.append(path)
        item = QListWidgetItem()
        row = QWidget()
        row_layout = QHBoxLayout(row)
        row_layout.setContentsMargins(4, 0, 4, 0)
        row_layout.addWidget(QLabel(path.name), 1)
        remove_btn = QPushButton("×")
        remove_btn.setFixedSize(20, 20)
        remove_btn.setStyleSheet("QPushButton { color: #c62828; font-weight: bold; }")
        remove_btn.clicked.connect(lambda _checked=False, p=path: self._remove_context_path(p))
        row_layout.addWidget(remove_btn)
        item.setSizeHint(row.sizeHint())
        self.context_docs_list.addItem(item)
        self.context_docs_list.setItemWidget(item, row)

    def _remove_context_path(self, path: Path):
        if path not in self._context_doc_paths:
            return
        idx = self._context_doc_paths.index(path)
        self._context_doc_paths.pop(idx)
        self.context_docs_list.takeItem(idx)

    def _compute_case_folder(self) -> str:
        """Return the case (matter) root folder derived from the session's input path.

        Walks up the deposition's path looking for a folder whose name starts with
        exactly 3 digits (the matter folder convention, e.g., '084 - Dudash').
        Falls back to the deposition's immediate parent folder if no match is found,
        or empty string if no input_path is available.
        """
        input_path = self._session.get("input_path", "")
        if not input_path:
            return ""
        parts = os.path.normpath(input_path).split(os.sep)
        for i in range(len(parts) - 1, -1, -1):
            if re.match(r'^\d{3}(\D|$)', parts[i]):
                return os.sep.join(parts[:i + 1])
        return os.path.dirname(input_path)

    def _on_add_context_files(self):
        initial_dir = self._compute_case_folder() or ""
        paths, _filter = QFileDialog.getOpenFileNames(
            self, "Add context documents", initial_dir,
            "Documents (*.pdf *.doc *.docx)",
        )
        for p in paths:
            path = Path(p)
            if path.suffix.lower() in self._CONTEXT_DOC_EXTS and path.exists():
                self._append_context_path(path)

    def _show_ctx_status(self, msg: str):
        self._ctx_status_label.setText(msg)
        self._ctx_status_clear_timer.start(3000)

    def accept(self):
        selected_topics = [
            item.text().strip()
            for item in self.topic_rows_in_order()
            if item.checkState() == Qt.Checked and item.text().strip()
        ]
        added_topics = [
            line.strip()
            for line in self.added_topics_edit.toPlainText().splitlines()
            if line.strip()
        ]
        if not selected_topics and not added_topics:
            QMessageBox.warning(
                self,
                "No topics selected",
                "Select at least one topic, or add a custom topic, before generating the summary.",
            )
            return
        context_doc_paths = [str(p.resolve()) for p in self._context_doc_paths if p.exists()]
        cfg = {
            "selected_topics": selected_topics,
            "added_topics": added_topics,
            "bullets_per_topic": self.bullets_spinbox.value(),
            "deponent_label": self.deponent_label_edit.text().strip() or "Deponent",
            "custom_rules": self.custom_rules_edit.toPlainText().strip(),
            "cross_check_enabled": self.cross_check_checkbox.isChecked(),
            "context_doc_paths": context_doc_paths,
            "bias": self.bias_combo.currentData() or "neutral",
            "bias_custom": (self.bias_custom_edit.text().strip()
                            if self.bias_combo.currentData() == "custom" else ""),
        }
        session_manager.update_user_config(self.session_path, cfg)
        super().accept()
