"""Standalone form widget for configuring a deposition summary.

Extracted from DepoSummaryConfigDialog so Wizard Mode can embed it inline
on the Settings page without duplicating ~250 LOC of widget setup.
"""

import os
import re
from pathlib import Path

from PySide6.QtCore import Qt, QSettings, QTimer
from PySide6.QtWidgets import (
    QCheckBox, QComboBox, QFileDialog, QHBoxLayout,
    QLabel, QLineEdit, QListWidget, QListWidgetItem, QMessageBox, QPlainTextEdit,
    QPushButton, QSpinBox, QSplitter, QVBoxLayout, QWidget,
)

from icharlotte_core.deposition import session_manager

_SETTINGS_KEY = "wizard/depo_settings/form_splitter_state"


class DepoSummaryConfigForm(QWidget):
    """All the configuration controls for a deposition summary, as an embeddable widget.

    Constructor takes a ``session_path`` string/Path.  Reads the session
    immediately on construction and populates the widgets.

    Public API
    ----------
    build_user_config() -> dict | None
        Validate inputs and return the config dict.  Returns None and shows a
        QMessageBox if validation fails.
    commit_user_config() -> bool
        Write the user_config to the session via session_manager.
        Returns True on success, False if validation fails.
    topic_rows_in_order()
        Yield each topic QListWidgetItem in current visual order.
    """

    def __init__(self, session_path, parent=None):
        super().__init__(parent)
        self.session_path = Path(session_path)
        self._session = session_manager.read_session(self.session_path)

        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(4)

        # Fixed header — sits above the splitter, not resizable.
        header_text = (
            f"Configure summary for <b>{self._session.get('deponent_name', '')}</b> "
            f"({self._session.get('deponent_type', '')}, "
            f"{self._session.get('deposition_date', 'date unknown')})"
        )
        header_label = QLabel(header_text)
        header_label.setContentsMargins(0, 0, 0, 4)
        root.addWidget(header_label)

        # ---- Inner splitter ----
        self.form_splitter = QSplitter(Qt.Orientation.Vertical)
        self.form_splitter.setObjectName("form_splitter")
        self.form_splitter.setChildrenCollapsible(False)

        # Section 0: Topics list
        topics_section = QWidget()
        topics_layout = QVBoxLayout(topics_section)
        topics_layout.setContentsMargins(0, 4, 0, 4)
        topics_layout.setSpacing(2)
        topics_lbl = QLabel("Topics (drag to reorder, uncheck to omit, double-click to rename):")
        topics_lbl.setContentsMargins(0, 0, 0, 2)
        topics_layout.addWidget(topics_lbl)
        self.topics_list = QListWidget()
        self.topics_list.setDragDropMode(QListWidget.InternalMove)
        self.topics_list.setSelectionMode(QListWidget.SingleSelection)
        self.topics_list.setDefaultDropAction(Qt.MoveAction)
        self.topics_list.setMinimumHeight(50)
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
        topics_layout.addWidget(self.topics_list)
        self.form_splitter.addWidget(topics_section)

        # Section 1: Audience + Tone row (non-collapsible)
        posture_section = QWidget()
        posture_layout = QVBoxLayout(posture_section)
        posture_layout.setContentsMargins(0, 0, 0, 0)
        posture_layout.setSpacing(4)

        # Audience sub-row
        audience_row = QHBoxLayout()
        audience_row.setContentsMargins(0, 0, 0, 0)
        audience_row.setSpacing(8)
        audience_row.addWidget(QLabel("Audience:"))
        self.audience_combo = QComboBox()
        for label, value in (
            ("Neutral", "neutral"),
            ("Plaintiff's Counsel", "pro_plaintiff"),
            ("Defense Counsel", "pro_defense"),
            ("Custom…", "custom"),
        ):
            self.audience_combo.addItem(label, value)
        audience_row.addWidget(self.audience_combo)
        self.audience_custom_edit = QLineEdit()
        self.audience_custom_edit.setPlaceholderText(
            "Describe the audience (e.g., 'Insurance carrier evaluating settlement value')"
        )
        self.audience_custom_edit.setVisible(False)
        audience_row.addWidget(self.audience_custom_edit, 1)
        posture_layout.addLayout(audience_row)

        # Tone sub-row
        tone_row = QHBoxLayout()
        tone_row.setContentsMargins(0, 0, 0, 0)
        tone_row.setSpacing(8)
        tone_row.addWidget(QLabel("Tone:"))
        self.tone_combo = QComboBox()
        for label, value in (
            ("Recitation (no editorializing)", "recitation"),
            ("Editorial (allow analysis)", "editorial"),
            ("Custom…", "custom"),
        ):
            self.tone_combo.addItem(label, value)
        tone_row.addWidget(self.tone_combo)
        self.tone_custom_edit = QLineEdit()
        self.tone_custom_edit.setPlaceholderText(
            "Describe the desired tone (e.g., 'Use plain-English, jury-presentation style')"
        )
        self.tone_custom_edit.setVisible(False)
        tone_row.addWidget(self.tone_custom_edit, 1)
        posture_layout.addLayout(tone_row)

        self.form_splitter.addWidget(posture_section)
        self.form_splitter.setCollapsible(1, False)

        self.audience_combo.currentIndexChanged.connect(self._on_audience_combo_changed)
        self.tone_combo.currentIndexChanged.connect(self._on_tone_combo_changed)

        # Section 2: Additional topics
        added_section = QWidget()
        added_layout = QVBoxLayout(added_section)
        added_layout.setContentsMargins(0, 4, 0, 4)
        added_layout.setSpacing(2)
        added_lbl = QLabel("Additional topics (one per line):")
        added_lbl.setContentsMargins(0, 0, 0, 2)
        added_layout.addWidget(added_lbl)
        self.added_topics_edit = QPlainTextEdit()
        self.added_topics_edit.setPlaceholderText(
            "One topic per line. These are appended after the checked topics above, in order."
        )
        self.added_topics_edit.setMinimumHeight(50)
        added_layout.addWidget(self.added_topics_edit)
        self.form_splitter.addWidget(added_section)

        # Section 3: Settings row (non-collapsible, fixed-ish single row)
        settings_section = QWidget()
        settings_layout = QHBoxLayout(settings_section)
        settings_layout.setContentsMargins(0, 0, 0, 0)
        settings_layout.setSpacing(8)
        settings_layout.addWidget(QLabel("Bullets per topic:"))
        self.bullets_spinbox = QSpinBox()
        self.bullets_spinbox.setRange(1, 15)
        self.bullets_spinbox.setValue(5)
        settings_layout.addWidget(self.bullets_spinbox)
        settings_layout.addSpacing(20)
        settings_layout.addWidget(QLabel("Deponent label:"))
        self.deponent_label_edit = QLineEdit("Plaintiff")
        settings_layout.addWidget(self.deponent_label_edit, 1)
        settings_layout.addSpacing(20)
        self.cross_check_checkbox = QCheckBox("Run cross-check pass")
        self.cross_check_checkbox.setChecked(True)
        settings_layout.addWidget(self.cross_check_checkbox)
        self.form_splitter.addWidget(settings_section)
        self.form_splitter.setCollapsible(3, False)

        # Section 4: Custom rules
        rules_section = QWidget()
        rules_layout = QVBoxLayout(rules_section)
        rules_layout.setContentsMargins(0, 4, 0, 4)
        rules_layout.setSpacing(2)
        rules_lbl = QLabel("Custom rules:")
        rules_lbl.setContentsMargins(0, 0, 0, 2)
        rules_layout.addWidget(rules_lbl)
        self.custom_rules_edit = QPlainTextEdit()
        self.custom_rules_edit.setPlaceholderText(
            "Any extra instructions for the summary (tense, citation style, things to avoid, etc.)."
        )
        self.custom_rules_edit.setMinimumHeight(50)
        rules_layout.addWidget(self.custom_rules_edit)
        self.form_splitter.addWidget(rules_section)

        # Section 5: Context documents
        ctx_section = QWidget()
        ctx_layout = QVBoxLayout(ctx_section)
        ctx_layout.setContentsMargins(0, 4, 0, 4)
        ctx_layout.setSpacing(2)

        ctx_header = QHBoxLayout()
        ctx_header.setContentsMargins(0, 0, 0, 0)
        ctx_header.setSpacing(8)
        ctx_hdr_lbl = QLabel("Context documents (drop .pdf, .doc, .docx here):")
        ctx_hdr_lbl.setContentsMargins(0, 0, 0, 2)
        ctx_header.addWidget(ctx_hdr_lbl)
        ctx_header.addStretch(1)
        self._ctx_add_btn = QPushButton("Add files…")
        self._ctx_add_btn.clicked.connect(self._on_add_context_files)
        ctx_header.addWidget(self._ctx_add_btn)
        ctx_layout.addLayout(ctx_header)

        self.context_docs_list = QListWidget()
        self.context_docs_list.setMinimumHeight(50)
        self.context_docs_list.setAcceptDrops(True)
        self.context_docs_list.dragEnterEvent = self._context_drag_enter
        self.context_docs_list.dragMoveEvent = self._context_drag_enter
        self.context_docs_list.dropEvent = self._context_drop
        ctx_layout.addWidget(self.context_docs_list)

        self._ctx_status_label = QLabel("")
        self._ctx_status_label.setStyleSheet("color: #c62828; font-style: italic; font-size: 11px;")
        ctx_layout.addWidget(self._ctx_status_label)

        self.form_splitter.addWidget(ctx_section)

        root.addWidget(self.form_splitter, 1)

        # ---- Deferred state init ----
        self._context_doc_paths: list[Path] = []
        self._ctx_status_clear_timer = QTimer(self)
        self._ctx_status_clear_timer.setSingleShot(True)
        self._ctx_status_clear_timer.timeout.connect(lambda: self._ctx_status_label.setText(""))

        # ---- Restore / initialise splitter sizes ----
        self._restore_splitter()
        self.form_splitter.splitterMoved.connect(self._save_splitter)

    # ---- Splitter persistence ----

    def _restore_splitter(self):
        settings = QSettings("iCharlotte", "iCharlotte")
        state = settings.value(_SETTINGS_KEY)
        if state:
            self.form_splitter.restoreState(state)
        else:
            # First-run defaults (pixels): topics=130, audience+tone=60, added=70,
            # settings=32, rules=80, context=80
            self.form_splitter.setSizes([130, 60, 70, 32, 80, 80])

    def _save_splitter(self):
        settings = QSettings("iCharlotte", "iCharlotte")
        settings.setValue(_SETTINGS_KEY, self.form_splitter.saveState())

    # ---- Public API ----

    def topic_rows_in_order(self):
        """Yield each topic QListWidgetItem in current visual order."""
        for i in range(self.topics_list.count()):
            yield self.topics_list.item(i)

    def build_user_config(self) -> dict | None:
        """Validate inputs and return the config dict.

        Returns None (and shows a QMessageBox) if validation fails.
        """
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
            return None
        context_doc_paths = [str(p.resolve()) for p in self._context_doc_paths if p.exists()]
        audience_value = self.audience_combo.currentData() or "neutral"
        tone_value = self.tone_combo.currentData() or "recitation"
        cfg = {
            "selected_topics": selected_topics,
            "added_topics": added_topics,
            "bullets_per_topic": self.bullets_spinbox.value(),
            "deponent_label": self.deponent_label_edit.text().strip() or "Deponent",
            "custom_rules": self.custom_rules_edit.toPlainText().strip(),
            "cross_check_enabled": self.cross_check_checkbox.isChecked(),
            "context_doc_paths": context_doc_paths,
            "audience": audience_value,
            "audience_custom": (self.audience_custom_edit.text().strip()
                                if audience_value == "custom" else ""),
            "tone": tone_value,
            "tone_custom": (self.tone_custom_edit.text().strip()
                            if tone_value == "custom" else ""),
        }
        return cfg

    def commit_user_config(self) -> bool:
        """Write user_config to the session file.

        Returns True on success, False if validation fails.
        """
        cfg = self.build_user_config()
        if cfg is None:
            return False
        session_manager.update_user_config(self.session_path, cfg)
        return True

    # ---- Internal helpers ----

    def _on_audience_combo_changed(self, _index):
        is_custom = self.audience_combo.currentData() == "custom"
        self.audience_custom_edit.setVisible(is_custom)
        if not is_custom:
            self.audience_custom_edit.clear()

    def _on_tone_combo_changed(self, _index):
        is_custom = self.tone_combo.currentData() == "custom"
        self.tone_custom_edit.setVisible(is_custom)
        if not is_custom:
            self.tone_custom_edit.clear()

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
        """Return the case (matter) root folder derived from the session's input path."""
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
