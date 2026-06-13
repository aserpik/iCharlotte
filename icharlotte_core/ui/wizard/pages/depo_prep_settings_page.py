"""Settings page for the Depo Prep wizard task."""
from __future__ import annotations

import json
import os
import tempfile
from pathlib import Path
from typing import List, Optional

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QCheckBox, QComboBox, QFileDialog, QGroupBox, QHBoxLayout, QLabel,
    QLineEdit, QListWidget, QListWidgetItem, QPlainTextEdit, QPushButton,
    QSizePolicy, QStackedWidget, QVBoxLayout, QWidget,
)

from ..file_drop import enable_file_drop
from ..registry import TaskSpec
from .depo_prep_topic_editor import TopicEditor
from .settings_page import SettingsPage


_STYLES = [
    ("discovery", "Discovery / Fact-gathering"),
    ("lockdown", "Lock-down (leading admissions)"),
    ("expert", "Expert challenge (Daubert-style)"),
    ("friendly", "Friendly (own client prep)"),
]

# Internal stack page indices.
_PAGE_SETUP = 0
_PAGE_REVIEW = 1


def _load_case_parties(case_root: str) -> List[str]:
    """Best-effort: return list of party names from CaseDataManager. Empty on any error."""
    try:
        from case_data_manager import CaseDataManager
    except ImportError:
        try:
            from Scripts.case_data_manager import CaseDataManager
        except ImportError:
            return []

    import re
    m = re.search(r"(\d{4})\D+(\d{3})", str(case_root or ""))
    if not m:
        return []
    file_number = f"{m.group(1)}.{m.group(2)}"

    try:
        cdm = CaseDataManager()
        parties = []
        for key in ("plaintiffs", "defendants"):
            v = cdm.get_value(file_number, key)
            if isinstance(v, list):
                parties.extend(str(x) for x in v)
            elif isinstance(v, str) and v.strip():
                parties.append(v.strip())
        return parties
    except Exception:
        return []


def _detach_layout(layout) -> None:
    """Recursively remove and delete every widget in ``layout`` (including those
    nested inside sub-layouts).

    The base SettingsPage builds a placeholder UI with buttons nested in
    sub-layouts (the file-button row and the Proceed-button row). A shallow
    strip leaves those buttons parented to the page with no layout slot, so they
    paint as orphans over our UI — the base 'Proceed' button is blue, which
    showed up as a stray blue block. Recursing fixes that.
    """
    while layout.count():
        item = layout.takeAt(0)
        w = item.widget()
        if w is not None:
            w.setParent(None)
            w.deleteLater()
            continue
        sub = item.layout()
        if sub is not None:
            _detach_layout(sub)


class DepoPrepSettingsPage(SettingsPage):
    """Custom settings page for Depo Prep.

    The page is an internal two-screen stack:
      * Setup screen — deponent, source files, instructions, per-topic content,
        and the "Analyze Sources" button.
      * Review screen — appears after Analyze; shows the proposed topics in an
        editable TopicEditor plus the "Generate Outline" button.

    Phase 1 is triggered on demand by "Analyze Sources", which emits
    ``analyze_requested``. TaskTab routes that to a settings-owned worker run
    (mirroring the speculative path) WITHOUT switching the *TaskTab* page, so the
    review screen can take over this page once Phase 1 completes. We do NOT reuse
    ``proceed_requested`` because that switches the tab to the Status page and
    routes ``awaiting_input`` to the generic deposition dialog instead of our
    embedded TopicEditor.

    ``phase2_requested(str)`` fires on "Generate Outline", carrying the
    session.json path; TaskTab connects it to ``advance_to_status_with_phase2``.
    """

    # Emitted when the user clicks "Analyze Sources" (Phase 1 trigger).
    analyze_requested = Signal()
    # Emitted when the user clicks "Generate Outline" (Phase 2 trigger).
    phase2_requested = Signal(str)

    def __init__(self, spec: TaskSpec, files, case_root: str | None = None, parent=None):
        super().__init__(spec, files=files, case_root=case_root, parent=parent)

        # Discard the base SettingsPage placeholder UI entirely; we render our
        # own. Recursively detach so no orphan base widgets paint over our page.
        outer = self.layout()
        _detach_layout(outer)
        for _attr in ("files_label", "files_list", "add_files_btn",
                      "remove_btn", "proceed_btn"):
            if hasattr(self, _attr):
                delattr(self, _attr)

        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        self._deponent_files: List[str] = []
        self._context_files: List[str] = []
        self._session_path: Optional[str] = None

        # Internal two-screen stack.
        self._stack = QStackedWidget()
        self._stack.addWidget(self._build_setup_page())    # index 0 = setup
        self._stack.addWidget(self._build_review_page())   # index 1 = review
        self._stack.setCurrentIndex(_PAGE_SETUP)
        outer.addWidget(self._stack)

        self._refresh_buttons()

    # ------------------------------------------------------------------ #
    # Screen builders
    # ------------------------------------------------------------------ #
    def _build_setup_page(self) -> QWidget:
        page = QWidget()
        v = QVBoxLayout(page)
        v.setContentsMargins(24, 24, 24, 24)
        v.setSpacing(10)

        # ---- Deponent ----
        deponent_box = QGroupBox("Deponent")
        dep_layout = QVBoxLayout(deponent_box)

        row1 = QHBoxLayout()
        row1.addWidget(QLabel("Name:"))
        self.deponent_name_combo = QComboBox()
        self.deponent_name_combo.setEditable(True)
        for name in _load_case_parties(self._case_root or ""):
            self.deponent_name_combo.addItem(name)
        self.deponent_name_combo.currentTextChanged.connect(self._refresh_buttons)
        row1.addWidget(self.deponent_name_combo, 1)
        dep_layout.addLayout(row1)

        row2 = QHBoxLayout()
        row2.addWidget(QLabel("Role:"))
        self.deponent_role_edit = QLineEdit()
        self.deponent_role_edit.setPlaceholderText("e.g., Plaintiff's treating orthopedist")
        row2.addWidget(self.deponent_role_edit, 1)
        dep_layout.addLayout(row2)

        v.addWidget(deponent_box)

        # ---- Sources ----
        sources_box = QGroupBox("Source files")
        s_layout = QVBoxLayout(sources_box)

        s_layout.addWidget(QLabel("Deponent's own materials:"))
        dep_files_row = QHBoxLayout()
        self.add_deponent_files_btn = QPushButton("+ Add files")
        self.add_deponent_files_btn.clicked.connect(self._on_add_deponent_files)
        dep_files_row.addWidget(self.add_deponent_files_btn)
        self.remove_deponent_files_btn = QPushButton("Remove selected")
        self.remove_deponent_files_btn.clicked.connect(
            lambda: self._remove_selected(self.deponent_files_list, self._deponent_files))
        dep_files_row.addWidget(self.remove_deponent_files_btn)
        dep_files_row.addStretch()
        s_layout.addLayout(dep_files_row)
        self.deponent_files_list = QListWidget()
        self.deponent_files_list.setSelectionMode(QListWidget.SelectionMode.ExtendedSelection)
        self.deponent_files_list.setMaximumHeight(80)
        enable_file_drop(self.deponent_files_list, self.add_deponent_files)
        s_layout.addWidget(self.deponent_files_list)

        s_layout.addWidget(QLabel("Case context:"))
        ctx_files_row = QHBoxLayout()
        self.add_context_files_btn = QPushButton("+ Add files")
        self.add_context_files_btn.clicked.connect(self._on_add_context_files)
        ctx_files_row.addWidget(self.add_context_files_btn)
        self.remove_context_files_btn = QPushButton("Remove selected")
        self.remove_context_files_btn.clicked.connect(
            lambda: self._remove_selected(self.context_files_list, self._context_files))
        ctx_files_row.addWidget(self.remove_context_files_btn)
        ctx_files_row.addStretch()
        s_layout.addLayout(ctx_files_row)
        self.context_files_list = QListWidget()
        self.context_files_list.setSelectionMode(QListWidget.SelectionMode.ExtendedSelection)
        self.context_files_list.setMaximumHeight(80)
        enable_file_drop(self.context_files_list, self.add_context_files)
        s_layout.addWidget(self.context_files_list)

        v.addWidget(sources_box)

        # ---- Instructions ----
        instr_box = QGroupBox("Instructions")
        i_layout = QVBoxLayout(instr_box)
        style_row = QHBoxLayout()
        style_row.addWidget(QLabel("Style:"))
        self.style_combo = QComboBox()
        for key, label in _STYLES:
            self.style_combo.addItem(label, key)
        style_row.addWidget(self.style_combo, 1)
        i_layout.addLayout(style_row)
        i_layout.addWidget(QLabel("Free-text strategy notes:"))
        self.free_text_edit = QPlainTextEdit()
        self.free_text_edit.setPlaceholderText(
            "Case theory, topics to emphasize, key admissions to extract, things to avoid...")
        self.free_text_edit.setMinimumHeight(80)
        i_layout.addWidget(self.free_text_edit)
        v.addWidget(instr_box)

        # ---- Per-topic content flags ----
        flags_box = QGroupBox("Per-topic content")
        f_layout = QVBoxLayout(flags_box)
        self.flag_strategic = QCheckBox("Strategic note (\"why this topic\")")
        self.flag_strategic.setChecked(True)
        self.flag_source_facts = QCheckBox("Key source facts (with citations)")
        self.flag_source_facts.setChecked(True)
        self.flag_impeachment = QCheckBox("Impeachment hooks / inconsistencies")
        self.flag_objection = QCheckBox("Anticipated objections + workaround phrasings")
        for cb in (self.flag_strategic, self.flag_source_facts,
                   self.flag_impeachment, self.flag_objection):
            f_layout.addWidget(cb)
        v.addWidget(flags_box)

        v.addStretch(1)

        # ---- Action row ----
        action_row = QHBoxLayout()
        action_row.addStretch()
        self.analyze_btn = QPushButton("Analyze Sources")
        self.analyze_btn.setStyleSheet(
            "background-color: #1976D2; color: white; font-weight: 600; padding: 8px 24px;")
        self.analyze_btn.clicked.connect(self._on_analyze_clicked)
        action_row.addWidget(self.analyze_btn)
        v.addLayout(action_row)

        return page

    def _build_review_page(self) -> QWidget:
        page = QWidget()
        v = QVBoxLayout(page)
        v.setContentsMargins(24, 24, 24, 24)
        v.setSpacing(10)

        self._review_header = QLabel("Proposed Topics")
        self._review_header.setStyleSheet("font-size: 15px; font-weight: 700;")
        v.addWidget(self._review_header)

        hint = QLabel(
            "Review the topics below. Check/uncheck to include, edit titles, "
            "drag to reorder, or add your own. Then click Generate Outline.")
        hint.setWordWrap(True)
        hint.setStyleSheet("color: #555;")
        v.addWidget(hint)

        self._phase1_status_label = QLabel("")
        self._phase1_status_label.setStyleSheet("color: #555; font-style: italic;")
        self._phase1_status_label.setWordWrap(True)
        v.addWidget(self._phase1_status_label)

        self.topic_editor = TopicEditor()
        self.topic_editor.setVisible(False)
        v.addWidget(self.topic_editor, 1)

        btn_row = QHBoxLayout()
        self.back_btn = QPushButton("< Back to settings")
        self.back_btn.clicked.connect(self._on_back_clicked)
        btn_row.addWidget(self.back_btn)
        btn_row.addStretch()
        self.generate_btn = QPushButton("Generate Outline")
        self.generate_btn.setStyleSheet(
            "background-color: #43A047; color: white; font-weight: 600; padding: 8px 24px;")
        self.generate_btn.setEnabled(False)
        self.generate_btn.clicked.connect(self._on_generate_clicked)
        btn_row.addWidget(self.generate_btn)
        v.addLayout(btn_row)

        return page

    # ------------------------------------------------------------------ #
    # Compatibility with base class
    # ------------------------------------------------------------------ #
    def _refresh_files_list(self) -> None:
        # Base SettingsPage calls this from __init__; ours does nothing because
        # we manage two separate lists.
        return

    @property
    def files(self) -> list[str]:
        # TaskTab uses self.files as the positional argv passed to the subprocess.
        # For Depo Prep, that's the config.json path after _on_analyze_clicked.
        return list(self._files)

    # ------------------------------------------------------------------ #
    # Public setters used in tests
    # ------------------------------------------------------------------ #
    def set_deponent_name(self, name: str) -> None:
        self.deponent_name_combo.setEditText(name)

    def set_deponent_role(self, role: str) -> None:
        self.deponent_role_edit.setText(role)

    def add_deponent_files(self, paths: List[str]) -> None:
        for p in paths:
            if p not in self._deponent_files:
                self._deponent_files.append(p)
                self.deponent_files_list.addItem(QListWidgetItem(os.path.basename(p)))
        self._refresh_buttons()

    def add_context_files(self, paths: List[str]) -> None:
        for p in paths:
            if p not in self._context_files:
                self._context_files.append(p)
                self.context_files_list.addItem(QListWidgetItem(os.path.basename(p)))
        self._refresh_buttons()

    # ------------------------------------------------------------------ #
    # Internals
    # ------------------------------------------------------------------ #
    def _on_add_deponent_files(self) -> None:
        paths, _ = QFileDialog.getOpenFileNames(
            self, "Add deponent's own materials", self._case_root or "", "All files (*.*)")
        if paths:
            self.add_deponent_files(paths)

    def _on_add_context_files(self) -> None:
        paths, _ = QFileDialog.getOpenFileNames(
            self, "Add case context", self._case_root or "", "All files (*.*)")
        if paths:
            self.add_context_files(paths)

    def _remove_selected(self, list_widget: QListWidget, store: List[str]) -> None:
        rows = sorted({i.row() for i in list_widget.selectedIndexes()}, reverse=True)
        for r in rows:
            if 0 <= r < len(store):
                store.pop(r)
                list_widget.takeItem(r)
        self._refresh_buttons()

    def _refresh_buttons(self) -> None:
        if not hasattr(self, "analyze_btn"):
            return
        has_name = bool(self.deponent_name_combo.currentText().strip())
        has_sources = bool(self._deponent_files or self._context_files)
        self.analyze_btn.setEnabled(has_name and has_sources)

    def _build_config_dict(self) -> dict:
        style = self.style_combo.currentData() or "discovery"
        return {
            "deponent_name": self.deponent_name_combo.currentText().strip(),
            "deponent_role": self.deponent_role_edit.text().strip(),
            "deponent_sources": list(self._deponent_files),
            "context_sources": list(self._context_files),
            "style": style,
            "free_text_notes": self.free_text_edit.toPlainText().strip(),
            "per_topic_flags": {
                "strategic_note": self.flag_strategic.isChecked(),
                "source_facts": self.flag_source_facts.isChecked(),
                "impeachment_hook": self.flag_impeachment.isChecked(),
                "objection_alts": self.flag_objection.isChecked(),
            },
            "case_root": self._case_root or "",
        }

    def _on_analyze_clicked(self) -> None:
        cfg = self._build_config_dict()
        # Persist to a temp config.json — the subprocess reads it.
        tmpdir = tempfile.mkdtemp(prefix="depo_prep_config_")
        cfg_path = Path(tmpdir) / "config.json"
        cfg_path.write_text(json.dumps(cfg, indent=2), encoding="utf-8")
        # Replace _files so TaskTab uses cfg_path as the positional argv.
        self._files = [str(cfg_path)]

        # Move to the review screen and show analysis-in-progress state.
        deponent = cfg["deponent_name"] or "deponent"
        self._review_header.setText(f"Proposed Topics — {deponent}")
        self._phase1_status_label.setStyleSheet("color: #1976D2; font-style: italic;")
        self._phase1_status_label.setText("Analyzing sources... this can take a minute.")
        self._phase1_status_label.setVisible(True)
        self.topic_editor.setVisible(False)
        self.generate_btn.setEnabled(False)
        self.analyze_btn.setEnabled(False)
        self._stack.setCurrentIndex(_PAGE_REVIEW)

        self.analyze_requested.emit()

    def to_dict(self) -> dict:
        return self._build_config_dict()

    # ---- Phase 1 completion hook ----
    def attach_worker(self, worker) -> bool:
        worker.status.connect(self._phase1_status_label.setText)
        worker.awaiting_input.connect(self._on_phase1_complete)
        worker.failed.connect(self._on_phase1_failed)
        return True

    def _on_phase1_complete(self, session_path: str) -> None:
        self._session_path = session_path
        # Ensure we're on the review screen (covers direct calls in tests).
        self._stack.setCurrentIndex(_PAGE_REVIEW)
        session_dir = Path(session_path).parent
        try:
            topics_payload = json.loads(
                (session_dir / "topics.json").read_text(encoding="utf-8"))
        except Exception as e:
            self._on_phase1_failed(f"Could not load topics.json: {e}")
            return
        self.topic_editor.set_topics(topics_payload.get("topics", []))
        self.topic_editor.setVisible(True)
        self.generate_btn.setEnabled(True)
        warning = topics_payload.get("warning")
        if warning:
            self._phase1_status_label.setText(warning)
            self._phase1_status_label.setStyleSheet("color: #E65100; font-style: italic;")
            self._phase1_status_label.setVisible(True)
        else:
            self._phase1_status_label.setVisible(False)

    def _on_phase1_failed(self, err: str) -> None:
        self._phase1_status_label.setText(f"Analysis failed: {err}")
        self._phase1_status_label.setStyleSheet("color: #C62828; font-style: italic;")
        self._phase1_status_label.setVisible(True)
        # Let the user go back and adjust; re-enable Analyze on the setup screen.
        self.analyze_btn.setEnabled(True)

    def _on_back_clicked(self) -> None:
        self._stack.setCurrentIndex(_PAGE_SETUP)
        self._refresh_buttons()

    def _on_generate_clicked(self) -> None:
        if not self._session_path:
            return
        # Persist current topic editor state back to topics.json.
        topics = self.topic_editor.get_topics()
        topics_json_path = Path(self._session_path).parent / "topics.json"
        topics_json_path.write_text(
            json.dumps({"topics": topics}, indent=2), encoding="utf-8")
        self.phase2_requested.emit(self._session_path)
