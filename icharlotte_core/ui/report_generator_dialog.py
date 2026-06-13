"""
Report Generator Dialog and Pipeline Worker.

Provides a QDialog for configuring report generation options (report type,
custom context documents, section assignment, skip flags) and a QThread
worker that runs the pipeline stages with progress signals.
"""

import os
import time
import logging
from typing import Dict, List, Optional

from PySide6.QtCore import Qt, Signal, QThread
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QFormLayout, QGroupBox,
    QComboBox, QCheckBox, QPushButton, QListWidget, QListWidgetItem,
    QFileDialog, QLabel, QMessageBox, QWidget, QScrollArea,
)

logger = logging.getLogger(__name__)

# Import section order from gather
from Scripts.report_generator.gather import SECTION_ORDER

from . import theme


class ReportGeneratorDialog(QDialog):
    """Dialog for configuring report generation before running the pipeline."""

    def __init__(self, file_number: str, parent=None):
        super().__init__(parent)
        self.file_number = file_number
        self.setWindowTitle(f"Generate Report — {file_number}")
        self.setMinimumWidth(550)
        self.setMinimumHeight(500)

        layout = QVBoxLayout(self)

        # --- Report Type ---
        type_layout = QFormLayout()
        self.type_combo = QComboBox()
        self.type_combo.addItems(["FSR", "STATUS"])
        self.type_combo.currentTextChanged.connect(self._on_type_changed)
        type_layout.addRow("Report Type:", self.type_combo)

        # Prior report (for STATUS)
        prior_row = QHBoxLayout()
        self.prior_path_label = QLabel("None selected")
        self.prior_path_label.setStyleSheet("color: #909090;")
        prior_row.addWidget(self.prior_path_label, 1)
        self.prior_browse_btn = QPushButton("Browse...")
        self.prior_browse_btn.clicked.connect(self._browse_prior_report)
        prior_row.addWidget(self.prior_browse_btn)
        self.prior_path = None
        type_layout.addRow("Prior Report:", prior_row)
        self.prior_browse_btn.setEnabled(False)
        self.prior_path_label.setEnabled(False)

        layout.addLayout(type_layout)

        # --- Context Documents ---
        doc_group = QGroupBox("Context Documents")
        doc_layout = QVBoxLayout(doc_group)

        self.override_check = QCheckBox("Use these documents instead of agent outputs")
        self.override_check.setToolTip(
            "When checked, the pipeline ignores raw agent outputs from AI OUTPUT folder "
            "and uses only the documents listed below."
        )
        doc_layout.addWidget(self.override_check)

        self.doc_list = QListWidget()
        self.doc_list.setMinimumHeight(100)
        doc_layout.addWidget(self.doc_list)

        doc_btn_layout = QHBoxLayout()
        add_btn = QPushButton("Add Documents...")
        add_btn.setStyleSheet("background-color: #4CAF50; color: white;")
        add_btn.clicked.connect(self._add_documents)
        doc_btn_layout.addWidget(add_btn)

        remove_btn = QPushButton("Remove Selected")
        remove_btn.setStyleSheet("background-color: #f44336; color: white;")
        remove_btn.clicked.connect(self._remove_selected)
        doc_btn_layout.addWidget(remove_btn)

        doc_btn_layout.addStretch()
        doc_layout.addLayout(doc_btn_layout)

        # --- Advanced: Section Assignment ---
        self.advanced_check = QCheckBox("Advanced: Assign documents to sections")
        self.advanced_check.toggled.connect(self._toggle_advanced)
        doc_layout.addWidget(self.advanced_check)

        self.assignment_widget = QWidget()
        self.assignment_layout = QVBoxLayout(self.assignment_widget)
        self.assignment_layout.setContentsMargins(0, 0, 0, 0)
        self.assignment_combos: Dict[str, QComboBox] = {}  # doc_path -> combo
        self.assignment_widget.setVisible(False)

        scroll = QScrollArea()
        scroll.setWidget(self.assignment_widget)
        scroll.setWidgetResizable(True)
        scroll.setMaximumHeight(150)
        self.assignment_scroll = scroll
        scroll.setVisible(False)
        doc_layout.addWidget(scroll)

        layout.addWidget(doc_group)

        # --- Skip Options ---
        skip_group = QGroupBox("Pipeline Options")
        skip_layout = QVBoxLayout(skip_group)
        self.skip_refine = QCheckBox("Skip Refine (use raw content, no LLM)")
        self.skip_polish = QCheckBox("Skip Polish (no final LLM pass)")
        self.skip_validate = QCheckBox("Skip Validate (no formatting check)")
        skip_layout.addWidget(self.skip_refine)
        skip_layout.addWidget(self.skip_polish)
        skip_layout.addWidget(self.skip_validate)
        layout.addWidget(skip_group)

        # --- Buttons ---
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        cancel_btn = QPushButton("Cancel")
        cancel_btn.clicked.connect(self.reject)
        btn_layout.addWidget(cancel_btn)

        generate_btn = theme.primary_button("Generate Report")
        generate_btn.clicked.connect(self._validate_and_accept)
        btn_layout.addWidget(generate_btn)
        layout.addLayout(btn_layout)

    def _on_type_changed(self, text: str):
        is_status = text == "STATUS"
        self.prior_browse_btn.setEnabled(is_status)
        self.prior_path_label.setEnabled(is_status)

    def _browse_prior_report(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Select Prior Report", "",
            "Word Documents (*.docx);;All Files (*.*)"
        )
        if path:
            self.prior_path = path
            self.prior_path_label.setText(os.path.basename(path))
            self.prior_path_label.setStyleSheet("color: #333;")

    def _add_documents(self):
        paths, _ = QFileDialog.getOpenFileNames(
            self, "Select Context Documents", "",
            "Documents (*.docx *.pdf *.txt);;All Files (*.*)"
        )
        for path in paths:
            # Avoid duplicates
            existing = [self.doc_list.item(i).data(Qt.ItemDataRole.UserRole)
                        for i in range(self.doc_list.count())]
            if path not in existing:
                item = QListWidgetItem(os.path.basename(path))
                item.setData(Qt.ItemDataRole.UserRole, path)
                item.setToolTip(path)
                self.doc_list.addItem(item)
        self._rebuild_assignment_panel()

    def _remove_selected(self):
        for item in self.doc_list.selectedItems():
            self.doc_list.takeItem(self.doc_list.row(item))
        self._rebuild_assignment_panel()

    def _toggle_advanced(self, checked: bool):
        self.assignment_widget.setVisible(checked)
        self.assignment_scroll.setVisible(checked)
        if checked:
            self._rebuild_assignment_panel()

    def _rebuild_assignment_panel(self):
        """Rebuild section-assignment combos for each document."""
        # Clear existing
        while self.assignment_layout.count():
            child = self.assignment_layout.takeAt(0)
            if child.widget():
                child.widget().deleteLater()
        self.assignment_combos.clear()

        section_choices = ["(Auto-assign)"] + list(SECTION_ORDER)

        for i in range(self.doc_list.count()):
            item = self.doc_list.item(i)
            path = item.data(Qt.ItemDataRole.UserRole)
            row = QHBoxLayout()
            label = QLabel(os.path.basename(path))
            label.setMinimumWidth(200)
            row.addWidget(label)
            combo = QComboBox()
            combo.addItems(section_choices)
            row.addWidget(combo)
            self.assignment_combos[path] = combo

            container = QWidget()
            container.setLayout(row)
            self.assignment_layout.addWidget(container)

    def _validate_and_accept(self):
        if self.type_combo.currentText() == "STATUS" and not self.prior_path:
            reply = QMessageBox.question(
                self, "No Prior Report",
                "STATUS reports typically need a prior report for comparison.\n"
                "Continue without one?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            )
            if reply == QMessageBox.StandardButton.No:
                return
        self.accept()

    def get_config(self) -> dict:
        """Return the configured options as a dict."""
        docs = []
        for i in range(self.doc_list.count()):
            item = self.doc_list.item(i)
            path = item.data(Qt.ItemDataRole.UserRole)
            docs.append(path)

        # Section assignments (only if advanced is checked)
        assignments = {}
        if self.advanced_check.isChecked():
            for path, combo in self.assignment_combos.items():
                section = combo.currentText()
                if section != "(Auto-assign)":
                    assignments[path] = section

        return {
            "file_number": self.file_number,
            "report_type": self.type_combo.currentText(),
            "prior_report_path": self.prior_path,
            "skip_refine": self.skip_refine.isChecked(),
            "skip_polish": self.skip_polish.isChecked(),
            "skip_validate": self.skip_validate.isChecked(),
            "context_documents": docs,
            "override_agent_outputs": self.override_check.isChecked(),
            "section_assignments": assignments,
        }


class ReportPipelineWorker(QThread):
    """QThread that runs the report generation pipeline with stage-level signals."""

    # Matches AgentRunner signal patterns for StatusWidget compatibility
    stage_started = Signal(str, int, int)    # name, number, total
    stage_completed = Signal(str, float)     # name, duration_sec
    stage_failed = Signal(str, str, bool)    # name, error, recoverable
    progress_update = Signal(int, str)       # percent, message
    log_update = Signal(str)                 # log line
    output_file = Signal(str)                # path to generated report
    finished_signal = Signal(bool)           # success

    STAGES = ["GATHER", "REFINE", "ASSEMBLE", "VALIDATE", "POLISH"]

    def __init__(self, config: dict, parent=None):
        super().__init__(parent)
        self.config = config
        self._cancelled = False

    def cancel(self):
        self._cancelled = True

    def run(self):
        try:
            self._run_pipeline()
        except Exception as e:
            logger.exception(f"Pipeline failed: {e}")
            self.log_update.emit(f"ERROR: {e}")
            self.finished_signal.emit(False)

    def _run_pipeline(self):
        total = len(self.STAGES)
        cfg = self.config

        # Initialize LLM caller
        llm_caller = None
        if not (cfg["skip_refine"] and cfg["skip_polish"]):
            try:
                from icharlotte_core.llm_config import LLMCaller
                llm_caller = LLMCaller()
                self.log_update.emit("LLM caller initialized")
            except Exception as e:
                self.log_update.emit(f"WARNING: LLM not available: {e}")

        # ── Stage 1: GATHER ──
        self.stage_started.emit("GATHER", 1, total)
        self.progress_update.emit(5, "Gathering case data...")
        t0 = time.time()

        try:
            from Scripts.report_generator.gather import gather_case_data, _extract_text
            gathered = gather_case_data(
                cfg["file_number"], cfg["report_type"], cfg.get("prior_report_path")
            )
            self.log_update.emit(
                f"Gathered {len(gathered['sections'])} sections, "
                f"{len(gathered['metadata'])} metadata fields"
            )

            # Apply custom document overrides
            if cfg.get("context_documents"):
                self._apply_custom_documents(gathered, cfg, _extract_text)

        except Exception as e:
            self.stage_failed.emit("GATHER", str(e), False)
            self.finished_signal.emit(False)
            return

        self.stage_completed.emit("GATHER", time.time() - t0)
        if self._cancelled:
            self.finished_signal.emit(False)
            return

        # ── Stage 2: REFINE ──
        self.stage_started.emit("REFINE", 2, total)
        t0 = time.time()

        if cfg["skip_refine"]:
            self.progress_update.emit(30, "Refine skipped")
            self.log_update.emit("REFINE skipped (using raw content)")
            from Scripts.report_generator.pipeline import _passthrough_sections
            refined = _passthrough_sections(gathered)
        else:
            self.progress_update.emit(20, "Refining sections with LLM...")
            try:
                from Scripts.report_generator.refine import refine_sections
                refined = refine_sections(gathered, llm_caller=llm_caller)
                for section, text in refined.items():
                    self.log_update.emit(f"  {section}: {len(text)} chars")
            except Exception as e:
                self.stage_failed.emit("REFINE", str(e), True)
                self.finished_signal.emit(False)
                return

        self.stage_completed.emit("REFINE", time.time() - t0)
        if self._cancelled:
            self.finished_signal.emit(False)
            return

        # ── Stage 3: ASSEMBLE ──
        self.stage_started.emit("ASSEMBLE", 3, total)
        self.progress_update.emit(50, "Assembling Word document...")
        t0 = time.time()

        try:
            from Scripts.report_generator.assemble import assemble_report
            assembled_path = assemble_report(gathered, refined)
            self.log_update.emit(f"Assembled: {assembled_path}")
        except Exception as e:
            self.stage_failed.emit("ASSEMBLE", str(e), False)
            self.finished_signal.emit(False)
            return

        self.stage_completed.emit("ASSEMBLE", time.time() - t0)
        if self._cancelled:
            self.finished_signal.emit(False)
            return

        # ── Stage 4: VALIDATE ──
        self.stage_started.emit("VALIDATE", 4, total)
        t0 = time.time()

        if cfg["skip_validate"]:
            self.progress_update.emit(70, "Validate skipped")
            self.log_update.emit("VALIDATE skipped")
        else:
            self.progress_update.emit(65, "Validating formatting...")
            try:
                from Scripts.report_generator.validate import validate_report
                validation = validate_report(assembled_path)
                self.log_update.emit(
                    f"Validation: {validation.pass_count} pass, {validation.fail_count} fail"
                )
                if validation.fail_count > 0:
                    self.log_update.emit(f"WARNING: {validation.fail_count} validation failures")
            except Exception as e:
                self.log_update.emit(f"Validate error (non-fatal): {e}")

        self.stage_completed.emit("VALIDATE", time.time() - t0)
        if self._cancelled:
            self.finished_signal.emit(False)
            return

        # ── Stage 5: POLISH ──
        self.stage_started.emit("POLISH", 5, total)
        t0 = time.time()

        if cfg["skip_polish"]:
            self.progress_update.emit(90, "Polish skipped")
            self.log_update.emit("POLISH skipped")
            final_path = assembled_path
        else:
            self.progress_update.emit(80, "Polishing transitions...")
            try:
                from Scripts.report_generator.polish import polish_report
                final_path = polish_report(
                    assembled_path, gathered, refined, llm_caller=llm_caller
                )
                self.log_update.emit(f"Polished: {final_path}")
            except Exception as e:
                self.stage_failed.emit("POLISH", str(e), True)
                # Fall back to assembled version
                final_path = assembled_path
                self.log_update.emit(f"Polish failed, using assembled version: {e}")

        self.stage_completed.emit("POLISH", time.time() - t0)

        # Post-polish validate
        if not cfg["skip_polish"] and not cfg["skip_validate"]:
            try:
                from Scripts.report_generator.validate import validate_report as validate_final
                post = validate_final(final_path)
                self.log_update.emit(
                    f"Post-polish: {post.pass_count} pass, {post.fail_count} fail"
                )
            except Exception as e:
                self.log_update.emit(f"Post-polish validate error: {e}")

        # Log generation
        try:
            from Scripts.report_generator.pipeline import _log_generation
            _log_generation(cfg["file_number"], final_path, cfg["report_type"], gathered)
        except Exception:
            pass

        self.progress_update.emit(100, "Done!")
        self.output_file.emit(final_path)
        self.finished_signal.emit(True)

    def _apply_custom_documents(self, gathered: dict, cfg: dict, extract_text_fn):
        """Apply user-selected documents to the gathered data."""
        docs = cfg["context_documents"]
        assignments = cfg.get("section_assignments", {})
        override = cfg.get("override_agent_outputs", False)

        if override:
            # Clear all existing section content
            gathered["sections"] = {}
            gathered["section_files"] = {}
            self.log_update.emit("Overriding agent outputs with custom documents")

        # Group documents: assigned vs auto-assign
        assigned_docs = {}   # section -> [texts]
        auto_docs = []       # texts for LLM to sort

        for path in docs:
            text = extract_text_fn(path)
            if not text or len(text.strip()) < 20:
                self.log_update.emit(f"  Skipping (no content): {os.path.basename(path)}")
                continue

            section = assignments.get(path)
            if section:
                assigned_docs.setdefault(section, []).append(
                    (os.path.basename(path), text)
                )
            else:
                auto_docs.append((os.path.basename(path), text))

        # Apply assigned documents to their sections
        for section, doc_list in assigned_docs.items():
            content = "\n\n".join(
                f"--- Source: {name} ---\n\n{text}" for name, text in doc_list
            )
            if section in gathered["sections"] and not override:
                gathered["sections"][section] += "\n\n" + content
            else:
                gathered["sections"][section] = content
            gathered["section_files"][section] = "; ".join(
                p for p in docs if assignments.get(p) == section
            )
            self.log_update.emit(f"  Assigned {len(doc_list)} doc(s) to {section}")

        # Auto-assign docs: concatenate into a special key for the refine stage
        if auto_docs:
            content = "\n\n".join(
                f"--- Source: {name} ---\n\n{text}" for name, text in auto_docs
            )
            # Add to all included sections as supplementary context,
            # or if overriding, put into a catch-all that refine will sort
            if override:
                # When overriding, put auto-assign docs into each included section
                # so the LLM can pick relevant content during refine
                for section in gathered.get("included_sections", SECTION_ORDER):
                    if section not in gathered["sections"]:
                        gathered["sections"][section] = content
                    # If already has assigned content, append
                    elif section not in assigned_docs:
                        gathered["sections"][section] = content
            else:
                # Append as supplementary context to existing sections
                for section in gathered.get("included_sections", []):
                    if section in gathered["sections"]:
                        gathered["sections"][section] += (
                            "\n\n--- Additional Context ---\n\n" + content
                        )
            self.log_update.emit(f"  {len(auto_docs)} doc(s) for auto-assignment")
