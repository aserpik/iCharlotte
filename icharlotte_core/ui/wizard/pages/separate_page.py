"""Wizard Mode "Separate Documents" task.

Settings (pick sensitivity, Analyze) -> Status (analyzing) -> Workbench
(embedded SeparatorWorkbench for review + split/merge). Backed by
SeparateAnalysisWorker, which runs Scripts/separate.py --headless to produce
the document map (and the Word index, as a side effect, exactly like Advanced
Mode).
"""
import json
import os
import re
import subprocess
import sys

from PySide6.QtCore import Qt, QThread, Signal
from PySide6.QtWidgets import (
    QFileDialog, QHBoxLayout, QLabel, QMessageBox, QSlider,
    QVBoxLayout, QWidget,
)

from icharlotte_core.config import SCRIPTS_DIR
from icharlotte_core import case_index_store
from icharlotte_core.ui.wizard import theme
from icharlotte_core.ui.wizard.pages.status_page import StatusPage
from icharlotte_core.ui.wizard.task_debug_helpers import (
    emit_debug,
    finish_debug_run,
    record_status,
    start_debug_run,
)
from icharlotte_core.ui.wizard.task_scaffold import WizardTaskContainer
from icharlotte_core.ui.separator_workbench import SeparatorWorkbench
from icharlotte_core.ui.wizard.file_picker import resolve_default_folder


class SeparateAnalysisWorker(QThread):
    progress = Signal(str)
    finished_analysis = Signal(bool, object)  # (success, list[dict] | error str)

    def __init__(self, pdf_path: str, sensitivity: int, parent=None):
        super().__init__(parent)
        self.pdf_path = pdf_path
        self.sensitivity = sensitivity

    @staticmethod
    def _parse_docs(stdout: str):
        match = re.search(r"JSON_MAP:\s*(.+)", stdout)
        if not match:
            return None
        json_path = match.group(1).strip()
        try:
            with open(json_path, "r", encoding="utf-8") as f:
                docs = json.load(f)
        except Exception:
            return None
        try:
            os.remove(json_path)
        except OSError:
            pass
        return docs

    def run(self):
        try:
            script_path = os.path.join(SCRIPTS_DIR, "separate.py")
            self.progress.emit(f"Analyzing {os.path.basename(self.pdf_path)}...")
            proc = subprocess.run(
                [sys.executable, script_path, "--headless",
                 "--sensitivity", str(self.sensitivity), self.pdf_path],
                capture_output=True, text=True, encoding="utf-8", errors="replace",
            )
            for line in (proc.stdout or "").splitlines():
                if line.strip():
                    self.progress.emit(line.strip())
            if proc.returncode != 0:
                self.finished_analysis.emit(
                    False, f"Analysis failed (exit {proc.returncode}). {proc.stderr[-500:]}")
                return
            docs = self._parse_docs(proc.stdout or "")
            if docs is None:
                self.finished_analysis.emit(
                    False, "Analysis completed but no document map was produced.")
                return
            self.finished_analysis.emit(True, docs)
        except Exception as e:
            self.finished_analysis.emit(False, str(e))


def _docs_from_workbench(workbench) -> list:
    """Read the current (possibly edited) document rows from a SeparatorWorkbench."""
    table = workbench.doc_table
    return [workbench._get_doc_from_row(row) for row in range(table.rowCount())]


PAGE_SETTINGS = 0
PAGE_STATUS = 1
PAGE_WORKBENCH = 2


class SeparateSettingsPage(QWidget):
    analyze_requested = Signal(int)  # sensitivity

    def __init__(self, pdf_path: str, parent=None):
        super().__init__(parent)
        self.pdf_path = pdf_path
        layout = QVBoxLayout(self)
        layout.setContentsMargins(theme.SPACE_XL, theme.SPACE_XL, theme.SPACE_XL, theme.SPACE_XL)
        layout.setSpacing(theme.SPACE_MD)

        layout.addWidget(theme.page_title("Separate Documents"))
        layout.addWidget(theme.helper_text(
            "iCharlotte will scan the PDF, identify the distinct documents inside it, "
            "and let you review, rename, split, and merge them. An index is also saved "
            "to NOTES/AI OUTPUT/INDEXES."))
        layout.addWidget(QLabel(f"File: {os.path.basename(pdf_path)}"))

        sens_row = QHBoxLayout()
        sens_row.addWidget(QLabel("Separation sensitivity:"))
        broad = QLabel("Broad"); broad.setStyleSheet("color:#666;font-size:11px;")
        sens_row.addWidget(broad)
        self.sensitivity_slider = QSlider(Qt.Orientation.Horizontal)
        self.sensitivity_slider.setMinimum(1)
        self.sensitivity_slider.setMaximum(3)
        self.sensitivity_slider.setValue(2)
        self.sensitivity_slider.setTickPosition(QSlider.TickPosition.TicksBelow)
        self.sensitivity_slider.setTickInterval(1)
        self.sensitivity_slider.setPageStep(1)
        self.sensitivity_slider.setFixedWidth(120)
        sens_row.addWidget(self.sensitivity_slider)
        fine = QLabel("Fine"); fine.setStyleSheet("color:#666;font-size:11px;")
        sens_row.addWidget(fine)
        sens_row.addStretch()
        layout.addLayout(sens_row)

        layout.addStretch(1)
        btn_row = QHBoxLayout()
        btn_row.addStretch()
        self.analyze_btn = theme.primary_button("Analyze")
        self.analyze_btn.clicked.connect(
            lambda: self.analyze_requested.emit(self.sensitivity_slider.value()))
        btn_row.addWidget(self.analyze_btn)
        layout.addLayout(btn_row)

    def to_dict(self) -> dict:
        return {"sensitivity": self.sensitivity_slider.value()}

    def from_dict(self, data: dict) -> None:
        if "sensitivity" in data:
            self.sensitivity_slider.setValue(int(data["sensitivity"]))


class SeparateWorkbenchPage(QWidget):
    """Hosts the SeparatorWorkbench plus a result banner + Open Folder button."""

    def __init__(self, parent=None):
        super().__init__(parent)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(theme.SPACE_MD, theme.SPACE_MD, theme.SPACE_MD, theme.SPACE_MD)
        layout.setSpacing(theme.SPACE_SM)

        self.banner = QLabel("")
        self.banner.setWordWrap(True)
        self.banner.setVisible(False)
        layout.addWidget(self.banner)

        self.workbench = SeparatorWorkbench()
        layout.addWidget(self.workbench, 1)

        btn_row = QHBoxLayout()
        self.open_folder_btn = theme.secondary_button("Open Output Folder")
        self.open_folder_btn.setVisible(False)
        self.open_folder_btn.clicked.connect(self._open_folder)
        btn_row.addWidget(self.open_folder_btn)
        btn_row.addStretch()
        layout.addLayout(btn_row)

        self._output_folder = ""
        self.workbench.processing_complete.connect(self._on_processing_complete)

    def _on_processing_complete(self, summary: dict):
        created = summary.get("created", [])
        errors = summary.get("errors", [])
        self._output_folder = summary.get("output_folder", "")
        text = f"Created {len(created)} file(s) in {os.path.basename(self._output_folder)}."
        if errors:
            text += f"  {len(errors)} error(s): " + "; ".join(errors[:3])
        self.banner.setText(text)
        self.banner.setVisible(True)
        self.open_folder_btn.setVisible(bool(self._output_folder))

    def _open_folder(self):
        if self._output_folder and os.path.isdir(self._output_folder):
            try:
                os.startfile(self._output_folder)  # Windows
            except Exception as e:
                QMessageBox.critical(self, "Open failed", f"Could not open folder:\n{e}")


class SeparateTaskTab(WizardTaskContainer):
    task_completed = Signal(dict)

    def __init__(self, spec, case_path: str, file_number: str, pdf_path: str, parent=None):
        super().__init__(spec, steps=["Settings", "Analyzing", "Review & Split"], parent=parent)
        self._case_path = case_path
        self._file_number = file_number
        self._pdf_path = pdf_path
        self._worker = None
        self._debug_run_id = None

        self.settings_page = SeparateSettingsPage(pdf_path)
        self.status_page = StatusPage()
        self.workbench_page = SeparateWorkbenchPage()
        self.workbench = self.workbench_page.workbench  # convenience for tests/wiring

        self.addWidget(self.settings_page)    # PAGE_SETTINGS
        self.addWidget(self.status_page)      # PAGE_STATUS
        self.addWidget(self.workbench_page)   # PAGE_WORKBENCH

        self.settings_page.analyze_requested.connect(self._start_analysis)
        self.status_page.cancel_requested.connect(self._on_cancel)
        self.workbench.reanalyze_requested.connect(self._start_analysis)
        self.workbench.processing_complete.connect(self._on_processing_complete)

    @property
    def spec(self):
        return self._spec

    @property
    def files(self) -> list:
        return [self._pdf_path]

    def _start_analysis(self, sensitivity: int):
        if self._worker is not None and self._worker.isRunning():
            return
        # Keep the settings-page slider in sync with whatever value triggered
        # this run (settings Analyze OR workbench Re-analyze) so the recorded
        # task settings reflect the sensitivity actually used.
        self.settings_page.sensitivity_slider.setValue(sensitivity)
        self.status_page.reset()
        self.status_page.progress_bar.setRange(0, 0)
        start_debug_run(
            self,
            source="wizard.separate.analysis",
            details={
                "case_path": self._case_path,
                "file_number": self._file_number,
                "pdf_path": self._pdf_path,
                "sensitivity": sensitivity,
            },
        )
        self.status_page.on_status("Analyzing...")
        record_status(self, "Analyzing...", source="wizard.ui")
        self.setCurrentIndex(PAGE_STATUS)
        worker = SeparateAnalysisWorker(self._pdf_path, sensitivity, parent=None)
        worker.progress.connect(self._on_worker_progress)
        worker.finished_analysis.connect(self._on_analysis_finished)
        worker.finished.connect(worker.deleteLater)
        self._worker = worker
        worker.start()

    def _on_worker_progress(self, message: str) -> None:
        self.status_page.on_status(message)
        record_status(self, message, source="wizard.separate.worker")

    def _on_analysis_finished(self, success: bool, payload: object):
        self._worker = None
        if not success:
            finish_debug_run(
                self,
                status="error",
                message=f"Task failed: {payload}",
                details={"error": str(payload)},
            )
            self.status_page.on_status(f"FAILED: {payload}")
            # Re-enable the workbench controls if this was a re-analyze attempt.
            self.workbench.set_busy(False)
            # Do NOT disconnect StatusPage's own cancel slot — the existing
            # cancel_requested -> _on_cancel wiring already returns to Settings.
            # Just relabel the button; reset() restores it on the next run.
            self.status_page.cancel_btn.setText("Back to Settings")
            self.status_page.cancel_btn.setEnabled(True)
            return
        docs = payload if isinstance(payload, list) else []
        emit_debug(
            self,
            phase="analysis_complete",
            message="Analysis complete",
            source="wizard.separate.worker",
            details={"document_count": len(docs)},
        )
        self.workbench.set_busy(False)
        self.workbench.load_docs(self._pdf_path, docs)
        self.setCurrentIndex(PAGE_WORKBENCH)
        self._persist_to_index(docs)

    def _on_cancel(self):
        emit_debug(
            self,
            phase="cancel",
            message="Back to settings requested",
            source="wizard.ui",
        )
        self.setCurrentIndex(PAGE_SETTINGS)

    def _on_processing_complete(self, summary: dict):
        from datetime import datetime
        self._persist_to_index(_docs_from_workbench(self.workbench))
        finish_debug_run(
            self,
            status="success",
            message="Task complete",
            details=dict(summary or {}),
        )
        self.task_completed.emit({
            "task_id": self._spec.task_id,
            "title": self._spec.title,
            "files": [self._pdf_path],
            "settings": self.settings_page.to_dict(),
            "output_path": summary.get("output_folder", ""),
            "completed_at": datetime.now().isoformat(timespec="seconds"),
        })

    def _persist_to_index(self, docs: list) -> None:
        """Persist this run's document map to the shared per-case index store so
        it is viewable later (Advanced Index tab / wizard reveal). Never fatal."""
        if not self._file_number:
            return
        try:
            case_index_store.upsert_pdf(self._file_number, self._pdf_path, docs)
        except Exception as e:  # persistence must never break the task flow
            try:
                from icharlotte_core.utils import log_event
                log_event(
                    f"[separate] failed to persist index for {self._pdf_path}: {e}",
                    "error",
                )
            except Exception:
                pass

    def closeEvent(self, event):
        if self._worker is not None and self._worker.isRunning():
            QMessageBox.information(
                self, "Task running",
                "Analysis is still running. Wait for it to finish before closing this tab.")
            event.ignore()
            return
        super().closeEvent(event)


def build_separate_tab(spec, case_path: str, file_number: str, parent=None):
    start_dir = resolve_default_folder(case_path, spec.default_folders)
    pdf_path, _ = QFileDialog.getOpenFileName(
        parent, "Select a PDF to separate", start_dir, "PDF files (*.pdf)")
    if not pdf_path:
        return None
    return SeparateTaskTab(
        spec=spec, case_path=case_path, file_number=file_number,
        pdf_path=pdf_path, parent=parent)
