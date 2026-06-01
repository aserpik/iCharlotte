"""InProcessTaskTab — Settings → Status → Output flow for in-process QThread workers.

Used by wizard tasks whose backend is an in-process QThread rather than a
subprocess script (Subpoena Tracker, Med Record Extractor, Respond to Discovery,
and custom task builders). Worker contract:

    progress = Signal(str)
    warning = Signal(str)
    finished_result = Signal(bool, str)  # (success, output_path_or_summary)
"""
from __future__ import annotations

import os
from typing import Callable, Optional

from PySide6.QtCore import QThread, Signal
from PySide6.QtWidgets import (
    QFileDialog,
    QHBoxLayout,
    QInputDialog,
    QLabel,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from icharlotte_core.ui.wizard import theme
from icharlotte_core.ui.wizard.pages.status_page import StatusPage
from icharlotte_core.ui.wizard.pages.output_page import OutputPage
from icharlotte_core.ui.wizard.task_scaffold import WizardTaskContainer


PAGE_SETTINGS = 0
PAGE_STATUS = 1
PAGE_OUTPUT = 2


# WorkerFactory signature: (case_path, file_number, settings_dict, parent) -> QThread
WorkerFactory = Callable[[str, str, dict, Optional[QWidget]], QThread]


class InProcessTaskTab(WizardTaskContainer):
    """Settings → Status → Output flow backed by an in-process QThread worker."""

    task_completed = Signal(dict)  # recent-tasks entry dict (parity with TaskTab)

    def __init__(
        self,
        spec,
        case_path: str,
        file_number: str,
        settings_widget: QWidget,
        output_widget: QWidget,
        worker_factory: WorkerFactory,
        auto_run: bool = False,
        parent: QWidget | None = None,
    ):
        super().__init__(spec, parent=parent)
        self._case_path = case_path
        self._file_number = file_number
        self._worker_factory = worker_factory
        self._worker: QThread | None = None

        self.settings_page = settings_widget
        self.status_page = StatusPage()
        self.output_page = output_widget

        self.addWidget(self.settings_page)  # 0
        self.addWidget(self.status_page)    # 1
        self.addWidget(self.output_page)    # 2

        # Settings widget contract: emits run_requested(dict).
        self.settings_page.run_requested.connect(self._on_run)
        self.status_page.cancel_requested.connect(self._on_cancel)
        # Output widget contract: emits rerun_requested() and edit_settings_requested().
        if hasattr(self.output_page, "rerun_requested"):
            self.output_page.rerun_requested.connect(self._on_rerun)
        if hasattr(self.output_page, "edit_settings_requested"):
            self.output_page.edit_settings_requested.connect(self._on_edit_settings)

        if auto_run:
            # Defer until after the widget is parented so log messages don't
            # spam during construction.
            from PySide6.QtCore import QTimer
            QTimer.singleShot(0, lambda: self._on_run({}))

    # ---- Public API parity with TaskTab ----

    @property
    def spec(self):
        return self._spec

    @property
    def files(self) -> list[str]:
        return []

    # ---- Transitions ----

    def _on_run(self, settings_dict: dict) -> None:
        self.status_page.reset()
        self.setCurrentIndex(PAGE_STATUS)
        self.status_page.on_status(f"Starting {self._spec.title}…")
        # In-process workers don't report % — keep the bar indeterminate.
        self.status_page.progress_bar.setRange(0, 0)
        self._last_settings = dict(settings_dict)
        self._start_worker(settings_dict)

    def _on_cancel(self) -> None:
        # In-process workers (SubpoenaTrackerWorker, MedRecordExtractorWorker)
        # don't expose a cooperative cancel. If the run hasn't started, just
        # bounce back to settings; otherwise wait it out.
        if self._worker is None:
            self.setCurrentIndex(PAGE_SETTINGS)
            return
        # Best-effort: call cancel() if the worker supports it.
        cancel = getattr(self._worker, "cancel", None)
        if callable(cancel):
            try:
                cancel()
            except Exception:
                pass
        else:
            QMessageBox.information(
                self,
                "Cannot cancel",
                "This task does not support cancellation. It will continue in the background.",
            )

    def _on_rerun(self) -> None:
        self._on_run(getattr(self, "_last_settings", {}))

    def _on_edit_settings(self) -> None:
        self.setCurrentIndex(PAGE_SETTINGS)

    # ---- Worker plumbing ----

    def _start_worker(self, settings_dict: dict) -> None:
        worker = self._worker_factory(
            self._case_path, self._file_number, settings_dict, self
        )
        worker.progress.connect(self.status_page.on_status)
        worker.warning.connect(lambda msg: self.status_page.on_status(f"WARNING: {msg}"))
        worker.finished_result.connect(self._on_worker_finished)
        self._worker = worker
        worker.start()

    def _on_worker_finished(self, success: bool, output_or_error: str) -> None:
        from datetime import datetime
        worker = self._worker
        self._worker = None

        if not success:
            self.status_page.on_status(f"FAILED: {output_or_error}")
            self.status_page.cancel_btn.setText("Back to Settings")
            self.status_page.cancel_btn.setEnabled(True)
            try:
                self.status_page.cancel_btn.clicked.disconnect()
            except RuntimeError:
                pass
            self.status_page.cancel_btn.clicked.connect(
                lambda: self.setCurrentIndex(PAGE_SETTINGS)
            )
            return

        settings = getattr(self, "_last_settings", {}) or {}
        if hasattr(self.output_page, "set_save_as_defaults"):
            default_dir = settings.get("save_default_dir", "")
            suggested_filename = settings.get("suggested_filename", "")
            if default_dir or suggested_filename:
                self.output_page.set_save_as_defaults(
                    default_dir=default_dir,
                    suggested_filename=suggested_filename,
                    required=True,
                )

        # Hand the result to the output page in whichever way it accepts.
        if hasattr(self.output_page, "show_result"):
            self.output_page.show_result(output_or_error)
        elif hasattr(self.output_page, "load_output"):
            self.output_page.load_output(output_or_error)
        self.setCurrentIndex(PAGE_OUTPUT)

        entry = {
            "task_id": self._spec.task_id,
            "title": self._spec.title,
            "files": [],
            "settings": getattr(self, "_last_settings", {}),
            "output_path": output_or_error if os.path.isfile(output_or_error or "") else "",
            "completed_at": datetime.now().isoformat(timespec="seconds"),
        }
        self.task_completed.emit(entry)


# ----------------------------- Settings widgets -----------------------------


class _RunnableSettingsPage(QWidget):
    """Base for in-process settings widgets — emits run_requested(dict)."""

    run_requested = Signal(dict)

    def to_dict(self) -> dict:
        return {}

    def from_dict(self, data: dict) -> None:
        return


class SubpoenaSettingsPage(_RunnableSettingsPage):
    """No inputs needed — single 'Run' button kicks off the scan."""

    def __init__(self, parent: QWidget | None = None):
        super().__init__(parent)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 24, 24, 24)
        layout.setSpacing(12)

        intro = theme.helper_text(
            "The Subpoena Tracker scans the case folder for issued subpoenas,"
            " received records, and medical chronologies, then produces a"
            " consolidated tracking report in NOTES/AI OUTPUT."
        )
        layout.addWidget(intro)

        layout.addStretch(1)

        btn_row = QHBoxLayout()
        btn_row.addStretch()
        self.run_btn = theme.primary_button("Run Subpoena Tracker")
        self.run_btn.clicked.connect(lambda: self.run_requested.emit({}))
        btn_row.addWidget(self.run_btn)
        layout.addLayout(btn_row)


class MedExtractorSettingsPage(_RunnableSettingsPage):
    """Prose text input + 'Extract' button."""

    def __init__(self, parent: QWidget | None = None):
        super().__init__(parent)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 24, 24, 24)
        layout.setSpacing(12)

        intro = theme.helper_text(
            "Paste medical chronology entries below. Each entry should include"
            " a date of treatment and provider/facility name.  Example: “On"
            " March 22, 2016, Plaintiff was evaluated by Vikram M Shaker, MD at"
            " Adventist Health for a CT of the abdomen…”"
        )
        layout.addWidget(intro)

        self.text_input = QPlainTextEdit()
        self.text_input.setPlaceholderText(
            "Paste entries here (one or more paragraphs)..."
        )
        self.text_input.setMinimumHeight(250)
        layout.addWidget(self.text_input, 1)

        btn_row = QHBoxLayout()
        btn_row.addStretch()
        self.extract_btn = theme.primary_button("Extract")
        self.extract_btn.clicked.connect(self._on_extract)
        btn_row.addWidget(self.extract_btn)
        layout.addLayout(btn_row)

    def _on_extract(self) -> None:
        text = self.text_input.toPlainText().strip()
        if not text:
            QMessageBox.warning(self, "Empty", "Please paste at least one entry.")
            return
        self.run_requested.emit({"text": text})


# ----------------------------- Output widgets -----------------------------


class MedExtractorOutputPage(QWidget):
    """Shows the extraction summary + an 'Open Output Folder' button."""

    rerun_requested = Signal()
    edit_settings_requested = Signal()

    def __init__(self, case_path: str, parent: QWidget | None = None):
        super().__init__(parent)
        self._case_path = case_path
        self._output_dir = os.path.join(case_path, "NOTES", "AI OUTPUT", "Med Record Extracts")

        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 24, 24, 24)
        layout.setSpacing(12)

        header = theme.page_title("Extraction complete")
        layout.addWidget(header)

        self.summary_view = QPlainTextEdit()
        self.summary_view.setReadOnly(True)
        self.summary_view.setStyleSheet(
            f"font-family: {theme.MONO}; font-size: {theme.FONT_BODY}px;"
        )
        layout.addWidget(self.summary_view, 1)

        btn_row = QHBoxLayout()
        self.open_folder_btn = theme.secondary_button("Open Output Folder")
        self.open_folder_btn.clicked.connect(self._on_open_folder)
        btn_row.addWidget(self.open_folder_btn)
        btn_row.addStretch()
        self.rerun_btn = theme.secondary_button("Run Again")
        self.rerun_btn.clicked.connect(self.rerun_requested.emit)
        btn_row.addWidget(self.rerun_btn)
        self.edit_settings_btn = theme.secondary_button("Edit Entries & Re-run")
        self.edit_settings_btn.clicked.connect(self.edit_settings_requested.emit)
        btn_row.addWidget(self.edit_settings_btn)
        layout.addLayout(btn_row)

    def show_result(self, summary_text: str) -> None:
        self.summary_view.setPlainText(summary_text or "(no summary returned)")

    def _on_open_folder(self) -> None:
        if not os.path.isdir(self._output_dir):
            QMessageBox.information(
                self,
                "Folder not found",
                f"Output folder does not exist yet:\n{self._output_dir}",
            )
            return
        try:
            os.startfile(self._output_dir)  # Windows
        except Exception as e:
            QMessageBox.critical(self, "Open failed", f"Could not open folder:\n{e}")


def _make_subpoena_output_page() -> OutputPage:
    """Build an OutputPage with the inputs/edit-settings buttons hidden.

    The base OutputPage exposes 'Re-run' and 'Edit Settings & Re-run', which is
    fine — InProcessTaskTab will wire them.
    """
    return OutputPage()


# ----------------------------- Factories used by iCharlotte.py -----------------------------


def build_subpoena_tab(spec, case_path: str, file_number: str, parent: QWidget | None) -> InProcessTaskTab:
    from icharlotte_core.subpoena_tracker import SubpoenaTrackerWorker

    def factory(cp, fn, settings, p):
        return SubpoenaTrackerWorker(cp, file_number=fn, parent=p)

    return InProcessTaskTab(
        spec=spec,
        case_path=case_path,
        file_number=file_number,
        settings_widget=SubpoenaSettingsPage(),
        output_widget=_make_subpoena_output_page(),
        worker_factory=factory,
        auto_run=True,
        parent=parent,
    )


def build_med_extractor_tab(spec, case_path: str, file_number: str, parent: QWidget | None) -> InProcessTaskTab:
    from icharlotte_core.med_record_extractor import MedRecordExtractorWorker

    def factory(cp, fn, settings, p):
        return MedRecordExtractorWorker(cp, fn, settings.get("text", ""), parent=p)

    return InProcessTaskTab(
        spec=spec,
        case_path=case_path,
        file_number=file_number,
        settings_widget=MedExtractorSettingsPage(),
        output_widget=MedExtractorOutputPage(case_path),
        worker_factory=factory,
        auto_run=False,
        parent=parent,
    )


def build_respond_to_discovery_tab(
    spec,
    case_path: str,
    file_number: str,
    parent: QWidget | None,
) -> InProcessTaskTab | None:
    from icharlotte_core.discovery.response_type_detector import (
        detect_type_from_filename,
        detect_type_from_text,
        resolve_detected_type,
    )
    from icharlotte_core.ui.wizard.file_picker import resolve_default_folder
    from icharlotte_core.ui.wizard.pages.respond_discovery_page import (
        RespondDiscoverySettingsPage,
        RespondDiscoveryWorker,
        read_first_page_text,
    )

    start_dir = resolve_default_folder(case_path, spec.default_folders)
    file_path, _ = QFileDialog.getOpenFileName(
        parent,
        "Select discovery to respond to",
        start_dir,
        "PDF files (*.pdf)",
    )
    if not file_path:
        return None

    filename_guess = detect_type_from_filename(file_path)
    text_guess = detect_type_from_text(read_first_page_text(file_path))
    resolution = resolve_detected_type(filename_guess, text_guess)
    detected_type = resolution.discovery_type
    if resolution.needs_user_choice:
        detected_type, ok = QInputDialog.getItem(
            parent,
            "Discovery Type",
            "Select the type of discovery:",
            ["FI", "SI", "RFA", "RPD"],
            1,
            False,
        )
        if not ok or not detected_type:
            return None

    def factory(cp, fn, settings, p):
        return RespondDiscoveryWorker(cp, fn, settings, parent=p)

    return InProcessTaskTab(
        spec=spec,
        case_path=case_path,
        file_number=file_number,
        settings_widget=RespondDiscoverySettingsPage(
            case_root=case_path,
            file_number=file_number,
            discovery_file=file_path,
            detected_type=detected_type or "SI",
        ),
        output_widget=_make_subpoena_output_page(),
        worker_factory=factory,
        auto_run=False,
        parent=parent,
    )


def build_oppose_motion_tab(
    spec,
    case_path: str,
    file_number: str,
    parent: QWidget | None,
):
    from icharlotte_core.ui.wizard.pages.oppose_motion_page import (
        build_oppose_motion_tab as _build,
    )

    return _build(
        spec=spec,
        case_path=case_path,
        file_number=file_number,
        parent=parent,
    )
