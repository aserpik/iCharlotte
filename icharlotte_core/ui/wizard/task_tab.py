"""TaskTab — QStackedWidget orchestrating Settings → Status → Output for one task.

Phase 5 wires in the real SubprocessWorker (replaces Phase 4 fake worker).
"""
from typing import List

from PySide6.QtCore import Signal
from PySide6.QtWidgets import QStackedWidget, QWidget

from icharlotte_core.ui.wizard.pages.settings_page import SettingsPage
from icharlotte_core.ui.wizard.pages.status_page import StatusPage
from icharlotte_core.ui.wizard.pages.output_page import OutputPage

PAGE_SETTINGS = 0
PAGE_STATUS = 1
PAGE_OUTPUT = 2


class TaskTab(QStackedWidget):
    """Stateful container for one running task. Owns its own worker."""

    closed = Signal()  # emitted when the tab is being removed
    task_completed = Signal(dict)  # recent-tasks entry dict

    def __init__(
        self,
        spec,
        files: List[str] | None = None,
        case_path: str = "",
        file_number: str = "",
        parent: QWidget | None = None,
    ):
        super().__init__(parent)
        self._spec = spec
        self._files: List[str] = list(files) if files else []
        self._case_path = case_path
        self._file_number = file_number
        self._worker = None
        self._worker_thread = None  # reserved if we move to QThread later

        self.settings_page = SettingsPage(spec, files=self._files)
        self.status_page = StatusPage()
        self.output_page = OutputPage()

        self.addWidget(self.settings_page)  # index 0 = PAGE_SETTINGS
        self.addWidget(self.status_page)    # index 1 = PAGE_STATUS
        self.addWidget(self.output_page)    # index 2 = PAGE_OUTPUT

        self.settings_page.proceed_requested.connect(self._on_proceed)
        self.status_page.cancel_requested.connect(self._on_cancel)
        self.output_page.edit_settings_requested.connect(self._on_edit_settings)
        self.output_page.rerun_requested.connect(self._on_rerun)

    # ---- Public API ----

    @property
    def files(self) -> List[str]:
        return list(self._files)

    @property
    def spec(self):
        return self._spec

    @property
    def current_page(self) -> int:
        return self.currentIndex()

    # ---- Transitions ----

    def _on_proceed(self, settings_dict: dict) -> None:
        self.status_page.reset()
        self.setCurrentIndex(PAGE_STATUS)
        self._start_run(settings_dict)

    def _on_cancel(self) -> None:
        if self._worker is not None:
            self._worker.cancel()
        else:
            self.setCurrentIndex(PAGE_SETTINGS)

    def _on_edit_settings(self) -> None:
        self.setCurrentIndex(PAGE_SETTINGS)

    def _on_rerun(self) -> None:
        self._on_proceed(self.settings_page.to_dict())

    def _show_output(self, output_path: str) -> None:
        self.output_page.load_output(output_path)
        self.setCurrentIndex(PAGE_OUTPUT)

    # ---- Worker (Phase 5 real subprocess) ----

    def _start_run(self, settings_dict: dict) -> None:
        from .runners.subprocess_worker import SubprocessWorker

        self.status_page.on_status(f"Starting {self._spec.title}…")
        self._worker = SubprocessWorker(
            script_name=self._spec.script_name,
            case_path=self._case_path,
            file_number=self._file_number,
            files=self._files,
            settings=settings_dict,
            parent=self,
        )
        self._worker.status.connect(self.status_page.on_status)
        self._worker.progress.connect(self.status_page.on_progress)
        self._worker.finished.connect(self._on_worker_finished)
        self._worker.failed.connect(self._on_worker_failed)
        self._worker.cancelled.connect(self._on_worker_cancelled)
        self._worker.start()

    def _on_worker_finished(self, output_path: str) -> None:
        from datetime import datetime
        self._worker = None
        entry = {
            "task_id": self._spec.task_id,
            "title": self._spec.title,
            "files": list(self._files),
            "settings": self.settings_page.to_dict(),
            "output_path": output_path,
            "completed_at": datetime.now().isoformat(timespec="seconds"),
        }
        self.task_completed.emit(entry)
        self._show_output(output_path)

    def _on_worker_failed(self, err: str) -> None:
        self._worker = None
        self.status_page.on_status(f"FAILED: {err}")
        self.status_page.cancel_btn.setText("Back to Settings")
        self.status_page.cancel_btn.setEnabled(True)
        try:
            self.status_page.cancel_btn.clicked.disconnect()
        except RuntimeError:
            pass
        self.status_page.cancel_btn.clicked.connect(lambda: self.setCurrentIndex(PAGE_SETTINGS))

    def _on_worker_cancelled(self) -> None:
        self._worker = None
        self.setCurrentIndex(PAGE_SETTINGS)
