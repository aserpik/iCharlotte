"""TaskTab — QStackedWidget orchestrating Settings → Status → Output for one task.

Phase 4 ships with a 'fake worker' that just sleeps and emits a synthetic
output path. Phase 5 replaces that with real subprocess-based runners.
"""
from typing import List

from PySide6.QtCore import QTimer, Signal
from PySide6.QtWidgets import QStackedWidget

from icharlotte_core.ui.wizard.pages.settings_page import SettingsPage
from icharlotte_core.ui.wizard.pages.status_page import StatusPage
from icharlotte_core.ui.wizard.pages.output_page import OutputPage

PAGE_SETTINGS = 0
PAGE_STATUS = 1
PAGE_OUTPUT = 2


class TaskTab(QStackedWidget):
    """Stateful container for one running task. Owns its own worker."""

    closed = Signal()  # emitted when the tab is being removed

    def __init__(
        self,
        spec,
        files=None,
        parent=None,
    ):
        super().__init__(parent)
        self._spec = spec
        self._files: List[str] = list(files) if files else []
        self._worker = None
        self._fake_worker_delay_ms = 2000  # Phase 4 fake-run duration

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
    def current_page(self) -> int:
        return self.currentIndex()

    # ---- Transitions ----

    def _on_proceed(self, settings_dict: dict) -> None:
        self.status_page.reset()
        self.setCurrentIndex(PAGE_STATUS)
        self._start_run(settings_dict)

    def _on_cancel(self) -> None:
        if self._worker is not None and hasattr(self._worker, "cancel"):
            self._worker.cancel()
        # Phase 4 fake worker has no cancel — just snap back to Settings.
        self._worker = None
        self.setCurrentIndex(PAGE_SETTINGS)

    def _on_edit_settings(self) -> None:
        self.setCurrentIndex(PAGE_SETTINGS)

    def _on_rerun(self) -> None:
        self._on_proceed(self.settings_page.to_dict())

    def _show_output(self, output_path: str) -> None:
        self.output_page.load_output(output_path)
        self.setCurrentIndex(PAGE_OUTPUT)

    # ---- Worker (Phase 4 fake) ----

    def _start_run(self, settings_dict: dict) -> None:
        self.status_page.on_status(f"Running {self._spec.title}…")
        self.status_page.on_status(f"Inputs: {len(self._files)} file(s)")
        # Phase 4 fake: after a short delay, "finish" with a stub path.
        delay = max(0, self._fake_worker_delay_ms)
        if self._files:
            stub_output = self._files[0]  # not a real .docx; replaced in Phase 5
        else:
            stub_output = ""
        QTimer.singleShot(delay, lambda: self._show_output(stub_output))
