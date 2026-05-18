"""MedChronSettingsPage — inline settings for the Med Chron Analysis task.

Mirrors DepositionSettingsPage:
1. Speculative Phase 1 (prep) launched as soon as the tab opens.
2. While extraction runs, a "Preparing chronology…" spinner is shown.
3. On AWAITING_INPUT, the MedChronConfigForm is built and swapped in.
4. User picks analyses + custom rows, clicks Proceed.
5. _on_proceed() validates via form.commit_user_config() then emits
   phase2_requested(session_path).
"""

from typing import Optional

from PySide6.QtCore import Signal
from PySide6.QtWidgets import (
    QLabel, QProgressBar, QSizePolicy, QStackedWidget, QVBoxLayout, QWidget,
)

from ..registry import TaskSpec
from .settings_page import SettingsPage


class MedChronSettingsPage(SettingsPage):
    """SettingsPage subclass that embeds MedChronConfigForm inline."""

    phase2_requested = Signal(str)  # carries session_path

    def __init__(
        self,
        spec: TaskSpec,
        files,
        case_root: str | None = None,
        parent: QWidget | None = None,
    ):
        super().__init__(spec, files=files, case_root=case_root, parent=parent)

        self._session_path: Optional[str] = None
        self._form = None

        # Page 0: "Preparing chronology…"
        prep_widget = QWidget()
        prep_layout = QVBoxLayout(prep_widget)
        prep_layout.setContentsMargins(8, 16, 8, 8)
        prep_layout.setSpacing(8)

        title = QLabel("Preparing chronology…")
        title.setStyleSheet("font-size: 13px; font-weight: 600; color: #1976D2;")
        prep_layout.addWidget(title)

        self._progress = QProgressBar()
        self._progress.setRange(0, 100)
        self._progress.setValue(0)
        prep_layout.addWidget(self._progress)

        self._small_status = QLabel("")
        self._small_status.setStyleSheet("color: #555; font-size: 11px;")
        self._small_status.setWordWrap(True)
        prep_layout.addWidget(self._small_status)
        prep_layout.addStretch()

        # Page 1: form placeholder (populated on phase 1 complete)
        self._form_placeholder = QWidget()
        self._form_placeholder_layout = QVBoxLayout(self._form_placeholder)
        self._form_placeholder_layout.setContentsMargins(0, 0, 0, 0)

        # Stack
        self._stack = QStackedWidget()
        self._stack.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding
        )
        self._stack.addWidget(prep_widget)            # index 0
        self._stack.addWidget(self._form_placeholder) # index 1
        self._stack.setCurrentIndex(0)

        outer = self.layout()
        # Remove the "Settings for … to be defined." placeholder body (index 3)
        item = outer.itemAt(3)
        if item is not None:
            w = item.widget()
            if w is not None:
                outer.removeWidget(w)
                w.deleteLater()
        outer.insertWidget(3, self._stack, 1)

        self.proceed_btn.setEnabled(False)

    def attach_worker(self, worker) -> bool:
        worker.status.connect(self._small_status.setText)
        worker.progress.connect(self._progress.setValue)
        worker.awaiting_input.connect(self._on_phase1_complete)
        worker.failed.connect(self._on_phase1_failed)
        return True

    def _on_phase1_complete(self, session_path: str) -> None:
        from icharlotte_core.ui.med_chron_config_form import MedChronConfigForm

        self._session_path = session_path
        try:
            self._form = MedChronConfigForm(session_path, parent=self._form_placeholder)
        except Exception as e:
            self._on_phase1_failed(f"Could not load analysis picker: {e}")
            return
        self._form_placeholder_layout.addWidget(self._form)
        self._stack.setCurrentIndex(1)
        self.proceed_btn.setEnabled(True)

    def _on_phase1_failed(self, err: str) -> None:
        self._small_status.setStyleSheet("color: #c62828; font-size: 11px;")
        self._small_status.setText(f"Preparation failed: {err}")

    def _on_proceed(self) -> None:
        if self._form is None:
            return
        if not self._form.commit_user_config():
            return  # validation failed; form showed error
        self.phase2_requested.emit(self._session_path)
