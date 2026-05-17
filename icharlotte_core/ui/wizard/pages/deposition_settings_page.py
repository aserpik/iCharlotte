"""DepositionSettingsPage — inline settings for the Summarize Depositions task.

Embeds DepoSummaryConfigForm directly on the Settings page and runs
Phase 1 (topic discovery) speculatively in the background as soon as
the tab opens.  By the time the user finishes reviewing or adjusting
the topic list, Phase 1 is typically already done.

Flow
----
1. TaskTab calls start_speculative_run() → SubprocessWorker is started.
2. Worker emits status/progress here (discovering label + progress bar).
3. When worker emits awaiting_input(session_path), we build the
   DepoSummaryConfigForm, swap the stack to it, and enable Proceed.
4. User adjusts topics / settings and clicks Proceed.
5. _on_proceed() calls form.commit_user_config(), then emits
   phase2_requested(session_path) → TaskTab.advance_to_status_with_phase2.
"""

from typing import Optional

from PySide6.QtCore import Signal
from PySide6.QtWidgets import (
    QLabel,
    QProgressBar,
    QSizePolicy,
    QStackedWidget,
    QVBoxLayout,
    QWidget,
)

from ..registry import TaskSpec
from .settings_page import SettingsPage


class DepositionSettingsPage(SettingsPage):
    """SettingsPage subclass that embeds DepoSummaryConfigForm inline.

    Phase 1 runs speculatively; the form appears once it completes.
    """

    # Emitted when the user has committed config and Phase 2 should start.
    # Carries the session_path string.
    phase2_requested = Signal(str)

    def __init__(
        self,
        spec: TaskSpec,
        files,
        case_root: str | None = None,
        parent: QWidget | None = None,
    ):
        super().__init__(spec, files=files, case_root=case_root, parent=parent)

        self._session_path: Optional[str] = None
        self._form = None  # DepoSummaryConfigForm; created on phase1 complete

        # --- Build the stacked widget that sits between the file list and
        #     the Proceed row ---

        # Page 0: discovering spinner
        discovering_widget = QWidget()
        discover_layout = QVBoxLayout(discovering_widget)
        discover_layout.setContentsMargins(8, 16, 8, 8)
        discover_layout.setSpacing(8)

        discovering_title = QLabel("Discovering topics…")
        discovering_title.setStyleSheet("font-size: 13px; font-weight: 600; color: #1976D2;")
        discover_layout.addWidget(discovering_title)

        self._progress = QProgressBar()
        self._progress.setRange(0, 100)
        self._progress.setValue(0)
        discover_layout.addWidget(self._progress)

        self._small_status = QLabel("")
        self._small_status.setStyleSheet("color: #555; font-size: 11px;")
        self._small_status.setWordWrap(True)
        discover_layout.addWidget(self._small_status)

        discover_layout.addStretch()

        # Page 1: config form placeholder (populated on phase1 complete)
        self._form_placeholder = QWidget()
        self._form_placeholder_layout = QVBoxLayout(self._form_placeholder)
        self._form_placeholder_layout.setContentsMargins(0, 0, 0, 0)

        # Assemble stacked widget
        self._stack = QStackedWidget()
        self._stack.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding
        )
        self._stack.addWidget(discovering_widget)    # index 0: discovering
        self._stack.addWidget(self._form_placeholder)  # index 1: config form
        self._stack.setCurrentIndex(0)

        # Insert the stack BEFORE the Proceed row.
        # SettingsPage.__init__ builds: files_label, btn_row, files_list,
        # body (stretch), proceed_row — so the stack goes after files_list
        # but before the Proceed row.
        # We find and remove the placeholder body label and insert stack instead.
        outer = self.layout()
        # Remove the "Settings for … to be defined." placeholder body (item at index 3)
        item = outer.itemAt(3)
        if item is not None:
            w = item.widget()
            if w is not None:
                outer.removeWidget(w)
                w.deleteLater()
        # Insert the stack at position 3 (before the Proceed row which is last)
        outer.insertWidget(3, self._stack, 1)

        # Proceed starts disabled — enabled after Phase 1 completes.
        self.proceed_btn.setEnabled(False)

    # ---- attach_worker override ----

    def attach_worker(self, worker) -> bool:
        """Wire Phase 1 status/progress/awaiting_input to our inline UI."""
        worker.status.connect(self._small_status.setText)
        worker.progress.connect(self._progress.setValue)
        worker.awaiting_input.connect(self._on_phase1_complete)
        worker.failed.connect(self._on_phase1_failed)
        return True

    # ---- Phase 1 completion ----

    def _on_phase1_complete(self, session_path: str) -> None:
        """Phase 1 done — build the config form and show it."""
        from icharlotte_core.ui.depo_summary_config_form import DepoSummaryConfigForm

        self._session_path = session_path

        # Build the form and add it to the placeholder container.
        try:
            self._form = DepoSummaryConfigForm(session_path, parent=self._form_placeholder)
        except Exception as e:
            self._on_phase1_failed(f"Could not load topic config: {e}")
            return

        self._form_placeholder_layout.addWidget(self._form)

        # Switch stack to the config form.
        self._stack.setCurrentIndex(1)

        # Enable Proceed now that the user can configure topics.
        self.proceed_btn.setEnabled(True)

    def _on_phase1_failed(self, err: str) -> None:
        """Phase 1 failed — show error on the discovering page."""
        self._small_status.setStyleSheet("color: #c62828; font-size: 11px;")
        self._small_status.setText(f"Phase 1 failed: {err}")
        # Proceed stays disabled.

    # ---- Override _on_proceed ----

    def _on_proceed(self) -> None:
        """Commit config and emit phase2_requested instead of the normal proceed_requested."""
        if self._form is None:
            return
        if not self._form.commit_user_config():
            return  # validation failed; form showed error
        self.phase2_requested.emit(self._session_path)
