"""MedChronSettingsPage — inline settings for the Med Chron Analysis task.

Mirrors DepositionSettingsPage:
1. Speculative Phase 1 (prep) launched as soon as the tab opens.
2. While extraction runs, a "Preparing chronology…" spinner is shown.
3. On AWAITING_INPUT, the MedChronConfigForm is built and swapped in.
4. User picks analyses + custom rows, clicks Proceed.
5. _on_proceed() validates via form.commit_user_config() then emits
   phase2_requested(session_path).

File swaps:
Med Chron is one-file-per-tab. The Phase 1 cache (narrative + full text +
provider name + file number) is keyed to the input file, so changing the
file on the settings page must invalidate that cache. Add/Remove here
replaces the single slot and emits ``restart_phase1_requested`` so TaskTab
can cancel the old worker and launch a fresh Phase 1 against the new file.
"""

from typing import Optional

from PySide6.QtCore import QSettings, Qt, Signal
from PySide6.QtWidgets import (
    QFileDialog, QLabel, QProgressBar, QSizePolicy, QSplitter,
    QStackedWidget, QVBoxLayout, QWidget,
)

from ..registry import TaskSpec
from .settings_page import SettingsPage


_FILES_SPLITTER_KEY = "med_chron_settings/files_splitter_sizes"


class MedChronSettingsPage(SettingsPage):
    """SettingsPage subclass that embeds MedChronConfigForm inline."""

    phase2_requested = Signal(str)  # carries session_path
    restart_phase1_requested = Signal(list)  # new files list (may be empty → cancel only)

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
        # File that the current Phase 1 cache was built from. Set in
        # attach_worker() and reset on file swap.
        self._prep_file: Optional[str] = self._files[0] if self._files else None

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

        # Pull the files list out of its fixed slot and put it + the stack into
        # a vertical splitter so the user can resize the Files pane against
        # the analyses panes below.
        outer.removeWidget(self.files_list)
        self.files_list.setMaximumHeight(16777215)
        self.files_list.setMinimumHeight(60)

        self._settings = QSettings("iCharlotte", "iCharlotte")
        self._outer_splitter = QSplitter(Qt.Orientation.Vertical)
        self._outer_splitter.setChildrenCollapsible(False)
        self._outer_splitter.setHandleWidth(6)
        self._outer_splitter.setStyleSheet(
            "QSplitter::handle { background: #e0e0e0; }"
            " QSplitter::handle:hover { background: #b0b0b0; }"
            " QSplitter::handle:pressed { background: #909090; }"
        )
        self._outer_splitter.addWidget(self.files_list)
        self._outer_splitter.addWidget(self._stack)
        self._outer_splitter.setStretchFactor(0, 0)
        self._outer_splitter.setStretchFactor(1, 1)
        self._restore_outer_splitter_sizes()
        self._outer_splitter.splitterMoved.connect(self._save_outer_splitter_sizes)

        outer.insertWidget(2, self._outer_splitter, 1)

        self.proceed_btn.setEnabled(False)

    def _restore_outer_splitter_sizes(self) -> None:
        saved = self._settings.value(_FILES_SPLITTER_KEY)
        if saved:
            try:
                sizes = [int(x) for x in saved]
                if len(sizes) == 2 and all(s > 0 for s in sizes):
                    self._outer_splitter.setSizes(sizes)
                    return
            except (TypeError, ValueError):
                pass
        self._outer_splitter.setSizes([150, 600])

    def _save_outer_splitter_sizes(self, *_args) -> None:
        self._settings.setValue(
            _FILES_SPLITTER_KEY, self._outer_splitter.sizes()
        )

    def attach_worker(self, worker) -> bool:
        # Record the file Phase 1 is running against so we can detect swaps.
        if getattr(worker, "files", None):
            self._prep_file = worker.files[0]
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

    # ---- File-swap handling ----

    def _on_add_files(self) -> None:
        """Override base behavior: Med Chron is one-file-per-tab.

        Picking a new file REPLACES the current selection rather than appending.
        If the new file differs from what Phase 1 ran on, restart Phase 1.
        """
        start_dir = self._case_root or ""
        paths, _ = QFileDialog.getOpenFileNames(
            self, "Select medical chronology", start_dir, "All files (*.*)"
        )
        if not paths:
            return
        new_file = paths[0]
        if self._files == [new_file]:
            return  # picked the same file again — nothing to do
        self._files = [new_file]
        self._refresh_files_list()
        self._maybe_restart_phase1()

    def _on_remove_files(self) -> None:
        super()._on_remove_files()
        self._maybe_restart_phase1()

    def _maybe_restart_phase1(self) -> None:
        """If the file slot now differs from the cached Phase 1 file, restart."""
        current_first = self._files[0] if self._files else None
        if current_first == self._prep_file:
            return  # nothing changed
        if current_first is None:
            # File removed but no replacement yet: cancel any in-flight prep
            # and show a holding message. Don't restart Phase 1 with empty input.
            self._reset_to_prep_state(message="Add a file to continue.")
            self.restart_phase1_requested.emit([])
            return
        self._reset_to_prep_state(message="Preparing chronology…")
        self.restart_phase1_requested.emit(list(self._files))

    def _reset_to_prep_state(self, message: str = "") -> None:
        """Tear down the form (if any) and return to the spinner page."""
        if self._form is not None:
            self._form.setParent(None)
            self._form.deleteLater()
            self._form = None
        self._session_path = None
        self._progress.setValue(0)
        self._small_status.setStyleSheet("color: #555; font-size: 11px;")
        self._small_status.setText(message)
        self._stack.setCurrentIndex(0)
        self.proceed_btn.setEnabled(False)
