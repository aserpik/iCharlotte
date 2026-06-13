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

from PySide6.QtCore import Qt, QSettings, Signal
from PySide6.QtWidgets import (
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QProgressBar,
    QPushButton,
    QSizePolicy,
    QSplitter,
    QStackedWidget,
    QVBoxLayout,
    QWidget,
)

from ..file_picker import resolve_default_folder
from ..registry import TaskSpec
from .settings_page import SettingsPage

_OUTER_KEY = "wizard/depo_settings/outer_splitter_state"


class DepositionSettingsPage(SettingsPage):
    """SettingsPage subclass that embeds DepoSummaryConfigForm inline.

    Phase 1 runs speculatively; the form appears once it completes.

    Layout
    ------
    The page is split into two panes by a vertical QSplitter:
      - Pane A (top):  Files box — label, Add Files / Remove buttons, files list.
      - Pane B (bottom): Phase indicator / config form (QStackedWidget).

    Both panes are re-parented widgets that SettingsPage.__init__ already
    created; the base-class references (files_label, files_list, add_files_btn,
    remove_btn, proceed_btn) still work unchanged.
    """

    # Emitted when the user has committed config and Phase 2 should start.
    # Carries the session_path string.
    phase2_requested = Signal(str)
    restart_phase1_requested = Signal(list)

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
        self._prep_file: Optional[str] = self._files[0] if self._files else None

        # ------------------------------------------------------------------ #
        # Strip everything from the base-class layout so we can rebuild it
        # with the splitter in the middle.
        # ------------------------------------------------------------------ #
        outer = self.layout()

        # The base layout order is:
        #   0: files_label
        #   1: file_btn_row (QHBoxLayout)
        #   2: files_list
        #   3: body placeholder label
        #   4: btn_row (QHBoxLayout, contains proceed_btn)
        #
        # We want to re-parent files_label, file_btn_row, files_list into a
        # "files section" widget, build the stacked widget for the configure
        # area, put both inside an outer QSplitter, and keep the Proceed row
        # (btn_row) at the very bottom outside the splitter.

        # 1. Remove everything from the outer layout (we rebuild it below).
        #    Collect the items we care about before clearing.
        files_label_widget = self.files_label          # QLabel
        files_list_widget = self.files_list            # QListWidget
        add_files_btn = self.add_files_btn             # QPushButton
        remove_btn = self.remove_btn                   # QPushButton
        proceed_btn = self.proceed_btn                 # QPushButton

        # Clear the layout items (do NOT deleteLater on the widgets — we reuse them).
        while outer.count():
            item = outer.takeAt(0)
            if item.widget():
                item.widget().setParent(None)
            # layouts (like file_btn_row) are just removed from outer; their
            # child widgets will be re-parented below.

        # 2. Build the files section widget.
        files_section = QWidget()
        files_layout = QVBoxLayout(files_section)
        files_layout.setContentsMargins(0, 4, 0, 4)
        files_layout.setSpacing(2)

        files_label_widget.setParent(files_section)
        files_layout.addWidget(files_label_widget)

        file_btn_row = QHBoxLayout()
        file_btn_row.setContentsMargins(0, 0, 0, 0)
        file_btn_row.setSpacing(8)
        add_files_btn.setParent(files_section)
        file_btn_row.addWidget(add_files_btn)
        remove_btn.setParent(files_section)
        file_btn_row.addWidget(remove_btn)
        file_btn_row.addStretch()
        files_layout.addLayout(file_btn_row)

        files_list_widget.setMaximumHeight(16777215)  # clear the 150px cap set by base
        files_list_widget.setParent(files_section)
        files_layout.addWidget(files_list_widget)

        # 3. Build the stacked widget (configure area, Pane B).

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
        self._stack.addWidget(discovering_widget)      # index 0: discovering
        self._stack.addWidget(self._form_placeholder)  # index 1: config form
        self._stack.setCurrentIndex(0)

        # 4. Outer splitter.
        self.outer_splitter = QSplitter(Qt.Orientation.Vertical)
        self.outer_splitter.setObjectName("outer_splitter")
        self.outer_splitter.setChildrenCollapsible(False)
        self.outer_splitter.addWidget(files_section)
        self.outer_splitter.addWidget(self._stack)

        # 5. Proceed row (re-create; proceed_btn already exists from base).
        proceed_btn.setParent(self)
        btn_row_widget = QWidget()
        btn_row_layout = QHBoxLayout(btn_row_widget)
        btn_row_layout.setContentsMargins(0, 0, 0, 0)
        btn_row_layout.addStretch()
        btn_row_layout.addWidget(proceed_btn)

        # 6. Rebuild the outer layout: margins, splitter, proceed row.
        outer.setContentsMargins(24, 24, 24, 24)
        outer.setSpacing(16)
        outer.addWidget(self.outer_splitter, 1)
        outer.addWidget(btn_row_widget)

        # 7. Restore splitter state (or set defaults).
        self._restore_outer_splitter()
        self.outer_splitter.splitterMoved.connect(self._save_outer_splitter)

        # Proceed starts disabled — enabled after Phase 1 completes.
        self.proceed_btn.setEnabled(False)
        if not self._files:
            self._reset_to_discovering_state("Add a transcript to continue.")

    # ---- Splitter persistence ----

    def _restore_outer_splitter(self):
        settings = QSettings("iCharlotte", "iCharlotte")
        state = settings.value(_OUTER_KEY)
        if state:
            self.outer_splitter.restoreState(state)
        else:
            # First-run default: files section ~150 px, rest of space to form.
            self.outer_splitter.setSizes([150, 400])

    def _save_outer_splitter(self):
        settings = QSettings("iCharlotte", "iCharlotte")
        settings.setValue(_OUTER_KEY, self.outer_splitter.saveState())

    # ---- attach_worker override ----

    def attach_worker(self, worker) -> bool:
        """Wire Phase 1 status/progress/awaiting_input to our inline UI."""
        if getattr(worker, "files", None):
            self._prep_file = worker.files[0]
        worker.status.connect(self._small_status.setText)
        worker.progress.connect(self._progress.setValue)
        worker.awaiting_input.connect(self._on_phase1_complete)
        worker.failed.connect(self._on_phase1_failed)
        return True

    # ---- File-swap handling ----

    def _on_add_files(self) -> None:
        """Pick one transcript for this tab and start/restart Phase 1.

        The underlying interactive deposition script is one-transcript per
        session, so replacing the slot keeps the existing per-tab behavior from
        the old launcher while allowing the user to choose the file on Settings.
        """
        start_dir = ""
        if self._case_root:
            start_dir = resolve_default_folder(
                self._case_root, self._spec.default_folders
            )
        paths, _ = QFileDialog.getOpenFileNames(
            self, "Select deposition transcript", start_dir, "All files (*.*)"
        )
        self.add_files(paths)

    def add_files(self, paths: list[str]) -> None:
        if not paths:
            return
        new_file = paths[0]
        if self._files == [new_file]:
            return
        self._files = [new_file]
        self._refresh_files_list()
        self._maybe_restart_phase1()

    def _on_remove_files(self) -> None:
        super()._on_remove_files()
        self._maybe_restart_phase1()

    def _maybe_restart_phase1(self) -> None:
        current_first = self._files[0] if self._files else None
        if current_first == self._prep_file:
            return
        if current_first is None:
            self._reset_to_discovering_state("Add a transcript to continue.")
            self.restart_phase1_requested.emit([])
            return
        self._reset_to_discovering_state("Discovering topics...")
        self.restart_phase1_requested.emit(list(self._files))

    def _reset_to_discovering_state(self, message: str = "") -> None:
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
