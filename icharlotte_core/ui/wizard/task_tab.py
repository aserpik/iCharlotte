"""TaskTab — QStackedWidget orchestrating Settings → Status → Output for one task."""
from typing import List, Optional

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
        self._awaiting_session_path: Optional[str] = None
        self._settings_owns_worker = False

        self.settings_page = spec.settings_page_cls(spec, files=self._files, case_root=case_path or None)
        self.status_page = StatusPage()
        self.output_page = OutputPage()

        self.addWidget(self.settings_page)  # index 0 = PAGE_SETTINGS
        self.addWidget(self.status_page)    # index 1 = PAGE_STATUS
        self.addWidget(self.output_page)    # index 2 = PAGE_OUTPUT

        self.settings_page.proceed_requested.connect(self._on_proceed)
        self.status_page.cancel_requested.connect(self._on_cancel)
        self.status_page.configure_requested.connect(self._on_configure_requested)
        self.output_page.edit_settings_requested.connect(self._on_edit_settings)
        self.output_page.rerun_requested.connect(self._on_rerun)

        # Opportunistically wire phase2_requested if the settings page exposes it.
        if hasattr(self.settings_page, "phase2_requested"):
            self.settings_page.phase2_requested.connect(self.advance_to_status_with_phase2)

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

    # ---- Speculative run (for SettingsPage subclasses that own the worker) ----

    def start_speculative_run(self) -> None:
        """Launch the worker immediately without switching pages.

        The settings page takes ownership via attach_worker() and drives the
        status/progress/awaiting_input signals itself.  TaskTab only wires the
        terminal signals (finished/failed/cancelled).
        """
        from .runners.subprocess_worker import SubprocessWorker

        self._settings_owns_worker = True
        self._awaiting_session_path = None

        worker = SubprocessWorker(
            script_name=self._spec.script_name,
            case_path=self._case_path,
            file_number=self._file_number,
            files=self._files,
            settings={},
            phase1_args=list(self._spec.phase1_args),
            phase2_flag=self._spec.phase2_flag,
            parent=self,
        )

        # Let the settings page own the mid-flight signals.
        self.settings_page.attach_worker(worker)

        # TaskTab only handles terminal signals.
        worker.finished.connect(self._on_worker_finished)
        worker.failed.connect(self._on_worker_failed)
        worker.cancelled.connect(self._on_worker_cancelled)

        self._worker = worker
        worker.start()

    def advance_to_status_with_phase2(self, session_path: str) -> None:
        """Switch to the Status page and kick off Phase 2.

        Called via the phase2_requested signal emitted by DepositionSettingsPage
        after the user commits their topic configuration. Handles two cases:

        - Initial speculative-run path: the worker started for Phase 1 is still
          alive; we just connect status signals and call resume_with_config.
        - Re-run path: _on_worker_finished cleared self._worker after the
          previous Phase 2 completed, but the session directory and the
          settings page's inline form are still intact. We build a fresh
          SubprocessWorker that skips Phase 1 and resumes against the
          existing session_path.
        """
        if self._worker is None:
            from .runners.subprocess_worker import SubprocessWorker

            self._worker = SubprocessWorker(
                script_name=self._spec.script_name,
                case_path=self._case_path,
                file_number=self._file_number,
                files=self._files,
                settings={},
                phase1_args=list(self._spec.phase1_args),
                phase2_flag=self._spec.phase2_flag,
                parent=self,
            )
            self._worker.finished.connect(self._on_worker_finished)
            self._worker.failed.connect(self._on_worker_failed)
            self._worker.cancelled.connect(self._on_worker_cancelled)

        self._settings_owns_worker = False
        self.status_page.reset()
        self.setCurrentIndex(PAGE_STATUS)

        # Wire the status/progress signals now that status page is visible.
        self._worker.status.connect(self.status_page.on_status)
        self._worker.progress.connect(self.status_page.on_progress)

        self.status_page.on_status("Running Phase 2 — generating summary…")
        self.status_page.progress_bar.setRange(0, 0)
        self._worker.resume_with_config(session_path)

    # ---- Worker ----

    # Scripts that ship their own internal multi-file dispatcher (party
    # grouping, consolidation passes, etc.) and would have those behaviors
    # broken by the generic per-file ParallelSubprocessWorker.
    _DISPATCHER_SCRIPTS = frozenset({"summarize_discovery.py"})

    @staticmethod
    def _pick_worker_cls(script_name: str, num_files: int) -> type:
        """Decide which runner class to use for ``script_name`` and
        ``num_files``.

        - Two-phase agents (AWAITING_INPUT/Phase 2) always run sequentially.
        - Single-file runs always run sequentially — process_document
          already handles append-into-existing-party-file natively.
        - Multi-file dispatcher-aware agents (summarize_discovery.py) use
          DispatcherSubprocessWorker so the script's own dispatcher
          handles party grouping + consolidation.
        - Everything else with >1 file uses the per-file parallel runner.
        """
        from .runners.subprocess_worker import SubprocessWorker
        from .runners.parallel_subprocess_worker import (
            ParallelSubprocessWorker,
            _TWO_PHASE_SCRIPTS,
        )
        from .runners.dispatcher_subprocess_worker import (
            DispatcherSubprocessWorker,
        )

        if num_files <= 1 or script_name in _TWO_PHASE_SCRIPTS:
            return SubprocessWorker
        if script_name in TaskTab._DISPATCHER_SCRIPTS:
            return DispatcherSubprocessWorker
        return ParallelSubprocessWorker

    def _start_run(self, settings_dict: dict) -> None:
        from .runners.subprocess_worker import SubprocessWorker
        from .runners.parallel_subprocess_worker import ParallelSubprocessWorker
        from .runners.dispatcher_subprocess_worker import (
            DispatcherSubprocessWorker,
        )

        self._awaiting_session_path = None
        self.status_page.on_status(f"Starting {self._spec.title}…")

        worker_cls = self._pick_worker_cls(
            self._spec.script_name, len(self._files)
        )

        if worker_cls is ParallelSubprocessWorker:
            self._worker = ParallelSubprocessWorker(
                script_name=self._spec.script_name,
                case_path=self._case_path,
                file_number=self._file_number,
                files=self._files,
                settings=settings_dict,
                max_concurrent=3,
                parent=self,
            )
        elif worker_cls is DispatcherSubprocessWorker:
            self._worker = DispatcherSubprocessWorker(
                script_name=self._spec.script_name,
                case_path=self._case_path,
                file_number=self._file_number,
                files=self._files,
                settings=settings_dict,
                parent=self,
            )
        else:
            self._worker = SubprocessWorker(
                script_name=self._spec.script_name,
                case_path=self._case_path,
                file_number=self._file_number,
                files=self._files,
                settings=settings_dict,
                phase1_args=list(self._spec.phase1_args),
                phase2_flag=self._spec.phase2_flag,
                parent=self,
            )

        if not self._settings_owns_worker:
            self._worker.status.connect(self.status_page.on_status)
            self._worker.progress.connect(self.status_page.on_progress)
            self._worker.awaiting_input.connect(self._on_worker_awaiting_input)
        self._worker.finished.connect(self._on_worker_finished)
        self._worker.failed.connect(self._on_worker_failed)
        self._worker.cancelled.connect(self._on_worker_cancelled)
        # Parallel runs stream per-file completions; the sequential worker
        # doesn't have this signal, so guard with hasattr.
        if hasattr(self._worker, "file_completed"):
            self._worker.file_completed.connect(self._on_worker_file_completed)

        self._total_files_for_run = len(self._files)
        self._files_done_for_run = 0
        self.output_page.set_progress_hint("")
        self._worker.start()

    def _on_worker_file_completed(self, output_path: str) -> None:
        """A single file's agent has finished in a parallel run.

        On the first completion we transition to the Output page so the user
        can start reviewing. Subsequent completions are appended to the
        picker without disrupting the currently-viewed file. The inline
        progress hint is updated to reflect how many are still running.
        """
        self._files_done_for_run += 1
        remaining = max(0, self._total_files_for_run - self._files_done_for_run)

        if self.currentIndex() != PAGE_OUTPUT:
            # First file done — switch to Output page so the user can read it.
            self.output_page.load_outputs([output_path])
            self.setCurrentIndex(PAGE_OUTPUT)
        else:
            # Already viewing outputs; just grow the picker.
            self.output_page.append_output(output_path)

        if remaining > 0:
            word = "file" if remaining == 1 else "files"
            self.output_page.set_progress_hint(
                f"Still processing {remaining} more {word}…"
            )
        else:
            self.output_page.set_progress_hint("")

    def _on_worker_finished(self, output_path: str) -> None:
        from datetime import datetime
        # Capture all produced outputs BEFORE we drop the worker reference.
        all_output_paths = list(getattr(self._worker, "output_paths", []) or [])
        if output_path and output_path not in all_output_paths:
            all_output_paths.append(output_path)

        self._worker = None
        self._awaiting_session_path = None
        self._settings_owns_worker = False
        entry = {
            "task_id": self._spec.task_id,
            "title": self._spec.title,
            "files": list(self._files),
            "settings": self.settings_page.to_dict(),
            "output_path": output_path,
            "output_paths": all_output_paths,
            "completed_at": datetime.now().isoformat(timespec="seconds"),
        }
        self.task_completed.emit(entry)

        # If the user has been viewing streamed outputs, the picker already
        # has them; just make sure any late-arriving paths are present and
        # the progress hint is cleared. Otherwise (sequential worker, or a
        # parallel run that finished before the first file_completed could
        # fire — shouldn't happen, but be defensive), do a full load.
        currently_on_output = self.currentIndex() == PAGE_OUTPUT
        existing = set(self.output_page.output_paths)
        for p in all_output_paths:
            if p not in existing:
                self.output_page.append_output(p)
        if not currently_on_output:
            if len(all_output_paths) > 1:
                self.output_page.load_outputs(all_output_paths)
            else:
                self.output_page.load_output(output_path)
            self.setCurrentIndex(PAGE_OUTPUT)
        self.output_page.set_progress_hint("")

    def _show_output(self, output_path: str) -> None:
        """Single-output convenience used by tests. Production code path is
        in _on_worker_finished which can surface multiple parallel outputs."""
        self.output_page.load_output(output_path)
        self.setCurrentIndex(PAGE_OUTPUT)

    def _on_worker_failed(self, err: str) -> None:
        self._worker = None
        self._awaiting_session_path = None
        self._settings_owns_worker = False
        self.output_page.set_progress_hint("")
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
        self._awaiting_session_path = None
        self._settings_owns_worker = False
        self.output_page.set_progress_hint("")
        self.setCurrentIndex(PAGE_SETTINGS)

    # ---- Two-phase deposition (legacy popup flow for Advanced Mode) ----

    def _on_worker_awaiting_input(self, session_path: str) -> None:
        """Phase 1 complete — store session path and switch status page to awaiting mode."""
        if self._settings_owns_worker:
            return  # settings page is handling this itself
        self._awaiting_session_path = session_path
        self.status_page.show_awaiting_input(session_path)
        self.status_page.pause_eta()

    def _on_configure_requested(self) -> None:
        """User clicked 'Configure Topics & Continue' — open DepoSummaryConfigDialog."""
        if self._worker is None or self._awaiting_session_path is None:
            return
        from PySide6.QtWidgets import QDialog
        from icharlotte_core.ui.depo_summary_config_dialog import DepoSummaryConfigDialog
        try:
            dlg = DepoSummaryConfigDialog(self._awaiting_session_path, parent=self)
        except Exception as e:
            self.status_page.on_status(f"Failed to open topic config dialog: {e}")
            return

        if dlg.exec() == QDialog.DialogCode.Accepted:
            self.status_page.resume_eta()
            # Re-show progress bar for Phase 2.
            self.status_page.awaiting_input_widget.setVisible(False)
            self.status_page.progress_bar.setRange(0, 0)
            self.status_page.status_label.setText("Running Phase 2 — generating summary…")
            self._worker.resume_with_config(self._awaiting_session_path)
        # On Cancel: leave the awaiting widget visible so the user can retry.
