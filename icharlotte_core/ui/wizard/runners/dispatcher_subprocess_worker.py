"""DispatcherSubprocessWorker — delegate multi-file runs to a script that
has its own internal dispatcher.

Why this exists
---------------
``ParallelSubprocessWorker`` assumes 1 input file → 1 output file and
gives every child a unique ``--output_path``. That's correct for
``summarize.py`` and friends but wrong for ``summarize_discovery.py``,
which:
  * groups multiple files from the same responding party into a single
    ``Discovery_Responses_{Party}.docx`` (via ``find_existing_party_file``),
  * runs a final consolidation LLM pass per unique output file (via
    ``consolidate_file``) when it's invoked in dispatcher mode
    (i.e. ``len(file_paths) > 1`` and no ``--output_path``).

Forcing ``--output_path`` breaks both behaviors. This worker preserves
them by running the script ONCE with all files passed positionally and
``--report-dir <temp>`` so the dispatcher can report back which output
files it ended up producing. After the process exits, we read the
``worker_*.txt`` reports the script writes for batch consolidation, and
surface the resulting (deduped) output paths to the OutputPage.

Trade-off vs. ParallelSubprocessWorker
--------------------------------------
The script's internal dispatcher runs children 3-at-a-time, but it does
not surface per-child completion to its own stdout in a way we can map
back to individual outputs mid-run. So ``file_completed`` only fires
after the whole batch is done (including consolidation), in mtime order.
The OutputPage still ends up with one entry per unique party file,
which is the user-visible behavior they wanted.
"""
import os
import re
import shutil
import sys
import tempfile
from typing import List, Optional

from PySide6.QtCore import QProcess, QTimer, Signal

from .base import BaseWorker


_PROGRESS_RE = re.compile(r"^PROGRESS:\s*(\d+)\s*(?::(.*))?$")


class DispatcherSubprocessWorker(BaseWorker):
    """Run a Scripts/*.py agent in single-subprocess "dispatcher" mode.

    Emits the same ``file_completed`` signal as ParallelSubprocessWorker
    so TaskTab can wire either runner identically.
    """

    file_completed = Signal(str)

    def __init__(
        self,
        script_name: str,
        case_path: str,
        file_number: str,
        files: List[str],
        settings: dict,
        parent=None,
    ):
        if not files:
            raise ValueError("DispatcherSubprocessWorker requires at least one input file.")
        super().__init__(case_path, file_number, files, settings, parent)
        self._script_name = script_name
        self._process: Optional[QProcess] = None
        self._stdout_buf = bytearray()
        self._report_dir: Optional[str] = None
        self._output_paths: List[str] = []

    # ---- Public API ----

    @property
    def output_paths(self) -> List[str]:
        return list(self._output_paths)

    # ---- Lifecycle ----

    def start(self) -> None:
        self._report_dir = tempfile.mkdtemp(prefix="wizard_disc_")
        argv = self._build_argv(self._report_dir)

        self._process = QProcess(self)
        self._process.setProcessChannelMode(QProcess.ProcessChannelMode.MergedChannels)
        self._process.readyReadStandardOutput.connect(self._drain_stdout)
        self._process.finished.connect(self._on_process_finished)
        self._process.errorOccurred.connect(self._on_process_error)

        self.status.emit(
            f"Starting {self._script_name} dispatcher on {len(self.files)} file(s)…"
        )
        self._process.start(sys.executable, argv)

    def cancel(self) -> None:
        super().cancel()
        if self._process is None:
            self._cleanup_report_dir()
            self.cancelled.emit()
            return
        self.status.emit("Cancellation requested — terminating dispatcher…")
        try:
            self._process.terminate()
        except RuntimeError:
            pass
        QTimer.singleShot(2000, self._hard_kill_if_running)

    def _hard_kill_if_running(self) -> None:
        if self._process is None:
            return
        try:
            if self._process.state() != QProcess.ProcessState.NotRunning:
                self._process.kill()
        except RuntimeError:
            pass

    # ---- Argv ----

    def _build_argv(self, report_dir: str) -> List[str]:
        """Compose the python argv. Exposed for tests."""
        return (
            ["-u", self._script_path()]
            + list(self.files)
            + ["--report-dir", report_dir]
        )

    def _script_path(self) -> str:
        # __file__ lives 5 levels deep:
        #   <root>/icharlotte_core/ui/wizard/runners/dispatcher_subprocess_worker.py
        here = os.path.abspath(__file__)
        repo_root = here
        for _ in range(5):
            repo_root = os.path.dirname(repo_root)
        return os.path.join(repo_root, "Scripts", self._script_name)

    # ---- Stdout / events ----

    def _drain_stdout(self) -> None:
        if self._process is None:
            return
        data = bytes(self._process.readAllStandardOutput())
        if not data:
            return
        self._stdout_buf.extend(data)
        while b"\n" in self._stdout_buf:
            line_bytes, _, rest = self._stdout_buf.partition(b"\n")
            self._stdout_buf = bytearray(rest)
            line = line_bytes.decode("utf-8", errors="replace").rstrip("\r")
            self._handle_line(line)

    def _handle_line(self, line: str) -> None:
        m = _PROGRESS_RE.match(line)
        if m:
            try:
                pct = max(0, min(100, int(m.group(1))))
                self.progress.emit(pct)
                msg = (m.group(2) or "").strip()
                if msg:
                    self.status.emit(msg)
            except ValueError:
                pass
            return
        if line:
            self.status.emit(line)

    def _on_process_finished(
        self, exit_code: int, exit_status: QProcess.ExitStatus
    ) -> None:
        self._drain_stdout()
        if self._stdout_buf:
            self._handle_line(self._stdout_buf.decode("utf-8", errors="replace"))
            self._stdout_buf = bytearray()

        if self.is_cancel_requested:
            self._cleanup_report_dir()
            self.cancelled.emit()
            return

        if exit_code != 0:
            self.failed.emit(f"Dispatcher exited with code {exit_code}.")
            self._cleanup_report_dir()
            return

        self._finalize_success()

    def _on_process_error(self, err: QProcess.ProcessError) -> None:
        if self.is_cancel_requested:
            return
        self.failed.emit(f"Process error: {err}")

    # ---- Output discovery ----

    def _collect_output_paths(self) -> List[str]:
        """Read ``worker_*.txt`` files in ``_report_dir`` and return the
        unique, existing output paths the dispatcher wrote, sorted by
        mtime (oldest first → newest last)."""
        if not self._report_dir or not os.path.isdir(self._report_dir):
            return []
        seen = set()
        results: List[str] = []
        try:
            entries = os.listdir(self._report_dir)
        except OSError:
            return []
        for name in entries:
            if not (name.startswith("worker_") and name.endswith(".txt")):
                continue
            try:
                with open(os.path.join(self._report_dir, name), "r") as f:
                    path = f.read().strip()
            except OSError:
                continue
            if not path or path in seen:
                continue
            if not os.path.isfile(path):
                continue
            seen.add(path)
            results.append(path)
        results.sort(key=lambda p: os.path.getmtime(p))
        return results

    def _finalize_success(self) -> None:
        paths = self._collect_output_paths()
        if not paths:
            self.failed.emit(
                "Dispatcher finished but produced no output files. "
                "Check the log above for per-file errors."
            )
            self._cleanup_report_dir()
            return

        for p in paths:
            self._output_paths.append(p)
            self.file_completed.emit(p)
        newest = paths[-1]  # mtime-sorted: last is newest
        self._cleanup_report_dir()
        self.finished.emit(newest)

    def _cleanup_report_dir(self) -> None:
        if self._report_dir and os.path.isdir(self._report_dir):
            try:
                shutil.rmtree(self._report_dir)
            except OSError:
                pass
        self._report_dir = None
