"""SubprocessWorker — runs an existing Scripts/*.py agent as a QProcess.

Stdout is parsed for:
  - lines starting with "PROGRESS:"       → progress(int) updates
  - lines starting with "AWAITING_INPUT:" → awaiting_input(session_path) held
                                            until process exits cleanly
  - all other lines                       → status(str) log emissions

Each agent is invoked with the document path as a **positional argument**
(no --file_number / --file flags).  The agents extract the file number from
the path themselves.

Multi-file behaviour: files are processed one at a time sequentially.  When
one finishes successfully we advance to the next, refreshing the pre-existing
output snapshot so we correctly detect the per-run .docx.

Two-phase deposition flow (summarize_deposition.py):
  Phase 1 — `python -u <script> <path>`
             Emits AWAITING_INPUT:<session_path> then exits with code 0.
             No .docx is produced yet.
  Phase 2 — `python -u <script> --phase=summary <session_path>`
             Produces the final .docx.
  After Phase 1 finishes, we emit awaiting_input(session_path) and pause.
  TaskTab calls resume_with_config(session_path) after the user picks topics.

Cancellation: cancel() calls QProcess.terminate() then kill() after 2 s,
then emits cancelled().
"""
import os
import re
import sys
from typing import List, Optional

from PySide6.QtCore import QProcess, QTimer

from .base import BaseWorker


_AI_OUTPUT_SUBPATH = os.path.join("NOTES", "AI Output")
_PROGRESS_RE = re.compile(r"^PROGRESS:\s*(\d+)\s*$")
_AWAITING_RE = re.compile(r"^AWAITING_INPUT:(.+)$")


class SubprocessWorker(BaseWorker):
    """Run a Scripts/*.py agent via QProcess, one file at a time."""

    def __init__(
        self,
        script_name: str,           # e.g. "summarize.py"
        case_path: str,
        file_number: str,
        files: List[str],
        settings: dict,
        parent=None,
    ):
        super().__init__(case_path, file_number, files, settings, parent)
        self._script_name = script_name
        self._process: Optional[QProcess] = None
        self._pre_existing_outputs: set = set()
        self._stdout_buf = bytearray()
        self._file_idx: int = 0
        self._awaiting_session_path: Optional[str] = None
        # Track the newest .docx seen across all file runs.
        self._newest_output: Optional[str] = None

    # ---- Lifecycle ----

    def start(self) -> None:
        self._file_idx = 0
        self._newest_output = None
        self._awaiting_session_path = None
        self._start_file(self.files[self._file_idx])

    def _start_file(self, file_path: str) -> None:
        """Snapshot outputs then launch Phase 1 for file_path."""
        self._pre_existing_outputs = self._scan_outputs()
        self._awaiting_session_path = None
        self._stdout_buf = bytearray()
        self._launch_process([self._script_path(), file_path])

    def _launch_process(self, extra_argv: List[str]) -> None:
        """Create and start a QProcess with `python -u <extra_argv>`."""
        if self._process is not None:
            try:
                self._process.disconnect()
            except Exception:
                pass
        self._process = QProcess(self)
        self._process.setProcessChannelMode(QProcess.ProcessChannelMode.MergedChannels)
        self._process.readyReadStandardOutput.connect(self._drain_stdout)
        self._process.finished.connect(self._on_process_finished)
        self._process.errorOccurred.connect(self._on_process_error)

        argv = ["-u"] + extra_argv
        self.status.emit(
            "Running: python "
            + " ".join(
                os.path.basename(a) if os.sep in a or "/" in a else a
                for a in argv
            )
        )
        self._process.start(sys.executable, argv)

    def _script_path(self) -> str:
        # __file__ lives 5 levels deep:
        #   <root>/icharlotte_core/ui/wizard/runners/subprocess_worker.py
        here = os.path.abspath(__file__)
        repo_root = here
        for _ in range(5):
            repo_root = os.path.dirname(repo_root)
        return os.path.join(repo_root, "Scripts", self._script_name)

    def cancel(self) -> None:
        super().cancel()
        if self._process is None:
            self.cancelled.emit()
            return
        self.status.emit("Cancellation requested…")
        self._process.terminate()
        QTimer.singleShot(2000, self._hard_kill_if_running)

    def _hard_kill_if_running(self) -> None:
        if self._process is not None and self._process.state() != QProcess.ProcessState.NotRunning:
            self.status.emit("Forcing kill…")
            self._process.kill()

    # ---- Phase 2 resume (deposition two-phase flow) ----

    def resume_with_config(self, session_path: str) -> None:
        """Start Phase 2: `python -u <script> --phase=summary <session_path>`."""
        self._awaiting_session_path = None
        self._stdout_buf = bytearray()
        # Refresh snapshot so we detect the Phase 2 .docx.
        self._pre_existing_outputs = self._scan_outputs()
        self._launch_process([self._script_path(), "--phase=summary", session_path])

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
        # PROGRESS:N
        m = _PROGRESS_RE.match(line)
        if m:
            try:
                raw = int(m.group(1))
                # Scale to aggregate progress across all files.
                scaled = (self._file_idx * 100 + raw) // max(len(self.files), 1)
                self.progress.emit(scaled)
            except ValueError:
                pass
            return

        # AWAITING_INPUT:<session_path>  — capture but do NOT emit yet.
        m = _AWAITING_RE.match(line)
        if m:
            self._awaiting_session_path = m.group(1).strip()
            return  # suppress from status log

        if line:
            self.status.emit(line)

    def _on_process_finished(self, exit_code: int, exit_status: QProcess.ExitStatus) -> None:
        # Drain any remaining bytes.
        self._drain_stdout()
        if self._stdout_buf:
            self._handle_line(self._stdout_buf.decode("utf-8", errors="replace"))
            self._stdout_buf.clear()

        if self.is_cancel_requested:
            self.cancelled.emit()
            return

        if exit_code != 0:
            self.failed.emit(f"Agent exited with code {exit_code}.")
            return

        # --- Two-phase deposition: Phase 1 complete ---
        if self._awaiting_session_path is not None:
            self.awaiting_input.emit(self._awaiting_session_path)
            return  # do NOT scan for .docx; do NOT advance to next file

        # --- Normal completion: look for new .docx ---
        new_outputs = self._scan_outputs() - self._pre_existing_outputs
        if not new_outputs:
            self.failed.emit("Agent reported success but no new .docx appeared in NOTES/AI Output/.")
            return

        newest = max(new_outputs, key=lambda p: os.path.getmtime(p))
        # Track the globally newest output across all file runs.
        if self._newest_output is None or os.path.getmtime(newest) > os.path.getmtime(self._newest_output):
            self._newest_output = newest

        # Advance to the next file, or finish.
        self._advance()

    def _advance(self) -> None:
        """Move to the next file, or emit finished if all files are done."""
        self._file_idx += 1
        if self._file_idx < len(self.files):
            self.status.emit(
                f"File {self._file_idx} of {len(self.files)} done — starting next…"
            )
            self._start_file(self.files[self._file_idx])
        else:
            self.finished.emit(self._newest_output or "")

    def _on_process_error(self, err: QProcess.ProcessError) -> None:
        if self.is_cancel_requested:
            return
        self.failed.emit(f"Process error: {err}")

    # ---- Output discovery ----

    def _scan_outputs(self) -> set:
        out_dir = os.path.join(self.case_path, _AI_OUTPUT_SUBPATH)
        if not os.path.isdir(out_dir):
            return set()
        results: set = set()
        for name in os.listdir(out_dir):
            if name.lower().endswith(".docx"):
                results.add(os.path.join(out_dir, name))
        return results
