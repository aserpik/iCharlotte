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
_PROGRESS_RE = re.compile(r"^PROGRESS:\s*(\d+)\s*(?::(.*))?$")
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
        phase1_args: List[str] | None = None,
        phase2_flag: str = "--phase=summary",
        parent=None,
    ):
        super().__init__(case_path, file_number, files, settings, parent)
        self._script_name = script_name
        self._phase1_args = list(phase1_args or [])
        self._phase2_flag = phase2_flag
        self._process: Optional[QProcess] = None
        # Map of path -> mtime captured BEFORE each run. A run "produced" a docx
        # if its mtime advanced (overwrite/append) or the path didn't exist
        # in the snapshot (new file). Path-only diff misses append-in-place,
        # which is the default behavior of summarize.py / summarize_deposition.py.
        self._pre_existing_outputs: dict = {}
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
        argv = [self._script_path()] + self._phase1_args + [file_path]
        self._launch_process(argv)

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
        """Start Phase 2 with the configured phase2_flag."""
        self._awaiting_session_path = None
        self._stdout_buf = bytearray()
        # Refresh snapshot so we detect the Phase 2 .docx.
        self._pre_existing_outputs = self._scan_outputs()
        self._launch_process([self._script_path(), self._phase2_flag, session_path])

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

    def _scaled_pct(self, raw_pct: int) -> int:
        total = max(1, len(self.files))
        scaled = (self._file_idx * 100 + raw_pct) / total
        return max(0, min(100, int(scaled)))

    def _handle_line(self, line: str) -> None:
        # PROGRESS:N  or  PROGRESS:N:message
        m = _PROGRESS_RE.match(line)
        if m:
            try:
                raw = int(m.group(1))
                self.progress.emit(self._scaled_pct(raw))
                msg = (m.group(2) or "").strip()
                if msg:
                    self.status.emit(msg)
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

        # --- Normal completion: detect produced .docx by mtime ---
        # A path counts as "produced" if it didn't exist in the pre-run snapshot
        # OR its mtime advanced during the run (covers append-in-place).
        current = self._scan_outputs()
        candidates = [
            (path, mtime)
            for path, mtime in current.items()
            if path not in self._pre_existing_outputs
            or mtime > self._pre_existing_outputs[path]
        ]
        if not candidates:
            self.failed.emit(
                "Agent reported success but no .docx in NOTES/AI Output/ was created or updated."
            )
            return

        newest_path, newest_mtime = max(candidates, key=lambda t: t[1])
        if (
            self._newest_output is None
            or newest_mtime > os.path.getmtime(self._newest_output)
        ):
            self._newest_output = newest_path

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

    def _scan_outputs(self) -> dict:
        """Return {path: mtime} for every .docx currently in NOTES/AI Output/."""
        out_dir = os.path.join(self.case_path, _AI_OUTPUT_SUBPATH)
        if not os.path.isdir(out_dir):
            return {}
        results: dict = {}
        for name in os.listdir(out_dir):
            if not name.lower().endswith(".docx"):
                continue
            path = os.path.join(out_dir, name)
            try:
                results[path] = os.path.getmtime(path)
            except OSError:
                pass
        return results
