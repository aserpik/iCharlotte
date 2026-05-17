"""SubprocessWorker — runs an existing Scripts/*.py agent as a QProcess.

Stdout is parsed for:
  - lines starting with "PROGRESS:" → progress(int) updates
  - all other lines             → status(str) log emissions

When the process exits with returncode 0, we look for the .docx the agent
wrote into <case>/NOTES/AI Output/ that didn't exist before the run started.
That path is emitted as finished(output_path).

If returncode != 0 or no new .docx appears, we emit failed(...).

Cancellation: cancel() calls QProcess.terminate() (and then kill() after
2 seconds if it hasn't exited), then emits cancelled().
"""
import os
import re
import time
from typing import List

from PySide6.QtCore import QProcess, QTimer

from .base import BaseWorker


_AI_OUTPUT_SUBPATH = os.path.join("NOTES", "AI Output")
_PROGRESS_RE = re.compile(r"^PROGRESS:\s*(\d+)\s*$")


class SubprocessWorker(BaseWorker):
    """Run a Scripts/*.py agent via QProcess."""

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
        self._process: QProcess | None = None
        self._pre_existing_outputs: set[str] = set()
        self._stdout_buf = bytearray()

    # ---- Lifecycle ----

    def start(self) -> None:
        self._pre_existing_outputs = self._scan_outputs()
        self._process = QProcess(self)
        self._process.setProcessChannelMode(QProcess.ProcessChannelMode.MergedChannels)
        self._process.readyReadStandardOutput.connect(self._drain_stdout)
        self._process.finished.connect(self._on_process_finished)
        self._process.errorOccurred.connect(self._on_process_error)

        # Build argv. Existing agents accept --file_number; some accept --files
        # (forward as repeated --file args; the agent ignores unknown flags
        # gracefully because we control which script we call).
        # __file__ lives 5 levels deep:
        #   <root>/icharlotte_core/ui/wizard/runners/subprocess_worker.py
        here = os.path.abspath(__file__)
        repo_root = here
        for _ in range(5):
            repo_root = os.path.dirname(repo_root)
        script_path = os.path.join(repo_root, "Scripts", self._script_name)
        argv = ["-u", script_path, "--file_number", self.file_number]
        for f in self.files:
            argv.extend(["--file", f])

        import sys
        self.status.emit(f"Running: python {' '.join(os.path.basename(a) if a == script_path else a for a in argv)}")
        self._process.start(sys.executable, argv)

    def cancel(self) -> None:
        super().cancel()
        if self._process is None:
            self.cancelled.emit()
            return
        self.status.emit("Cancellation requested…")
        self._process.terminate()
        # Hard-kill backup after 2s.
        QTimer.singleShot(2000, self._hard_kill_if_running)

    def _hard_kill_if_running(self) -> None:
        if self._process is not None and self._process.state() != QProcess.ProcessState.NotRunning:
            self.status.emit("Forcing kill…")
            self._process.kill()

    # ---- Stdout / events ----

    def _drain_stdout(self) -> None:
        if self._process is None:
            return
        data = bytes(self._process.readAllStandardOutput())
        if not data:
            return
        self._stdout_buf.extend(data)
        # Process complete lines.
        while b"\n" in self._stdout_buf:
            line_bytes, _, rest = self._stdout_buf.partition(b"\n")
            self._stdout_buf = bytearray(rest)
            line = line_bytes.decode("utf-8", errors="replace").rstrip("\r")
            self._handle_line(line)

    def _handle_line(self, line: str) -> None:
        m = _PROGRESS_RE.match(line)
        if m:
            try:
                self.progress.emit(int(m.group(1)))
            except ValueError:
                pass
            return
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
        new_outputs = self._scan_outputs() - self._pre_existing_outputs
        if not new_outputs:
            self.failed.emit("Agent reported success but no new .docx appeared in NOTES/AI Output/.")
            return
        # Pick the newest of the new outputs.
        newest = max(new_outputs, key=lambda p: os.path.getmtime(p))
        self.finished.emit(newest)

    def _on_process_error(self, err: QProcess.ProcessError) -> None:
        if self.is_cancel_requested:
            return
        self.failed.emit(f"Process error: {err}")

    # ---- Output discovery ----

    def _scan_outputs(self) -> set[str]:
        out_dir = os.path.join(self.case_path, _AI_OUTPUT_SUBPATH)
        if not os.path.isdir(out_dir):
            return set()
        results: set[str] = set()
        for name in os.listdir(out_dir):
            if name.lower().endswith(".docx"):
                results.add(os.path.join(out_dir, name))
        return results
