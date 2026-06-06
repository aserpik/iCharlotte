"""CaseAgentWorker - runs case-level Scripts/*.py agents by file number."""
from __future__ import annotations

import os
import sys
from collections import deque
from typing import Iterable

from PySide6.QtCore import QProcess, QTimer

from .base import BaseWorker


class CaseAgentWorker(BaseWorker):
    """Run an existing case-level agent script with the loaded file number.

    Unlike SubprocessWorker, this worker does not receive document paths and
    does not scan NOTES/AI Output for produced docx files. Complaint and docket
    update multiple case-level artifacts, so callers summarize current case
    state after the process exits.
    """

    def __init__(
        self,
        script_name: str,
        case_path: str,
        file_number: str,
        extra_flags: Iterable[str] | None = None,
        recent_line_limit: int = 80,
        parent=None,
    ):
        super().__init__(
            case_path=case_path,
            file_number=file_number,
            files=[],
            settings={},
            parent=parent,
        )
        self._script_name = script_name
        self._extra_flags = list(extra_flags or [])
        self._process: QProcess | None = None
        self._stdout_buf = bytearray()
        self._recent_lines = deque(maxlen=max(1, int(recent_line_limit)))

    @property
    def recent_lines(self) -> list[str]:
        return list(self._recent_lines)

    def command_argv(self) -> list[str]:
        return [self._script_path(), self.file_number, "--headless", *self._extra_flags]

    def _script_path(self) -> str:
        here = os.path.abspath(__file__)
        repo_root = here
        for _ in range(5):
            repo_root = os.path.dirname(repo_root)
        return os.path.join(repo_root, "Scripts", self._script_name)

    def start(self) -> None:
        argv = ["-u", *self.command_argv()]
        self.status.emit(
            "Running: python "
            + " ".join(
                os.path.basename(a) if os.sep in a or "/" in a else a
                for a in argv
            )
        )
        self._process = QProcess(self)
        self._process.setProcessChannelMode(QProcess.ProcessChannelMode.MergedChannels)
        self._process.readyReadStandardOutput.connect(self._drain_stdout)
        self._process.finished.connect(self._on_process_finished)
        self._process.errorOccurred.connect(self._on_process_error)
        self._process.start(sys.executable, argv)

    def cancel(self) -> None:
        super().cancel()
        if self._process is None:
            self.cancelled.emit()
            return
        self.status.emit("Cancellation requested.")
        self._process.terminate()
        QTimer.singleShot(2000, self._hard_kill_if_running)

    def _hard_kill_if_running(self) -> None:
        if (
            self._process is not None
            and self._process.state() != QProcess.ProcessState.NotRunning
        ):
            self.status.emit("Forcing kill.")
            self._process.kill()

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
        text = (line or "").strip()
        if not text:
            return
        self._recent_lines.append(text)
        self.status.emit(text)

    def _on_process_finished(self, exit_code: int, exit_status: QProcess.ExitStatus) -> None:
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
        self.finished.emit("")

    def _on_process_error(self, err: QProcess.ProcessError) -> None:
        if self.is_cancel_requested:
            return
        self.failed.emit(f"Process error: {err}")
