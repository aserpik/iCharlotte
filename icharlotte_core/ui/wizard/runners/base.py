"""BaseWorker — common signal surface + cancellation contract for wizard runners."""
from typing import List

from PySide6.QtCore import QObject, Signal


class BaseWorker(QObject):
    """Abstract worker for wizard tasks.

    Subclasses implement start() (which kicks off a QProcess or QThread) and call
    self._on_status / _on_progress / _on_finished / _on_failed as the work proceeds.
    Cancellation is cooperative: cancel() flips a flag; subclasses decide how to
    honor it (e.g., terminating a QProcess, polling the flag in a loop).
    """

    status = Signal(str)         # one log line
    progress = Signal(int)       # 0-100
    finished = Signal(str)       # output_path (.docx)
    failed = Signal(str)         # error message
    cancelled = Signal()         # emitted after cancel takes effect
    awaiting_input = Signal(str) # session_path emitted when agent pauses awaiting user config (depositions Phase 1 done)

    def __init__(
        self,
        case_path: str,
        file_number: str,
        files: List[str],
        settings: dict,
        parent: QObject | None = None,
    ):
        super().__init__(parent)
        self.case_path = case_path
        self.file_number = file_number
        self.files = list(files)
        self.settings = dict(settings)
        self._cancel_requested = False

    @property
    def is_cancel_requested(self) -> bool:
        return self._cancel_requested

    def cancel(self) -> None:
        """Request cancellation. Subclasses may override to take additional action."""
        self._cancel_requested = True

    def start(self) -> None:
        """Subclasses must implement."""
        raise NotImplementedError
