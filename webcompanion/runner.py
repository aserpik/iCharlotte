"""Qt-free subprocess driver speaking the wizard stdout protocol."""
import os
import subprocess
import sys
import threading
from typing import Callable, List, Optional

from .protocol import ParsedLine, parse_line

_CREATE_NO_WINDOW = 0x08000000 if os.name == "nt" else 0


class ScriptRunner:
    """Run ``python -u <argv...>``; stream parsed stdout lines to a callback.

    on_event(ParsedLine) fires for every stdout line; on_exit(returncode)
    fires exactly once after EOF. Both run on the reader thread — callers
    must do their own locking.
    """

    def __init__(
        self,
        argv: List[str],
        on_event: Callable[[ParsedLine], None],
        on_exit: Callable[[int], None],
        cwd: Optional[str] = None,
    ):
        self._argv = list(argv)
        self._on_event = on_event
        self._on_exit = on_exit
        self._cwd = cwd
        self._proc: Optional[subprocess.Popen] = None

    def start(self) -> None:
        self._proc = subprocess.Popen(
            [sys.executable, "-u"] + self._argv,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="replace",
            cwd=self._cwd,
            creationflags=_CREATE_NO_WINDOW,
        )
        threading.Thread(target=self._read_loop, daemon=True).start()

    def _read_loop(self) -> None:
        proc = self._proc
        for line in proc.stdout:
            self._on_event(parse_line(line.rstrip("\r\n")))
        rc = proc.wait()
        self._on_exit(rc)

    def cancel(self) -> None:
        """Terminate, then hard-kill after 2 s if still alive."""
        proc = self._proc
        if proc is None or proc.poll() is not None:
            return
        proc.terminate()

        def _kill():
            if proc.poll() is None:
                proc.kill()

        threading.Timer(2.0, _kill).start()
