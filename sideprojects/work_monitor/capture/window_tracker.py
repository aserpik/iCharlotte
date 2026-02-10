"""Active window tracking via GetForegroundWindow() polling."""

import threading
import ctypes
import ctypes.wintypes
import os
import time
import logging
from datetime import datetime

import win32gui
import win32process

from ..db import ActivityDB
from ..case_extractor import extract_from_multiple
from .. import config

logger = logging.getLogger("work_monitor.window_tracker")

try:
    import psutil
    _HAS_PSUTIL = True
except ImportError:
    _HAS_PSUTIL = False

user32 = ctypes.windll.user32


class WindowTracker(threading.Thread):
    """Polls the foreground window every 1 second and records switches."""

    def __init__(self, db: ActivityDB, stop_event: threading.Event,
                 idle_flag: threading.Event = None):
        super().__init__(daemon=True, name="WindowTracker")
        self.db = db
        self.stop_event = stop_event
        self.idle_flag = idle_flag

        self._current_hwnd = None
        self._current_app = ""
        self._current_title = ""
        self._current_event_id = None
        self._switch_time = None
        self._lock = threading.Lock()

    def run(self):
        logger.info("WindowTracker started")
        while not self.stop_event.is_set():
            try:
                self._poll()
            except Exception as e:
                logger.debug(f"Poll error: {e}")

            # Sleep in 1-second chunks for responsive shutdown
            elapsed = 0.0
            while elapsed < config.WINDOW_POLL_INTERVAL:
                if self.stop_event.is_set():
                    return
                time.sleep(min(1.0, config.WINDOW_POLL_INTERVAL - elapsed))
                elapsed += 1.0

        # Close final window event
        self._close_current_event()
        logger.info("WindowTracker stopped")

    def _poll(self):
        hwnd = user32.GetForegroundWindow()
        if hwnd == self._current_hwnd or hwnd == 0:
            return

        now = datetime.now()

        # Close previous window event
        self._close_current_event(now)

        # Get new window info
        try:
            title = win32gui.GetWindowText(hwnd)
        except Exception:
            title = ""

        app_name = self._get_app_name(hwnd)
        case_number = extract_from_multiple(title)

        event_id = self.db.insert_window_event(
            timestamp=now.isoformat(),
            app_name=app_name,
            window_title=title,
            case_number=case_number,
        )

        with self._lock:
            self._current_hwnd = hwnd
            self._current_app = app_name
            self._current_title = title
            self._current_event_id = event_id
            self._switch_time = now

    def _close_current_event(self, now=None):
        """Update duration of the current window event."""
        if self._current_event_id and self._switch_time:
            now = now or datetime.now()
            duration = (now - self._switch_time).total_seconds()
            try:
                self.db.update_window_event_duration(
                    self._current_event_id, duration
                )
            except Exception:
                pass

    def _get_app_name(self, hwnd: int) -> str:
        """Get process name for a window handle."""
        try:
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            if _HAS_PSUTIL:
                return psutil.Process(pid).name()
            else:
                import win32api
                handle = win32api.OpenProcess(0x0410, False, pid)
                try:
                    name = os.path.basename(
                        win32process.GetModuleFileNameEx(handle, 0)
                    )
                    return name
                finally:
                    win32api.CloseHandle(handle)
        except Exception:
            return "unknown"

    def get_current_app(self) -> str:
        """Thread-safe access to current app name."""
        with self._lock:
            return self._current_app or "unknown"

    def get_current_title(self) -> str:
        """Thread-safe access to current window title."""
        with self._lock:
            return self._current_title or ""
