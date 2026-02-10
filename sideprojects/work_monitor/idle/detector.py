"""Idle detection via GetLastInputInfo and workstation lock detection."""

import threading
import ctypes
import ctypes.wintypes
import time
import logging
from datetime import datetime

from ..db import ActivityDB
from .. import config

logger = logging.getLogger("work_monitor.idle")


class LASTINPUTINFO(ctypes.Structure):
    _fields_ = [
        ('cbSize', ctypes.wintypes.UINT),
        ('dwTime', ctypes.wintypes.DWORD),
    ]


class IdleDetector(threading.Thread):
    """Detects user idle state via GetLastInputInfo and lock detection."""

    def __init__(self, db: ActivityDB, stop_event: threading.Event,
                 idle_flag: threading.Event,
                 on_return_callback=None):
        super().__init__(daemon=True, name="IdleDetector")
        self.db = db
        self.stop_event = stop_event
        self.idle_flag = idle_flag
        self.on_return_callback = on_return_callback

        self._is_idle = False
        self._idle_start = None
        self._current_idle_type = None

    def run(self):
        logger.info("IdleDetector started")
        while not self.stop_event.is_set():
            try:
                idle_seconds = self._get_idle_time()
                is_locked = self._is_workstation_locked()

                if idle_seconds >= config.IDLE_THRESHOLD and not self._is_idle:
                    # Transition to idle
                    self._is_idle = True
                    self._idle_start = datetime.now()
                    self._current_idle_type = "locked" if is_locked else "inactive"
                    self.idle_flag.set()
                    logger.info(f"User idle ({self._current_idle_type})")

                elif idle_seconds < config.IDLE_THRESHOLD and self._is_idle:
                    # Return from idle
                    self._end_idle()
                    self.idle_flag.clear()
                    logger.info("User returned from idle")

                    if self.on_return_callback:
                        try:
                            self.on_return_callback()
                        except Exception as e:
                            logger.debug(f"Return callback error: {e}")

            except Exception as e:
                logger.debug(f"Idle check error: {e}")

            # Check every IDLE_CHECK_INTERVAL seconds, sleeping in 1s chunks
            elapsed = 0.0
            while elapsed < config.IDLE_CHECK_INTERVAL:
                if self.stop_event.is_set():
                    # Record final idle period if currently idle
                    if self._is_idle:
                        self._end_idle()
                    logger.info("IdleDetector stopped")
                    return
                time.sleep(1.0)
                elapsed += 1.0

    def _get_idle_time(self) -> float:
        """Get seconds since last user input."""
        lii = LASTINPUTINFO()
        lii.cbSize = ctypes.sizeof(LASTINPUTINFO)
        ctypes.windll.user32.GetLastInputInfo(ctypes.byref(lii))
        millis = ctypes.windll.kernel32.GetTickCount() - lii.dwTime
        return millis / 1000.0

    def _is_workstation_locked(self) -> bool:
        """Check if the workstation is locked via OpenInputDesktop."""
        try:
            desktop = ctypes.windll.user32.OpenInputDesktop(0, False, 0x0001)
            if desktop:
                ctypes.windll.user32.CloseDesktop(desktop)
                return False
            return True
        except Exception:
            return False

    def _end_idle(self):
        """Record the idle period in the database."""
        if self._idle_start:
            now = datetime.now()
            duration = (now - self._idle_start).total_seconds()

            try:
                self.db.insert_idle_period(
                    start_time=self._idle_start.isoformat(),
                    end_time=now.isoformat(),
                    idle_type=self._current_idle_type,
                    duration_seconds=duration,
                )
            except Exception as e:
                logger.debug(f"Failed to record idle period: {e}")

        self._is_idle = False
        self._idle_start = None
        self._current_idle_type = None
