"""Clipboard monitoring via pyperclip polling."""

import threading
import time
import logging
import hashlib
from datetime import datetime

import pyperclip

from ..db import ActivityDB
from ..case_extractor import extract_case_number

logger = logging.getLogger("work_monitor.clipboard")


class ClipboardMonitor(threading.Thread):
    """Monitors clipboard changes by polling every 2 seconds via pyperclip."""

    def __init__(self, db: ActivityDB, stop_event: threading.Event,
                 window_tracker):
        super().__init__(daemon=True, name="ClipboardMonitor")
        self.db = db
        self.stop_event = stop_event
        self.window_tracker = window_tracker
        self._last_hash = None

    def run(self):
        logger.info("ClipboardMonitor started")

        # Capture initial clipboard state so we don't log pre-existing content
        try:
            text = pyperclip.paste()
            if text:
                self._last_hash = hashlib.md5(
                    text.encode('utf-8', errors='replace')
                ).hexdigest()
        except Exception:
            pass

        while not self.stop_event.is_set():
            try:
                self._poll()
            except Exception as e:
                logger.debug(f"Clipboard poll error: {e}")

            if self.stop_event.wait(timeout=2.0):
                break

        logger.info("ClipboardMonitor stopped")

    def _poll(self):
        """Check if clipboard content has changed."""
        try:
            text = pyperclip.paste()
        except Exception:
            return

        if not text:
            return

        content_hash = hashlib.md5(
            text.encode('utf-8', errors='replace')
        ).hexdigest()
        if content_hash == self._last_hash:
            return
        self._last_hash = content_hash

        content_preview = text[:500]
        case_number = extract_case_number(content_preview) if content_preview else None

        self.db.insert_clipboard_event(
            timestamp=datetime.now().isoformat(),
            content_type="text",
            content_preview=content_preview,
            source_app=self.window_tracker.get_current_app(),
            source_window=self.window_tracker.get_current_title(),
            destination_app=None,
            destination_window=None,
            case_number=case_number,
        )
