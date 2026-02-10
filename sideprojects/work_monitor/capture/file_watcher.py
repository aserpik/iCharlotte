"""File system event monitoring via watchdog."""

import threading
import time
import os
import logging
from datetime import datetime

from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

from ..db import ActivityDB
from ..case_extractor import extract_from_multiple
from .. import config

logger = logging.getLogger("work_monitor.file_watcher")


class CaseFileHandler(FileSystemEventHandler):
    """Handles file system events and writes to DB."""

    def __init__(self, db: ActivityDB):
        self.db = db
        self._last_events: dict[tuple, float] = {}

    def _should_record(self, path: str, event_type: str) -> bool:
        """Debounce and filter events."""
        # Skip directories
        if os.path.isdir(path):
            return False

        # Filter by extension
        ext = os.path.splitext(path)[1].lower()
        if ext not in config.LEGAL_EXTENSIONS:
            return False

        # Skip temp files
        basename = os.path.basename(path)
        if basename.startswith('~') or basename.startswith('.'):
            return False

        # Debounce: skip if same event within 2 seconds
        key = (path, event_type)
        now = time.time()
        if key in self._last_events and now - self._last_events[key] < 2.0:
            return False
        self._last_events[key] = now

        # Prune stale debounce entries every 1000 events
        if len(self._last_events) > 1000:
            cutoff = now - 10.0
            self._last_events = {
                k: v for k, v in self._last_events.items() if v > cutoff
            }

        return True

    def on_created(self, event):
        if self._should_record(event.src_path, "open"):
            self._record("open", event.src_path)

    def on_modified(self, event):
        if self._should_record(event.src_path, "save"):
            self._record("save", event.src_path)

    def on_moved(self, event):
        if self._should_record(event.dest_path, "rename"):
            self._record("rename", event.dest_path)

    def on_deleted(self, event):
        if self._should_record(event.src_path, "delete"):
            self._record("delete", event.src_path)

    def _record(self, event_type: str, file_path: str):
        try:
            self.db.insert_file_event(
                timestamp=datetime.now().isoformat(),
                event_type=event_type,
                file_path=file_path,
                app_name=None,
                case_number=extract_from_multiple(file_path),
            )
        except Exception as e:
            logger.debug(f"Failed to record file event: {e}")


class FileWatcher(threading.Thread):
    """Runs watchdog observers on configured directories."""

    def __init__(self, db: ActivityDB, stop_event: threading.Event):
        super().__init__(daemon=True, name="FileWatcher")
        self.db = db
        self.stop_event = stop_event
        self._observer = Observer()

    def run(self):
        logger.info("FileWatcher started")
        handler = CaseFileHandler(self.db)
        watched = 0

        for folder in config.WATCHED_FOLDERS:
            if os.path.exists(folder):
                try:
                    self._observer.schedule(handler, folder, recursive=True)
                    watched += 1
                    logger.info(f"Watching: {folder}")
                except Exception as e:
                    logger.warning(f"Cannot watch {folder}: {e}")

        if watched == 0:
            logger.warning("No directories to watch")
            # Still stay alive so service doesn't think we died
            while not self.stop_event.is_set():
                time.sleep(1.0)
            return

        self._observer.start()

        while not self.stop_event.is_set():
            time.sleep(1.0)

        self._observer.stop()
        self._observer.join(timeout=5)
        logger.info("FileWatcher stopped")
