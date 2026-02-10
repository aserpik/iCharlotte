"""
Work Activity Monitor - Main Service

Run: python -m sideprojects.work_monitor.service
Designed for Windows Task Scheduler (trigger: on logon, restart on failure).
"""

import signal
import sys
import os
import threading
import queue
import time
import logging

from . import config
from .db import ActivityDB
from .capture.window_tracker import WindowTracker
from .capture.screenshot import ScreenshotCapture, ScreenshotAnalyzer
from .capture.clipboard import ClipboardMonitor
from .capture.file_watcher import FileWatcher
from .capture.outlook_monitor import OutlookMonitor
from .idle.detector import IdleDetector
from .cleanup import run_cleanup

# Configure logging to file
config.ensure_dirs()
logging.basicConfig(
    filename=config.SERVICE_LOG,
    level=logging.INFO,
    format="%(asctime)s [%(name)s] %(levelname)s: %(message)s",
)
logger = logging.getLogger("work_monitor.service")

# Also log to console when running interactively
console = logging.StreamHandler()
console.setLevel(logging.INFO)
console.setFormatter(logging.Formatter("%(asctime)s [%(name)s] %(message)s", "%H:%M:%S"))
logging.getLogger("work_monitor").addHandler(console)


# Thread factory functions for restart capability
def _make_threads(db, stop_event, idle_flag, analysis_queue):
    """Create all capture threads. Returns dict of name -> thread."""
    window_tracker = WindowTracker(db, stop_event, idle_flag)

    screenshot_capture = ScreenshotCapture(
        db, stop_event, analysis_queue, window_tracker, idle_flag
    )
    screenshot_analyzer = ScreenshotAnalyzer(db, stop_event, analysis_queue)
    clipboard_monitor = ClipboardMonitor(db, stop_event, window_tracker)
    file_watcher = FileWatcher(db, stop_event)
    outlook_monitor = OutlookMonitor(db, stop_event)

    idle_detector = IdleDetector(
        db, stop_event, idle_flag,
        on_return_callback=screenshot_capture.trigger_burst,
    )

    return {
        "WindowTracker": window_tracker,
        "ScreenshotCapture": screenshot_capture,
        "ScreenshotAnalyzer": screenshot_analyzer,
        "ClipboardMonitor": clipboard_monitor,
        "FileWatcher": file_watcher,
        "OutlookMonitor": outlook_monitor,
        "IdleDetector": idle_detector,
    }


def main():
    logger.info("=" * 60)
    logger.info("Work Activity Monitor starting")

    # Initialize
    config.ensure_dirs()
    db = ActivityDB()

    # Shared coordination objects
    stop_event = threading.Event()
    idle_flag = threading.Event()
    analysis_queue = queue.Queue(maxsize=100)

    # Signal handlers for graceful shutdown
    def shutdown(signum=None, frame=None):
        logger.info("Shutdown signal received")
        stop_event.set()

    signal.signal(signal.SIGINT, shutdown)
    signal.signal(signal.SIGTERM, shutdown)

    # Create and start all threads
    threads = _make_threads(db, stop_event, idle_flag, analysis_queue)

    for name, t in threads.items():
        t.start()
        logger.info(f"Started {name}")

    # Run cleanup once at startup
    try:
        run_cleanup(db)
    except Exception as e:
        logger.warning(f"Startup cleanup failed: {e}")

    # Main loop
    last_cleanup = time.time()

    try:
        while not stop_event.is_set():
            time.sleep(10)

            # Periodic cleanup
            if time.time() - last_cleanup > config.CLEANUP_INTERVAL:
                try:
                    run_cleanup(db)
                except Exception as e:
                    logger.warning(f"Cleanup failed: {e}")
                last_cleanup = time.time()

            # Monitor thread health
            for name, t in list(threads.items()):
                if not t.is_alive() and not stop_event.is_set():
                    logger.warning(f"{name} died unexpectedly")

    except KeyboardInterrupt:
        shutdown()

    # Wait for threads to finish
    logger.info("Waiting for threads to stop...")
    for name, t in threads.items():
        t.join(timeout=5)
        if t.is_alive():
            logger.warning(f"{name} did not stop within 5 seconds")
        else:
            logger.info(f"{name} stopped")

    logger.info("Work Activity Monitor stopped")


if __name__ == "__main__":
    main()
