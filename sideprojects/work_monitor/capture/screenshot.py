"""Screenshot capture and Gemini Vision analysis."""

import threading
import queue
import time
import os
import logging
from datetime import datetime

import mss
from PIL import Image

from ..db import ActivityDB
from ..case_extractor import extract_from_multiple
from .. import config

logger = logging.getLogger("work_monitor.screenshot")


class ScreenshotCapture(threading.Thread):
    """Captures screenshots at fixed intervals."""

    def __init__(self, db: ActivityDB, stop_event: threading.Event,
                 analysis_queue: queue.Queue, window_tracker,
                 idle_flag: threading.Event = None):
        super().__init__(daemon=True, name="ScreenshotCapture")
        self.db = db
        self.stop_event = stop_event
        self.analysis_queue = analysis_queue
        self.window_tracker = window_tracker
        self.idle_flag = idle_flag
        self._burst_remaining = 0

    def run(self):
        logger.info("ScreenshotCapture started")
        while not self.stop_event.is_set():
            # Skip during idle (unless burst mode)
            if self.idle_flag and self.idle_flag.is_set() and self._burst_remaining <= 0:
                time.sleep(1.0)
                continue

            try:
                self._capture()
            except Exception as e:
                logger.debug(f"Capture error: {e}")

            if self._burst_remaining > 0:
                self._burst_remaining -= 1
                interval = config.BURST_SCREENSHOT_INTERVAL
            else:
                interval = config.SCREENSHOT_INTERVAL

            # Sleep in 1-second chunks
            elapsed = 0.0
            while elapsed < interval:
                if self.stop_event.is_set():
                    return
                time.sleep(1.0)
                elapsed += 1.0

        logger.info("ScreenshotCapture stopped")

    def _capture(self):
        timestamp = datetime.now()
        filename = f"screenshot_{timestamp.strftime('%Y%m%d_%H%M%S')}.jpg"
        filepath = os.path.join(config.SCREENSHOT_DIR, filename)

        with mss.mss() as sct:
            monitor = sct.monitors[1]  # Primary monitor
            sct_img = sct.grab(monitor)
            img = Image.frombytes("RGB", sct_img.size, sct_img.bgra, "raw", "BGRX")
            img.save(filepath, "JPEG", quality=config.SCREENSHOT_QUALITY)

        app_name = self.window_tracker.get_current_app()
        window_title = self.window_tracker.get_current_title()
        case_number = extract_from_multiple(window_title)

        row_id = self.db.insert_screenshot(
            timestamp=timestamp.isoformat(),
            file_path=filepath,
            vision_description=None,
            app_name=app_name,
            window_title=window_title,
            case_number=case_number,
        )

        # Queue for async vision analysis
        try:
            self.analysis_queue.put_nowait((row_id, filepath, app_name))
        except queue.Full:
            logger.debug("Analysis queue full, skipping vision analysis")

    def trigger_burst(self):
        """Called by idle detector on return from idle."""
        self._burst_remaining = config.BURST_SCREENSHOT_COUNT
        logger.info(f"Burst mode: {self._burst_remaining} screenshots queued")


class ScreenshotAnalyzer(threading.Thread):
    """Processes queued screenshots through Gemini Vision API."""

    def __init__(self, db: ActivityDB, stop_event: threading.Event,
                 analysis_queue: queue.Queue):
        super().__init__(daemon=True, name="ScreenshotAnalyzer")
        self.db = db
        self.stop_event = stop_event
        self.analysis_queue = analysis_queue
        self._client = None

    def _get_client(self):
        if self._client is None:
            try:
                from google import genai
                api_key = os.environ.get("GEMINI_API_KEY")
                if api_key:
                    self._client = genai.Client(api_key=api_key)
            except ImportError:
                logger.warning("google-genai not installed, vision analysis disabled")
        return self._client

    def run(self):
        logger.info("ScreenshotAnalyzer started")
        while not self.stop_event.is_set():
            try:
                row_id, filepath, app_name = self.analysis_queue.get(timeout=5.0)
                self._analyze(row_id, filepath, app_name)
            except queue.Empty:
                continue
            except Exception as e:
                logger.debug(f"Analyzer error: {e}")

        logger.info("ScreenshotAnalyzer stopped")

    def _analyze(self, row_id: int, filepath: str, app_name: str):
        client = self._get_client()
        if not client or not os.path.exists(filepath):
            return

        try:
            with open(filepath, "rb") as f:
                uploaded = client.files.upload(
                    file=f,
                    config={
                        'display_name': f"work_screenshot_{row_id}",
                        'mime_type': 'image/jpeg',
                    }
                )

            prompt = (
                "Describe what the user is doing in this screenshot in 2-3 sentences. "
                f"The active application is {app_name}. "
                "Focus on: what document/webpage is open, what action is being performed, "
                "and any visible case numbers in format ####.###. "
                "Be factual and concise."
            )

            response = client.models.generate_content(
                model=config.VISION_MODEL,
                contents=[uploaded, prompt],
            )

            if response and response.text:
                description = response.text.strip()
                self.db.update_screenshot_description(row_id, description)
                logger.debug(f"Screenshot {row_id} analyzed: {description[:80]}...")
        except Exception as e:
            logger.debug(f"Vision analysis failed for screenshot {row_id}: {e}")
