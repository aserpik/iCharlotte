"""Cleanup routines for old screenshots and log records."""

import os
import time
import logging
from datetime import datetime, timedelta

from . import config
from .db import ActivityDB

logger = logging.getLogger("work_monitor.cleanup")


def run_cleanup(db: ActivityDB):
    """Run all cleanup tasks."""
    cleanup_screenshots(db)
    cleanup_old_records(db)


def cleanup_screenshots(db: ActivityDB):
    """Delete screenshot files older than SCREENSHOT_RETENTION_HOURS."""
    if not os.path.exists(config.SCREENSHOT_DIR):
        return

    cutoff = time.time() - (config.SCREENSHOT_RETENTION_HOURS * 3600)
    removed = 0

    for filename in os.listdir(config.SCREENSHOT_DIR):
        filepath = os.path.join(config.SCREENSHOT_DIR, filename)
        try:
            if os.path.isfile(filepath) and os.path.getmtime(filepath) < cutoff:
                os.remove(filepath)
                db.clear_screenshot_path(filepath)
                removed += 1
        except Exception as e:
            logger.debug(f"Failed to remove {filepath}: {e}")

    if removed:
        logger.info(f"Cleaned up {removed} old screenshots")


def cleanup_old_records(db: ActivityDB):
    """Delete event records older than LOG_RETENTION_DAYS."""
    cutoff_date = (
        datetime.now() - timedelta(days=config.LOG_RETENTION_DAYS)
    ).strftime("%Y-%m-%d")

    try:
        db.delete_records_before(cutoff_date)
        logger.info(f"Cleaned up records before {cutoff_date}")
    except Exception as e:
        logger.debug(f"Record cleanup failed: {e}")
