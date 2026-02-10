"""Central configuration for the Work Activity Monitor."""

import os

# Base paths
PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
DATA_DIR = os.path.join(PROJECT_ROOT, "data", "work_monitor")
DB_PATH = os.path.join(DATA_DIR, "db", "activity.db")
SCREENSHOT_DIR = os.path.join(DATA_DIR, "screenshots")
DAILY_REPORT_DIR = os.path.join(DATA_DIR, "reports", "daily")
WEEKLY_REPORT_DIR = os.path.join(DATA_DIR, "reports", "weekly")
SERVICE_LOG = os.path.join(DATA_DIR, "service.log")

# Case folders root (used for case number extraction from paths, not for file watching)
CASE_FOLDERS_ROOT = r"Z:\Shared\Current Clients"

# File watcher directories -- LOCAL folders only.
# The shared Z: drive is excluded because watchdog captures ALL users' events,
# not just yours. Your activity on shared files is still tracked via window titles,
# screenshots, and clipboard monitoring.
WATCHED_FOLDERS = [
    os.path.expanduser("~/Desktop"),
    os.path.expanduser("~/Documents"),
    os.path.expanduser("~/Downloads"),
]

# Capture intervals (seconds)
WINDOW_POLL_INTERVAL = 1.0
SCREENSHOT_INTERVAL = 45.0
IDLE_CHECK_INTERVAL = 5.0
IDLE_THRESHOLD = 300              # 5 minutes
OUTLOOK_POLL_INTERVAL = 30
BURST_SCREENSHOT_COUNT = 3
BURST_SCREENSHOT_INTERVAL = 10.0

# Retention
SCREENSHOT_RETENTION_HOURS = 48
LOG_RETENTION_DAYS = 30

# Screenshot settings
SCREENSHOT_QUALITY = 60  # JPEG quality (lower = smaller files)

# Gemini Vision model for screenshot analysis
VISION_MODEL = "gemini-2.0-flash"

# File extensions to track
LEGAL_EXTENSIONS = {
    '.pdf', '.docx', '.doc', '.xlsx', '.xls', '.msg', '.eml',
    '.txt', '.rtf', '.pptx', '.csv', '.jpg', '.jpeg', '.png', '.tiff',
}

# Cleanup interval (seconds)
CLEANUP_INTERVAL = 3600  # 1 hour


def ensure_dirs():
    """Create all required data directories."""
    for d in [
        os.path.dirname(DB_PATH),
        SCREENSHOT_DIR,
        DAILY_REPORT_DIR,
        WEEKLY_REPORT_DIR,
    ]:
        os.makedirs(d, exist_ok=True)
