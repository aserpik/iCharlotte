"""Thread-safe SQLite database for work activity events."""

import sqlite3
import threading
import os
from datetime import datetime, timedelta

from . import config

SCHEMA_SQL = """
CREATE TABLE IF NOT EXISTS window_events (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    timestamp DATETIME NOT NULL,
    app_name TEXT,
    window_title TEXT,
    duration_seconds REAL,
    case_number TEXT
);

CREATE TABLE IF NOT EXISTS screenshots (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    timestamp DATETIME NOT NULL,
    file_path TEXT,
    vision_description TEXT,
    app_name TEXT,
    window_title TEXT,
    case_number TEXT
);

CREATE TABLE IF NOT EXISTS clipboard_events (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    timestamp DATETIME NOT NULL,
    content_type TEXT,
    content_preview TEXT,
    source_app TEXT,
    source_window TEXT,
    destination_app TEXT,
    destination_window TEXT,
    case_number TEXT
);

CREATE TABLE IF NOT EXISTS file_events (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    timestamp DATETIME NOT NULL,
    event_type TEXT,
    file_path TEXT,
    app_name TEXT,
    case_number TEXT
);

CREATE TABLE IF NOT EXISTS document_sessions (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    timestamp_start DATETIME,
    timestamp_end DATETIME,
    file_path TEXT,
    app_name TEXT,
    duration_seconds REAL,
    save_count INTEGER DEFAULT 0,
    case_number TEXT
);

CREATE TABLE IF NOT EXISTS email_events (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    timestamp DATETIME NOT NULL,
    event_type TEXT,
    subject TEXT,
    recipients TEXT,
    folder TEXT,
    has_attachment BOOLEAN,
    body_preview TEXT,
    case_number TEXT
);

CREATE TABLE IF NOT EXISTS idle_periods (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    start_time DATETIME NOT NULL,
    end_time DATETIME,
    idle_type TEXT,
    duration_seconds REAL
);

CREATE TABLE IF NOT EXISTS report_metadata (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    report_type TEXT NOT NULL,
    report_date TEXT NOT NULL,
    file_path TEXT,
    generated_at DATETIME,
    summary TEXT
);

CREATE INDEX IF NOT EXISTS idx_window_events_ts ON window_events(timestamp);
CREATE INDEX IF NOT EXISTS idx_screenshots_ts ON screenshots(timestamp);
CREATE INDEX IF NOT EXISTS idx_clipboard_events_ts ON clipboard_events(timestamp);
CREATE INDEX IF NOT EXISTS idx_file_events_ts ON file_events(timestamp);
CREATE INDEX IF NOT EXISTS idx_email_events_ts ON email_events(timestamp);
CREATE INDEX IF NOT EXISTS idx_idle_periods_start ON idle_periods(start_time);
CREATE INDEX IF NOT EXISTS idx_window_events_case ON window_events(case_number);
CREATE INDEX IF NOT EXISTS idx_email_events_case ON email_events(case_number);
CREATE UNIQUE INDEX IF NOT EXISTS idx_email_events_dedup
    ON email_events(event_type, subject, recipients);
"""


class ActivityDB:
    """Thread-safe SQLite database for work activity events."""

    def __init__(self, db_path: str = None):
        self.db_path = db_path or config.DB_PATH
        self._write_lock = threading.Lock()
        self._init_db()

    def _get_connection(self) -> sqlite3.Connection:
        """Create a new connection (each thread should get its own)."""
        conn = sqlite3.connect(self.db_path, timeout=10)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA journal_mode=WAL")
        conn.execute("PRAGMA busy_timeout=5000")
        return conn

    def _init_db(self):
        """Create all tables if they don't exist."""
        os.makedirs(os.path.dirname(self.db_path), exist_ok=True)
        conn = self._get_connection()
        try:
            conn.executescript(SCHEMA_SQL)
            conn.commit()
        finally:
            conn.close()

    # ---- Write methods (all acquire _write_lock) ----

    def insert_window_event(self, *, timestamp, app_name, window_title,
                            duration_seconds=None, case_number=None) -> int:
        with self._write_lock:
            conn = self._get_connection()
            try:
                cursor = conn.execute(
                    "INSERT INTO window_events (timestamp, app_name, window_title, "
                    "duration_seconds, case_number) VALUES (?, ?, ?, ?, ?)",
                    (timestamp, app_name, window_title, duration_seconds, case_number)
                )
                conn.commit()
                return cursor.lastrowid
            finally:
                conn.close()

    def update_window_event_duration(self, event_id: int, duration_seconds: float):
        """Update the duration of a window event (called on window switch)."""
        with self._write_lock:
            conn = self._get_connection()
            try:
                conn.execute(
                    "UPDATE window_events SET duration_seconds = ? WHERE id = ?",
                    (duration_seconds, event_id)
                )
                conn.commit()
            finally:
                conn.close()

    def insert_screenshot(self, *, timestamp, file_path, vision_description=None,
                          app_name=None, window_title=None, case_number=None) -> int:
        with self._write_lock:
            conn = self._get_connection()
            try:
                cursor = conn.execute(
                    "INSERT INTO screenshots (timestamp, file_path, vision_description, "
                    "app_name, window_title, case_number) VALUES (?, ?, ?, ?, ?, ?)",
                    (timestamp, file_path, vision_description, app_name,
                     window_title, case_number)
                )
                conn.commit()
                return cursor.lastrowid
            finally:
                conn.close()

    def update_screenshot_description(self, screenshot_id: int, description: str):
        """Update vision description after async Gemini analysis."""
        with self._write_lock:
            conn = self._get_connection()
            try:
                conn.execute(
                    "UPDATE screenshots SET vision_description = ? WHERE id = ?",
                    (description, screenshot_id)
                )
                conn.commit()
            finally:
                conn.close()

    def insert_clipboard_event(self, *, timestamp, content_type=None,
                               content_preview=None, source_app=None,
                               source_window=None, destination_app=None,
                               destination_window=None, case_number=None) -> int:
        with self._write_lock:
            conn = self._get_connection()
            try:
                cursor = conn.execute(
                    "INSERT INTO clipboard_events (timestamp, content_type, content_preview, "
                    "source_app, source_window, destination_app, destination_window, "
                    "case_number) VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
                    (timestamp, content_type, content_preview, source_app,
                     source_window, destination_app, destination_window, case_number)
                )
                conn.commit()
                return cursor.lastrowid
            finally:
                conn.close()

    def insert_file_event(self, *, timestamp, event_type, file_path,
                          app_name=None, case_number=None) -> int:
        with self._write_lock:
            conn = self._get_connection()
            try:
                cursor = conn.execute(
                    "INSERT INTO file_events (timestamp, event_type, file_path, "
                    "app_name, case_number) VALUES (?, ?, ?, ?, ?)",
                    (timestamp, event_type, file_path, app_name, case_number)
                )
                conn.commit()
                return cursor.lastrowid
            finally:
                conn.close()

    def insert_document_session(self, *, timestamp_start, timestamp_end=None,
                                file_path=None, app_name=None,
                                duration_seconds=None, save_count=0,
                                case_number=None) -> int:
        with self._write_lock:
            conn = self._get_connection()
            try:
                cursor = conn.execute(
                    "INSERT INTO document_sessions (timestamp_start, timestamp_end, "
                    "file_path, app_name, duration_seconds, save_count, case_number) "
                    "VALUES (?, ?, ?, ?, ?, ?, ?)",
                    (timestamp_start, timestamp_end, file_path, app_name,
                     duration_seconds, save_count, case_number)
                )
                conn.commit()
                return cursor.lastrowid
            finally:
                conn.close()

    def insert_email_event(self, *, timestamp, event_type, subject=None,
                           recipients=None, folder=None, has_attachment=False,
                           body_preview=None, case_number=None) -> int:
        with self._write_lock:
            conn = self._get_connection()
            try:
                cursor = conn.execute(
                    "INSERT OR IGNORE INTO email_events (timestamp, event_type, subject, "
                    "recipients, folder, has_attachment, body_preview, case_number) "
                    "VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
                    (timestamp, event_type, subject, recipients, folder,
                     has_attachment, body_preview, case_number)
                )
                conn.commit()
                return cursor.lastrowid
            finally:
                conn.close()

    def insert_idle_period(self, *, start_time, end_time=None,
                           idle_type=None, duration_seconds=None) -> int:
        with self._write_lock:
            conn = self._get_connection()
            try:
                cursor = conn.execute(
                    "INSERT INTO idle_periods (start_time, end_time, idle_type, "
                    "duration_seconds) VALUES (?, ?, ?, ?)",
                    (start_time, end_time, idle_type, duration_seconds)
                )
                conn.commit()
                return cursor.lastrowid
            finally:
                conn.close()

    def insert_report_metadata(self, *, report_type, report_date, file_path=None,
                               generated_at=None, summary=None) -> int:
        with self._write_lock:
            conn = self._get_connection()
            try:
                cursor = conn.execute(
                    "INSERT INTO report_metadata (report_type, report_date, file_path, "
                    "generated_at, summary) VALUES (?, ?, ?, ?, ?)",
                    (report_type, report_date, file_path, generated_at, summary)
                )
                conn.commit()
                return cursor.lastrowid
            finally:
                conn.close()

    # ---- Read methods (no lock needed, WAL allows concurrent reads) ----

    def get_daily_events(self, date_str: str) -> dict:
        """Get all events for a given date (YYYY-MM-DD)."""
        conn = self._get_connection()
        try:
            result = {}
            for table in ['window_events', 'screenshots', 'clipboard_events',
                          'file_events', 'document_sessions', 'email_events',
                          'idle_periods']:
                ts_col = 'start_time' if table == 'idle_periods' else (
                    'timestamp_start' if table == 'document_sessions' else 'timestamp'
                )
                rows = conn.execute(
                    f"SELECT * FROM {table} WHERE {ts_col} LIKE ? ORDER BY {ts_col}",
                    (f"{date_str}%",)
                ).fetchall()
                result[table] = [dict(row) for row in rows]
            return result
        finally:
            conn.close()

    def get_weekly_events(self, start_date: str, end_date: str) -> dict:
        """Get all events between two dates (inclusive)."""
        conn = self._get_connection()
        try:
            result = {}
            for table in ['window_events', 'screenshots', 'clipboard_events',
                          'file_events', 'document_sessions', 'email_events',
                          'idle_periods']:
                ts_col = 'start_time' if table == 'idle_periods' else (
                    'timestamp_start' if table == 'document_sessions' else 'timestamp'
                )
                rows = conn.execute(
                    f"SELECT * FROM {table} WHERE {ts_col} >= ? AND {ts_col} < ? "
                    f"ORDER BY {ts_col}",
                    (start_date, end_date)
                ).fetchall()
                result[table] = [dict(row) for row in rows]
            return result
        finally:
            conn.close()

    # ---- Cleanup methods ----

    def clear_screenshot_path(self, file_path: str):
        """Null out file_path for a deleted screenshot (keeps vision description)."""
        with self._write_lock:
            conn = self._get_connection()
            try:
                conn.execute(
                    "UPDATE screenshots SET file_path = NULL WHERE file_path = ?",
                    (file_path,)
                )
                conn.commit()
            finally:
                conn.close()

    def delete_records_before(self, date_str: str):
        """Delete event records older than date_str (YYYY-MM-DD)."""
        with self._write_lock:
            conn = self._get_connection()
            try:
                for table, ts_col in [
                    ('window_events', 'timestamp'),
                    ('screenshots', 'timestamp'),
                    ('clipboard_events', 'timestamp'),
                    ('file_events', 'timestamp'),
                    ('document_sessions', 'timestamp_start'),
                    ('email_events', 'timestamp'),
                    ('idle_periods', 'start_time'),
                ]:
                    conn.execute(
                        f"DELETE FROM {table} WHERE {ts_col} < ?",
                        (date_str,)
                    )
                conn.commit()
            finally:
                conn.close()
