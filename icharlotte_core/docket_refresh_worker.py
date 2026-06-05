"""
Docket Refresh Worker — Background thread that automatically re-downloads
court dockets when hearing dates pass, keeping the master case list current.

After a hearing date passes + a grace period (default 2 days), the worker
spawns docket.py --headless for the case.  docket.py itself updates the DB
and JSON with new hearing/trial dates once the download completes.
"""

import os
import sys
import json
import subprocess
import time
from datetime import datetime, date, timedelta
from PySide6.QtCore import QThread, Signal

from icharlotte_core.master_db import MasterCaseDatabase
from icharlotte_core.config import GEMINI_DATA_DIR
from icharlotte_core.utils import log_event, parse_hearing_data


class DocketRefreshWorker(QThread):
    """Background worker that monitors hearing dates and triggers docket refreshes."""

    # Signals
    docket_refreshed = Signal(str)   # file_number — docket download completed
    status = Signal(str)             # status message for logging
    error = Signal(str)              # error message

    # Configuration
    CHECK_INTERVAL = 86400           # Check once per day (seconds)
    GRACE_DAYS = 2                   # Days to wait after hearing before refreshing
    MAX_CONCURRENT = 2               # Max concurrent docket downloads
    PROCESS_TIMEOUT = 600            # 10 minutes per docket download
    STARTUP_DELAY = 90               # Defer first check N seconds after launch
                                     # so the docket-scraper storm doesn't coincide
                                     # with startup UI build + Z: scan (crash window)

    def __init__(self, db: MasterCaseDatabase = None):
        super().__init__()
        self.db = db or MasterCaseDatabase()
        self._running = True

    def run(self):
        """Main loop — check periodically for cases needing docket refresh."""
        log_event("[DocketRefresh] Worker started.")
        self.status.emit("Docket refresh worker started")

        # Defer the first check so the startup docket-scraper storm (Selenium
        # subprocesses) does not coincide with launch, when the main thread is busy
        # building the UI and scanning Z:. The 2026-06-04 crash cluster happened
        # during that window. Interruptible so request_stop() is still prompt.
        for _ in range(self.STARTUP_DELAY):
            if not self._running:
                log_event("[DocketRefresh] Worker stopped before first check.")
                return
            time.sleep(1)

        while self._running:
            try:
                self._check_and_refresh()
            except Exception as e:
                msg = f"[DocketRefresh] Error in check cycle: {e}"
                log_event(msg, "error")
                self.error.emit(msg)

            # Sleep in small increments so we can stop promptly
            for _ in range(self.CHECK_INTERVAL):
                if not self._running:
                    break
                time.sleep(1)

        log_event("[DocketRefresh] Worker stopped.")

    def request_stop(self):
        self._running = False

    # ------------------------------------------------------------------
    # Core logic
    # ------------------------------------------------------------------

    def _check_and_refresh(self):
        """Scan all cases, find those needing a docket refresh, and run downloads."""
        cases = self.db.get_all_cases()
        today = date.today()
        threshold = today - timedelta(days=self.GRACE_DAYS)
        candidates = []

        for case in cases:
            file_number = case["file_number"]
            try:
                if self._needs_refresh(case, file_number, today, threshold):
                    candidates.append(file_number)
            except Exception as e:
                log_event(f"[DocketRefresh] Error checking {file_number}: {e}", "error")

        if not candidates:
            return

        log_event(f"[DocketRefresh] {len(candidates)} case(s) need docket refresh: {candidates}")
        self.status.emit(f"Refreshing dockets for {len(candidates)} case(s)")
        self._run_docket_downloads(candidates)

    def _needs_refresh(self, case, file_number, today, threshold):
        """Determine whether a case's docket should be re-downloaded.

        Returns True when:
        1. The case has at least one hearing date that has passed the grace period.
        2. We haven't already downloaded a docket AFTER that hearing date passed.
        3. The case has a venue_county set (otherwise scraper would fail).
        """
        # --- Collect all hearing dates for this case ---
        all_dates = self._get_all_hearing_dates(case, file_number)
        if not all_dates:
            return False

        # Find the most recent hearing that is past the grace period
        most_recent_past = None
        has_future_hearing = False
        for d in all_dates:
            if d <= threshold:
                if most_recent_past is None or d > most_recent_past:
                    most_recent_past = d
            if d >= today:
                has_future_hearing = True

        if most_recent_past is None:
            return False  # No hearings have passed the grace period yet

        # If there are still future hearings on file, we might still want to
        # refresh (the court may have added new ones), but only if we haven't
        # refreshed since the most recent past hearing.
        # If NO future hearings remain, definitely refresh.

        # --- Check last docket download date ---
        last_download_str = case.get("last_docket_download", "")
        if last_download_str:
            try:
                last_download = datetime.strptime(last_download_str, "%Y-%m-%d").date()
                # Already downloaded after the hearing passed — skip
                if last_download > most_recent_past:
                    return False
            except ValueError:
                pass  # Bad date format, proceed with refresh

        # --- Check venue exists (required for docket scraper) ---
        json_path = os.path.join(GEMINI_DATA_DIR, f"{file_number}.json")
        if not os.path.exists(json_path):
            return False

        try:
            with open(json_path, "r", encoding="utf-8") as f:
                vars_data = json.load(f)
        except (json.JSONDecodeError, OSError):
            return False

        venue = self._get_json_val(vars_data, "venue_county")
        if not venue or str(venue).lower() in ("null", "n/a", "none", ""):
            return False

        return True

    def _get_all_hearing_dates(self, case, file_number):
        """Return a list of date objects for all hearing/trial dates on file."""
        dates = []

        # 1. Parse other_hearings from JSON
        json_path = os.path.join(GEMINI_DATA_DIR, f"{file_number}.json")
        if os.path.exists(json_path):
            try:
                with open(json_path, "r", encoding="utf-8") as f:
                    vars_data = json.load(f)
                oh = self._get_json_val(vars_data, "other_hearings") or self._get_json_val(vars_data, "other_hearing")
                if oh and str(oh).lower() != "none":
                    parsed = parse_hearing_data(str(oh))
                    for h in parsed:
                        if h["date_obj"].year < 9999:
                            dates.append(h["date_obj"].date())
            except (json.JSONDecodeError, OSError):
                pass

        # 2. DB next_hearing_date
        nhd = case.get("next_hearing_date", "")
        if nhd and nhd.lower() != "none":
            try:
                dates.append(datetime.strptime(nhd, "%Y-%m-%d").date())
            except ValueError:
                pass

        # 3. DB trial_date (also counts as a hearing that, if passed, warrants refresh)
        td = case.get("trial_date", "")
        if td and td.lower() != "none":
            try:
                dates.append(datetime.strptime(td, "%Y-%m-%d").date())
            except ValueError:
                pass

        # Deduplicate
        return list(set(dates))

    @staticmethod
    def _get_json_val(vars_data, key):
        v = vars_data.get(key)
        if isinstance(v, dict) and "value" in v:
            return v["value"]
        return v or ""

    # ------------------------------------------------------------------
    # Subprocess management
    # ------------------------------------------------------------------

    def _run_docket_downloads(self, file_numbers):
        """Run docket.py for each file number, respecting concurrency limit."""
        scripts_dir = os.path.join(os.getcwd(), "Scripts")
        docket_script = os.path.join(scripts_dir, "docket.py")

        if not os.path.exists(docket_script):
            self.error.emit(f"[DocketRefresh] docket.py not found at {docket_script}")
            return

        queue = list(file_numbers)
        active = {}  # pid -> (file_num, process, start_time)

        while (queue or active) and self._running:
            # Fill slots
            while queue and len(active) < self.MAX_CONCURRENT and self._running:
                fn = queue.pop(0)
                try:
                    args = [sys.executable, docket_script, fn, "--headless"]
                    kwargs = {}
                    if os.name == "nt":
                        kwargs["creationflags"] = subprocess.CREATE_NO_WINDOW
                    p = subprocess.Popen(args, **kwargs)
                    active[p.pid] = (fn, p, time.time())
                    log_event(f"[DocketRefresh] Started docket download for {fn} (pid {p.pid})")
                    self.status.emit(f"Downloading docket for {fn}...")
                except Exception as e:
                    log_event(f"[DocketRefresh] Failed to start docket for {fn}: {e}", "error")
                    self.error.emit(f"Failed to start docket for {fn}: {e}")

            # Check active processes
            to_remove = []
            for pid, (fn, p, start_time) in active.items():
                elapsed = time.time() - start_time
                if p.poll() is not None:
                    if p.returncode == 0:
                        log_event(f"[DocketRefresh] Docket refresh completed for {fn}")
                        self.docket_refreshed.emit(fn)
                    else:
                        log_event(f"[DocketRefresh] Docket refresh failed for {fn} (exit {p.returncode})", "warning")
                    to_remove.append(pid)
                elif elapsed > self.PROCESS_TIMEOUT:
                    log_event(f"[DocketRefresh] Docket download timed out for {fn}, terminating", "warning")
                    p.terminate()
                    to_remove.append(pid)

            for pid in to_remove:
                del active[pid]

            if queue or active:
                time.sleep(2)
