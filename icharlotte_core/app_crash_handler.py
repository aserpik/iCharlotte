"""
Application-Level Crash Handler for iCharlotte

Provides comprehensive crash logging for the main GUI application with:
- Full exception context capture
- Qt-specific exception handling
- QThread worker exception tracking
- System state capture (memory, open files, Qt state)
- Rotating crash logs
- Startup/shutdown logging
- Stderr capture for C-level / COM / WebEngine fatal messages
- Heartbeat file for detecting silent deaths (GPU crash, OS kill, etc.)
- Previous-session crash detection on startup
"""

import os
import sys
import datetime
import traceback
import threading
import atexit
import platform
import json
import logging
import signal
import ctypes
from logging.handlers import RotatingFileHandler
from typing import Optional, Callable, Any
from functools import wraps

# =============================================================================
# Log Directory Configuration
# =============================================================================

LOG_DIR = os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
    'logs'
)

CRASH_LOG_DIR = os.path.join(LOG_DIR, 'crashes')

# Ensure directories exist
for directory in [LOG_DIR, CRASH_LOG_DIR]:
    if not os.path.exists(directory):
        try:
            os.makedirs(directory)
        except OSError:
            pass


# =============================================================================
# Rotating Log Setup
# =============================================================================

def setup_rotating_log(name: str, filename: str, max_bytes: int = 10*1024*1024, backup_count: int = 5) -> logging.Logger:
    """Set up a rotating file logger."""
    log_path = os.path.join(LOG_DIR, filename)

    logger = logging.getLogger(name)
    logger.setLevel(logging.DEBUG)
    logger.handlers = []  # Clear existing handlers

    handler = RotatingFileHandler(
        log_path,
        maxBytes=max_bytes,
        backupCount=backup_count,
        encoding='utf-8'
    )
    handler.setLevel(logging.DEBUG)

    formatter = logging.Formatter(
        '%(asctime)s [%(levelname)s] %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    handler.setFormatter(formatter)
    logger.addHandler(handler)

    return logger


# Main application logger
app_logger = setup_rotating_log('icharlotte.app', 'app.log')

# Crash-specific logger (separate file for easy finding)
crash_logger = setup_rotating_log('icharlotte.crash', 'app_crash.log', max_bytes=5*1024*1024, backup_count=3)

# Heartbeat and stderr paths
HEARTBEAT_PATH = os.path.join(CRASH_LOG_DIR, 'heartbeat.json')
STDERR_LOG_PATH = os.path.join(LOG_DIR, 'stderr.log')


# =============================================================================
# Stderr Capture — catches C-level, COM, and WebEngine fatal messages
# =============================================================================

class StderrCapture:
    """Tee stderr to both console and a log file.

    C-level fatal exceptions (e.g. Windows code 0x8001010e), Qt/WebEngine
    GPU errors, and Chromium crash messages are written to stderr and never
    reach Python's exception system.  This captures them to a rotated log.
    """

    def __init__(self, log_path: str, original_stderr):
        self._original = original_stderr
        self._log_path = log_path
        self._file = None
        try:
            self._file = open(log_path, 'a', encoding='utf-8', errors='replace')
            self._file.write(f"\n{'='*60}\nStderr capture started: {datetime.datetime.now().isoformat()}\n{'='*60}\n")
            self._file.flush()
        except Exception:
            pass

    def write(self, text):
        # Always forward to original stderr
        if self._original:
            try:
                self._original.write(text)
            except Exception:
                pass
        # Also write to log file
        if self._file and text.strip():
            try:
                self._file.write(text)
                self._file.flush()
            except Exception:
                pass

    def flush(self):
        if self._original:
            try:
                self._original.flush()
            except Exception:
                pass
        if self._file:
            try:
                self._file.flush()
            except Exception:
                pass

    def fileno(self):
        if self._original:
            return self._original.fileno()
        raise OSError("no fileno")

    def isatty(self):
        return False

    def close(self):
        if self._file:
            try:
                self._file.close()
            except Exception:
                pass

    @property
    def encoding(self):
        return getattr(self._original, 'encoding', 'utf-8')


def install_stderr_capture():
    """Install stderr capture.  Returns the capture object."""
    capture = StderrCapture(STDERR_LOG_PATH, sys.stderr)
    sys.stderr = capture

    # Also redirect the C runtime's stderr (fd 2) to the same log file
    # so that C-level fprintf(stderr, ...) from Qt/Chromium/COM is captured
    try:
        stderr_fd_log = STDERR_LOG_PATH + '.fd2'
        c_stderr = os.open(stderr_fd_log, os.O_WRONLY | os.O_CREAT | os.O_APPEND)
        os.dup2(c_stderr, 2)
        os.close(c_stderr)
    except Exception:
        pass

    return capture


# =============================================================================
# Heartbeat — detects silent crashes on next startup
# =============================================================================

def _write_heartbeat(pid: int, checkpoints: list, context: dict):
    """Write heartbeat JSON atomically."""
    data = {
        'pid': pid,
        'timestamp': datetime.datetime.now().isoformat(),
        'clean_shutdown': False,
        'checkpoints': checkpoints[-20:] if checkpoints else [],
        'context': {k: str(v) for k, v in (context or {}).items()},
    }
    # Get memory info
    try:
        import psutil
        mem = psutil.Process(pid).memory_info()
        data['rss_mb'] = round(mem.rss / 1024 / 1024, 2)
    except Exception:
        pass
    tmp = HEARTBEAT_PATH + '.tmp'
    try:
        with open(tmp, 'w', encoding='utf-8') as f:
            json.dump(data, f, default=str)
        os.replace(tmp, HEARTBEAT_PATH)
    except Exception:
        pass


def _mark_clean_shutdown():
    """Mark the heartbeat as clean shutdown."""
    try:
        if os.path.exists(HEARTBEAT_PATH):
            with open(HEARTBEAT_PATH, 'r', encoding='utf-8') as f:
                data = json.load(f)
            data['clean_shutdown'] = True
            data['shutdown_time'] = datetime.datetime.now().isoformat()
            with open(HEARTBEAT_PATH, 'w', encoding='utf-8') as f:
                json.dump(data, f, default=str)
    except Exception:
        pass


def check_previous_crash() -> Optional[dict]:
    """Check if the previous session crashed (no clean shutdown).

    Call this at startup BEFORE installing the new heartbeat.
    Returns crash info dict if previous session died, else None.
    """
    try:
        if not os.path.exists(HEARTBEAT_PATH):
            return None
        with open(HEARTBEAT_PATH, 'r', encoding='utf-8') as f:
            data = json.load(f)
        if data.get('clean_shutdown'):
            return None
        # Previous session didn't shut down cleanly
        return data
    except Exception:
        return None


def _log_previous_crash(prev: dict):
    """Write a crash report for the previous session's silent death."""
    crash_time = datetime.datetime.now()
    last_heartbeat = prev.get('timestamp', 'unknown')
    pid = prev.get('pid', 'unknown')

    # Collect stderr evidence (Python-level + C-level fd2)
    stderr_tail = ''
    for spath in [STDERR_LOG_PATH, STDERR_LOG_PATH + '.fd2']:
        try:
            if os.path.exists(spath):
                with open(spath, 'r', encoding='utf-8', errors='replace') as f:
                    lines = f.readlines()
                    if lines:
                        stderr_tail += f'\n--- {os.path.basename(spath)} ---\n'
                        stderr_tail += ''.join(lines[-50:])
        except Exception:
            pass

    # Collect faulthandler evidence
    faulthandler_tail = ''
    fh_path = os.path.join(CRASH_LOG_DIR, 'faulthandler.log')
    try:
        if os.path.exists(fh_path):
            with open(fh_path, 'r', encoding='utf-8', errors='replace') as f:
                lines = f.readlines()
                faulthandler_tail = ''.join(lines[-80:]) if lines else ''
    except Exception:
        pass

    # Check Windows Event Log for the crashed PID
    wer_info = ''
    if platform.system() == 'Windows' and pid != 'unknown':
        try:
            wer_info = _query_windows_error_reports(pid, last_heartbeat)
        except Exception:
            pass

    # Build report
    report_lines = [
        "=" * 70,
        "SILENT CRASH DETECTED (previous session)",
        "=" * 70,
        f"Detection Time:   {crash_time.isoformat()}",
        f"Last Heartbeat:   {last_heartbeat}",
        f"Previous PID:     {pid}",
        f"Memory at death:  {prev.get('rss_mb', 'N/A')} MB RSS",
        "",
    ]

    if prev.get('checkpoints'):
        report_lines.append("Last Checkpoints:")
        for cp in prev['checkpoints'][-10:]:
            ts = cp.get('time', '?')
            desc = cp.get('description', '?')
            report_lines.append(f"  [{ts}] {desc}")
        report_lines.append("")

    if prev.get('context'):
        report_lines.append("Context:")
        for k, v in prev['context'].items():
            report_lines.append(f"  {k}: {v}")
        report_lines.append("")

    if stderr_tail.strip():
        report_lines.append("-" * 70)
        report_lines.append("STDERR (last 50 lines before crash):")
        report_lines.append("-" * 70)
        report_lines.append(stderr_tail)
        report_lines.append("")

    if faulthandler_tail.strip():
        report_lines.append("-" * 70)
        report_lines.append("FAULTHANDLER LOG (tail):")
        report_lines.append("-" * 70)
        report_lines.append(faulthandler_tail)
        report_lines.append("")

    if wer_info:
        report_lines.append("-" * 70)
        report_lines.append("WINDOWS ERROR REPORTING:")
        report_lines.append("-" * 70)
        report_lines.append(wer_info)
        report_lines.append("")

    full_text = '\n'.join(report_lines)

    # Write to crash log
    crash_logger.critical(full_text)

    # Also write a standalone crash report file
    timestamp = crash_time.strftime("%Y%m%d_%H%M%S")
    txt_path = os.path.join(CRASH_LOG_DIR, f"silent_crash_{timestamp}_{pid}.txt")
    try:
        with open(txt_path, 'w', encoding='utf-8') as f:
            f.write(full_text)
    except Exception:
        pass

    # Write JSON version
    json_path = os.path.join(CRASH_LOG_DIR, f"silent_crash_{timestamp}_{pid}.json")
    try:
        json_report = {
            'event': 'silent_crash',
            'detection_time': crash_time.isoformat(),
            'last_heartbeat': last_heartbeat,
            'pid': pid,
            'rss_mb': prev.get('rss_mb'),
            'checkpoints': prev.get('checkpoints', []),
            'context': prev.get('context', {}),
            'stderr_tail': stderr_tail[-2000:] if stderr_tail else '',
            'faulthandler_tail': faulthandler_tail[-2000:] if faulthandler_tail else '',
            'wer_info': wer_info[:2000] if wer_info else '',
        }
        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(json_report, f, indent=2, default=str)
    except Exception:
        pass

    return txt_path


def _query_windows_error_reports(pid, last_heartbeat_str: str) -> str:
    """Query Windows Event Log for application crash events near the heartbeat time."""
    try:
        import subprocess
        # Query Application Error events (Event ID 1000) from the last 24 hours
        result = subprocess.run(
            [
                'powershell', '-NoProfile', '-Command',
                'Get-WinEvent -FilterHashtable @{LogName="Application";Id=1000} '
                '-MaxEvents 10 -ErrorAction SilentlyContinue | '
                'Select-Object TimeCreated,Message | Format-List'
            ],
            capture_output=True, text=True, timeout=10
        )
        if result.stdout.strip():
            return result.stdout.strip()[:3000]
        # Also check for WER events (Event ID 1001)
        result2 = subprocess.run(
            [
                'powershell', '-NoProfile', '-Command',
                'Get-WinEvent -FilterHashtable @{LogName="Application";Id=1001} '
                '-MaxEvents 5 -ErrorAction SilentlyContinue | '
                'Select-Object TimeCreated,Message | Format-List'
            ],
            capture_output=True, text=True, timeout=10
        )
        if result2.stdout.strip():
            return result2.stdout.strip()[:3000]
    except Exception:
        pass
    return ''


# =============================================================================
# Application Crash Handler
# =============================================================================

class AppCrashHandler:
    """
    Comprehensive crash handler for the main iCharlotte application.

    Features:
    - Global exception hook for uncaught exceptions
    - Thread exception handling
    - Qt-specific exception handling
    - Checkpoint tracking for context
    - System state capture
    - Detailed crash reports (JSON + human-readable)
    """

    _instance: Optional['AppCrashHandler'] = None

    def __init__(self):
        """Initialize the crash handler."""
        self.start_time = datetime.datetime.now()
        self.checkpoints = []
        self.context = {}
        self._original_excepthook = None
        self._original_thread_excepthook = None
        self._installed = False
        self._qt_app = None
        self._main_window = None
        self._heartbeat_timer = None
        self._stderr_capture = None
        self._clean_shutdown = False

    @classmethod
    def get_instance(cls) -> 'AppCrashHandler':
        """Get or create the singleton instance."""
        if cls._instance is None:
            cls._instance = AppCrashHandler()
        return cls._instance

    def install(self, qt_app=None, main_window=None):
        """
        Install the crash handler.

        Args:
            qt_app: QApplication instance for Qt state capture
            main_window: Main window for UI state capture
        """
        if self._installed:
            return

        self._qt_app = qt_app
        self._main_window = main_window

        # --- Check for previous session crash BEFORE overwriting heartbeat ---
        prev_crash = check_previous_crash()
        if prev_crash:
            _log_previous_crash(prev_crash)
            app_logger.warning(
                f"Previous session (PID {prev_crash.get('pid')}) crashed silently. "
                f"Last heartbeat: {prev_crash.get('timestamp')}. See logs/crashes/."
            )

        # --- Install stderr capture ---
        self._stderr_capture = install_stderr_capture()

        # --- Store original hooks ---
        self._original_excepthook = sys.excepthook

        # Install our hook
        sys.excepthook = self._exception_hook

        # Handle thread exceptions (Python 3.8+)
        if hasattr(threading, 'excepthook'):
            self._original_thread_excepthook = threading.excepthook
            threading.excepthook = self._thread_exception_hook

        # --- Install signal handlers for SIGTERM, SIGBREAK, SIGABRT ---
        for sig in (signal.SIGTERM, signal.SIGABRT):
            try:
                signal.signal(sig, self._signal_handler)
            except (OSError, ValueError):
                pass
        # Windows-specific: SIGBREAK (Ctrl+Break / taskkill)
        if hasattr(signal, 'SIGBREAK'):
            try:
                signal.signal(signal.SIGBREAK, self._signal_handler)
            except (OSError, ValueError):
                pass

        # --- Register cleanup ---
        atexit.register(self._on_exit)

        # --- Write initial heartbeat ---
        _write_heartbeat(os.getpid(), self.checkpoints, self.context)

        # --- Start heartbeat thread (writes every 5 seconds) ---
        self._start_heartbeat_thread()

        self._installed = True
        self._log_startup()

    def set_qt_references(self, qt_app=None, main_window=None):
        """Update Qt references after installation."""
        if qt_app:
            self._qt_app = qt_app
        if main_window:
            self._main_window = main_window

    def _start_heartbeat_thread(self):
        """Start a daemon thread that writes heartbeat every 5 seconds."""
        def _beat():
            while not self._clean_shutdown:
                _write_heartbeat(os.getpid(), self.checkpoints, self.context)
                # Sleep in small increments so we can stop quickly
                for _ in range(50):  # 5 seconds total
                    if self._clean_shutdown:
                        break
                    threading.Event().wait(0.1)
        t = threading.Thread(target=_beat, name='heartbeat', daemon=True)
        t.start()

    def _signal_handler(self, signum, frame):
        """Handle termination signals — log before death."""
        sig_name = signal.Signals(signum).name if hasattr(signal, 'Signals') else str(signum)
        crash_time = datetime.datetime.now()
        uptime = (crash_time - self.start_time).total_seconds()

        # Write crash report
        crash_logger.critical("=" * 70)
        crash_logger.critical(f"SIGNAL RECEIVED: {sig_name} (signal {signum})")
        crash_logger.critical(f"Time: {crash_time.isoformat()}")
        crash_logger.critical(f"Uptime: {uptime:.1f}s")
        crash_logger.critical("=" * 70)

        # Dump all thread tracebacks
        crash_logger.critical("Thread tracebacks at signal time:")
        for thread_id, stack in sys._current_frames().items():
            thread_name = 'unknown'
            for t in threading.enumerate():
                if t.ident == thread_id:
                    thread_name = t.name
                    break
            crash_logger.critical(f"\n--- Thread {thread_name} (id={thread_id}) ---")
            for line in traceback.format_stack(stack):
                crash_logger.critical(line.rstrip())

        if self.checkpoints:
            crash_logger.critical("Last Checkpoints:")
            for cp in self.checkpoints[-10:]:
                crash_logger.critical(f"  [{cp['time']}] {cp['description']}")

        # Write a standalone report
        timestamp = crash_time.strftime("%Y%m%d_%H%M%S")
        txt_path = os.path.join(CRASH_LOG_DIR, f"signal_crash_{timestamp}_{os.getpid()}.txt")
        try:
            with open(txt_path, 'w', encoding='utf-8') as f:
                f.write(f"Signal {sig_name} received at {crash_time.isoformat()}\n")
                f.write(f"Uptime: {uptime:.1f}s\n\n")
                for thread_id, stack in sys._current_frames().items():
                    thread_name = 'unknown'
                    for t in threading.enumerate():
                        if t.ident == thread_id:
                            thread_name = t.name
                            break
                    f.write(f"\n--- Thread {thread_name} (id={thread_id}) ---\n")
                    f.write(''.join(traceback.format_stack(stack)))
        except Exception:
            pass

        # Mark heartbeat so we know it was a signal, not silent
        try:
            data = {
                'pid': os.getpid(),
                'timestamp': crash_time.isoformat(),
                'clean_shutdown': False,
                'death_cause': f'signal:{sig_name}',
                'checkpoints': self.checkpoints[-20:],
            }
            with open(HEARTBEAT_PATH, 'w', encoding='utf-8') as f:
                json.dump(data, f, default=str)
        except Exception:
            pass

        # Re-raise to default handler
        signal.signal(signum, signal.SIG_DFL)
        os.kill(os.getpid(), signum)

    def checkpoint(self, description: str, **extra):
        """
        Record a checkpoint for crash context.

        Args:
            description: Description of current operation
            **extra: Additional key-value pairs for context
        """
        entry = {
            'time': datetime.datetime.now().isoformat(),
            'description': description,
            **extra
        }
        self.checkpoints.append(entry)

        # Keep last 100 checkpoints
        if len(self.checkpoints) > 100:
            self.checkpoints = self.checkpoints[-100:]

        # Also log to app logger
        app_logger.debug(f"CHECKPOINT: {description}")

    def add_context(self, key: str, value: Any):
        """Add persistent context for crash reports."""
        self.context[key] = value

    def log_info(self, message: str):
        """Log an info message."""
        app_logger.info(message)

    def log_warning(self, message: str):
        """Log a warning message."""
        app_logger.warning(message)

    def log_error(self, message: str, exc_info: bool = False):
        """Log an error message."""
        if exc_info:
            app_logger.error(message, exc_info=True)
        else:
            app_logger.error(message)

    def log_debug(self, message: str):
        """Log a debug message."""
        app_logger.debug(message)

    def _log_startup(self):
        """Log application startup information."""
        startup_info = self._get_system_info()
        startup_info['event'] = 'startup'
        startup_info['start_time'] = self.start_time.isoformat()

        # Log to app logger
        app_logger.info("=" * 70)
        app_logger.info("iCharlotte Application Starting")
        app_logger.info("=" * 70)
        app_logger.info(f"Start Time: {self.start_time.isoformat()}")
        app_logger.info(f"Python: {sys.version.split()[0]}")
        app_logger.info(f"Platform: {platform.platform()}")
        app_logger.info(f"PID: {os.getpid()}")
        app_logger.info(f"Working Directory: {os.getcwd()}")

        # Log Qt version if available
        try:
            from PySide6 import __version__ as pyside_version
            app_logger.info(f"PySide6: {pyside_version}")
        except ImportError:
            pass

        # Log memory info
        mem_info = self._get_memory_info()
        if mem_info:
            app_logger.info(f"Memory: {mem_info.get('rss_mb', 'N/A')} MB RSS")

        app_logger.info("=" * 70)

    def _on_exit(self):
        """Log normal application shutdown and mark heartbeat as clean."""
        self._clean_shutdown = True
        uptime = (datetime.datetime.now() - self.start_time).total_seconds()

        app_logger.info("=" * 70)
        app_logger.info("iCharlotte Application Shutting Down (clean)")
        app_logger.info(f"Uptime: {uptime:.1f} seconds ({uptime/60:.1f} minutes)")
        app_logger.info("=" * 70)

        # Mark heartbeat as clean shutdown
        _mark_clean_shutdown()

    def _exception_hook(self, exc_type, exc_value, exc_tb):
        """Handle uncaught exceptions in main thread."""
        self._write_crash_report(exc_type, exc_value, exc_tb, 'main_thread')

        # Also dump all thread stacks for context
        self._dump_all_threads()

        # Try to show error message box if Qt is available
        self._try_show_error_dialog(exc_type, exc_value)

        # Call original hook
        if self._original_excepthook:
            self._original_excepthook(exc_type, exc_value, exc_tb)

    def _thread_exception_hook(self, args):
        """Handle uncaught exceptions in threads."""
        thread_name = args.thread.name if args.thread else 'unknown'
        self._write_crash_report(
            args.exc_type,
            args.exc_value,
            args.exc_traceback,
            f'thread:{thread_name}'
        )

        # Call original hook
        if self._original_thread_excepthook:
            self._original_thread_excepthook(args)

    def _dump_all_threads(self):
        """Dump tracebacks of all running threads to crash log."""
        try:
            crash_logger.critical("--- All thread tracebacks ---")
            for thread_id, stack in sys._current_frames().items():
                thread_name = 'unknown'
                for t in threading.enumerate():
                    if t.ident == thread_id:
                        thread_name = t.name
                        break
                crash_logger.critical(f"\n  Thread {thread_name} (id={thread_id}):")
                for line in traceback.format_stack(stack):
                    crash_logger.critical(f"    {line.rstrip()}")
        except Exception:
            pass

    def _write_crash_report(self, exc_type, exc_value, exc_tb, source: str):
        """Write a comprehensive crash report."""
        crash_time = datetime.datetime.now()
        uptime = (crash_time - self.start_time).total_seconds()

        # Format traceback
        tb_lines = traceback.format_exception(exc_type, exc_value, exc_tb)
        formatted_tb = "".join(tb_lines)

        # Build crash report
        report = {
            'event': 'crash',
            'source': source,
            'crash_time': crash_time.isoformat(),
            'uptime_seconds': uptime,
            'exception': {
                'type': exc_type.__name__,
                'message': str(exc_value),
                'traceback': formatted_tb
            },
            'context': {
                'checkpoints': self.checkpoints[-20:],  # Last 20 checkpoints
                'additional': self.context.copy()
            },
            'system': self._get_system_info(),
            'qt_state': self._get_qt_state()
        }

        # Write JSON crash report
        json_path = self._get_crash_path('.json')
        self._write_json_report(json_path, report)

        # Write human-readable crash report
        txt_path = self._get_crash_path('.txt')
        self._write_readable_report(txt_path, report)

        # Log to crash logger
        crash_logger.critical("=" * 70)
        crash_logger.critical(f"CRASH DETECTED: {exc_type.__name__}: {exc_value}")
        crash_logger.critical(f"Source: {source}")
        crash_logger.critical(f"Crash Time: {crash_time.isoformat()}")
        crash_logger.critical(f"Uptime: {uptime:.1f}s")
        crash_logger.critical(f"Report saved to: {txt_path}")
        crash_logger.critical("=" * 70)
        crash_logger.critical(formatted_tb)

        # Also log last checkpoints
        if self.checkpoints:
            crash_logger.critical("Last Checkpoints:")
            for cp in self.checkpoints[-10:]:
                crash_logger.critical(f"  [{cp['time']}] {cp['description']}")

    def _get_crash_path(self, ext: str) -> str:
        """Generate crash report file path."""
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"app_crash_{timestamp}_{os.getpid()}{ext}"
        return os.path.join(CRASH_LOG_DIR, filename)

    def _write_json_report(self, path: str, report: dict):
        """Write JSON crash report."""
        try:
            with open(path, 'w', encoding='utf-8') as f:
                json.dump(report, f, indent=2, default=str)
        except Exception as e:
            crash_logger.error(f"Failed to write JSON crash report: {e}")

    def _write_readable_report(self, path: str, report: dict):
        """Write human-readable crash report."""
        try:
            with open(path, 'w', encoding='utf-8') as f:
                f.write("=" * 70 + "\n")
                f.write("iCharlotte Application Crash Report\n")
                f.write("=" * 70 + "\n\n")

                f.write(f"Crash Time: {report['crash_time']}\n")
                f.write(f"Uptime: {report['uptime_seconds']:.1f} seconds\n")
                f.write(f"Source: {report['source']}\n\n")

                f.write("-" * 70 + "\n")
                f.write("EXCEPTION\n")
                f.write("-" * 70 + "\n")
                f.write(f"Type: {report['exception']['type']}\n")
                f.write(f"Message: {report['exception']['message']}\n\n")
                f.write("Traceback:\n")
                f.write(report['exception']['traceback'])
                f.write("\n")

                # Checkpoints
                if report['context']['checkpoints']:
                    f.write("-" * 70 + "\n")
                    f.write("RECENT CHECKPOINTS\n")
                    f.write("-" * 70 + "\n")
                    for cp in report['context']['checkpoints']:
                        f.write(f"  [{cp['time']}] {cp['description']}\n")
                    f.write("\n")

                # Additional context
                if report['context']['additional']:
                    f.write("-" * 70 + "\n")
                    f.write("ADDITIONAL CONTEXT\n")
                    f.write("-" * 70 + "\n")
                    for k, v in report['context']['additional'].items():
                        f.write(f"  {k}: {v}\n")
                    f.write("\n")

                # System info
                f.write("-" * 70 + "\n")
                f.write("SYSTEM INFORMATION\n")
                f.write("-" * 70 + "\n")
                sys_info = report['system']
                f.write(f"Python: {sys_info.get('python_version', 'N/A')}\n")
                f.write(f"Platform: {sys_info.get('platform', 'N/A')}\n")
                f.write(f"PID: {sys_info.get('pid', 'N/A')}\n")

                if sys_info.get('memory'):
                    mem = sys_info['memory']
                    f.write(f"Memory: {mem.get('rss_mb', 'N/A')} MB RSS, ")
                    f.write(f"{mem.get('vms_mb', 'N/A')} MB VMS\n")

                # Qt state
                qt_state = report.get('qt_state', {})
                if qt_state:
                    f.write("\n")
                    f.write("-" * 70 + "\n")
                    f.write("QT STATE\n")
                    f.write("-" * 70 + "\n")
                    for k, v in qt_state.items():
                        f.write(f"  {k}: {v}\n")

        except Exception as e:
            crash_logger.error(f"Failed to write readable crash report: {e}")

    def _get_system_info(self) -> dict:
        """Gather system information."""
        info = {
            'python_version': sys.version.split()[0],
            'platform': platform.platform(),
            'pid': os.getpid(),
            'cwd': os.getcwd(),
            'memory': self._get_memory_info()
        }

        # Add PySide6 version
        try:
            from PySide6 import __version__ as pyside_version
            info['pyside6_version'] = pyside_version
        except ImportError:
            pass

        return info

    def _get_memory_info(self) -> Optional[dict]:
        """Get current memory usage."""
        try:
            import psutil
            process = psutil.Process(os.getpid())
            mem = process.memory_info()
            return {
                'rss_mb': round(mem.rss / 1024 / 1024, 2),
                'vms_mb': round(mem.vms / 1024 / 1024, 2)
            }
        except ImportError:
            pass
        except Exception:
            pass
        return None

    def _get_qt_state(self) -> dict:
        """Capture Qt application state."""
        state = {}

        try:
            if self._qt_app:
                state['app_name'] = self._qt_app.applicationName() or 'iCharlotte'

            if self._main_window:
                state['window_visible'] = self._main_window.isVisible()
                state['window_size'] = f"{self._main_window.width()}x{self._main_window.height()}"

                # Get current tab if available
                try:
                    if hasattr(self._main_window, 'tab_widget'):
                        current_tab = self._main_window.tab_widget.currentIndex()
                        state['current_tab'] = current_tab
                except Exception:
                    pass

                # Get file number if available
                try:
                    if hasattr(self._main_window, 'file_number'):
                        state['file_number'] = self._main_window.file_number
                except Exception:
                    pass

        except Exception as e:
            state['error'] = str(e)

        return state

    def _try_show_error_dialog(self, exc_type, exc_value):
        """Try to show an error dialog to the user."""
        try:
            from PySide6.QtWidgets import QMessageBox, QApplication

            app = QApplication.instance()
            if app:
                msg = QMessageBox()
                msg.setIcon(QMessageBox.Icon.Critical)
                msg.setWindowTitle("iCharlotte - Critical Error")
                msg.setText(f"An unexpected error occurred:\n\n{exc_type.__name__}: {exc_value}")
                msg.setInformativeText("A crash report has been saved to the logs folder.")
                msg.setDetailedText(f"Check logs/crashes/ for detailed crash reports.\n\nLog directory: {CRASH_LOG_DIR}")
                msg.exec()
        except Exception:
            pass  # Don't fail if we can't show the dialog


# =============================================================================
# QThread Exception Wrapper
# =============================================================================

def wrap_qthread_run(run_method: Callable) -> Callable:
    """
    Decorator to wrap QThread.run() methods with exception handling.

    Usage:
        class MyThread(QThread):
            @wrap_qthread_run
            def run(self):
                # Your code here
    """
    @wraps(run_method)
    def wrapped(self, *args, **kwargs):
        try:
            return run_method(self, *args, **kwargs)
        except Exception as e:
            handler = AppCrashHandler.get_instance()

            # Log the exception
            exc_type, exc_value, exc_tb = sys.exc_info()
            thread_name = getattr(self, 'objectName', lambda: 'QThread')()

            handler._write_crash_report(
                exc_type,
                exc_value,
                exc_tb,
                f'QThread:{thread_name}'
            )

            # Re-raise to allow normal Qt exception handling
            raise

    return wrapped


class SafeThread:
    """
    Mixin class for QThread subclasses that adds automatic exception logging.

    Usage:
        class MyThread(SafeThread, QThread):
            def do_work(self):
                # Your actual work here
                pass
    """

    def run(self):
        """Override run with exception handling."""
        try:
            self.do_work()
        except Exception as e:
            handler = AppCrashHandler.get_instance()
            exc_type, exc_value, exc_tb = sys.exc_info()

            thread_name = self.objectName() if hasattr(self, 'objectName') else 'SafeThread'
            handler._write_crash_report(
                exc_type,
                exc_value,
                exc_tb,
                f'SafeThread:{thread_name}'
            )
            raise

    def do_work(self):
        """Override this method with your actual work."""
        raise NotImplementedError("Subclasses must implement do_work()")


# =============================================================================
# Convenience Functions
# =============================================================================

def install_crash_handler(qt_app=None, main_window=None) -> AppCrashHandler:
    """
    Install the application crash handler.

    Args:
        qt_app: QApplication instance
        main_window: Main window instance

    Returns:
        The AppCrashHandler instance
    """
    handler = AppCrashHandler.get_instance()
    handler.install(qt_app, main_window)
    return handler


def checkpoint(description: str, **extra):
    """
    Record a checkpoint for crash context.

    Args:
        description: What the application is currently doing
        **extra: Additional context key-value pairs
    """
    handler = AppCrashHandler.get_instance()
    handler.checkpoint(description, **extra)


def add_context(key: str, value: Any):
    """Add context that will be included in crash reports."""
    handler = AppCrashHandler.get_instance()
    handler.add_context(key, value)


def log_info(message: str):
    """Log an info message to the application log."""
    handler = AppCrashHandler.get_instance()
    handler.log_info(message)


def log_warning(message: str):
    """Log a warning message to the application log."""
    handler = AppCrashHandler.get_instance()
    handler.log_warning(message)


def log_error(message: str, exc_info: bool = False):
    """Log an error message to the application log."""
    handler = AppCrashHandler.get_instance()
    handler.log_error(message, exc_info)


def log_debug(message: str):
    """Log a debug message to the application log."""
    handler = AppCrashHandler.get_instance()
    handler.log_debug(message)


# =============================================================================
# Exception Decorator for GUI Callbacks
# =============================================================================

def safe_slot(func: Callable) -> Callable:
    """
    Decorator for Qt slot methods that catches and logs exceptions.

    Usage:
        @safe_slot
        def on_button_clicked(self):
            # This exception will be logged instead of crashing silently
            raise ValueError("Something went wrong")
    """
    @wraps(func)
    def wrapper(*args, **kwargs):
        try:
            return func(*args, **kwargs)
        except Exception as e:
            handler = AppCrashHandler.get_instance()
            exc_type, exc_value, exc_tb = sys.exc_info()

            # Log but don't write full crash report for slot errors
            handler.log_error(f"Exception in slot {func.__name__}: {exc_type.__name__}: {exc_value}", exc_info=True)

            # Also checkpoint for context
            handler.checkpoint(f"Slot exception: {func.__name__}", error=str(e))

            # Re-raise by default (can be changed to silent fail if needed)
            raise

    return wrapper
