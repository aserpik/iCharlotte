"""
Crash Handler for iCharlotte Agents

Provides robust crash logging with full context capture for debugging
agent failures that would otherwise be silent.
"""

import os
import sys
import datetime
import traceback
import threading
import atexit
import platform
from typing import Optional, Callable
import json

# =============================================================================
# Crash Log Configuration
# =============================================================================

CRASH_LOG_DIR = os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
    'logs',
    'crashes'
)

# Ensure crash log directory exists
if not os.path.exists(CRASH_LOG_DIR):
    try:
        os.makedirs(CRASH_LOG_DIR)
    except OSError:
        CRASH_LOG_DIR = os.path.dirname(os.path.abspath(__file__))


# =============================================================================
# Crash Handler Class
# =============================================================================

class CrashHandler:
    """
    Global crash handler that captures unhandled exceptions and writes
    detailed crash reports.

    Usage:
        handler = CrashHandler("MedRecord", input_file="/path/to/file.pdf")
        handler.install()

        # Your code here - any unhandled exception will be logged
    """

    _instance: Optional['CrashHandler'] = None

    def __init__(
        self,
        agent_name: str,
        input_file: str = "",
        additional_context: dict = None
    ):
        """
        Initialize the crash handler.

        Args:
            agent_name: Name of the agent (e.g., "MedRecord", "Discovery")
            input_file: Path to input file being processed
            additional_context: Any additional context to include in crash reports
        """
        self.agent_name = agent_name
        self.input_file = input_file
        self.additional_context = additional_context or {}
        self.start_time = datetime.datetime.now()
        self.last_checkpoint = ""
        self.progress_history = []
        self._original_excepthook = None
        self._original_thread_excepthook = None
        self._installed = False

    @classmethod
    def get_instance(cls) -> Optional['CrashHandler']:
        """Get the global crash handler instance."""
        return cls._instance

    def install(self):
        """Install the crash handler as the global exception hook."""
        if self._installed:
            return

        CrashHandler._instance = self

        # Store original hooks
        self._original_excepthook = sys.excepthook

        # Install our hook
        sys.excepthook = self._exception_hook

        # Handle thread exceptions (Python 3.8+)
        if hasattr(threading, 'excepthook'):
            self._original_thread_excepthook = threading.excepthook
            threading.excepthook = self._thread_exception_hook

        # Register cleanup on normal exit
        atexit.register(self._cleanup)

        self._installed = True
        self._log_startup()

    def uninstall(self):
        """Uninstall the crash handler."""
        if not self._installed:
            return

        if self._original_excepthook:
            sys.excepthook = self._original_excepthook

        if self._original_thread_excepthook and hasattr(threading, 'excepthook'):
            threading.excepthook = self._original_thread_excepthook

        self._installed = False
        CrashHandler._instance = None

    def checkpoint(self, description: str, progress: int = None):
        """
        Record a checkpoint for crash context.

        Args:
            description: Description of current operation
            progress: Optional progress percentage
        """
        timestamp = datetime.datetime.now().isoformat()
        checkpoint = {
            "time": timestamp,
            "description": description,
            "progress": progress
        }
        self.progress_history.append(checkpoint)
        self.last_checkpoint = description

        # Keep only last 50 checkpoints to avoid memory bloat
        if len(self.progress_history) > 50:
            self.progress_history = self.progress_history[-50:]

    def add_context(self, key: str, value):
        """Add additional context for crash reports."""
        self.additional_context[key] = value

    def _log_startup(self):
        """Log agent startup to crash log."""
        log_file = self._get_log_path("startup")

        startup_info = {
            "event": "startup",
            "agent": self.agent_name,
            "input_file": self.input_file,
            "start_time": self.start_time.isoformat(),
            "python_version": sys.version,
            "platform": platform.platform(),
            "pid": os.getpid(),
            "context": self.additional_context
        }

        self._write_log(log_file, startup_info)

    def _exception_hook(self, exc_type, exc_value, exc_tb):
        """Handle uncaught exceptions."""
        self._write_crash_report(exc_type, exc_value, exc_tb, "main")

        # Also print to stderr for any capture (may fail if pipe is broken)
        try:
            print(f"\n{'='*60}", file=sys.stderr)
            print(f"CRASH: {self.agent_name} agent crashed!", file=sys.stderr)
            print(f"Error: {exc_type.__name__}: {exc_value}", file=sys.stderr)
            print(f"Crash log written to: {CRASH_LOG_DIR}", file=sys.stderr)
            print(f"{'='*60}\n", file=sys.stderr)
        except OSError:
            pass  # stderr pipe broken

        # Call original hook
        if self._original_excepthook:
            self._original_excepthook(exc_type, exc_value, exc_tb)

    def _thread_exception_hook(self, args):
        """Handle uncaught exceptions in threads."""
        self._write_crash_report(
            args.exc_type,
            args.exc_value,
            args.exc_traceback,
            f"thread:{args.thread.name if args.thread else 'unknown'}"
        )

        # Call original hook
        if self._original_thread_excepthook:
            self._original_thread_excepthook(args)

    def _write_crash_report(self, exc_type, exc_value, exc_tb, source: str):
        """Write a detailed crash report."""
        crash_time = datetime.datetime.now()
        log_file = self._get_log_path("crash")

        # Format traceback
        tb_lines = traceback.format_exception(exc_type, exc_value, exc_tb)
        formatted_tb = "".join(tb_lines)

        # Collect memory info if available
        memory_info = self._get_memory_info()

        crash_report = {
            "event": "crash",
            "agent": self.agent_name,
            "source": source,
            "crash_time": crash_time.isoformat(),
            "uptime_seconds": (crash_time - self.start_time).total_seconds(),
            "exception": {
                "type": exc_type.__name__,
                "message": str(exc_value),
                "traceback": formatted_tb
            },
            "context": {
                "input_file": self.input_file,
                "last_checkpoint": self.last_checkpoint,
                "progress_history": self.progress_history[-10:],  # Last 10 checkpoints
                "additional": self.additional_context
            },
            "system": {
                "python_version": sys.version,
                "platform": platform.platform(),
                "pid": os.getpid(),
                "memory": memory_info
            }
        }

        self._write_log(log_file, crash_report)

        # Also write a human-readable version
        readable_file = self._get_log_path("crash", ext=".txt")
        self._write_readable_crash_report(readable_file, crash_report)

    def _write_readable_crash_report(self, file_path: str, report: dict):
        """Write a human-readable crash report."""
        try:
            with open(file_path, 'w', encoding='utf-8') as f:
                f.write("=" * 70 + "\n")
                f.write(f"iCharlotte Agent Crash Report\n")
                f.write("=" * 70 + "\n\n")

                f.write(f"Agent: {report['agent']}\n")
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

                f.write("-" * 70 + "\n")
                f.write("CONTEXT\n")
                f.write("-" * 70 + "\n")
                f.write(f"Input File: {report['context']['input_file']}\n")
                f.write(f"Last Checkpoint: {report['context']['last_checkpoint']}\n\n")

                if report['context']['progress_history']:
                    f.write("Recent Progress:\n")
                    for cp in report['context']['progress_history']:
                        f.write(f"  [{cp['time']}] {cp['description']}")
                        if cp.get('progress') is not None:
                            f.write(f" ({cp['progress']}%)")
                        f.write("\n")
                    f.write("\n")

                if report['context']['additional']:
                    f.write("Additional Context:\n")
                    for k, v in report['context']['additional'].items():
                        f.write(f"  {k}: {v}\n")
                    f.write("\n")

                f.write("-" * 70 + "\n")
                f.write("SYSTEM\n")
                f.write("-" * 70 + "\n")
                f.write(f"PID: {report['system']['pid']}\n")
                f.write(f"Platform: {report['system']['platform']}\n")
                f.write(f"Python: {report['system']['python_version'].split()[0]}\n")

                if report['system']['memory']:
                    mem = report['system']['memory']
                    f.write(f"Memory: {mem.get('rss_mb', 'N/A')} MB RSS, "
                           f"{mem.get('vms_mb', 'N/A')} MB VMS\n")

        except Exception as e:
            try:
                print(f"Failed to write readable crash report: {e}", file=sys.stderr)
            except OSError:
                pass  # stderr pipe broken

    def _get_log_path(self, event_type: str, ext: str = ".json") -> str:
        """Generate a log file path."""
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"{self.agent_name}_{event_type}_{timestamp}_{os.getpid()}{ext}"
        return os.path.join(CRASH_LOG_DIR, filename)

    def _write_log(self, file_path: str, data: dict):
        """Write data to log file."""
        try:
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=2, default=str)
        except Exception as e:
            try:
                print(f"Failed to write crash log: {e}", file=sys.stderr)
            except OSError:
                pass  # stderr pipe broken

    def _get_memory_info(self) -> Optional[dict]:
        """Get current memory usage info."""
        try:
            import psutil
            process = psutil.Process(os.getpid())
            mem = process.memory_info()
            return {
                "rss_mb": round(mem.rss / 1024 / 1024, 2),
                "vms_mb": round(mem.vms / 1024 / 1024, 2)
            }
        except ImportError:
            pass
        except Exception:
            pass
        return None

    def _cleanup(self):
        """Cleanup on normal exit."""
        pass  # Could log successful completion if needed


# =============================================================================
# Convenience Functions
# =============================================================================

def install_crash_handler(
    agent_name: str,
    input_file: str = "",
    **context
) -> CrashHandler:
    """
    Install a crash handler for an agent script.

    Args:
        agent_name: Name of the agent
        input_file: Path to input file being processed
        **context: Additional context key-value pairs

    Returns:
        The installed CrashHandler instance
    """
    handler = CrashHandler(agent_name, input_file, context)
    handler.install()
    return handler


def checkpoint(description: str, progress: int = None):
    """
    Record a checkpoint in the global crash handler.

    Args:
        description: Description of current operation
        progress: Optional progress percentage
    """
    handler = CrashHandler.get_instance()
    if handler:
        handler.checkpoint(description, progress)


def add_crash_context(key: str, value):
    """
    Add context to the global crash handler.

    Args:
        key: Context key
        value: Context value
    """
    handler = CrashHandler.get_instance()
    if handler:
        handler.add_context(key, value)


# =============================================================================
# Memory Monitor Integration
# =============================================================================

class MemoryGuard:
    """
    Context manager that monitors memory usage and raises an exception
    if memory exceeds threshold.
    """

    def __init__(
        self,
        warn_threshold_mb: int = 1500,
        abort_threshold_mb: int = 2500,
        check_interval_items: int = 10,
        logger: Callable = None
    ):
        """
        Initialize memory guard.

        Args:
            warn_threshold_mb: Memory threshold for warnings
            abort_threshold_mb: Memory threshold for aborting
            check_interval_items: Check memory every N items
            logger: Optional logger function
        """
        self.warn_threshold_mb = warn_threshold_mb
        self.abort_threshold_mb = abort_threshold_mb
        self.check_interval = check_interval_items
        self.logger = logger or print
        self.item_count = 0
        self._has_psutil = False

        try:
            import psutil
            self._has_psutil = True
        except ImportError:
            pass

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        pass

    def check(self, force: bool = False) -> bool:
        """
        Check memory usage. Call this periodically during processing.

        Args:
            force: Force check regardless of interval

        Returns:
            True if memory is OK, False if approaching limit

        Raises:
            MemoryError: If memory exceeds abort threshold
        """
        self.item_count += 1

        if not force and self.item_count % self.check_interval != 0:
            return True

        if not self._has_psutil:
            return True

        try:
            import psutil
            import gc

            process = psutil.Process(os.getpid())
            rss_mb = process.memory_info().rss / 1024 / 1024

            if rss_mb >= self.abort_threshold_mb:
                # Force garbage collection
                gc.collect()
                rss_mb = process.memory_info().rss / 1024 / 1024

                if rss_mb >= self.abort_threshold_mb:
                    checkpoint(f"Memory abort: {rss_mb:.0f}MB")
                    raise MemoryError(
                        f"Memory usage ({rss_mb:.0f}MB) exceeded abort threshold "
                        f"({self.abort_threshold_mb}MB)"
                    )

            elif rss_mb >= self.warn_threshold_mb:
                self.logger(f"WARNING: High memory usage: {rss_mb:.0f}MB")
                gc.collect()
                return False

            return True

        except ImportError:
            return True
        except MemoryError:
            raise
        except Exception:
            return True
