"""
Unified Logging System for iCharlotte Agents

Provides structured logging with pass-level progress tracking,
rotating log files, and a protocol for UI parsing.
"""

import os
import sys
import logging
import datetime
import time
from typing import Optional
from logging.handlers import RotatingFileHandler
from dataclasses import dataclass


# =============================================================================
# Log Directory Configuration
# =============================================================================

LOG_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'logs')

# Ensure log directory exists
if not os.path.exists(LOG_DIR):
    try:
        os.makedirs(LOG_DIR)
    except OSError:
        LOG_DIR = os.path.dirname(os.path.abspath(__file__))  # Fallback


# =============================================================================
# Structured Output Protocol
# =============================================================================

"""
The AgentLogger outputs structured messages that can be parsed by the UI.

Protocol Format:
    PASS_START:<pass_name>:<current>:<total>
    PASS_COMPLETE:<pass_name>:<status>:<duration_sec>
    PASS_FAILED:<pass_name>:<error>:<recoverable>
    PROGRESS:<percent>:<message>
    OUTPUT_FILE:<path>
    ERROR:<message>
    WARNING:<message>

Examples:
    PASS_START:Extraction:1:3
    PROGRESS:35:Extracting page 5 of 14
    PASS_COMPLETE:Extraction:success:12.5
    PASS_FAILED:CrossCheck:Model timeout:recoverable
    OUTPUT_FILE:C:\\path\\to\\output.docx
"""


@dataclass
class PassInfo:
    """Information about a processing pass."""
    name: str
    number: int
    total: int
    start_time: float = 0.0
    end_time: float = 0.0
    status: str = "pending"  # pending, in_progress, success, failed
    error: Optional[str] = None

    @property
    def duration(self) -> float:
        """Duration in seconds."""
        if self.end_time > 0 and self.start_time > 0:
            return self.end_time - self.start_time
        return 0.0


class AgentLogger:
    """
    Unified logger for iCharlotte agents with pass tracking.

    Features:
    - Structured output protocol for UI parsing
    - Pass-level progress tracking
    - Rotating log files
    - Thread-safe logging
    - Stdout and file output

    Usage:
        logger = AgentLogger("Discovery", file_number="1234.567")
        logger.pass_start("Extraction", 1, 3)
        logger.progress(25, "Extracting page 5 of 20")
        logger.pass_complete("Extraction", success=True)
    """

    def __init__(
        self,
        agent_name: str,
        file_number: Optional[str] = None,
        log_to_file: bool = True,
        log_to_stdout: bool = True
    ):
        """
        Initialize the agent logger.

        Args:
            agent_name: Name of the agent (e.g., "Discovery", "Deposition").
            file_number: Optional case file number for context.
            log_to_file: Whether to write to log file.
            log_to_stdout: Whether to write to stdout (for UI capture).
        """
        self.agent_name = agent_name
        self.file_number = file_number
        self.log_to_file = log_to_file
        self.log_to_stdout = log_to_stdout

        # Pass tracking
        self.current_pass: Optional[PassInfo] = None
        self.pass_history: list = []

        # Set up file logging
        self._file_logger = None
        if log_to_file:
            self._setup_file_logger()

    def _setup_file_logger(self):
        """Set up rotating file logger."""
        log_file = os.path.join(
            LOG_DIR,
            f"icharlotte_{datetime.datetime.now().strftime('%Y-%m-%d')}.log"
        )

        self._file_logger = logging.getLogger(f"icharlotte.{self.agent_name}")
        self._file_logger.setLevel(logging.DEBUG)

        # Remove existing handlers
        self._file_logger.handlers = []

        # Rotating handler: 10MB max, keep 7 files
        handler = RotatingFileHandler(
            log_file,
            maxBytes=10 * 1024 * 1024,
            backupCount=7,
            encoding='utf-8'
        )
        handler.setLevel(logging.DEBUG)

        formatter = logging.Formatter(
            '%(asctime)s [%(name)s] %(levelname)s - %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )
        handler.setFormatter(formatter)

        self._file_logger.addHandler(handler)

    def _format_prefix(self) -> str:
        """Format the log prefix with agent name and file number."""
        if self.file_number:
            return f"[{self.agent_name}] [{self.file_number}]"
        return f"[{self.agent_name}]"

    def _safe_print(self, message: str):
        """Print to stdout with encoding safety and broken pipe handling."""
        if not self.log_to_stdout:
            return

        try:
            print(message, flush=True)
        except UnicodeEncodeError:
            try:
                encoded = message.encode(
                    sys.stdout.encoding or 'utf-8',
                    errors='replace'
                ).decode(sys.stdout.encoding or 'utf-8')
                print(encoded, flush=True)
            except OSError:
                pass  # stdout pipe broken
            except Exception:
                try:
                    print(message.encode('ascii', errors='replace').decode('ascii'), flush=True)
                except OSError:
                    pass  # stdout pipe broken
        except OSError:
            pass  # stdout pipe broken (common when running multiple agents)

    def _log(self, level: str, message: str, structured: str = None):
        """
        Internal logging method.

        Args:
            level: Log level (debug, info, warning, error).
            message: Human-readable message.
            structured: Optional structured protocol message.
        """
        prefix = self._format_prefix()
        full_message = f"{prefix} {message}"

        # Log to file
        if self._file_logger:
            log_method = getattr(self._file_logger, level, self._file_logger.info)
            log_method(message)

        # Output structured message first (for UI parsing)
        if structured:
            self._safe_print(structured)

        # Output human-readable message
        self._safe_print(full_message)

    # =========================================================================
    # Standard Logging Methods
    # =========================================================================

    def debug(self, message: str):
        """Log a debug message."""
        self._log("debug", message)

    def info(self, message: str, progress: int = None):
        """
        Log an info message.

        Args:
            message: The message to log.
            progress: Optional progress percentage (0-100).
        """
        structured = None
        if progress is not None:
            structured = f"PROGRESS:{progress}:{message}"

        self._log("info", message, structured)

    def warning(self, message: str):
        """Log a warning message."""
        structured = f"WARNING:{message}"
        self._log("warning", message, structured)

    def error(self, message: str, exc_info: bool = False):
        """
        Log an error message.

        Args:
            message: The error message.
            exc_info: Whether to include exception traceback.
        """
        structured = f"ERROR:{message}"

        if exc_info:
            import traceback
            tb = traceback.format_exc()
            message = f"{message}\n{tb}"

        self._log("error", message, structured)

    # =========================================================================
    # Pass Tracking Methods
    # =========================================================================

    def pass_start(self, pass_name: str, pass_number: int, total_passes: int):
        """
        Log the start of a processing pass.

        Args:
            pass_name: Name of the pass (e.g., "Extraction", "Summary").
            pass_number: Current pass number (1-indexed).
            total_passes: Total number of passes.
        """
        # Save previous pass if exists
        if self.current_pass:
            self.pass_history.append(self.current_pass)

        # Create new pass info
        self.current_pass = PassInfo(
            name=pass_name,
            number=pass_number,
            total=total_passes,
            start_time=time.time(),
            status="in_progress"
        )

        structured = f"PASS_START:{pass_name}:{pass_number}:{total_passes}"
        message = f"Starting {pass_name} pass ({pass_number}/{total_passes})"
        self._log("info", message, structured)

    def pass_complete(self, pass_name: str, success: bool = True, details: str = None):
        """
        Log the completion of a processing pass.

        Args:
            pass_name: Name of the pass.
            success: Whether the pass succeeded.
            details: Optional additional details.
        """
        duration = 0.0

        if self.current_pass and self.current_pass.name == pass_name:
            self.current_pass.end_time = time.time()
            self.current_pass.status = "success" if success else "failed"
            duration = self.current_pass.duration
            self.pass_history.append(self.current_pass)
            self.current_pass = None

        status = "success" if success else "failed"
        structured = f"PASS_COMPLETE:{pass_name}:{status}:{duration:.1f}"

        message = f"{pass_name} pass {status}"
        if duration > 0:
            message += f" ({duration:.1f}s)"
        if details:
            message += f" - {details}"

        self._log("info", message, structured)

    def pass_failed(self, pass_name: str, error: str, recoverable: bool = True):
        """
        Log a failed processing pass.

        Args:
            pass_name: Name of the pass.
            error: Error message.
            recoverable: Whether the failure can be retried.
        """
        if self.current_pass and self.current_pass.name == pass_name:
            self.current_pass.end_time = time.time()
            self.current_pass.status = "failed"
            self.current_pass.error = error
            self.pass_history.append(self.current_pass)
            self.current_pass = None

        recoverable_str = "recoverable" if recoverable else "fatal"
        structured = f"PASS_FAILED:{pass_name}:{error}:{recoverable_str}"
        message = f"{pass_name} pass failed: {error} ({recoverable_str})"

        self._log("error", message, structured)

    # =========================================================================
    # Progress Tracking Methods
    # =========================================================================

    def progress(self, percent: int, message: str):
        """
        Log progress update.

        Args:
            percent: Progress percentage (0-100).
            message: Progress message.
        """
        structured = f"PROGRESS:{percent}:{message}"
        self._log("info", message, structured)

    def output_file(self, file_path: str):
        """
        Log the path to an output file.

        Args:
            file_path: Path to the output file.
        """
        structured = f"OUTPUT_FILE:{file_path}"
        message = f"Saved to: {file_path}"
        self._log("info", message, structured)

    # =========================================================================
    # Utility Methods
    # =========================================================================

    def get_pass_summary(self) -> dict:
        """
        Get a summary of all passes.

        Returns:
            Dictionary with pass statistics.
        """
        total = len(self.pass_history)
        if self.current_pass:
            total += 1

        successful = sum(1 for p in self.pass_history if p.status == "success")
        failed = sum(1 for p in self.pass_history if p.status == "failed")
        total_duration = sum(p.duration for p in self.pass_history)

        return {
            "total_passes": total,
            "successful": successful,
            "failed": failed,
            "total_duration": total_duration,
            "pass_details": [
                {
                    "name": p.name,
                    "status": p.status,
                    "duration": p.duration,
                    "error": p.error
                }
                for p in self.pass_history
            ]
        }


# =============================================================================
# Compatibility Layer
# =============================================================================

def create_legacy_log_event(agent_name: str, log_file: str = None):
    """
    Create a legacy log_event function compatible with existing agent scripts.

    Args:
        agent_name: Name of the agent.
        log_file: Optional legacy log file path.

    Returns:
        A log_event function with the old signature.
    """
    logger = AgentLogger(agent_name)

    # Also set up legacy file logger if path provided
    if log_file:
        legacy_logger = logging.getLogger(f"legacy.{agent_name}")
        legacy_logger.setLevel(logging.INFO)
        handler = logging.FileHandler(log_file, encoding='utf-8')
        handler.setFormatter(logging.Formatter(
            '%(asctime)s - %(levelname)s - %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        ))
        legacy_logger.addHandler(handler)

        def log_event(message: str, level: str = "info"):
            # Log to legacy file
            if level == "info":
                legacy_logger.info(message)
            elif level == "error":
                legacy_logger.error(message)
            elif level == "warning":
                legacy_logger.warning(message)

            # Also log via new system
            if level == "error":
                logger.error(message)
            elif level == "warning":
                logger.warning(message)
            else:
                logger.info(message)

        return log_event

    # No legacy file, just use new logger
    def log_event(message: str, level: str = "info"):
        if level == "error":
            logger.error(message)
        elif level == "warning":
            logger.warning(message)
        else:
            logger.info(message)

    return log_event


# =============================================================================
# Global Instance
# =============================================================================

_default_logger: Optional[AgentLogger] = None


def get_logger(agent_name: str = "Default", file_number: str = None) -> AgentLogger:
    """
    Get or create an agent logger.

    Args:
        agent_name: Name of the agent.
        file_number: Optional case file number.

    Returns:
        AgentLogger instance.
    """
    global _default_logger

    if _default_logger is None or _default_logger.agent_name != agent_name:
        _default_logger = AgentLogger(agent_name, file_number)
    elif file_number and _default_logger.file_number != file_number:
        _default_logger.file_number = file_number

    return _default_logger
