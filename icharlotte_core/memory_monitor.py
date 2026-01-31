"""
Memory Monitoring Module for iCharlotte

Provides memory usage tracking with configurable thresholds
to prevent crashes on large document processing.
"""

import os
import gc
import sys
from dataclasses import dataclass
from typing import Optional, Callable
from contextlib import contextmanager
import time


# =============================================================================
# Memory Status
# =============================================================================

@dataclass
class MemoryStatus:
    """Current memory usage status."""
    current_mb: float
    peak_mb: float
    available_mb: float
    percent_used: float
    warn_threshold_mb: float
    abort_threshold_mb: float

    @property
    def is_warning(self) -> bool:
        """Memory usage is above warning threshold."""
        return self.current_mb >= self.warn_threshold_mb

    @property
    def should_abort(self) -> bool:
        """Memory usage is above abort threshold."""
        return self.current_mb >= self.abort_threshold_mb

    @property
    def status_level(self) -> str:
        """Get status level: 'ok', 'warning', 'critical'."""
        if self.should_abort:
            return "critical"
        elif self.is_warning:
            return "warning"
        return "ok"

    def __str__(self) -> str:
        return (f"Memory: {self.current_mb:.1f}MB / {self.available_mb:.1f}MB "
                f"({self.percent_used:.1f}%) [{self.status_level}]")


@dataclass
class OperationStats:
    """Statistics for a tracked operation."""
    name: str
    start_memory_mb: float
    end_memory_mb: float
    peak_memory_mb: float
    duration_seconds: float

    @property
    def memory_delta_mb(self) -> float:
        """Change in memory usage during operation."""
        return self.end_memory_mb - self.start_memory_mb

    def __str__(self) -> str:
        return (f"Operation '{self.name}': "
                f"{self.duration_seconds:.2f}s, "
                f"Memory: {self.start_memory_mb:.1f}MB -> {self.end_memory_mb:.1f}MB "
                f"(delta: {self.memory_delta_mb:+.1f}MB, peak: {self.peak_memory_mb:.1f}MB)")


# =============================================================================
# Memory Monitor
# =============================================================================

class MemoryMonitor:
    """
    Monitor memory usage and prevent crashes on large operations.

    Features:
    - Real-time memory tracking
    - Configurable warning and abort thresholds
    - Operation tracking with before/after comparison
    - Automatic garbage collection suggestions

    Usage:
        monitor = MemoryMonitor(warn_threshold_mb=1500, abort_threshold_mb=2000)

        # Check memory status
        status = monitor.check()
        if status.should_abort:
            raise MemoryLimitError("Memory limit exceeded")

        # Track an operation
        with monitor.track_operation("Extract PDF"):
            process_large_pdf(file_path)
    """

    def __init__(
        self,
        warn_threshold_mb: float = 1500,
        abort_threshold_mb: float = 2000,
        logger: Callable = None
    ):
        """
        Initialize the memory monitor.

        Args:
            warn_threshold_mb: Memory threshold for warnings (MB).
            abort_threshold_mb: Memory threshold to abort operations (MB).
            logger: Optional logging function.
        """
        self.warn_threshold_mb = warn_threshold_mb
        self.abort_threshold_mb = abort_threshold_mb
        self.logger = logger or print

        # Track peak memory
        self._peak_memory_mb = 0.0

        # Operation history
        self._operation_history: list = []

        # Try to import psutil for accurate memory tracking
        self._psutil_available = False
        try:
            import psutil
            self._psutil_available = True
            self._process = psutil.Process(os.getpid())
        except ImportError:
            pass

    def _log(self, message: str, level: str = "info"):
        """Log a message."""
        if callable(self.logger):
            self.logger(message)
        elif hasattr(self.logger, level):
            getattr(self.logger, level)(message)
        else:
            try:
                print(message, flush=True)
            except OSError:
                pass  # stdout pipe broken

    def _get_memory_mb(self) -> float:
        """Get current memory usage in MB."""
        if self._psutil_available:
            try:
                import psutil
                return self._process.memory_info().rss / (1024 * 1024)
            except Exception:
                pass

        # Fallback: use sys.getsizeof on garbage collector tracked objects
        # This is less accurate but works without psutil
        gc.collect()
        try:
            # Get size of all tracked objects
            total = sum(sys.getsizeof(obj) for obj in gc.get_objects())
            return total / (1024 * 1024)
        except Exception:
            return 0.0

    def _get_available_memory_mb(self) -> float:
        """Get available system memory in MB."""
        if self._psutil_available:
            try:
                import psutil
                return psutil.virtual_memory().available / (1024 * 1024)
            except Exception:
                pass

        # Fallback: assume 4GB available
        return 4096.0

    def check(self) -> MemoryStatus:
        """
        Check current memory status.

        Returns:
            MemoryStatus object with current usage and thresholds.
        """
        current_mb = self._get_memory_mb()
        available_mb = self._get_available_memory_mb()

        # Update peak
        if current_mb > self._peak_memory_mb:
            self._peak_memory_mb = current_mb

        # Calculate percentage (of abort threshold)
        percent_used = (current_mb / self.abort_threshold_mb) * 100 if self.abort_threshold_mb > 0 else 0

        return MemoryStatus(
            current_mb=current_mb,
            peak_mb=self._peak_memory_mb,
            available_mb=available_mb,
            percent_used=percent_used,
            warn_threshold_mb=self.warn_threshold_mb,
            abort_threshold_mb=self.abort_threshold_mb
        )

    def should_abort(self) -> bool:
        """
        Check if current memory usage requires aborting.

        Returns:
            True if memory is above abort threshold.
        """
        return self.check().should_abort

    def is_warning(self) -> bool:
        """
        Check if current memory usage is at warning level.

        Returns:
            True if memory is above warning threshold.
        """
        return self.check().is_warning

    def suggest_gc(self) -> bool:
        """
        Suggest running garbage collection based on memory status.

        Returns:
            True if GC was run.
        """
        status = self.check()

        if status.is_warning:
            self._log(f"Memory warning: {status.current_mb:.1f}MB. Running garbage collection...")
            gc.collect()
            new_status = self.check()
            freed = status.current_mb - new_status.current_mb
            self._log(f"GC freed {freed:.1f}MB. Current: {new_status.current_mb:.1f}MB")
            return True

        return False

    def force_gc(self) -> float:
        """
        Force garbage collection and return memory freed.

        Returns:
            Memory freed in MB.
        """
        before = self._get_memory_mb()
        gc.collect()
        after = self._get_memory_mb()
        return before - after

    @contextmanager
    def track_operation(self, operation_name: str, abort_on_threshold: bool = True):
        """
        Context manager to track memory usage of an operation.

        Args:
            operation_name: Name of the operation for logging.
            abort_on_threshold: Whether to raise exception if abort threshold hit.

        Yields:
            MemoryStatus at start of operation.

        Raises:
            MemoryLimitError: If abort threshold is exceeded and abort_on_threshold is True.

        Usage:
            with monitor.track_operation("Process PDF") as start_status:
                process_pdf(file_path)
        """
        from icharlotte_core.exceptions import MemoryLimitError

        # Record start
        start_time = time.time()
        start_status = self.check()
        peak_during = start_status.current_mb

        self._log(f"Starting operation '{operation_name}' at {start_status.current_mb:.1f}MB")

        try:
            yield start_status

            # Check periodically during operation? Not easily done with context manager.
            # The caller should call check() periodically for long operations.

        finally:
            # Record end
            end_time = time.time()
            end_status = self.check()
            duration = end_time - start_time

            # Update peak if current is higher
            if end_status.current_mb > peak_during:
                peak_during = end_status.current_mb

            # Create stats
            stats = OperationStats(
                name=operation_name,
                start_memory_mb=start_status.current_mb,
                end_memory_mb=end_status.current_mb,
                peak_memory_mb=peak_during,
                duration_seconds=duration
            )

            self._operation_history.append(stats)
            self._log(str(stats))

            # Check thresholds
            if abort_on_threshold and end_status.should_abort:
                raise MemoryLimitError(
                    f"Memory limit exceeded during '{operation_name}'",
                    current_mb=end_status.current_mb,
                    limit_mb=self.abort_threshold_mb
                )

    def get_operation_history(self) -> list:
        """Get list of tracked operations."""
        return self._operation_history.copy()

    def get_peak_memory(self) -> float:
        """Get peak memory usage in MB."""
        return self._peak_memory_mb

    def reset_peak(self):
        """Reset peak memory tracking."""
        self._peak_memory_mb = self._get_memory_mb()

    def get_summary(self) -> dict:
        """
        Get summary of memory usage and operations.

        Returns:
            Dictionary with memory statistics.
        """
        status = self.check()
        return {
            "current_mb": status.current_mb,
            "peak_mb": self._peak_memory_mb,
            "available_mb": status.available_mb,
            "warn_threshold_mb": self.warn_threshold_mb,
            "abort_threshold_mb": self.abort_threshold_mb,
            "status_level": status.status_level,
            "operations_tracked": len(self._operation_history),
            "total_operation_time": sum(op.duration_seconds for op in self._operation_history)
        }


# =============================================================================
# Global Instance
# =============================================================================

_default_monitor: Optional[MemoryMonitor] = None


def get_monitor(warn_threshold_mb: float = 1500, abort_threshold_mb: float = 2000) -> MemoryMonitor:
    """
    Get or create the global memory monitor.

    Args:
        warn_threshold_mb: Warning threshold in MB.
        abort_threshold_mb: Abort threshold in MB.

    Returns:
        MemoryMonitor instance.
    """
    global _default_monitor

    if _default_monitor is None:
        _default_monitor = MemoryMonitor(warn_threshold_mb, abort_threshold_mb)

    return _default_monitor


def check_memory() -> MemoryStatus:
    """Check memory status using global monitor."""
    return get_monitor().check()


def should_abort() -> bool:
    """Check if memory is above abort threshold."""
    return get_monitor().should_abort()


@contextmanager
def track_memory(operation_name: str):
    """Track memory for an operation using global monitor."""
    with get_monitor().track_operation(operation_name) as status:
        yield status
