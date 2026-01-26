"""
Unified Exception Hierarchy for iCharlotte

Provides consistent exception handling patterns across all agents
with retry decorators and error classification.
"""

import functools
import time
import traceback
from typing import Callable, List, Optional, Type, Any


# =============================================================================
# Base Exceptions
# =============================================================================

class ICharlotteException(Exception):
    """Base exception for all iCharlotte errors."""

    def __init__(self, message: str, details: Optional[dict] = None):
        super().__init__(message)
        self.message = message
        self.details = details or {}

    def __str__(self):
        if self.details:
            return f"{self.message} | Details: {self.details}"
        return self.message


# =============================================================================
# Document Processing Exceptions
# =============================================================================

class ExtractionError(ICharlotteException):
    """Raised when text extraction fails."""

    def __init__(self, message: str, file_path: str = "", page: int = None):
        super().__init__(message, {"file_path": file_path, "page": page})
        self.file_path = file_path
        self.page = page


class OCRError(ExtractionError):
    """Raised when OCR processing fails."""
    pass


class UnsupportedFormatError(ExtractionError):
    """Raised when document format is not supported."""
    pass


# =============================================================================
# LLM Exceptions
# =============================================================================

class LLMError(ICharlotteException):
    """Base exception for LLM-related errors."""

    def __init__(self, message: str, provider: str = "", model: str = "",
                 recoverable: bool = True):
        super().__init__(message, {"provider": provider, "model": model})
        self.provider = provider
        self.model = model
        self.recoverable = recoverable


class LLMRateLimitError(LLMError):
    """Raised when hitting API rate limits. Signals to wait and retry."""

    def __init__(self, message: str, provider: str = "", model: str = "",
                 retry_after: float = None):
        super().__init__(message, provider, model, recoverable=True)
        self.retry_after = retry_after or 60.0


class LLMModelUnavailableError(LLMError):
    """Model is unavailable. Try a fallback model."""

    def __init__(self, message: str, provider: str = "", model: str = ""):
        super().__init__(message, provider, model, recoverable=True)


class LLMQuotaExceededError(LLMError):
    """API quota exceeded. May need to wait longer or switch providers."""

    def __init__(self, message: str, provider: str = ""):
        super().__init__(message, provider, recoverable=False)


class LLMResponseError(LLMError):
    """Response was blocked, empty, or invalid."""

    def __init__(self, message: str, provider: str = "", model: str = "",
                 reason: str = ""):
        super().__init__(message, provider, model, recoverable=True)
        self.reason = reason


class LLMTimeoutError(LLMError):
    """Request timed out."""

    def __init__(self, message: str, provider: str = "", model: str = "",
                 timeout: float = None):
        super().__init__(message, provider, model, recoverable=True)
        self.timeout = timeout


# =============================================================================
# Processing Pass Exceptions
# =============================================================================

class PassFailedError(ICharlotteException):
    """A processing pass (extraction, summary, cross-check) failed."""

    def __init__(self, pass_name: str, message: str, recoverable: bool = True,
                 partial_result: Any = None):
        super().__init__(message, {"pass_name": pass_name, "recoverable": recoverable})
        self.pass_name = pass_name
        self.recoverable = recoverable
        self.partial_result = partial_result


class ExtractionPassError(PassFailedError):
    """Extraction pass failed."""

    def __init__(self, message: str, recoverable: bool = True, partial_result: Any = None):
        super().__init__("Extraction", message, recoverable, partial_result)


class SummaryPassError(PassFailedError):
    """Summary pass failed."""

    def __init__(self, message: str, recoverable: bool = True, partial_result: Any = None):
        super().__init__("Summary", message, recoverable, partial_result)


class CrossCheckPassError(PassFailedError):
    """Cross-check pass failed."""

    def __init__(self, message: str, recoverable: bool = True, partial_result: Any = None):
        super().__init__("CrossCheck", message, recoverable, partial_result)


# =============================================================================
# Validation Exceptions
# =============================================================================

class ValidationError(ICharlotteException):
    """Output validation failed."""

    def __init__(self, message: str, errors: List[str] = None,
                 warnings: List[str] = None):
        super().__init__(message, {"errors": errors, "warnings": warnings})
        self.errors = errors or []
        self.warnings = warnings or []


class MissingMetadataError(ValidationError):
    """Required metadata is missing from output."""
    pass


class InsufficientContentError(ValidationError):
    """Output content is too short or missing required sections."""
    pass


# =============================================================================
# Resource Exceptions
# =============================================================================

class MemoryLimitError(ICharlotteException):
    """Approaching or exceeded memory limits."""

    def __init__(self, message: str, current_mb: float = None,
                 limit_mb: float = None):
        super().__init__(message, {"current_mb": current_mb, "limit_mb": limit_mb})
        self.current_mb = current_mb
        self.limit_mb = limit_mb


class FileLockError(ICharlotteException):
    """File is locked and cannot be accessed."""

    def __init__(self, message: str, file_path: str = ""):
        super().__init__(message, {"file_path": file_path})
        self.file_path = file_path


# =============================================================================
# Retry Decorators
# =============================================================================

def retry_on_error(
    max_retries: int = 3,
    backoff_seconds: float = 2.0,
    backoff_multiplier: float = 2.0,
    exceptions: tuple = (Exception,),
    on_retry: Callable = None,
    logger: Callable = None
):
    """
    Decorator for retrying a function on specific exceptions.

    Args:
        max_retries: Maximum number of retry attempts.
        backoff_seconds: Initial delay between retries.
        backoff_multiplier: Multiplier for exponential backoff.
        exceptions: Tuple of exception types to catch and retry.
        on_retry: Optional callback called on each retry with (attempt, exception).
        logger: Optional logger function for retry messages.

    Returns:
        Decorated function with retry logic.
    """
    def decorator(func: Callable):
        @functools.wraps(func)
        def wrapper(*args, **kwargs):
            last_exception = None
            delay = backoff_seconds

            for attempt in range(max_retries + 1):
                try:
                    return func(*args, **kwargs)
                except exceptions as e:
                    last_exception = e

                    if attempt < max_retries:
                        if logger:
                            logger(f"Attempt {attempt + 1} failed: {e}. Retrying in {delay}s...")

                        if on_retry:
                            on_retry(attempt + 1, e)

                        time.sleep(delay)
                        delay *= backoff_multiplier
                    else:
                        if logger:
                            logger(f"All {max_retries + 1} attempts failed.")

            raise last_exception

        return wrapper
    return decorator


def retry_on_llm_error(
    max_retries: int = 3,
    backoff_seconds: float = 2.0,
    fallback_models: List[str] = None,
    logger: Callable = None
):
    """
    Decorator specifically for LLM calls with model fallback support.

    Args:
        max_retries: Maximum retries per model.
        backoff_seconds: Initial delay between retries.
        fallback_models: List of fallback model names to try.
        logger: Optional logger function.

    Returns:
        Decorated function with LLM-specific retry logic.
    """
    def decorator(func: Callable):
        @functools.wraps(func)
        def wrapper(*args, **kwargs):
            # Track which models to try
            models = fallback_models or []
            last_exception = None

            # Try the function with default model first
            for attempt in range(max_retries):
                try:
                    return func(*args, **kwargs)
                except LLMRateLimitError as e:
                    if logger:
                        logger(f"Rate limited: {e}. Waiting {e.retry_after}s...")
                    time.sleep(e.retry_after)
                    last_exception = e
                except LLMModelUnavailableError as e:
                    if logger:
                        logger(f"Model unavailable: {e}")
                    last_exception = e
                    break  # Try fallback models
                except LLMError as e:
                    if e.recoverable:
                        if logger:
                            logger(f"LLM error (recoverable): {e}")
                        time.sleep(backoff_seconds * (attempt + 1))
                        last_exception = e
                    else:
                        raise

            # Try fallback models
            for model in models:
                if logger:
                    logger(f"Trying fallback model: {model}")

                # Inject model into kwargs
                kwargs['model'] = model

                for attempt in range(max_retries):
                    try:
                        return func(*args, **kwargs)
                    except LLMError as e:
                        if logger:
                            logger(f"Fallback {model} attempt {attempt + 1} failed: {e}")
                        last_exception = e
                        if not e.recoverable:
                            break
                        time.sleep(backoff_seconds)

            if last_exception:
                raise last_exception
            raise LLMError("All model attempts failed", recoverable=False)

        return wrapper
    return decorator


def retry_on_file_lock(
    max_retries: int = 10,
    delay_seconds: float = 0.5,
    logger: Callable = None
):
    """
    Decorator for retrying file operations when files are locked.

    Args:
        max_retries: Maximum number of retry attempts.
        delay_seconds: Delay between retries.
        logger: Optional logger function.

    Returns:
        Decorated function with file lock retry logic.
    """
    def decorator(func: Callable):
        @functools.wraps(func)
        def wrapper(*args, **kwargs):
            last_exception = None

            for attempt in range(max_retries):
                try:
                    return func(*args, **kwargs)
                except (PermissionError, IOError, FileLockError) as e:
                    last_exception = e
                    if logger:
                        logger(f"File locked (attempt {attempt + 1}/{max_retries}): {e}")
                    time.sleep(delay_seconds)

            if last_exception:
                raise FileLockError(
                    f"File still locked after {max_retries} attempts",
                    file_path=str(getattr(last_exception, 'filename', 'unknown'))
                )

        return wrapper
    return decorator


# =============================================================================
# Exception Handlers
# =============================================================================

class ExceptionHandler:
    """
    Centralized exception handler for agents.

    Provides consistent logging and classification of errors.
    """

    def __init__(self, logger: Callable = None):
        """
        Initialize the exception handler.

        Args:
            logger: Logger function to use. Defaults to print.
        """
        self.logger = logger or print

    def handle(self, exception: Exception, context: str = "") -> dict:
        """
        Handle an exception and return structured error info.

        Args:
            exception: The exception to handle.
            context: Optional context string for the error.

        Returns:
            Dictionary with error information.
        """
        error_info = {
            "type": type(exception).__name__,
            "message": str(exception),
            "context": context,
            "recoverable": getattr(exception, 'recoverable', False),
            "traceback": traceback.format_exc()
        }

        # Add specific details based on exception type
        if isinstance(exception, ICharlotteException):
            error_info["details"] = exception.details

        # Log the error
        level = "warning" if error_info["recoverable"] else "error"
        self.logger(f"[{level.upper()}] {context}: {exception}")

        return error_info

    def classify(self, exception: Exception) -> str:
        """
        Classify an exception into a category.

        Args:
            exception: The exception to classify.

        Returns:
            Category string.
        """
        if isinstance(exception, ExtractionError):
            return "extraction"
        elif isinstance(exception, LLMError):
            return "llm"
        elif isinstance(exception, PassFailedError):
            return "processing"
        elif isinstance(exception, ValidationError):
            return "validation"
        elif isinstance(exception, MemoryLimitError):
            return "resource"
        elif isinstance(exception, FileLockError):
            return "file_access"
        else:
            return "unknown"

    def should_retry(self, exception: Exception) -> bool:
        """
        Determine if an operation should be retried based on the exception.

        Args:
            exception: The exception to check.

        Returns:
            True if the operation should be retried.
        """
        if isinstance(exception, ICharlotteException):
            return getattr(exception, 'recoverable', False)

        # Standard exceptions that might be retryable
        if isinstance(exception, (TimeoutError, ConnectionError)):
            return True

        if isinstance(exception, (PermissionError, IOError)):
            return True  # File might become available

        return False
