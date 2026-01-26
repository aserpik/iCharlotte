"""
iCharlotte Calendar Agent Module

Monitors outgoing emails to calendartestai@gmail.com and automatically
creates Google Calendar entries based on document type and legal deadlines.
"""

# Lazy imports to handle missing optional dependencies (like google-api-python-client)
# This allows the module to be imported even when Google libraries aren't installed

def __getattr__(name):
    """Lazy import attributes to handle missing dependencies gracefully."""
    if name == 'GoogleCalendarClient':
        try:
            from .gcal_client import GoogleCalendarClient
            return GoogleCalendarClient
        except ImportError as e:
            raise ImportError(
                f"GoogleCalendarClient requires google-api-python-client. "
                f"Install with: pip install google-api-python-client google-auth-httplib2 google-auth-oauthlib"
            ) from e
    elif name == 'DeadlineCalculator':
        from .deadline_calculator import DeadlineCalculator
        return DeadlineCalculator
    elif name == 'AttachmentClassifier':
        from .attachment_classifier import AttachmentClassifier
        return AttachmentClassifier
    elif name == 'CalendarMonitorWorker':
        from .calendar_monitor import CalendarMonitorWorker
        return CalendarMonitorWorker
    raise AttributeError(f"module {__name__!r} has no attribute {name!r}")


__all__ = [
    'GoogleCalendarClient',
    'DeadlineCalculator',
    'AttachmentClassifier',
    'CalendarMonitorWorker',
]
