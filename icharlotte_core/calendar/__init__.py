"""
iCharlotte Calendar Agent Module

Monitors outgoing emails to calendartestai@gmail.com and automatically
creates Google Calendar entries based on document type and legal deadlines.
"""

from .gcal_client import GoogleCalendarClient
from .deadline_calculator import DeadlineCalculator
from .attachment_classifier import AttachmentClassifier
from .calendar_monitor import CalendarMonitorWorker

__all__ = [
    'GoogleCalendarClient',
    'DeadlineCalculator',
    'AttachmentClassifier',
    'CalendarMonitorWorker',
]
