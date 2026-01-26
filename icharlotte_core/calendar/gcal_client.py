"""
Google Calendar API Client for iCharlotte Calendar Agent.

Handles OAuth2 authentication and calendar event creation.
"""

import os
from datetime import datetime, timedelta
from pathlib import Path
from typing import Optional, List, Dict, Any

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

from icharlotte_core.ui.logs_tab import LogManager


# OAuth scopes for calendar access
SCOPES = ['https://www.googleapis.com/auth/calendar.events']

# Path to credentials and token files
CALENDAR_DIR = Path(__file__).parent
CREDENTIALS_FILE = CALENDAR_DIR / 'credentials.json'
TOKEN_FILE = CALENDAR_DIR / 'token.json'


class GoogleCalendarClient:
    """
    Google Calendar API client with OAuth2 authentication.

    Usage:
        client = GoogleCalendarClient()
        if client.authenticate():
            client.create_event(
                title="MSJ Opposition Due",
                date=datetime(2026, 3, 3),
                description="File: 3200.284\\nMotion for Summary Judgment Opposition",
                file_number="3200.284"
            )
    """

    def __init__(self, calendar_id: str = 'primary'):
        """
        Initialize the Google Calendar client.

        Args:
            calendar_id: The calendar ID to use (default: 'primary' for main calendar)
        """
        self.calendar_id = calendar_id
        self.service = None
        self.credentials = None
        self.log = LogManager()

    def authenticate(self) -> bool:
        """
        Authenticate with Google Calendar API using OAuth2.

        On first run, opens a browser for user authorization.
        Subsequent runs use the stored token.

        Returns:
            True if authentication successful, False otherwise
        """
        try:
            # Check for existing token
            if TOKEN_FILE.exists():
                self.credentials = Credentials.from_authorized_user_file(
                    str(TOKEN_FILE), SCOPES
                )

            # Refresh or get new credentials if needed
            if not self.credentials or not self.credentials.valid:
                if self.credentials and self.credentials.expired and self.credentials.refresh_token:
                    self.log.add_log("Calendar", "Refreshing Google Calendar token...")
                    self.credentials.refresh(Request())
                else:
                    if not CREDENTIALS_FILE.exists():
                        self.log.add_log("Calendar", f"ERROR: credentials.json not found at {CREDENTIALS_FILE}")
                        return False

                    self.log.add_log("Calendar", "Opening browser for Google Calendar authorization...")
                    flow = InstalledAppFlow.from_client_secrets_file(
                        str(CREDENTIALS_FILE), SCOPES
                    )
                    self.credentials = flow.run_local_server(port=0)

                # Save the credentials for future runs
                with open(TOKEN_FILE, 'w') as token:
                    token.write(self.credentials.to_json())
                self.log.add_log("Calendar", "Google Calendar token saved")

            # Build the calendar service
            self.service = build('calendar', 'v3', credentials=self.credentials)
            self.log.add_log("Calendar", "Google Calendar authenticated successfully")
            return True

        except Exception as e:
            self.log.add_log("Calendar", f"Authentication error: {e}")
            return False

    def create_event(
        self,
        title: str,
        date: datetime,
        description: str = "",
        file_number: str = "",
        email_link: str = "",
        all_day: bool = True,
        reminder_minutes: int = 1440  # 24 hours before
    ) -> Optional[str]:
        """
        Create a calendar event.

        Args:
            title: Event title (will be prefixed with file number if provided)
            date: Event date
            description: Event description
            file_number: Case file number (####.###)
            email_link: Link to the original email
            all_day: If True, creates all-day event; otherwise timed event
            reminder_minutes: Minutes before event to send reminder

        Returns:
            Event ID if successful, None otherwise
        """
        if not self.service:
            if not self.authenticate():
                return None

        try:
            # Format title with file number
            if file_number:
                event_title = f"[{file_number}] {title}"
            else:
                event_title = title

            # Build description with all relevant info
            desc_parts = []
            if file_number:
                desc_parts.append(f"File Number: {file_number}")
            if description:
                desc_parts.append(description)
            if email_link:
                desc_parts.append(f"\nOriginal Email: {email_link}")

            full_description = "\n".join(desc_parts)

            # Build event body
            if all_day:
                event = {
                    'summary': event_title,
                    'description': full_description,
                    'start': {
                        'date': date.strftime('%Y-%m-%d'),
                        'timeZone': 'America/Los_Angeles',
                    },
                    'end': {
                        'date': date.strftime('%Y-%m-%d'),
                        'timeZone': 'America/Los_Angeles',
                    },
                    'reminders': {
                        'useDefault': False,
                        'overrides': [
                            {'method': 'popup', 'minutes': reminder_minutes},
                        ],
                    },
                }
            else:
                # Timed event (9 AM - 10 AM)
                start_time = date.replace(hour=9, minute=0, second=0)
                end_time = date.replace(hour=10, minute=0, second=0)
                event = {
                    'summary': event_title,
                    'description': full_description,
                    'start': {
                        'dateTime': start_time.isoformat(),
                        'timeZone': 'America/Los_Angeles',
                    },
                    'end': {
                        'dateTime': end_time.isoformat(),
                        'timeZone': 'America/Los_Angeles',
                    },
                    'reminders': {
                        'useDefault': False,
                        'overrides': [
                            {'method': 'popup', 'minutes': reminder_minutes},
                        ],
                    },
                }

            # Create the event
            created_event = self.service.events().insert(
                calendarId=self.calendar_id,
                body=event
            ).execute()

            event_id = created_event.get('id')
            self.log.add_log("Calendar", f"Created event: {event_title} on {date.strftime('%Y-%m-%d')}")
            return event_id

        except HttpError as e:
            self.log.add_log("Calendar", f"Failed to create event: {e}")
            return None
        except Exception as e:
            self.log.add_log("Calendar", f"Error creating event: {e}")
            return None

    def create_deadline_events(
        self,
        deadlines: List[Dict[str, Any]],
        file_number: str = "",
        email_link: str = ""
    ) -> List[str]:
        """
        Create multiple calendar events from a list of deadlines.

        Args:
            deadlines: List of deadline dicts with keys:
                - title: str
                - date: datetime
                - description: str (optional)
            file_number: Case file number
            email_link: Link to original email

        Returns:
            List of created event IDs
        """
        event_ids = []

        for deadline in deadlines:
            event_id = self.create_event(
                title=deadline.get('title', 'Deadline'),
                date=deadline['date'],
                description=deadline.get('description', ''),
                file_number=file_number,
                email_link=email_link
            )
            if event_id:
                event_ids.append(event_id)

        return event_ids

    def delete_event(self, event_id: str) -> bool:
        """
        Delete a calendar event.

        Args:
            event_id: The event ID to delete

        Returns:
            True if successful, False otherwise
        """
        if not self.service:
            if not self.authenticate():
                return False

        try:
            self.service.events().delete(
                calendarId=self.calendar_id,
                eventId=event_id
            ).execute()
            self.log.add_log("Calendar", f"Deleted event: {event_id}")
            return True
        except HttpError as e:
            self.log.add_log("Calendar", f"Failed to delete event: {e}")
            return False
