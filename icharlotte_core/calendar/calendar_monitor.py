"""
Calendar Monitor Worker for iCharlotte Calendar Agent.

Polls Outlook Sent Items for emails to calendartestai@gmail.com,
classifies attachments, calculates legal deadlines, and creates
Google Calendar events.
"""

import os
import re
import tempfile
from datetime import datetime, timedelta
from typing import List, Dict, Any, Optional

# PySide6 is optional - provide fallback for testing
try:
    from PySide6.QtCore import QThread, Signal
    PYSIDE6_AVAILABLE = True
except ImportError:
    PYSIDE6_AVAILABLE = False
    # Create a mock QThread for testing without Qt
    class QThread:
        def __init__(self):
            pass
        def msleep(self, ms):
            import time
            time.sleep(ms / 1000.0)
    Signal = lambda *args: property(lambda self: None)

from icharlotte_core.ui.logs_tab import LogManager
from icharlotte_core.calendar.deadline_calculator import DeadlineCalculator
from icharlotte_core.calendar.attachment_classifier import AttachmentClassifier

# Conditional import for GoogleCalendarClient (requires google-api-python-client)
try:
    from icharlotte_core.calendar.gcal_client import GoogleCalendarClient
    GCAL_AVAILABLE = True
except ImportError:
    GoogleCalendarClient = None
    GCAL_AVAILABLE = False


class CalendarMonitorWorker(QThread):
    """
    Background worker that monitors Outlook Sent Items for calendar-related emails.

    Watches for emails:
    - FROM: aserpik@bordinsemmer.com
    - TO: calendartestai@gmail.com

    For each matching email:
    1. Extracts file number from subject (####.###)
    2. Scans email body for dates (last 3 emails if thread)
    3. Classifies attachments (correspondence, discovery, motion, etc.)
    4. Calculates appropriate deadlines
    5. Creates Google Calendar events

    Signals:
        calendar_event_created(str, str): (file_number, event_title)
        error(str): Error message
        status(str): Status update message
    """

    # Signals - only define when PySide6 is available
    if PYSIDE6_AVAILABLE:
        calendar_event_created = Signal(str, str)  # (file_number, event_title)
        error = Signal(str)
        status = Signal(str)

    # Configuration
    POLL_INTERVAL = 30  # seconds
    FILE_NUMBER_REGEX = r'(\d{4})[._\-\u2013\u2014](\d{3})'
    TARGET_RECIPIENT = "calendartestai@gmail.com"
    SENDER_DOMAIN = "bordinsemmer.com"

    # Track processed emails and conversations to avoid duplicates
    _processed_emails = set()
    _processed_conversations = set()  # Track ConversationID to avoid duplicate threads

    def __init__(self):
        super().__init__()
        self.stop_requested = False
        self.log = LogManager()
        # GoogleCalendarClient may not be available if google libraries aren't installed
        self.gcal_client = GoogleCalendarClient() if GCAL_AVAILABLE else None
        self.deadline_calc = DeadlineCalculator()
        self.classifier = AttachmentClassifier()

    def request_stop(self):
        """Signal the worker to stop on the next iteration."""
        self.stop_requested = True

    def run(self):
        """Main worker loop."""
        import pythoncom

        try:
            pythoncom.CoInitialize()
            self.log.add_log("Calendar", "Calendar monitor started")
            self.status.emit("Calendar monitor started")

            # Authenticate with Google Calendar on startup
            if not self.gcal_client.authenticate():
                self.error.emit("Failed to authenticate with Google Calendar")
                return

            while not self.stop_requested:
                try:
                    self._poll_sent_items()
                except Exception as e:
                    self.log.add_log("Calendar", f"Poll error: {e}")
                    self.error.emit(f"Poll error: {e}")

                # Sleep in small increments to allow quick stop
                for _ in range(self.POLL_INTERVAL):
                    if self.stop_requested:
                        break
                    self.msleep(1000)

        except Exception as e:
            self.log.add_log("Calendar", f"Monitor error: {e}")
            self.error.emit(f"Monitor error: {e}")
        finally:
            pythoncom.CoUninitialize()
            self.log.add_log("Calendar", "Calendar monitor stopped")
            self.status.emit("Calendar monitor stopped")

    def _poll_sent_items(self):
        """Check Sent Items for matching emails."""
        import win32com.client

        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            mapi = outlook.GetNamespace("MAPI")

            # Get Sent Items folder (olFolderSentMail = 5)
            sent_folder = mapi.GetDefaultFolder(5)

            # Filter items from last 24 hours
            cutoff_time = datetime.now() - timedelta(hours=24)
            cutoff_str = cutoff_time.strftime("%m/%d/%Y %H:%M %p")

            restriction = f"[SentOn] >= '{cutoff_str}'"
            items = sent_folder.Items.Restrict(restriction)

            for item in items:
                if self.stop_requested:
                    break

                try:
                    self._process_email(item)
                except Exception as e:
                    self.log.add_log("Calendar", f"Email processing error: {e}")

        except Exception as e:
            self.log.add_log("Calendar", f"Outlook access error: {e}")
            self.error.emit(f"Outlook access error: {e}")

    def _process_email(self, item):
        """Process a single email item."""
        try:
            entry_id = item.EntryID
            subject = str(item.Subject or "")[:50]

            # Skip if already processed
            if entry_id in self._processed_emails:
                return

            # Also check ConversationID to avoid processing FW:/RE: duplicates
            conversation_id = self._get_conversation_id(item)
            if conversation_id in self._processed_conversations:
                self.log.add_log("Calendar", f"Skipping duplicate conversation: {subject}")
                self._processed_emails.add(entry_id)
                return

            # Check sender
            sender_email = ""
            try:
                sender_email = item.SenderEmailAddress or ""
                if not sender_email or "@" not in sender_email:
                    sender_email = item.Sender.GetExchangeUser().PrimarySmtpAddress
            except:
                pass

            if self.SENDER_DOMAIN.lower() not in sender_email.lower():
                return

            # Check recipient
            recipient_emails = self._get_recipient_emails(item)
            if not any(self.TARGET_RECIPIENT.lower() in addr.lower() for addr in recipient_emails):
                return

            # This is a matching email - process it
            self.log.add_log("Calendar", f"Processing: {subject}")

            # Extract file number
            file_number = self._extract_file_number(str(item.Subject or ""))
            if not file_number:
                self.log.add_log("Calendar", f"No file number in subject: {item.Subject}")
                # Still process - might have dates in body
                file_number = "NO_FILE"

            # Get email body (handle threads - last 3 emails)
            body = self._get_thread_body(item, max_emails=3)

            # Extract dates from body
            body_dates = self._extract_dates_from_body(body)

            # Log date count only if dates were found
            if body_dates:
                self.log.add_log("Calendar", f"Found {len(body_dates)} date(s) in email body")

            # Create email link (Outlook protocol)
            email_link = f"outlook:{entry_id}"

            # Process attachments (only document types)
            deadlines = []
            has_attachment = item.Attachments.Count > 0

            if has_attachment:
                deadlines = self._process_attachments(item, file_number, email_link)

            # Check if we processed deposition notices - if so, skip body date extraction
            # Deposition notices are self-contained and we don't want extra dates from the email body
            has_deposition = any('Deposition of' in d.get('title', '') for d in deadlines)

            # If no document-based deadlines were created, use dates from body
            if not deadlines and body_dates:
                # Use LLM to extract descriptive title from email
                action_title = self._extract_action_title(body, str(item.Subject or ""))
                for date in body_dates:
                    deadlines.append({
                        'title': action_title,
                        'date': date,
                        'description': f'Scheduled from email\n\nSubject: {item.Subject}\n\nBody excerpt:\n{body[:300]}...',
                    })

            # Also add body dates if they're different from document deadlines
            # BUT skip this if we have deposition notices (they're self-contained)
            if deadlines and body_dates and not has_deposition:
                existing_dates = {d['date'].date() for d in deadlines if d.get('date')}
                for date in body_dates:
                    if date.date() not in existing_dates:
                        action_title = self._extract_action_title(body, str(item.Subject or ""))
                        deadlines.append({
                            'title': action_title,
                            'date': date,
                            'description': f'Additional date from email\n\nSubject: {item.Subject}',
                        })

            # Create calendar events
            events_created = 0
            for deadline in deadlines:
                event_id = self.gcal_client.create_event(
                    title=deadline['title'],
                    date=deadline['date'],
                    description=deadline.get('description', ''),
                    file_number=file_number,
                    email_link=email_link,
                    all_day=deadline.get('all_day', True)  # Default to all-day unless specified
                )
                if event_id:
                    events_created += 1
                    self.calendar_event_created.emit(file_number, deadline['title'])

            # Mark as processed (both email and conversation)
            self._processed_emails.add(entry_id)
            self._processed_conversations.add(conversation_id)

            if events_created > 0:
                self.log.add_log("Calendar", f"[{file_number}] Created {events_created} calendar event(s)")
                self.status.emit(f"Created {events_created} event(s) for {file_number}")
            else:
                self.log.add_log("Calendar", f"[{file_number}] No dates/deadlines found to calendar")
                self.status.emit(f"Processed {file_number} - no dates found")

        except AttributeError as e:
            # Item doesn't have expected properties - log it for debugging
            self.log.add_log("Calendar", f"AttributeError processing email: {e}")
        except Exception as e:
            self.log.add_log("Calendar", f"Processing error: {e}")
            import traceback
            self.log.add_log("Calendar", f"Traceback: {traceback.format_exc()[:500]}")

    def _get_recipient_emails(self, item) -> List[str]:
        """Get all recipient email addresses."""
        emails = []
        try:
            for recipient in item.Recipients:
                try:
                    email = recipient.AddressEntry.GetExchangeUser().PrimarySmtpAddress
                except:
                    email = recipient.Address or ""
                if email:
                    emails.append(email.lower())
        except:
            to_field = str(item.To or "").lower()
            emails = [to_field]
        return emails

    def _extract_file_number(self, subject: str) -> str:
        """Extract file number from subject line."""
        # Normalize dashes
        subject = subject.replace('\u2013', '-').replace('\u2014', '-')
        match = re.search(self.FILE_NUMBER_REGEX, subject)
        if match:
            return f"{match.group(1)}.{match.group(2)}"
        return ""

    def _get_thread_body(self, item, max_emails: int = 1) -> str:
        """
        Get email body, extracting only the FIRST (most recent) message.

        For date extraction, we only want the current message, not quoted
        replies or forwarded content that may contain stale relative dates.
        """
        body = str(item.Body or "")

        # Common thread separators - split on these to get individual messages
        separators = [
            r'-{3,}\s*Original Message\s*-{3,}',
            r'From:.*?Sent:.*?To:.*?Subject:',
            r'On .+? wrote:',
            r'_{5,}',
            r'-{5,}',
            r'From:\s+\S+@',  # Start of forwarded email header
        ]

        # Find the earliest separator position to extract only the first message
        first_part = body
        for sep in separators:
            match = re.search(sep, body, flags=re.IGNORECASE | re.DOTALL)
            if match:
                # Take only the text BEFORE the separator
                candidate = body[:match.start()].strip()
                if len(candidate) > 20:  # Ensure we have substantial content
                    if len(candidate) < len(first_part):
                        first_part = candidate

        # If the first part is very short, it might just be a greeting - use more content
        if len(first_part) < 50 and len(body) > 50:
            first_part = body[:500]  # Just take first 500 chars as fallback

        return first_part

    def _extract_dates_from_body(self, body: str) -> List[datetime]:
        """
        Extract dates from email body using natural language parsing.

        Handles:
        - Explicit dates: "January 26, 2026", "1/26/2026"
        - Relative dates: "tomorrow", "next Monday", "in 2 weeks"
        - Natural language: "end of month", "next Friday at 3pm"
        """
        import dateparser

        dates = []
        now = datetime.now()

        # First, try to find explicit date patterns
        # Order matters: patterns with year first, then patterns without year
        # The pattern without year uses negative lookahead to avoid matching dates that have a year
        explicit_patterns = [
            r'(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2}(?:st|nd|rd|th)?,?\s*\d{4}',
            r'\d{1,2}[/-]\d{1,2}[/-]\d{2,4}',
            # Month DD without year - negative lookahead to skip if followed by year
            r'(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2}(?:st|nd|rd|th)?(?!,?\s*\d{4})',
        ]

        for pattern in explicit_patterns:
            matches = re.findall(pattern, body, re.IGNORECASE)
            for match in matches:
                parsed = dateparser.parse(match, settings={
                    'PREFER_DATES_FROM': 'future',
                    'RELATIVE_BASE': now
                })
                if parsed and parsed > now and parsed not in dates:
                    dates.append(parsed)

        # Then, look for natural language date expressions
        # Process body to track what's already matched
        body_lower = body.lower()
        matched_spans = []  # Track matched text positions to avoid double-matching

        # Order matters - check multi-word phrases first (e.g., "day after tomorrow" before "tomorrow")
        natural_patterns = [
            (r'\b(day\s+after\s+tomorrow)\b', 'day after tomorrow'),
            (r'\b(the\s+day\s+after\s+tomorrow)\b', 'day after tomorrow'),
            (r'(?<!after\s)\b(tomorrow)\b', 'tomorrow'),  # Negative lookbehind to skip "after tomorrow"
            (r'\b(today)\b', 'today'),
            (r'\b(next\s+(?:monday|tuesday|wednesday|thursday|friday|saturday|sunday))\b', None),
            (r'\b(this\s+(?:monday|tuesday|wednesday|thursday|friday|saturday|sunday))\b', None),
            (r'\b(next\s+week)\b', None),
            (r'\b(in\s+\d+\s+(?:days?|weeks?|months?))\b', None),
            (r'\b(end\s+of\s+(?:month|week))\b', None),
            (r'\b(\d+\s+(?:days?|weeks?|months?)\s+from\s+(?:now|today))\b', None),
            (r'\b(two\s+days?\s+from\s+(?:now|today))\b', None),
            (r'\b(a\s+week\s+from\s+(?:now|today))\b', None),
        ]

        # Check if "day after tomorrow" exists - if so, don't match plain "tomorrow"
        has_day_after_tomorrow = 'day after tomorrow' in body_lower

        # Day name to weekday mapping (Monday=0, Sunday=6)
        day_names = {
            'monday': 0, 'tuesday': 1, 'wednesday': 2, 'thursday': 3,
            'friday': 4, 'saturday': 5, 'sunday': 6
        }

        for pattern, skip_if in natural_patterns:
            # Skip plain "tomorrow" if "day after tomorrow" exists
            if skip_if == 'tomorrow' and has_day_after_tomorrow:
                continue

            matches = re.findall(pattern, body, re.IGNORECASE)
            for match in matches:
                parsed = None

                # Handle "next Monday", "next Tuesday", etc. manually since dateparser doesn't
                next_day_match = re.match(r'next\s+(\w+)', match, re.IGNORECASE)
                if next_day_match:
                    day_name = next_day_match.group(1).lower()
                    if day_name in day_names:
                        target_weekday = day_names[day_name]
                        current_weekday = now.weekday()
                        days_ahead = target_weekday - current_weekday
                        if days_ahead <= 0:  # Target day already passed this week
                            days_ahead += 7
                        parsed = now + timedelta(days=days_ahead)

                # Handle "this Monday", "this Tuesday", etc.
                this_day_match = re.match(r'this\s+(\w+)', match, re.IGNORECASE)
                if this_day_match and not parsed:
                    day_name = this_day_match.group(1).lower()
                    if day_name in day_names:
                        target_weekday = day_names[day_name]
                        current_weekday = now.weekday()
                        days_ahead = target_weekday - current_weekday
                        if days_ahead < 0:  # Target day already passed this week
                            days_ahead += 7
                        elif days_ahead == 0:  # It's today
                            days_ahead = 0  # Keep it as today
                        parsed = now + timedelta(days=days_ahead)

                # Fall back to dateparser for other patterns
                if not parsed:
                    parsed = dateparser.parse(match, settings={
                        'PREFER_DATES_FROM': 'future',
                        'RELATIVE_BASE': now
                    })

                if parsed and parsed >= now and parsed not in dates:
                    # For "today", only add if there's a time component or it's explicitly mentioned
                    if 'today' in match.lower() and parsed.date() == now.date():
                        continue  # Skip "today" unless there's a specific time
                    dates.append(parsed)
                    self.log.add_log("Calendar", f"Parsed date '{match}' -> {parsed.strftime('%Y-%m-%d')}")

        # Also try to extract any date-like phrases using dateparser's search
        try:
            # Look for date expressions in the first part of the email (before signature)
            # Split on common signature indicators
            main_body = body.split('\n\n')[0] if '\n\n' in body else body[:500]

            # Try parsing sentences that might contain dates
            sentences = re.split(r'[.!?\n]', main_body)
            for sentence in sentences:
                sentence = sentence.strip()
                if len(sentence) < 10 or len(sentence) > 200:
                    continue

                # Look for calendar-related keywords
                if any(kw in sentence.lower() for kw in ['calendar', 'schedule', 'meeting', 'call', 'deadline', 'due', 'hearing', 'trial', 'deposition']):
                    parsed = dateparser.parse(sentence, settings={
                        'PREFER_DATES_FROM': 'future',
                        'RELATIVE_BASE': now
                    })
                    if parsed and parsed > now and parsed not in dates:
                        dates.append(parsed)
        except:
            pass

        return sorted(set(dates))

    def _parse_date(self, date_str: str) -> Optional[datetime]:
        """Parse a date string using dateparser."""
        import dateparser

        try:
            parsed = dateparser.parse(date_str, settings={
                'PREFER_DATES_FROM': 'future',
                'RELATIVE_BASE': datetime.now()
            })
            return parsed
        except:
            return None

    def _extract_action_title(self, body: str, subject: str) -> str:
        """
        Use LLM to extract a short descriptive action/title from email content.

        Returns something like "Call with Joe" or "Meeting with opposing counsel"
        instead of generic "Date from Email".
        """
        try:
            from icharlotte_core.llm import LLMHandler
            from icharlotte_core.config import API_KEYS

            prompt = f"""Extract a short calendar event title (3-6 words) from this email.
The title should describe the action/event being scheduled.

Examples:
- "Call with Joe"
- "Meeting with opposing counsel"
- "Deadline for discovery responses"
- "Client meeting"
- "Deposition prep call"

Subject: {subject}

Email body:
{body[:500]}

Return ONLY the short event title, nothing else. If no clear action is mentioned, return "Scheduled Event"."""

            # Use Gemini for quick extraction
            response = LLMHandler.generate(
                provider="Gemini",
                model="gemini-3-flash-preview",
                system_prompt="You extract short calendar event titles from emails.",
                user_prompt=prompt,
                file_contents="",
                settings={
                    'temperature': 0.1,
                    'max_tokens': 50,
                    'stream': False
                }
            )

            title = response.strip().strip('"\'')
            # Validate it's reasonably short
            if len(title) > 50:
                title = title[:47] + "..."
            if not title or len(title) < 3:
                title = "Scheduled Event"

            self.log.add_log("Calendar", f"Extracted title: {title}")
            return title

        except Exception as e:
            self.log.add_log("Calendar", f"LLM title extraction error: {e}")
            return "Scheduled Event"

    def _get_conversation_id(self, item) -> str:
        """Get the ConversationID to group related emails."""
        try:
            return item.ConversationID
        except:
            # Fallback to EntryID if ConversationID not available
            return item.EntryID

    def _process_attachments(self, item, file_number: str, email_link: str) -> List[Dict[str, Any]]:
        """
        Process email attachments and generate deadlines.
        """
        deadlines = []
        email_date = item.SentOn.replace(tzinfo=None) if hasattr(item.SentOn, 'replace') else datetime.now()

        for i in range(1, item.Attachments.Count + 1):
            try:
                attachment = item.Attachments.Item(i)
                filename = attachment.FileName

                # Skip non-document attachments
                ext = os.path.splitext(filename)[1].lower()
                if ext not in ['.pdf', '.doc', '.docx', '.txt']:
                    continue

                self.log.add_log("Calendar", f"Classifying attachment: {filename}")

                # Classify the attachment
                temp_path, classification = self.classifier.classify_from_outlook_attachment(attachment)

                try:
                    doc_type = classification.get('doc_type', 'correspondence')
                    motion_type = classification.get('motion_type', 'standard')
                    hearing_date = classification.get('hearing_date')
                    summary = classification.get('summary', '')

                    self.log.add_log("Calendar", f"Classified as: {doc_type} / {motion_type}")

                    # Generate deadlines based on document type
                    if doc_type == 'correspondence':
                        # Calendar dates mentioned in correspondence
                        for date in classification.get('dates_found', []):
                            if date > datetime.now():
                                deadlines.append({
                                    'title': 'Deadline from Correspondence',
                                    'date': date,
                                    'description': f'Date from: {filename}\n{summary}',
                                })

                    elif doc_type == 'deposition_notice':
                        # Deposition notice - calendar the deposition date/time with deponent name
                        deposition_date = classification.get('deposition_date')
                        deposition_time = classification.get('deposition_time')  # e.g., "09:00", "14:30"
                        deponent_name = classification.get('deponent_name', 'Unknown')

                        if deposition_date:
                            # Combine date and time if time is available
                            if deposition_time:
                                try:
                                    hour, minute = map(int, deposition_time.split(':'))
                                    deposition_datetime = deposition_date.replace(hour=hour, minute=minute)
                                    time_str = deposition_datetime.strftime('%I:%M %p')
                                except:
                                    deposition_datetime = deposition_date
                                    time_str = "Time TBD"
                            else:
                                deposition_datetime = deposition_date
                                time_str = "Time TBD"

                            deadlines.append({
                                'title': f"Deposition of {deponent_name}",
                                'date': deposition_datetime,
                                'all_day': deposition_time is None,  # Timed event if we have a time
                                'description': f"Deposition Notice\nDeponent: {deponent_name}\nTime: {time_str}\n\nDocument: {filename}\n{summary}",
                            })
                            self.log.add_log("Calendar", f"Deposition: {deponent_name} on {deposition_date.strftime('%Y-%m-%d')} at {time_str}")
                        else:
                            self.log.add_log("Calendar", f"No deposition date found in: {filename}")

                        # Skip processing dates_found for deposition notices - only use the extracted deposition date
                        # Clear dates_found to prevent extra calendar entries
                        classification['dates_found'] = []

                    elif doc_type == 'discovery_request':
                        # Discovery response due: 30 days + service extension
                        deadline = self.deadline_calc.get_discovery_response_deadline(
                            request_date=email_date,
                            service_method='electronic'
                        )
                        deadline['title'] = f"Discovery Response Due ({filename})"
                        deadline['description'] += f"\n\nDocument: {filename}\n{summary}"
                        deadlines.append(deadline)

                    elif doc_type == 'discovery_response':
                        # Motion to compel deadline: 45 days
                        deadline = self.deadline_calc.get_motion_to_compel_deadline(
                            response_date=email_date
                        )
                        deadline['title'] = f"Motion to Compel Deadline ({filename})"
                        deadline['description'] += f"\n\nDocument: {filename}\n{summary}"
                        deadlines.append(deadline)

                    elif doc_type == 'motion':
                        # Opposition deadline
                        if hearing_date:
                            opp_deadline = self.deadline_calc.get_opposition_deadline(
                                motion_type=motion_type or 'standard',
                                hearing_date=hearing_date,
                                service_method='electronic'
                            )
                            if opp_deadline:
                                opp_deadline['title'] = f"Opposition Due - {motion_type.upper() if motion_type else 'Motion'}"
                                opp_deadline['description'] += f"\n\nDocument: {filename}\n{summary}"
                                deadlines.append(opp_deadline)

                            # Also add hearing date
                            deadlines.append({
                                'title': f"Hearing - {motion_type.upper() if motion_type else 'Motion'}",
                                'date': hearing_date,
                                'description': f"Hearing on {motion_type or 'motion'}\n\nDocument: {filename}\n{summary}",
                            })
                        else:
                            self.log.add_log("Calendar", f"No hearing date found in motion: {filename}")

                    elif doc_type == 'opposition':
                        # Reply deadline
                        if hearing_date:
                            reply_deadline = self.deadline_calc.get_reply_deadline(
                                motion_type=motion_type or 'standard',
                                hearing_date=hearing_date,
                                service_method='electronic'
                            )
                            if reply_deadline:
                                reply_deadline['title'] = f"Reply Due - {motion_type.upper() if motion_type else 'Motion'}"
                                reply_deadline['description'] += f"\n\nDocument: {filename}\n{summary}"
                                deadlines.append(reply_deadline)
                        else:
                            self.log.add_log("Calendar", f"No hearing date found in opposition: {filename}")

                    elif doc_type == 'reply':
                        # Reply received - might want to calendar the hearing
                        if hearing_date:
                            deadlines.append({
                                'title': f"Hearing - {motion_type.upper() if motion_type else 'Motion'}",
                                'date': hearing_date,
                                'description': f"Reply received. Hearing on {motion_type or 'motion'}\n\nDocument: {filename}\n{summary}",
                            })

                finally:
                    # Clean up temp file
                    if temp_path and os.path.exists(temp_path):
                        try:
                            os.remove(temp_path)
                        except:
                            pass

            except Exception as e:
                self.log.add_log("Calendar", f"Attachment processing error: {e}")

        return deadlines
