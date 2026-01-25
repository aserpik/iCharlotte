"""
Sent Items Monitor - Polls Outlook Sent Items for emails to serpiklaw@gmail.com
with body starting with "AS -" and creates todos in the Master List.
"""

import re
from datetime import datetime, timedelta
from PyQt6.QtCore import QThread, pyqtSignal

from icharlotte_core.master_db import MasterCaseDatabase
from icharlotte_core.ui.logs_tab import LogManager
from icharlotte_core.llm import LLMHandler
from icharlotte_core.config import API_KEYS


class SentItemsMonitorWorker(QThread):
    """
    Background worker that polls Outlook Sent Items every 2 minutes,
    detects emails to serpiklaw@gmail.com with body starting with "AS -",
    extracts the file number from the subject, and creates a todo.
    """

    # Signals
    todo_created = pyqtSignal(str, str)  # (file_number, todo_text)
    error = pyqtSignal(str)  # error message
    status = pyqtSignal(str)  # status message

    # Configuration
    POLL_INTERVAL = 30  # seconds
    FILE_NUMBER_REGEX = r'(\d{4})[._\-\u2013\u2014](\d{3})'  # Include en-dash and em-dash

    # Target email/prefix/assignee configurations
    # Each tuple: (email_address, body_prefix, assignee_initials)
    TARGETS = [
        ("serpiklaw@gmail.com", "AS -", "AS"),
        ("fmottola@bordinsemmer.com", "FM -", "FM"),
        ("hproshyan@bordinsemmer.com", "HP -", "HP"),
        ("cetmekjian@bordinsemmer.com", "CE -", "CE"),
    ]

    # Filler words to strip from todo text
    FILLER_WORDS = ['please', 'pls', 'thnx', 'thanks', 'thx', 'ty']

    @staticmethod
    def _normalize_dashes(text: str) -> str:
        """Convert en-dash and em-dash to regular hyphen."""
        return text.replace('\u2013', '-').replace('\u2014', '-').replace('\u2012', '-')

    def __init__(self, db: MasterCaseDatabase = None):
        super().__init__()
        self.db = db or MasterCaseDatabase()
        self.stop_requested = False

    def request_stop(self):
        """Signal the worker to stop on the next iteration."""
        self.stop_requested = True

    def run(self):
        """Main worker loop."""
        import pythoncom

        try:
            pythoncom.CoInitialize()
            self.status.emit("Email monitor started")

            while not self.stop_requested:
                try:
                    self._poll_sent_items()
                except Exception as e:
                    self.error.emit(f"Poll error: {e}")

                # Sleep in small increments to allow quick stop
                for _ in range(self.POLL_INTERVAL):
                    if self.stop_requested:
                        break
                    self.msleep(1000)  # 1 second

        except Exception as e:
            self.error.emit(f"Monitor error: {e}")
        finally:
            pythoncom.CoUninitialize()
            self.status.emit("Email monitor stopped")

    def _poll_sent_items(self):
        """Check Sent Items for matching emails."""
        import win32com.client

        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            mapi = outlook.GetNamespace("MAPI")

            # Get Sent Items folder (olFolderSentMail = 5)
            sent_folder = mapi.GetDefaultFolder(5)

            # Filter items from last 24 hours for performance
            cutoff_time = datetime.now() - timedelta(hours=24)
            cutoff_str = cutoff_time.strftime("%m/%d/%Y %H:%M %p")

            # Use Restrict to filter by sent time
            restriction = f"[SentOn] >= '{cutoff_str}'"
            items = sent_folder.Items.Restrict(restriction)

            for item in items:
                if self.stop_requested:
                    break

                try:
                    self._process_email(item)
                except Exception:
                    pass  # Continue processing other emails

        except Exception as e:
            self.error.emit(f"Outlook access error: {e}")

    def _process_email(self, item):
        """Process a single email item."""
        try:
            # Get unique identifier for this email
            entry_id = item.EntryID

            # Skip if already processed
            if self.db.is_email_processed(entry_id):
                return

            # Get all recipient email addresses
            recipient_emails = []
            try:
                for recipient in item.Recipients:
                    email_addr = ""
                    try:
                        # SMTP address is in the AddressEntry
                        email_addr = recipient.AddressEntry.GetExchangeUser().PrimarySmtpAddress
                    except:
                        # Fallback to Address property (works for non-Exchange)
                        email_addr = recipient.Address or ""
                    if email_addr:
                        recipient_emails.append(email_addr.lower())
            except:
                # Fallback to simple To field
                to_field = str(item.To or "").lower()
                recipient_emails = [to_field]

            # Check body (normalize dashes first)
            body = str(item.Body or "").strip()
            body_normalized = self._normalize_dashes(body)

            # Find matching target configuration
            matched_target = None
            for target_email, body_prefix, assignee in self.TARGETS:
                # Check if email matches any recipient
                email_matches = any(target_email.lower() in addr for addr in recipient_emails)
                if not email_matches:
                    continue

                # Check if body starts with the expected prefix
                if body_normalized.upper().startswith(body_prefix.upper()):
                    matched_target = (target_email, body_prefix, assignee)
                    break

            if not matched_target:
                return

            target_email, body_prefix, assignee = matched_target
            subject = str(item.Subject or "")

            # Extract file number from subject
            subject = str(item.Subject or "")
            file_number = self._extract_file_number(subject)

            if not file_number:
                self.status.emit(f"Skipped email - no file number in subject: {subject[:50]}")
                return

            # Verify case exists in database
            case = self.db.get_case(file_number)
            if not case:
                self.status.emit(f"Skipped email - case not found: {file_number}")
                return

            # Extract and clean todo text (use normalized body and matched prefix)
            todo_text = self._extract_todo_text(body_normalized, body_prefix)

            if not todo_text:
                self.status.emit(f"Skipped email - empty todo text after prefix")
                return

            # Create the todo
            today = datetime.now().strftime("%Y-%m-%d")
            self.db.add_todo(
                file_number=file_number,
                item=todo_text,
                color='yellow',
                created_date=today
            )

            # Assign to the matched assignee automatically
            todos = self.db.get_todos(file_number)
            if todos:
                # Get the most recently added todo (should be first due to ORDER BY id DESC)
                latest_todo = todos[0]
                self.db.update_todo_assignment(
                    latest_todo['id'],
                    assignee,
                    datetime.now().strftime("%m/%d/%y")
                )

            # Mark email as processed
            self.db.mark_email_processed(entry_id, file_number, todo_text)

            # Log to Email Monitor category
            LogManager().add_log("Email Monitor", f"[{file_number}] → {assignee}: {todo_text}")

            # Emit signal
            self.todo_created.emit(file_number, todo_text)
            self.status.emit(f"Created todo for {file_number}: {todo_text[:40]}...")

        except AttributeError:
            # Item doesn't have expected properties (e.g., meeting request)
            pass

    def _extract_file_number(self, subject: str) -> str:
        """Extract file number from subject line using regex."""
        match = re.search(self.FILE_NUMBER_REGEX, subject)
        if match:
            # Normalize to ####.###
            return f"{match.group(1)}.{match.group(2)}"
        return ""

    def _extract_todo_text(self, body: str, body_prefix: str = "AS -") -> str:
        """
        Extract and clean todo text from email body using LLM.

        Body format example:
            AS - Can you please work on finishing the FSR I started on this case?

        Returns:
            Cleaned action-only todo text: "Finish the FSR"
        """
        # Remove the prefix (case insensitive)
        text = body.strip()
        if text.upper().startswith(body_prefix.upper()):
            text = text[len(body_prefix):].strip()

        # Take only the first line (main instruction)
        lines = text.split('\n')
        first_line = lines[0].strip() if lines else ""

        if not first_line:
            return ""

        # Try LLM extraction first
        extracted = self._extract_action_with_llm(first_line)
        if extracted:
            return extracted

        # Fallback to rule-based extraction
        return self._extract_todo_text_fallback(first_line)

    def _extract_action_with_llm(self, text: str) -> str:
        """Use LLM to extract just the action from email text."""
        if not API_KEYS.get("Gemini"):
            return ""

        system_prompt = """You are a task extraction assistant. Extract ONLY the core action/task from the given text.

Rules:
- Remove all conversational language (can you, please, would you, etc.)
- Remove references to "this case" or "the case"
- Remove phrases like "I started", "I need you to", "work on"
- Keep the essential action and its object (e.g., "finish the FSR", "call the plaintiff")
- Use imperative form (start with a verb)
- Capitalize the first letter
- Be concise - just the action, nothing else
- Output ONLY the extracted action, no explanation"""

        user_prompt = f"Extract the action from: {text}"

        try:
            result = LLMHandler.generate(
                provider="Gemini",
                model="gemini-2.0-flash",
                system_prompt=system_prompt,
                user_prompt=user_prompt,
                file_contents="",
                settings={"temperature": 0.1, "max_tokens": 50, "stream": False}
            )
            # Clean up the result
            result = result.strip().strip('"').strip("'")
            if result:
                return result
        except Exception as e:
            LogManager().add_log("Email Monitor", f"LLM extraction failed: {e}")

        return ""

    def _extract_todo_text_fallback(self, first_line: str) -> str:
        """Fallback rule-based extraction when LLM is unavailable."""
        # Remove filler words
        words = first_line.split()
        cleaned_words = []
        for word in words:
            # Check if word (stripped of punctuation) is a filler word
            word_clean = re.sub(r'[^\w]', '', word.lower())
            if word_clean not in self.FILLER_WORDS:
                cleaned_words.append(word)

        result = ' '.join(cleaned_words)

        # Clean up punctuation artifacts (trailing commas, multiple spaces)
        result = re.sub(r'\s*,\s*$', '', result)  # Remove trailing comma
        result = re.sub(r'\s+', ' ', result)  # Normalize whitespace
        result = result.strip()

        # Capitalize first letter
        if result:
            result = result[0].upper() + result[1:]

        return result
