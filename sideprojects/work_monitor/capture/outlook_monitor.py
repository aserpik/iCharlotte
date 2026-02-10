"""Outlook email event monitoring via COM API."""

import threading
import time
import logging
from datetime import datetime, timedelta

from ..db import ActivityDB
from ..case_extractor import extract_case_number
from .. import config

logger = logging.getLogger("work_monitor.outlook")


class OutlookMonitor(threading.Thread):
    """Polls Outlook for email events (sent, received)."""

    def __init__(self, db: ActivityDB, stop_event: threading.Event):
        super().__init__(daemon=True, name="OutlookMonitor")
        self.db = db
        self.stop_event = stop_event
        self._processed_ids: set[str] = set()

    def run(self):
        import pythoncom
        logger.info("OutlookMonitor started")
        try:
            pythoncom.CoInitialize()

            while not self.stop_event.is_set():
                try:
                    self._poll_sent()
                    self._poll_inbox()
                except Exception as e:
                    logger.debug(f"Outlook poll error: {e}")

                # Sleep in 1-second chunks (pattern from sent_items_monitor.py)
                for _ in range(config.OUTLOOK_POLL_INTERVAL):
                    if self.stop_event.is_set():
                        return
                    time.sleep(1.0)
        finally:
            pythoncom.CoUninitialize()
            logger.info("OutlookMonitor stopped")

    def _get_outlook(self):
        import win32com.client
        outlook = win32com.client.Dispatch("Outlook.Application")
        return outlook.GetNamespace("MAPI")

    def _poll_sent(self):
        """Poll Sent Items for recently sent emails."""
        mapi = self._get_outlook()
        sent_folder = mapi.GetDefaultFolder(5)  # olFolderSentMail

        cutoff = datetime.now() - timedelta(hours=24)
        cutoff_str = cutoff.strftime("%m/%d/%Y %H:%M %p")
        items = sent_folder.Items.Restrict(f"[SentOn] >= '{cutoff_str}'")

        for item in items:
            if self.stop_event.is_set():
                return
            self._process_email(item, "sent")

    def _poll_inbox(self):
        """Poll Inbox for recently received emails."""
        mapi = self._get_outlook()
        inbox = mapi.GetDefaultFolder(6)  # olFolderInbox

        cutoff = datetime.now() - timedelta(hours=24)
        cutoff_str = cutoff.strftime("%m/%d/%Y %H:%M %p")
        items = inbox.Items.Restrict(f"[ReceivedTime] >= '{cutoff_str}'")

        for item in items:
            if self.stop_event.is_set():
                return
            self._process_email(item, "received")

    def _process_email(self, item, event_type: str):
        try:
            entry_id = item.EntryID
            key = f"{event_type}:{entry_id}"
            if key in self._processed_ids:
                return
            self._processed_ids.add(key)

            subject = str(getattr(item, 'Subject', '') or "")
            body = str(getattr(item, 'Body', '') or "")

            # Detect reply/forward from subject prefix
            subject_stripped = subject.strip().upper()
            if subject_stripped.startswith("RE:"):
                event_type = "replied"
            elif subject_stripped.startswith(("FW:", "FWD:")):
                event_type = "forwarded"

            # Get recipients -- prefer display name, fall back to address
            recipients = []
            try:
                for r in item.Recipients:
                    name = str(getattr(r, 'Name', '') or "")
                    if name:
                        recipients.append(name)
                    else:
                        addr = str(getattr(r, 'Address', '') or "")
                        if addr:
                            recipients.append(addr)
            except Exception:
                to_field = str(getattr(item, 'To', '') or "")
                if to_field:
                    recipients = [to_field]

            has_attachment = False
            try:
                has_attachment = item.Attachments.Count > 0
            except Exception:
                pass

            # Extract case number from subject first, then body
            case_number = extract_case_number(subject) or extract_case_number(body[:200])

            self.db.insert_email_event(
                timestamp=datetime.now().isoformat(),
                event_type=event_type,
                subject=subject,
                recipients=", ".join(recipients),
                folder=getattr(item.Parent, 'Name', '') if hasattr(item, 'Parent') else "",
                has_attachment=has_attachment,
                body_preview=body[:500],
                case_number=case_number,
            )
        except AttributeError:
            pass  # Non-mail items (meeting requests, etc.)
        except Exception as e:
            logger.debug(f"Email processing error: {e}")

    def _prune_processed_ids(self):
        """Keep processed IDs set bounded (called periodically)."""
        if len(self._processed_ids) > 5000:
            self._processed_ids.clear()
