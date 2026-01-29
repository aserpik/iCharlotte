"""
Sent Items Monitor - Polls Outlook Sent Items for emails to serpiklaw@gmail.com
with body starting with "AS -" and creates todos in the Master List.
"""

import os
import re
import tempfile
from datetime import datetime, timedelta
from PySide6.QtCore import QThread, Signal

from icharlotte_core.master_db import MasterCaseDatabase
from icharlotte_core.ui.logs_tab import LogManager
from icharlotte_core.llm import LLMHandler
from icharlotte_core.config import API_KEYS
from icharlotte_core.document_processor import DocumentProcessor


class SentItemsMonitorWorker(QThread):
    """
    Background worker that polls Outlook Sent Items every 2 minutes,
    detects emails to serpiklaw@gmail.com with body starting with "AS -",
    extracts the file number from the subject, and creates a todo.
    """

    # Signals
    todo_created = Signal(str, str)  # (file_number, todo_text)
    error = Signal(str)  # error message
    status = Signal(str)  # status message

    # Configuration
    POLL_INTERVAL = 30  # seconds
    FILE_NUMBER_REGEX = r'(\d{4})[._\-\u2013\u2014](\d{3})'  # Include en-dash and em-dash

    # Target email/prefix/assignee configurations
    # Each tuple: (email_address, body_prefix, assignee_initials)
    TARGETS = [
        ("serpiklaw@gmail.com", "AS -", "AS"),
        ("fmottola@bordinsemmer.com", "FM -- ", "FM"),
        ("hproshyan@bordinsemmer.com", "HP -- ", "HP"),
        ("cetmekjian@bordinsemmer.com", "CE -- ", "CE"),
    ]

    # Filler words to strip from todo text
    FILLER_WORDS = ['please', 'pls', 'thnx', 'thanks', 'thx', 'ty']

    # Max number of prior emails in thread to analyze for context
    MAX_THREAD_EMAILS = 3

    # Patterns to detect email boundaries in forwarded/replied content
    EMAIL_BOUNDARY_PATTERNS = [
        r'^-{3,}\s*Original Message\s*-{3,}',  # -----Original Message-----
        r'^From:\s+.+',  # From: Someone
        r'^On .+wrote:$',  # On Mon, Jan 1, 2024, Someone wrote:
        r'^\s*_{5,}\s*$',  # _____________________
    ]

    @staticmethod
    def _normalize_dashes(text: str) -> str:
        """Convert en-dash and em-dash to regular hyphen."""
        return text.replace('\u2013', '-').replace('\u2014', '-').replace('\u2012', '-')

    @staticmethod
    def _extract_attachment_names(item) -> list:
        """Extract attachment filenames from an Outlook item."""
        attachments = []
        try:
            for i in range(1, item.Attachments.Count + 1):
                att = item.Attachments.Item(i)
                filename = str(att.FileName or "")
                if filename:
                    attachments.append(filename)
        except Exception:
            pass
        return attachments

    @staticmethod
    def _extract_subject_context(subject: str) -> str:
        """
        Extract meaningful context from email subject.
        Strips file numbers, RE:/FW: prefixes, and common noise.

        Example:
            "FW: 3200.284 - Plaintiff's Responses to Discovery"
            -> "Plaintiff's Responses to Discovery"
        """
        text = subject.strip()

        # Remove RE:/FW: prefixes (can be multiple, case insensitive)
        text = re.sub(r'^(RE|FW|FWD):\s*', '', text, flags=re.IGNORECASE)
        text = re.sub(r'^(RE|FW|FWD):\s*', '', text, flags=re.IGNORECASE)  # Handle nested

        # Remove file number patterns (####.### or ####-### etc.)
        text = re.sub(r'\d{4}[._\-\u2013\u2014]\d{3}', '', text)

        # Clean up separators left behind
        text = re.sub(r'^\s*[-:]\s*', '', text)  # Leading dash/colon
        text = re.sub(r'\s*[-:]\s*$', '', text)  # Trailing dash/colon
        text = re.sub(r'\s+', ' ', text).strip()

        return text

    def _is_forward(self, subject: str) -> bool:
        """Check if the email is a forward based on subject line."""
        subject_upper = subject.upper().strip()
        return subject_upper.startswith('FW:') or subject_upper.startswith('FWD:')

    def _parse_thread_emails(self, body: str, my_instruction: str) -> list:
        """
        Parse forwarded/replied email content to extract context from prior emails.

        Returns a list of dicts with 'subject' and 'snippet' for up to MAX_THREAD_EMAILS
        prior emails in the thread.
        """
        emails = []

        # Remove my instruction from the beginning
        remaining = body[len(my_instruction):].strip() if body.startswith(my_instruction) else body

        # Split by email boundaries
        combined_pattern = '|'.join(f'({p})' for p in self.EMAIL_BOUNDARY_PATTERNS)

        # Find all email boundaries
        boundaries = list(re.finditer(combined_pattern, remaining, re.MULTILINE | re.IGNORECASE))

        if not boundaries:
            return emails

        # Extract content between boundaries (each represents a prior email)
        for i, match in enumerate(boundaries[:self.MAX_THREAD_EMAILS]):
            start = match.start()
            end = boundaries[i + 1].start() if i + 1 < len(boundaries) else len(remaining)

            email_block = remaining[start:end].strip()

            # Extract subject from this email block
            subject_match = re.search(r'^Subject:\s*(.+?)$', email_block, re.MULTILINE | re.IGNORECASE)
            subject = subject_match.group(1).strip() if subject_match else ""

            # Extract a content snippet (first few lines after headers, max 200 chars)
            lines = email_block.split('\n')
            content_lines = []
            past_headers = False
            for line in lines:
                line_stripped = line.strip()
                # Skip header lines
                if re.match(r'^(From|To|Cc|Sent|Date|Subject):\s', line_stripped, re.IGNORECASE):
                    past_headers = True
                    continue
                if past_headers and line_stripped:
                    content_lines.append(line_stripped)
                    if len(' '.join(content_lines)) > 200:
                        break

            snippet = ' '.join(content_lines)[:200]

            if subject or snippet:
                emails.append({
                    'subject': subject,
                    'snippet': snippet
                })

        return emails

    # File extensions to skip (images, signatures, etc.)
    SKIP_EXTENSIONS = {'.jpg', '.jpeg', '.png', '.gif', '.bmp', '.ico', '.tiff', '.webp'}
    # Max chars to extract from first page of attachment
    MAX_ATTACHMENT_TEXT = 500

    def _extract_attachment_text(self, item) -> list:
        """
        Extract first page text from non-image attachments.

        Returns a list of dicts with 'filename' and 'first_page_text'.
        Skips image files and signature-like attachments.
        """
        results = []
        temp_dir = None

        try:
            # Create temp directory for attachment extraction
            temp_dir = tempfile.mkdtemp(prefix="icharlotte_att_")
            processor = DocumentProcessor()

            for i in range(1, item.Attachments.Count + 1):
                try:
                    att = item.Attachments.Item(i)
                    filename = str(att.FileName or "")

                    if not filename:
                        continue

                    # Get file extension
                    ext = os.path.splitext(filename)[1].lower()

                    # Skip image files
                    if ext in self.SKIP_EXTENSIONS:
                        continue

                    # Skip small files (likely signatures) - check size if available
                    try:
                        if att.Size < 5000:  # Less than 5KB, likely signature
                            continue
                    except:
                        pass

                    # Skip files with signature-like names
                    name_lower = filename.lower()
                    if any(sig in name_lower for sig in ['signature', 'image00', 'logo', 'banner']):
                        continue

                    # Only process PDFs and DOCX for now
                    if ext not in ['.pdf', '.docx', '.doc']:
                        # Still include filename for context, just no text
                        results.append({
                            'filename': filename,
                            'first_page_text': ''
                        })
                        continue

                    # Save attachment to temp file
                    temp_path = os.path.join(temp_dir, filename)
                    att.SaveAsFile(temp_path)

                    # Extract first page text
                    first_page_text = self._extract_first_page(temp_path, processor)

                    results.append({
                        'filename': filename,
                        'first_page_text': first_page_text[:self.MAX_ATTACHMENT_TEXT] if first_page_text else ''
                    })

                    # Clean up temp file immediately
                    try:
                        os.remove(temp_path)
                    except:
                        pass

                except Exception as e:
                    LogManager().add_log("Email Monitor", f"Error extracting attachment {i}: {e}")
                    continue

        except Exception as e:
            LogManager().add_log("Email Monitor", f"Error in attachment extraction: {e}")

        finally:
            # Clean up temp directory
            if temp_dir:
                try:
                    import shutil
                    shutil.rmtree(temp_dir, ignore_errors=True)
                except:
                    pass

        return results

    def _extract_first_page(self, file_path: str, processor: DocumentProcessor) -> str:
        """
        Extract text from just the first page of a document.

        For PDFs, extracts only page 1.
        For DOCX, extracts first ~1000 chars.
        """
        ext = os.path.splitext(file_path)[1].lower()

        try:
            if ext == '.pdf':
                from pypdf import PdfReader
                reader = PdfReader(file_path)
                if len(reader.pages) > 0:
                    text = reader.pages[0].extract_text() or ""
                    # If native extraction failed, try OCR on first page
                    if len(text.strip()) < 50 and hasattr(processor, '_ocr_page'):
                        ocr_text = processor._ocr_page(file_path, 0)
                        if ocr_text and len(ocr_text.strip()) > len(text.strip()):
                            text = ocr_text
                    return text.strip()

            elif ext in ['.docx', '.doc']:
                from docx import Document as DocxDocument
                doc = DocxDocument(file_path)
                text_parts = []
                char_count = 0
                for para in doc.paragraphs:
                    text_parts.append(para.text)
                    char_count += len(para.text)
                    if char_count > 1000:  # Roughly first page
                        break
                return '\n'.join(text_parts).strip()

        except Exception as e:
            LogManager().add_log("Email Monitor", f"Error extracting first page from {file_path}: {e}")

        return ""

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

            # Gather context for smarter todo extraction
            context = {
                'subject_context': self._extract_subject_context(subject),
                'is_forward': self._is_forward(subject),
                'attachments': [],
                'thread_emails': []
            }

            # Extract attachment info (names + first page text for docs)
            try:
                context['attachments'] = self._extract_attachment_text(item)
            except Exception as e:
                LogManager().add_log("Email Monitor", f"Attachment extraction failed: {e}")
                # Fallback to just names
                context['attachments'] = [{'filename': n, 'first_page_text': ''}
                                          for n in self._extract_attachment_names(item)]

            # Parse thread emails only if this is a forward
            if context['is_forward']:
                try:
                    # Get my instruction (first line after prefix)
                    my_text = body_normalized
                    if my_text.upper().startswith(body_prefix.upper()):
                        my_text = my_text[len(body_prefix):].strip()
                    my_instruction = my_text.split('\n')[0].strip() if my_text else ""

                    context['thread_emails'] = self._parse_thread_emails(body_normalized,
                                                                          body_prefix + " " + my_instruction)
                except Exception as e:
                    LogManager().add_log("Email Monitor", f"Thread parsing failed: {e}")

            # Extract and clean todo text with full context
            todo_text = self._extract_todo_text(body_normalized, body_prefix, context)

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

    def _extract_todo_text(self, body: str, body_prefix: str = "AS -", context: dict = None) -> str:
        """
        Extract and clean todo text from email body using LLM with full context.

        Args:
            body: Full email body text
            body_prefix: The prefix like "AS -" or "FM -- "
            context: Dict with 'subject_context', 'is_forward', 'attachments', 'thread_emails'

        Returns:
            Context-aware todo text that makes sense standalone.
        """
        context = context or {}

        # Remove the prefix (case insensitive)
        text = body.strip()
        if text.upper().startswith(body_prefix.upper()):
            text = text[len(body_prefix):].strip()

        # Take only the first line (main instruction)
        lines = text.split('\n')
        first_line = lines[0].strip() if lines else ""

        if not first_line:
            return ""

        # Try LLM extraction with context
        extracted = self._extract_action_with_context(first_line, context)
        if extracted:
            return extracted

        # Fallback to rule-based extraction (no context awareness)
        return self._extract_todo_text_fallback(first_line)

    def _extract_action_with_context(self, instruction: str, context: dict) -> str:
        """
        Use LLM to synthesize a self-contained todo description using all available context.

        Args:
            instruction: The user's instruction (e.g., "please review and summarize")
            context: Dict with subject_context, attachments, thread_emails, is_forward

        Returns:
            A self-contained todo description that makes sense without the email.
        """
        if not API_KEYS.get("Gemini"):
            return ""

        # Build context sections for the prompt
        context_parts = []

        # Subject context
        subject_ctx = context.get('subject_context', '').strip()
        if subject_ctx:
            context_parts.append(f"Email subject: {subject_ctx}")

        # Attachment info
        attachments = context.get('attachments', [])
        if attachments:
            att_info = []
            for att in attachments:
                filename = att.get('filename', '')
                first_page = att.get('first_page_text', '').strip()
                if filename:
                    if first_page:
                        # Truncate first page text for prompt
                        first_page_truncated = first_page[:300] + "..." if len(first_page) > 300 else first_page
                        att_info.append(f"  - {filename}\n    First page: {first_page_truncated}")
                    else:
                        att_info.append(f"  - {filename}")
            if att_info:
                context_parts.append("Attachments:\n" + "\n".join(att_info))

        # Thread emails (for forwards)
        thread_emails = context.get('thread_emails', [])
        if thread_emails:
            thread_info = []
            for i, email in enumerate(thread_emails[:3], 1):
                subj = email.get('subject', '').strip()
                snippet = email.get('snippet', '').strip()
                if subj or snippet:
                    thread_info.append(f"  Prior email {i}:")
                    if subj:
                        thread_info.append(f"    Subject: {subj}")
                    if snippet:
                        thread_info.append(f"    Content: {snippet[:150]}...")
            if thread_info:
                context_parts.append("Email thread context:\n" + "\n".join(thread_info))

        # Build the full prompt
        context_str = "\n\n".join(context_parts) if context_parts else ""

        system_prompt = """You are a task extraction assistant for a legal case management system.

Your job is to create a SELF-CONTAINED todo item description that will make sense when viewed later without the original email.

CRITICAL RULES:
1. The todo must be understandable WITHOUT seeing the email
2. Include WHAT needs to be acted upon (document name, topic, etc.) - don't just say "review" without specifying what
3. Use the context (subject, attachments, thread) to identify what the instruction refers to
4. Remove conversational language (please, can you, thanks, etc.)
5. Use imperative form (start with a verb)
6. Be concise but complete (usually 5-15 words)
7. Capitalize the first letter

EXAMPLES:
- Instruction: "review and summarize" + Attachment: "Plaintiff's Responses to Discovery.pdf"
  Output: "Review and summarize Plaintiff's Responses to Discovery"

- Instruction: "respond to this" + Subject: "Defendant's Motion to Compel"
  Output: "Respond to Defendant's Motion to Compel"

- Instruction: "follow up" + Thread: "RE: Deposition scheduling for Dr. Smith"
  Output: "Follow up on Dr. Smith deposition scheduling"

- Instruction: "draft a response" + Attachment: "Form Interrogatories.pdf" + First page mentions "Set Two"
  Output: "Draft responses to Form Interrogatories Set Two"

Output ONLY the todo text, no explanation or quotes."""

        user_prompt = f"Instruction from email: {instruction}"
        if context_str:
            user_prompt += f"\n\nAvailable context:\n{context_str}"

        try:
            result = LLMHandler.generate(
                provider="Gemini",
                model="gemini-2.0-flash",
                system_prompt=system_prompt,
                user_prompt=user_prompt,
                file_contents="",
                settings={"temperature": 0.2, "max_tokens": 100, "stream": False}
            )
            # Clean up the result
            result = result.strip().strip('"').strip("'")
            if result:
                return result
        except Exception as e:
            LogManager().add_log("Email Monitor", f"LLM context extraction failed: {e}")

        # Fallback to simple extraction if context-aware fails
        return self._extract_action_simple(instruction)

    def _extract_action_simple(self, text: str) -> str:
        """Simple LLM extraction without context (fallback)."""
        if not API_KEYS.get("Gemini"):
            return ""

        system_prompt = """Extract the core action/task from the text. Remove conversational language.
Use imperative form. Output ONLY the action, no explanation."""

        try:
            result = LLMHandler.generate(
                provider="Gemini",
                model="gemini-2.0-flash",
                system_prompt=system_prompt,
                user_prompt=f"Extract action: {text}",
                file_contents="",
                settings={"temperature": 0.1, "max_tokens": 50, "stream": False}
            )
            result = result.strip().strip('"').strip("'")
            if result:
                return result
        except Exception:
            pass

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
