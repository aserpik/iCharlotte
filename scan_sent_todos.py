"""
Scan Sent Items for Proposed To-Dos

Scans the last 2 weeks of sent emails from Outlook, extracts potential
to-do items using the same logic as the email monitor, and presents
them for user selection before adding to the master list.
"""

import re
import sys
from datetime import datetime, timedelta
from typing import List, Dict, Optional

from PySide6.QtWidgets import (
    QApplication, QDialog, QVBoxLayout, QHBoxLayout, QLabel,
    QPushButton, QListWidget, QListWidgetItem, QCheckBox,
    QMessageBox, QProgressDialog, QWidget, QFrame
)
from PySide6.QtCore import Qt, QThread, Signal

from icharlotte_core.master_db import MasterCaseDatabase
from icharlotte_core.llm import LLMHandler
from icharlotte_core.config import API_KEYS


class ProposedTodo:
    """Represents a proposed todo item extracted from an email."""
    def __init__(self, file_number: str, raw_text: str, cleaned_text: str,
                 assignee: str, subject: str, sent_date: datetime, entry_id: str):
        self.file_number = file_number
        self.raw_text = raw_text
        self.cleaned_text = cleaned_text
        self.assignee = assignee
        self.subject = subject
        self.sent_date = sent_date
        self.entry_id = entry_id


class EmailScanWorker(QThread):
    """Background worker to scan sent emails."""
    progress = Signal(str)
    found_todo = Signal(object)  # ProposedTodo
    finished = Signal(int)  # count
    error = Signal(str)

    # Same configuration as SentItemsMonitorWorker
    FILE_NUMBER_REGEX = r'(\d{4})[._\-\u2013\u2014](\d{3})'
    TARGETS = [
        ("serpiklaw@gmail.com", "AS -", "AS"),
        ("fmottola@bordinsemmer.com", "FM -- ", "FM"),
        ("hproshyan@bordinsemmer.com", "HP -- ", "HP"),
        ("cetmekjian@bordinsemmer.com", "CE -- ", "CE"),
    ]
    FILLER_WORDS = ['please', 'pls', 'thnx', 'thanks', 'thx', 'ty']

    def __init__(self, days_back: int = 14):
        super().__init__()
        self.days_back = days_back
        self.db = MasterCaseDatabase()
        self.stop_requested = False

    @staticmethod
    def _normalize_dashes(text: str) -> str:
        """Convert en-dash and em-dash to regular hyphen."""
        return text.replace('\u2013', '-').replace('\u2014', '-').replace('\u2012', '-')

    def run(self):
        import pythoncom
        import win32com.client

        try:
            pythoncom.CoInitialize()
            self.progress.emit("Connecting to Outlook...")

            outlook = win32com.client.Dispatch("Outlook.Application")
            mapi = outlook.GetNamespace("MAPI")
            sent_folder = mapi.GetDefaultFolder(5)  # olFolderSentMail

            # Filter items from last N days
            cutoff_time = datetime.now() - timedelta(days=self.days_back)
            cutoff_str = cutoff_time.strftime("%m/%d/%Y %H:%M %p")
            restriction = f"[SentOn] >= '{cutoff_str}'"

            self.progress.emit(f"Scanning emails from last {self.days_back} days...")
            items = sent_folder.Items.Restrict(restriction)

            count = 0
            total = items.Count

            for idx, item in enumerate(items):
                if self.stop_requested:
                    break

                try:
                    self.progress.emit(f"Processing {idx + 1}/{total}...")
                    proposed = self._process_email(item)
                    if proposed:
                        self.found_todo.emit(proposed)
                        count += 1
                except Exception:
                    pass  # Skip problematic emails

            pythoncom.CoUninitialize()
            self.finished.emit(count)

        except Exception as e:
            self.error.emit(f"Scan error: {e}")
            self.finished.emit(0)

    def _process_email(self, item) -> Optional[ProposedTodo]:
        """Process a single email and return ProposedTodo if it matches criteria."""
        try:
            entry_id = item.EntryID

            # Skip if already processed by the monitor
            if self.db.is_email_processed(entry_id):
                return None

            # Get all recipient email addresses
            recipient_emails = []
            try:
                for recipient in item.Recipients:
                    email_addr = ""
                    try:
                        email_addr = recipient.AddressEntry.GetExchangeUser().PrimarySmtpAddress
                    except:
                        email_addr = recipient.Address or ""
                    if email_addr:
                        recipient_emails.append(email_addr.lower())
            except:
                to_field = str(item.To or "").lower()
                recipient_emails = [to_field]

            # Check body
            body = str(item.Body or "").strip()
            body_normalized = self._normalize_dashes(body)

            # Find matching target configuration
            matched_target = None
            for target_email, body_prefix, assignee in self.TARGETS:
                email_matches = any(target_email.lower() in addr for addr in recipient_emails)
                if not email_matches:
                    continue
                if body_normalized.upper().startswith(body_prefix.upper()):
                    matched_target = (target_email, body_prefix, assignee)
                    break

            if not matched_target:
                return None

            target_email, body_prefix, assignee = matched_target
            subject = str(item.Subject or "")

            # Extract file number
            file_number = self._extract_file_number(subject)
            if not file_number:
                return None

            # Verify case exists
            case = self.db.get_case(file_number)
            if not case:
                return None

            # Extract todo text
            raw_text = self._get_first_line(body_normalized, body_prefix)
            if not raw_text:
                return None

            cleaned_text = self._extract_action_with_llm(raw_text)
            if not cleaned_text:
                cleaned_text = self._extract_todo_text_fallback(raw_text)

            if not cleaned_text:
                return None

            # Get sent date
            try:
                sent_date = datetime(
                    item.SentOn.year, item.SentOn.month, item.SentOn.day,
                    item.SentOn.hour, item.SentOn.minute
                )
            except:
                sent_date = datetime.now()

            return ProposedTodo(
                file_number=file_number,
                raw_text=raw_text,
                cleaned_text=cleaned_text,
                assignee=assignee,
                subject=subject,
                sent_date=sent_date,
                entry_id=entry_id
            )

        except AttributeError:
            return None

    def _extract_file_number(self, subject: str) -> str:
        """Extract file number from subject line."""
        match = re.search(self.FILE_NUMBER_REGEX, subject)
        if match:
            return f"{match.group(1)}.{match.group(2)}"
        return ""

    def _get_first_line(self, body: str, body_prefix: str) -> str:
        """Get the first line of the body after removing the prefix."""
        text = body.strip()
        if text.upper().startswith(body_prefix.upper()):
            text = text[len(body_prefix):].strip()
        lines = text.split('\n')
        return lines[0].strip() if lines else ""

    def _extract_action_with_llm(self, text: str) -> str:
        """Use LLM to extract the action from email text."""
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
            result = result.strip().strip('"').strip("'")
            return result if result else ""
        except:
            return ""

    def _extract_todo_text_fallback(self, first_line: str) -> str:
        """Fallback rule-based extraction."""
        words = first_line.split()
        cleaned_words = []
        for word in words:
            word_clean = re.sub(r'[^\w]', '', word.lower())
            if word_clean not in self.FILLER_WORDS:
                cleaned_words.append(word)

        result = ' '.join(cleaned_words)
        result = re.sub(r'\s*,\s*$', '', result)
        result = re.sub(r'\s+', ' ', result)
        result = result.strip()

        if result:
            result = result[0].upper() + result[1:]
        return result


class TodoItemWidget(QWidget):
    """Widget for displaying a proposed todo item with checkbox."""
    def __init__(self, proposed: ProposedTodo, parent=None):
        super().__init__(parent)
        self.proposed = proposed

        layout = QHBoxLayout(self)
        layout.setContentsMargins(5, 5, 5, 5)

        # Checkbox
        self.checkbox = QCheckBox()
        self.checkbox.setChecked(True)  # Default to checked
        layout.addWidget(self.checkbox)

        # Info layout
        info_layout = QVBoxLayout()
        info_layout.setSpacing(2)

        # File number and assignee
        header = QLabel(f"<b>{proposed.file_number}</b> → {proposed.assignee}")
        header.setStyleSheet("color: #1565c0;")
        info_layout.addWidget(header)

        # Cleaned todo text
        todo_label = QLabel(proposed.cleaned_text)
        todo_label.setStyleSheet("font-size: 13px; padding-left: 10px;")
        todo_label.setWordWrap(True)
        info_layout.addWidget(todo_label)

        # Date and subject (smaller)
        date_str = proposed.sent_date.strftime("%m/%d/%y %H:%M")
        meta = QLabel(f"<i>Sent: {date_str} | Subject: {proposed.subject[:50]}...</i>")
        meta.setStyleSheet("color: gray; font-size: 10px; padding-left: 10px;")
        info_layout.addWidget(meta)

        layout.addLayout(info_layout, 1)

    def is_selected(self) -> bool:
        return self.checkbox.isChecked()


class SentTodoScannerDialog(QDialog):
    """Dialog for scanning and selecting todos from sent emails."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Scan Sent Items for To-Dos")
        self.setMinimumSize(700, 500)
        self.db = MasterCaseDatabase()
        self.proposed_todos: List[ProposedTodo] = []
        self.item_widgets: List[TodoItemWidget] = []

        self.setup_ui()

    def setup_ui(self):
        layout = QVBoxLayout(self)

        # Header
        header = QLabel("Scan your sent emails for potential to-do items")
        header.setStyleSheet("font-size: 14px; font-weight: bold; margin-bottom: 10px;")
        layout.addWidget(header)

        # Info
        info = QLabel("This will scan the last 2 weeks of sent emails using the same\n"
                      "logic as the email monitor (looking for 'AS -', 'FM -- ', etc. prefixes).")
        info.setStyleSheet("color: gray; margin-bottom: 10px;")
        layout.addWidget(info)

        # Scan button
        self.scan_btn = QPushButton("Scan Sent Items (Last 2 Weeks)")
        self.scan_btn.setStyleSheet("""
            QPushButton {
                background-color: #2196f3;
                color: white;
                padding: 10px 20px;
                border-radius: 5px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #1976d2;
            }
            QPushButton:disabled {
                background-color: #bdbdbd;
            }
        """)
        self.scan_btn.clicked.connect(self.start_scan)
        layout.addWidget(self.scan_btn)

        # Status label
        self.status_label = QLabel("")
        self.status_label.setStyleSheet("color: #666; margin: 5px 0;")
        layout.addWidget(self.status_label)

        # Results list
        self.results_frame = QFrame()
        self.results_frame.setFrameShape(QFrame.Shape.StyledPanel)
        self.results_frame.setVisible(False)
        results_layout = QVBoxLayout(self.results_frame)

        # Select all / none buttons
        btn_layout = QHBoxLayout()
        self.select_all_btn = QPushButton("Select All")
        self.select_all_btn.clicked.connect(self.select_all)
        self.select_none_btn = QPushButton("Select None")
        self.select_none_btn.clicked.connect(self.select_none)
        btn_layout.addWidget(self.select_all_btn)
        btn_layout.addWidget(self.select_none_btn)
        btn_layout.addStretch()
        results_layout.addLayout(btn_layout)

        # List widget
        self.list_widget = QListWidget()
        self.list_widget.setSpacing(5)
        results_layout.addWidget(self.list_widget)

        layout.addWidget(self.results_frame, 1)

        # Bottom buttons
        bottom_layout = QHBoxLayout()
        bottom_layout.addStretch()

        self.add_btn = QPushButton("Add Selected To-Dos")
        self.add_btn.setStyleSheet("""
            QPushButton {
                background-color: #4caf50;
                color: white;
                padding: 10px 20px;
                border-radius: 5px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #388e3c;
            }
            QPushButton:disabled {
                background-color: #bdbdbd;
            }
        """)
        self.add_btn.setEnabled(False)
        self.add_btn.clicked.connect(self.add_selected_todos)
        bottom_layout.addWidget(self.add_btn)

        self.close_btn = QPushButton("Close")
        self.close_btn.clicked.connect(self.reject)
        bottom_layout.addWidget(self.close_btn)

        layout.addLayout(bottom_layout)

    def start_scan(self):
        """Start the email scan."""
        self.scan_btn.setEnabled(False)
        self.list_widget.clear()
        self.proposed_todos.clear()
        self.item_widgets.clear()
        self.results_frame.setVisible(False)
        self.add_btn.setEnabled(False)

        self.worker = EmailScanWorker(days_back=14)
        self.worker.progress.connect(self.on_progress)
        self.worker.found_todo.connect(self.on_found_todo)
        self.worker.finished.connect(self.on_scan_finished)
        self.worker.error.connect(self.on_error)
        self.worker.start()

    def on_progress(self, message: str):
        self.status_label.setText(message)

    def on_found_todo(self, proposed: ProposedTodo):
        self.proposed_todos.append(proposed)

    def on_scan_finished(self, count: int):
        self.scan_btn.setEnabled(True)

        if count == 0:
            self.status_label.setText("No new to-do items found in sent emails.")
            return

        self.status_label.setText(f"Found {count} proposed to-do items. Select which ones to add:")
        self.results_frame.setVisible(True)

        # Populate list
        for proposed in self.proposed_todos:
            item = QListWidgetItem()
            widget = TodoItemWidget(proposed)
            self.item_widgets.append(widget)

            item.setSizeHint(widget.sizeHint())
            self.list_widget.addItem(item)
            self.list_widget.setItemWidget(item, widget)

        self.add_btn.setEnabled(True)

    def on_error(self, message: str):
        self.scan_btn.setEnabled(True)
        self.status_label.setText(f"Error: {message}")
        QMessageBox.warning(self, "Scan Error", message)

    def select_all(self):
        for widget in self.item_widgets:
            widget.checkbox.setChecked(True)

    def select_none(self):
        for widget in self.item_widgets:
            widget.checkbox.setChecked(False)

    def add_selected_todos(self):
        """Add selected todos to the master list."""
        selected = [w.proposed for w in self.item_widgets if w.is_selected()]

        if not selected:
            QMessageBox.information(self, "No Selection", "Please select at least one to-do item to add.")
            return

        confirm = QMessageBox.question(
            self,
            "Confirm Add",
            f"Add {len(selected)} to-do items to the master list?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )

        if confirm != QMessageBox.StandardButton.Yes:
            return

        # Add todos
        added = 0
        for proposed in selected:
            try:
                today = datetime.now().strftime("%Y-%m-%d")
                self.db.add_todo(
                    file_number=proposed.file_number,
                    item=proposed.cleaned_text,
                    color='yellow',
                    created_date=today
                )

                # Assign to the appropriate person
                todos = self.db.get_todos(proposed.file_number)
                if todos:
                    latest_todo = todos[0]
                    self.db.update_todo_assignment(
                        latest_todo['id'],
                        proposed.assignee,
                        datetime.now().strftime("%m/%d/%y")
                    )

                # Mark email as processed so it won't show up again
                self.db.mark_email_processed(
                    proposed.entry_id,
                    proposed.file_number,
                    proposed.cleaned_text
                )

                added += 1
            except Exception as e:
                print(f"Error adding todo: {e}")

        QMessageBox.information(
            self,
            "Success",
            f"Added {added} to-do items to the master list.\n\n"
            "Refresh the Master List tab to see the new items."
        )

        self.accept()


def main():
    """Run the scanner dialog."""
    app = QApplication.instance()
    standalone = False

    if not app:
        app = QApplication(sys.argv)
        standalone = True

    dialog = SentTodoScannerDialog()
    dialog.exec()

    if standalone:
        sys.exit(0)


if __name__ == "__main__":
    main()
