import os
import json
import sys
from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, 
    QPlainTextEdit, QListWidget, QListWidgetItem, QLabel, QMessageBox, QSplitter,
    QCheckBox, QLineEdit, QTextBrowser,
    QProgressBar, QDialog, QFrame, QScrollArea, QSizePolicy
)
from PyQt6.QtGui import QTextCursor, QDesktopServices, QColor, QFont, QIcon, QPainter, QPen, QBrush
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QSize, QTimer, QRect

try:
    import win32com.client
    import pythoncom
    import win32gui
except ImportError:
    win32com = None
    pythoncom = None
    win32gui = None

import markdown
import re
import datetime

from ..config import GEMINI_DATA_DIR
from ..llm import LLMHandler, LLMWorker
from ..email_manager import EmailSyncWorker, EmailDatabase
from ..utils import format_date_to_mm_dd_yyyy

class CacheWorker(QThread):
    finished = pyqtSignal(str) # Returns cache name
    error = pyqtSignal(str)

    def __init__(self, provider, model, content, system_instruction=None):
        super().__init__()
        self.provider = provider
        self.model = model
        self.content = content
        self.system_instruction = system_instruction

    def run(self):
        try:
            # We use a TTL of 60 minutes for the chat session
            name = LLMHandler.create_cache(
                self.provider, self.model, self.content, 
                system_instruction=self.system_instruction,
                ttl_minutes=60
            )
            if name:
                self.finished.emit(name)
            else:
                self.error.emit("Failed to create cache (returned None)")
        except Exception as e:
            self.error.emit(str(e))

class ComposeDialog(QDialog):
    def __init__(self, parent=None, default_subject="", default_to="", default_body=""):
        super().__init__(parent)
        self.setWindowTitle("Compose Email")
        self.resize(600, 500)
        layout = QVBoxLayout(self)
        
        layout.addWidget(QLabel("To:"))
        self.to_input = QLineEdit(default_to)
        layout.addWidget(self.to_input)
        
        layout.addWidget(QLabel("Subject:"))
        self.subject_input = QLineEdit(default_subject)
        layout.addWidget(self.subject_input)
        
        layout.addWidget(QLabel("Body:"))
        self.body_input = QPlainTextEdit(default_body)
        layout.addWidget(self.body_input)
        
        btn_layout = QHBoxLayout()
        self.send_btn = QPushButton("Send")
        self.send_btn.clicked.connect(self.accept)
        btn_layout.addWidget(self.send_btn)
        
        cancel_btn = QPushButton("Cancel")
        cancel_btn.clicked.connect(self.reject)
        btn_layout.addWidget(cancel_btn)
        layout.addLayout(btn_layout)

    def get_data(self):
        return {
            'to': self.to_input.text(),
            'subject': self.subject_input.text(),
            'body': self.body_input.toPlainText()
        }

class EmailChatDialog(QDialog):
    open_email = pyqtSignal(dict) # Emits the full email data

    def __init__(self, parent=None, email_list=None):
        super().__init__(parent)
        self.setWindowTitle("Email Intelligence (Gemini)")
        self.resize(600, 700)
        
        self.email_list = email_list or []
        self.email_map = {} # Index -> Email Data
        self.context_string = self._build_context()
        self.cache_name = None
        self.current_stream_text = ""
        
        layout = QVBoxLayout(self)
        
        self.history = QTextBrowser()
        self.history.setOpenExternalLinks(False) 
        self.history.setOpenLinks(False)
        self.history.anchorClicked.connect(self.on_link_clicked)
        layout.addWidget(self.history)
        
        input_layout = QHBoxLayout()
        self.input_field = QLineEdit()
        self.input_field.setPlaceholderText("Initializing AI context...")
        self.input_field.setEnabled(False) # Wait for cache
        self.input_field.returnPressed.connect(self.send_query)
        input_layout.addWidget(self.input_field)
        
        self.send_btn = QPushButton("Send")
        self.send_btn.setEnabled(False)
        self.send_btn.clicked.connect(self.send_query)
        input_layout.addWidget(self.send_btn)
        
        layout.addLayout(input_layout)
        
        # Initial greeting
        self.append_message("System", f"Loaded {len(self.email_list)} emails. Optimizing for speed...")
        
        self.system_prompt = (
            "You are an expert legal assistant analyzing a database of emails for a specific case. "
            "The user is Andrei Serpik (email: aserpik@bordinsemmer.com). "
            "When the user says 'I', 'me', or 'my', they are referring to Andrei Serpik. "
            "You have access to the full text of all emails, numbered sequentially (e.g., EMAIL 1). "
            "Answer the user's questions based ONLY on the provided email context. "
            "CRITICAL: When citing information, ALWAYS reference the email number using the format '[Email X]'. "
            "Example: 'The settlement was discussed in [Email 5] and confirmed in [Email 8].' "
            "If the answer is not in the emails, state that clearly."
        )

        # Start Cache Creation
        self.cache_worker = CacheWorker(
            "Gemini", "gemini-3-flash-preview", 
            self.context_string, 
            system_instruction=self.system_prompt
        ) 
        
        self.cache_worker.finished.connect(self.on_cache_ready)
        self.cache_worker.error.connect(self.on_cache_error)
        self.cache_worker.start()

    def on_cache_ready(self, name):
        self.cache_name = name
        self.append_message("System", "Context cached! Responses will be faster.")
        self.input_field.setPlaceholderText("Ask a question about these emails...")
        self.input_field.setEnabled(True)
        self.send_btn.setEnabled(True)
        self.input_field.setFocus()

    def on_cache_error(self, err):
        self.append_message("System", f"Note: Context caching might be limited for this preview model. Switching to standard mode (slightly slower).")
        self.cache_name = None
        self.input_field.setPlaceholderText("Ask a question...")
        self.input_field.setEnabled(True)
        self.send_btn.setEnabled(True)

    def _build_context(self):
        parts = []
        for i, email in enumerate(self.email_list):
            idx = i + 1
            self.email_map[idx] = email
            body = (email.get('body_text') or "").strip()
            entry = (
                f"--- EMAIL {idx} ---\n"
                f"Date: {email['received_time']}\n"
                f"From: {email['sender']}\n"
                f"Subject: {email['subject']}\n"
                f"Body:\n{body}\n"
            )
            parts.append(entry)
        return "\n".join(parts)

    def append_message(self, role, text):
        color = "#0000FF" if role == "User" else "#008800"
        
        if role == "Gemini":
            self.render_gemini_block(text, color)
        else:
            self.history.append(f"<b style='color:{color}'>{role}:</b> {text}<br>")

    def render_gemini_block(self, text, color="#008800"):
        # Convert Markdown to HTML
        html_content = markdown.markdown(text, extensions=['fenced_code', 'tables', 'nl2br'])
        # Linkify citations
        html_content = re.sub(r'\[Email (\d+)\]', r'<a href="email:\1">[Email \1]</a>', html_content)
        
        self.history.append(f"<b style='color:{color}'>Gemini:</b>")
        self.history.insertHtml(html_content)
        self.history.append("<br>")

    def on_link_clicked(self, url):
        scheme = url.scheme()
        if scheme == "email":
            try:
                idx = int(url.path())
                email_data = self.email_map.get(idx)
                if email_data:
                    self.open_email.emit(email_data)
                else:
                    QMessageBox.warning(self, "Error", f"Email {idx} not found in map.")
            except ValueError:
                pass
        else:
            # Open real links in browser
            QDesktopServices.openUrl(url)

    def send_query(self):
        query = self.input_field.text().strip()
        if not query:
            return
            
        self.append_message("User", query)
        self.input_field.clear()
        self.input_field.setEnabled(False)
        self.send_btn.setEnabled(False)
        
        # Prepare for streaming response
        self.current_stream_text = ""
        self.history.append(f"<b style='color:#008800'>Gemini:</b>")
        
        # Capture start position for streaming text (after the label)
        cursor = self.history.textCursor()
        cursor.movePosition(QTextCursor.MoveOperation.End)
        self.stream_start_position = cursor.position()
        
        # Settings
        settings = {
            'temperature': 1.0,
            'max_tokens': -1,
            'thinking_level': "None",
            'stream': True,
            'cache_name': self.cache_name
        }
        
        self.worker = LLMWorker(
            provider="Gemini",
            model="gemini-3-flash-preview", 
            system=self.system_prompt,
            user=query,
            files=self.context_string,
            settings=settings
        )
        self.worker.new_token.connect(self.on_token)
        self.worker.finished.connect(self.on_response_finished)
        self.worker.error.connect(self.on_error)
        self.worker.start()
        
    def on_token(self, token):
        self.current_stream_text += token
        # Append plain text for speed
        cursor = self.history.textCursor()
        cursor.movePosition(QTextCursor.MoveOperation.End)
        cursor.insertText(token)
        self.history.setTextCursor(cursor)
        
    def on_response_finished(self, full_text):
        # Replace the raw streamed text with rendered HTML
        cursor = self.history.textCursor()
        
        # 1. Select the raw text range
        cursor.setPosition(self.stream_start_position)
        cursor.movePosition(QTextCursor.MoveOperation.End, QTextCursor.MoveMode.KeepAnchor)
        
        # 2. Remove it
        cursor.removeSelectedText()
        
        # 3. Insert Rendered HTML
        # Helper to render
        html_content = markdown.markdown(full_text, extensions=['fenced_code', 'tables', 'nl2br'])
        html_content = re.sub(r'\[Email (\d+)\]', r'<a href="email:\1">[Email \1]</a>', html_content)
        
        cursor.insertHtml(html_content)
        self.history.append("<br>") # Spacing after message
        
        self.input_field.setEnabled(True)
        self.send_btn.setEnabled(True)
        self.input_field.setFocus()

    def on_error(self, err):
        self.append_message("Error", str(err))
        self.input_field.setEnabled(True)
        self.send_btn.setEnabled(True)

class EmailWorker(QThread):
    finished = pyqtSignal()
    error = pyqtSignal(str)

    def __init__(self, recipient, subject, body):
        super().__init__()
        self.recipient = recipient
        self.subject = subject
        self.body = body

    def run(self):
        try:
            if not win32com or not pythoncom:
                raise ImportError("pywin32 module not found.")

            pythoncom.CoInitialize()
            outlook = win32com.client.Dispatch('Outlook.Application')
            mail = outlook.CreateItem(0) # 0 = olMailItem
            mail.To = self.recipient
            mail.Subject = self.subject
            mail.Body = self.body
            mail.Send()
            pythoncom.CoUninitialize()
            self.finished.emit()
        except Exception as e:
            self.error.emit(str(e))

class AvatarWidget(QWidget):
    def __init__(self, name, size=40, parent=None):
        super().__init__(parent)
        self.setFixedSize(size, size)
        self.name = name
        self.color = self._generate_color(name)

    def _generate_color(self, name):
        # Simple hash to color
        hash_val = sum(ord(c) for c in name)
        hue = hash_val % 360
        return QColor.fromHsl(hue, 150, 100)

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        
        # Draw Circle
        painter.setBrush(QBrush(self.color))
        painter.setPen(Qt.PenStyle.NoPen)
        painter.drawEllipse(0, 0, self.width(), self.height())
        
        # Draw Initials
        initials = "".join([n[0] for n in self.name.split()[:2]]).upper()
        painter.setPen(Qt.GlobalColor.white)
        font = painter.font()
        font.setBold(True)
        font.setPixelSize(self.width() // 2)
        painter.setFont(font)
        painter.drawText(self.rect(), Qt.AlignmentFlag.AlignCenter, initials)

class EmailListItem(QWidget):
    def __init__(self, email_data, parent=None):
        super().__init__(parent)
        self.email_data = email_data
        
        layout = QVBoxLayout(self)
        layout.setContentsMargins(10, 8, 10, 8)
        layout.setSpacing(2)
        
        # Row 1: Sender Name & Date
        row1 = QHBoxLayout()
        sender_label = QLabel(email_data.get('sender', 'Unknown'))
        sender_label.setStyleSheet("font-weight: bold; font-size: 14px; color: #202020;")
        row1.addWidget(sender_label)
        
        row1.addStretch()
        
        # Format Date
        # Assume ISO or similar: YYYY-MM-DD HH:MM:SS
        raw_date = email_data.get('received_time', '')
        date_str = format_date_to_mm_dd_yyyy(raw_date)
        
        date_label = QLabel(date_str)
        date_label.setStyleSheet("color: #0078D4; font-size: 12px;") # Outlook Blueish
        row1.addWidget(date_label)
        
        if email_data.get('has_attachments'):
             attach_icon = QLabel("📎")
             attach_icon.setStyleSheet("color: #666; font-size: 12px;")
             row1.addWidget(attach_icon)

        layout.addLayout(row1)
        
        # Row 2: Subject
        subject = email_data.get('subject', '(No Subject)')
        subject_label = QLabel(subject)
        subject_label.setStyleSheet("font-size: 13px; color: #333;")
        subject_label.setWordWrap(False)
        layout.addWidget(subject_label)
        
        # Row 3: Preview (Body Snippet)
        body = email_data.get('body_text', '').replace('\\n', ' ').strip()
        preview_label = QLabel(body[:100])
        preview_label.setStyleSheet("color: #666; font-size: 12px;")
        preview_label.setWordWrap(False)
        layout.addWidget(preview_label)

class EmailTab(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.current_case_number = None
        self.db = None
        self.setup_ui()

    def setup_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(0,0,0,0)
        main_layout.setSpacing(0)
        
        # --- Toolbar (Top Bar) ---
        toolbar = QFrame()
        toolbar.setStyleSheet("background-color: #f3f2f1; border-bottom: 1px solid #e1dfdd; padding: 4px;")
        toolbar.setFixedHeight(50)  # Fix height to single line
        tb_layout = QHBoxLayout(toolbar)
        tb_layout.setContentsMargins(10, 5, 10, 5)
        
        self.sync_btn = QPushButton("Sync")
        self.sync_btn.clicked.connect(self.start_sync)
        # Use flat Outlook-style buttons if possible, or standard clean buttons
        self.sync_btn.setStyleSheet("""
            QPushButton { border: none; background: transparent; padding: 5px; font-weight: bold; color: #0078D4; }
            QPushButton:hover { background-color: #eaeaea; }
        """)
        tb_layout.addWidget(self.sync_btn)
        
        self.full_sync_cb = QCheckBox("Full")
        tb_layout.addWidget(self.full_sync_cb)
        
        tb_layout.addSpacing(20)
        
        search_container = QFrame()
        search_container.setStyleSheet("background-color: white; border: 1px solid #ccc; border-radius: 4px;")
        search_container.setFixedHeight(32) # Fix search box height
        sc_layout = QHBoxLayout(search_container)
        sc_layout.setContentsMargins(5, 0, 5, 0) # Reduce margins
        
        self.search_bar = QLineEdit()
        self.search_bar.setPlaceholderText("Search")
        self.search_bar.setStyleSheet("border: none;")
        self.search_bar.returnPressed.connect(self.perform_search)
        sc_layout.addWidget(self.search_bar)
        
        search_icon = QPushButton("🔍")
        search_icon.setStyleSheet("border: none; background: transparent;")
        search_icon.clicked.connect(self.perform_search)
        sc_layout.addWidget(search_icon)
        
        tb_layout.addWidget(search_container)
        
        tb_layout.addSpacing(10)
        self.hide_sent_cb = QCheckBox("Hide Sent")
        self.hide_sent_cb.stateChanged.connect(self.perform_search)
        tb_layout.addWidget(self.hide_sent_cb)
        
        tb_layout.addStretch()
        
        self.compose_btn = QPushButton("New Email")
        self.compose_btn.setStyleSheet("""
            QPushButton { background-color: #0078D4; color: white; border: none; padding: 6px 12px; border-radius: 2px; font-weight: bold; }
            QPushButton:hover { background-color: #106EBE; }
        """)
        self.compose_btn.clicked.connect(self.open_compose)
        tb_layout.addWidget(self.compose_btn)

        tb_layout.addSpacing(10)

        self.ask_gemini_btn = QPushButton("Ask Gemini")
        self.ask_gemini_btn.clicked.connect(self.open_gemini_chat)
        self.ask_gemini_btn.setStyleSheet("color: #8E24AA; border: 1px solid #8E24AA; padding: 5px 10px; border-radius: 2px; font-weight: bold;")
        tb_layout.addWidget(self.ask_gemini_btn)
        
        main_layout.addWidget(toolbar, 0) # Stretch 0
        
        # --- Progress Bar ---
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setFixedHeight(3)
        self.progress_bar.setStyleSheet("QProgressBar { border: none; background: #f0f0f0; } QProgressBar::chunk { background-color: #0078D4; }")
        main_layout.addWidget(self.progress_bar)

        self.status_label = QLabel("")
        self.status_label.setStyleSheet("padding: 2px 10px; color: #666;")
        main_layout.addWidget(self.status_label)

        # --- Split View ---
        splitter = QSplitter(Qt.Orientation.Horizontal)
        splitter.setHandleWidth(1)
        splitter.setStyleSheet("QSplitter::handle { background-color: #e1dfdd; }")
        main_layout.addWidget(splitter, 1) # Stretch 1 to take remaining space
        
        # === Left: Email List ===
        list_container = QWidget()
        list_container.setStyleSheet("background-color: white;")
        list_layout = QVBoxLayout(list_container)
        list_layout.setContentsMargins(0,0,0,0)
        list_layout.setSpacing(0)
        
        # List Header "Results / By Date"
        list_header = QFrame()
        list_header.setStyleSheet("border-bottom: 1px solid #e1dfdd; padding: 8px;")
        lh_layout = QHBoxLayout(list_header)
        lh_layout.setContentsMargins(5, 0, 5, 0)
        
        results_lbl = QLabel("Results")
        results_lbl.setStyleSheet("font-weight: bold; font-size: 14px; color: #333;")
        lh_layout.addWidget(results_lbl)
        lh_layout.addStretch()
        date_sort_btn = QPushButton("By Date ↓")
        date_sort_btn.setStyleSheet("border: none; color: #0078D4;")
        lh_layout.addWidget(date_sort_btn)
        
        list_layout.addWidget(list_header)
        
        self.email_list_widget = QListWidget()
        self.email_list_widget.setFrameShape(QFrame.Shape.NoFrame)
        self.email_list_widget.setStyleSheet("""
            QListWidget::item { border-bottom: 1px solid #f0f0f0; }
            QListWidget::item:selected { background-color: #c7e0f4; border-left: 4px solid #0078D4; }
            QListWidget::item:hover { background-color: #f3f2f1; }
        """)
        self.email_list_widget.itemSelectionChanged.connect(self.on_email_selected)
        list_layout.addWidget(self.email_list_widget)
        
        splitter.addWidget(list_container)
        
        # === Right: Reading Pane ===
        read_container = QWidget()
        read_container.setStyleSheet("background-color: white;")
        read_layout = QVBoxLayout(read_container)
        read_layout.setContentsMargins(0,0,0,0)
        read_layout.setSpacing(0)
        
        # Placeholder for when no email is selected
        self.placeholder_label = QLabel("Select an item to read")
        self.placeholder_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.placeholder_label.setStyleSheet("color: #888; font-size: 16px;")
        read_layout.addWidget(self.placeholder_label)

        # Actual Content Area (Hidden initially)
        self.content_area = QWidget()
        self.content_area.setVisible(False)
        content_layout = QVBoxLayout(self.content_area)
        content_layout.setContentsMargins(20, 20, 20, 20)
        content_layout.setSpacing(15)
        
        # 1. Subject Header
        self.subject_label = QLabel("Subject Line Here")
        self.subject_label.setWordWrap(True)
        self.subject_label.setStyleSheet("font-size: 20px; font-weight: 500; color: #202020;")
        content_layout.addWidget(self.subject_label)
        
        # 2. Sender Info Bar
        info_bar = QHBoxLayout()
        
        self.avatar_container = QWidget()
        av_layout = QVBoxLayout(self.avatar_container)
        av_layout.setContentsMargins(0,0,0,0)
        info_bar.addWidget(self.avatar_container)
        
        sender_col = QVBoxLayout()
        sender_col.setSpacing(2)
        
        self.from_label = QLabel("Sender Name")
        self.from_label.setStyleSheet("font-size: 14px; font-weight: bold; color: #202020;")
        sender_col.addWidget(self.from_label)
        
        self.to_cc_label = QLabel("To: Receiver; Cc: Others")
        self.to_cc_label.setStyleSheet("color: #666; font-size: 12px;")
        sender_col.addWidget(self.to_cc_label)
        
        info_bar.addLayout(sender_col)
        info_bar.addStretch()
        
        # Action Buttons (Reply, etc.)
        actions_layout = QHBoxLayout()
        actions_layout.setSpacing(5)
        
        def create_action_btn(text, icon_char, slot):
            btn = QPushButton(f" {icon_char} {text}")
            btn.setStyleSheet("""
                QPushButton { background: transparent; color: #444; border: 1px solid #ddd; border-radius: 3px; padding: 4px 10px; }
                QPushButton:hover { background-color: #f3f2f1; }
            """)
            btn.clicked.connect(slot)
            return btn
            
        self.reply_btn = create_action_btn("Reply", "↩", lambda: self.handle_reply(False))
        actions_layout.addWidget(self.reply_btn)
        
        self.reply_all_btn = create_action_btn("Reply All", "«", lambda: self.handle_reply(True))
        actions_layout.addWidget(self.reply_all_btn)
        
        self.forward_btn = create_action_btn("Forward", "→", self.handle_forward)
        actions_layout.addWidget(self.forward_btn)
        
        info_bar.addLayout(actions_layout)
        
        # Date on far right? Or part of action bar? Outlook puts date under actions or next to them.
        # Let's put date above actions or next to them.
        # We'll put date in a column to the right of actions for now.
        
        self.date_detail_label = QLabel("Fri 12/26/2025 1:08 PM")
        self.date_detail_label.setStyleSheet("color: #666; font-size: 12px; margin-left: 10px;")
        info_bar.addWidget(self.date_detail_label)
        
        content_layout.addLayout(info_bar)
        
        # 3. Attachments Area (Conditional)
        self.attachment_frame = QFrame()
        self.attachment_frame.setStyleSheet("background-color: #fff; border: 1px solid #e1dfdd; border-radius: 4px; padding: 5px;")
        self.attachment_frame.setVisible(False)
        att_layout = QHBoxLayout(self.attachment_frame)
        att_layout.setContentsMargins(5, 5, 5, 5)
        
        # Placeholder icon for PDF
        pdf_icon = QLabel("📄") 
        pdf_icon.setStyleSheet("font-size: 20px; color: #d0021b;")
        att_layout.addWidget(pdf_icon)
        
        self.attachment_label = QLabel("Attachments available (details not synced)")
        self.attachment_label.setStyleSheet("font-weight: bold; color: #333;")
        att_layout.addWidget(self.attachment_label)
        att_layout.addStretch()
        
        content_layout.addWidget(self.attachment_frame)

        # 4. Body
        self.body_view = QTextBrowser()
        self.body_view.setFrameShape(QFrame.Shape.NoFrame)
        self.body_view.setOpenExternalLinks(True)
        self.body_view.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.body_view.customContextMenuRequested.connect(self.show_body_context_menu)
        content_layout.addWidget(self.body_view)
        
        read_layout.addWidget(self.content_area)
        
        splitter.addWidget(read_container)
        splitter.setSizes([350, 800])

    def show_body_context_menu(self, pos):
        menu = self.body_view.createStandardContextMenu()
        
        cursor = self.body_view.textCursor()
        if cursor.hasSelection():
            menu.addSeparator()
            send_to_note_action = menu.addAction("Send to NoteTaker")
            send_to_note_action.triggered.connect(self.send_selection_to_notetaker)
            
        menu.exec(self.body_view.mapToGlobal(pos))

    def send_selection_to_notetaker(self):
        text = self.body_view.textCursor().selectedText().strip()
        if not text:
            return
            
        file_num = self.get_file_number()
        if not file_num:
            return

        # Prepare metadata for snippet
        row_item = self.email_list_widget.currentItem()
        email_data = {}
        if row_item:
            widget = self.email_list_widget.itemWidget(row_item)
            if widget:
                email_data = widget.email_data

        snippet = {
            "type": "email_snippet",
            "content": text,
            "source": f"Email from {email_data.get('sender', 'Unknown')} regarding '{email_data.get('subject', 'No Subject')}'",
            "date": email_data.get('received_time', 'Unknown'),
            "timestamp": str(datetime.datetime.now())
        }

        # Save to a snippets file in case directory
        case_dir = os.path.join(GEMINI_DATA_DIR, file_num)
        os.makedirs(case_dir, exist_ok=True)
        snippets_path = os.path.join(case_dir, "snippets.json")
        
        try:
            snippets = []
            if os.path.exists(snippets_path):
                with open(snippets_path, 'r', encoding='utf-8') as f:
                    snippets = json.load(f)
            
            snippets.append(snippet)
            
            with open(snippets_path, 'w', encoding='utf-8') as f:
                json.dump(snippets, f, indent=4)
                
            self.status_label.setText(f"Snippet saved to NoteTaker ({len(text)} chars)")
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Failed to save snippet: {e}")

    def get_file_number(self):
        # Try to get file number from MainWindow
        main_win = self.window()
        if main_win and hasattr(main_win, 'file_number'):
            return main_win.file_number
        return None

    def check_db_init(self):
        file_num = self.get_file_number()
        if not file_num:
            return False
            
        if self.current_case_number != file_num:
            self.current_case_number = file_num
            self.db = EmailDatabase(file_num)
            self.perform_search() # Load initial
        return True

    def start_sync(self):
        file_num = self.get_file_number()
        if not file_num:
            QMessageBox.warning(self, "Error", "No case file number loaded.")
            return

        self.sync_btn.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setRange(0, 0) # Indeterminate
        
        self.worker = EmailSyncWorker(file_num, full_sync=self.full_sync_cb.isChecked())
        self.worker.progress.connect(self.on_sync_progress)
        self.worker.finished.connect(self.on_sync_finished)
        self.worker.error.connect(self.on_sync_error)
        self.worker.start()

    def on_sync_progress(self, msg):
        self.status_label.setText(msg)

    def on_sync_finished(self):
        self.status_label.setText("Sync Complete")
        self.progress_bar.setVisible(False)
        self.sync_btn.setEnabled(True)
        self.check_db_init()
        self.perform_search() # Refresh list

    def on_sync_error(self, err):
        self.status_label.setText(f"Error: {err}")
        self.progress_bar.setVisible(False)
        self.sync_btn.setEnabled(True)
        QMessageBox.critical(self, "Sync Error", str(err))

    def perform_search(self):
        if not self.check_db_init():
            return

        query = self.search_bar.text().strip()
        results = self.db.search_emails(query if query else None)
        
        # Filter Sent Emails if Checkbox is Checked
        if self.hide_sent_cb.isChecked():
            filtered_results = []
            for email in results:
                # 1. Check folder path
                folder = email.get('folder_path', '').lower()
                if "sent items" in folder:
                    continue
                
                # 2. Check Sender (Simple heuristic for "Me")
                sender = email.get('sender', '').lower()
                sender_email = email.get('sender_email', '').lower()
                
                if "aserpik" in sender_email or "serpik, andrei" in sender:
                    continue
                    
                filtered_results.append(email)
            results = filtered_results

        self.email_list_widget.clear()
        
        for email in results:
            item = QListWidgetItem(self.email_list_widget)
            item.setSizeHint(QSize(0, 85)) # Adjust height for rows
            
            widget = EmailListItem(email)
            self.email_list_widget.setItemWidget(item, widget)

    def on_email_selected(self):
        item = self.email_list_widget.currentItem()
        if not item:
            self.placeholder_label.setVisible(True)
            self.content_area.setVisible(False)
            return
            
        widget = self.email_list_widget.itemWidget(item)
        if not widget: return
        
        self.placeholder_label.setVisible(False)
        self.content_area.setVisible(True)
        self.display_email(widget.email_data)

    def display_email(self, email_data):
        self.current_email_data = email_data # Store for reply/forward
        
        # 1. Header Fields
        self.subject_label.setText(email_data.get('subject', '(No Subject)'))
        self.from_label.setText(email_data.get('sender', 'Unknown'))
        
        # To/CC
        to_str = email_data.get('to_recipients', '')
        cc_str = email_data.get('cc_recipients', '')
        recipients = f"To: {to_str}"
        if cc_str:
            recipients += f"; Cc: {cc_str}"
        self.to_cc_label.setText(recipients)
        
        # Date
        self.date_detail_label.setText(email_data.get('received_time', ''))
        
        # Avatar
        # Clear old avatar
        if self.avatar_container.layout().count():
            self.avatar_container.layout().takeAt(0).widget().deleteLater()
        
        avatar = AvatarWidget(email_data.get('sender', 'Unknown'), size=45)
        self.avatar_container.layout().addWidget(avatar)

        # 2. Attachments
        has_attachments = email_data.get('has_attachments', False)
        self.attachment_frame.setVisible(has_attachments)

        # 3. Body
        body_html = email_data.get('body_html', '')
        # Basic cleanup or injection of style
        if body_html and len(body_html) > 50: 
            self.body_view.setHtml(body_html)
        else:
            text = email_data.get('body_text', '')
            self.body_view.setPlainText(text)

    def handle_reply(self, reply_all=False):
        if not hasattr(self, 'current_email_data'): return
        
        data = self.current_email_data
        subject = data.get('subject', '')
        if not subject.lower().startswith("re:"):
            subject = "RE: " + subject
            
        sender_email = data.get('sender_email', '')
        # If reply all, add 'to' and 'cc' to recipient list
        to = sender_email
        if reply_all:
             others = data.get('to_recipients', '') + ";" + data.get('cc_recipients', '')
             to += ";" + others
        
        # Open Compose
        dlg = ComposeDialog(self, default_subject=subject, default_to=to)
        if dlg.exec():
            new_data = dlg.get_data()
            self.send_email(new_data['to'], new_data['subject'], new_data['body'])

    def handle_forward(self):
        if not hasattr(self, 'current_email_data'): return
        data = self.current_email_data
        
        subject = data.get('subject', '')
        if not subject.lower().startswith("fw:"):
            subject = "FW: " + subject
            
        # Construct forwarded body
        original_body = data.get('body_text', '')
        forward_header = (
            f"\n\n\n---------- Forwarded message ---------\n"
            f"From: {data.get('sender')}\n"
            f"Date: {data.get('received_time')}\n"
            f"Subject: {data.get('subject')}\n"
            f"To: {data.get('to_recipients')}\n\n"
        )
        body = forward_header + original_body
        
        dlg = ComposeDialog(self, default_subject=subject, default_body=body)
        if dlg.exec():
            new_data = dlg.get_data()
            self.send_email(new_data['to'], new_data['subject'], new_data['body'])

    def load_case_variables(self):
        file_num = self.get_file_number()
        if not file_num:
            return {}
        
        json_path = os.path.join(GEMINI_DATA_DIR, f"{file_num}.json")
        if not os.path.exists(json_path):
            return {}
            
        try:
            with open(json_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
                # Flatten the structure: {key: "val"} or {key: {value: "val"}} -> {key: "val"}
                flat_data = {}
                for k, v in data.items():
                    if isinstance(v, dict) and "value" in v:
                        flat_data[k] = v["value"]
                    else:
                        flat_data[k] = v
                return flat_data
        except Exception as e:
            print(f"Error loading vars: {e}")
            return {}

    def _parse_date(self, date_str):
        if not date_str: return None
        formats = [
            "%Y-%m-%d", "%m/%d/%Y", "%m-%d-%Y", "%d/%m/%Y", 
            "%B %d, %Y", "%b %d, %Y"
        ]
        for fmt in formats:
            try:
                return datetime.datetime.strptime(date_str.strip(), fmt)
            except ValueError:
                continue
        return None

    def _format_date_long(self, date_str):
        dt = self._parse_date(date_str)
        if dt:
            return dt.strftime("%B %d, %Y")
        return date_str # Fallback

    def generate_status_report(self):
        vars = self.load_case_variables()
        if not vars:
            QMessageBox.warning(self, "Error", "No case variables found. Please ensure variables are loaded for this case.")
            return

        # Store vars for the callback
        self._pending_report_vars = vars
        
        proc_hist = str(vars.get("procedural_history", "")).strip()
        
        if not proc_hist:
            # No history, just proceed immediately
            self._finalize_status_report("No future hearings are currently scheduled.")
            return

        # Prepare LLM request
        self.status_label.setText("Generating procedural summary...")
        self.status_report_btn.setEnabled(False) 

        today_str = datetime.datetime.now().strftime("%B %d, %Y")
        
        system_prompt = (
            "You are an expert legal assistant. Your task is to extract future hearing dates from a procedural history text "
            "and write them as clear, complete sentences."
        )
        
        user_prompt = (
            f"Today is {today_str}. Review the following procedural history text. "
            "Identify any hearing dates that occur AFTER today. "
            "For each future hearing, generate a complete, professional sentence stating what the hearing is and when it is scheduled. "
            "Example: 'The Court has scheduled a Case Management Conference for March 25, 2026.' "
            "Combine these sentences into a single paragraph. "
            "If there are no future hearings found, return exactly: 'No future hearings are currently scheduled.'\n\n"
            f"Text to analyze:\n{proc_hist}"
        )

        settings = {
            'temperature': 1.0,
            'max_tokens': -1,
            'stream': False
        }

        # Use gemini-3-flash-preview as default mode
        self.llm_worker = LLMWorker(
            "Gemini", "gemini-3-flash-preview", 
            system_prompt, user_prompt, "", settings
        )
        self.llm_worker.finished.connect(self._finalize_status_report)
        self.llm_worker.error.connect(lambda e: self._finalize_status_report(f"Error parsing history: {e}"))
        self.llm_worker.start()

    def _finalize_status_report(self, proc_status_text):
        self.status_label.setText("")
        self.status_report_btn.setEnabled(True)
        
        # Clean up any potential error prefix from lambda if needed
        if not proc_status_text:
             proc_status_text = "No future hearings are currently scheduled."

        vars = self._pending_report_vars
        
        # Helper to get var with default
        def get(k): return str(vars.get(k, "")).strip()

        # Gather Data
        adjuster_email = get("adjuster_email")
        client_email = get("Client_Email")
        case_name = get("Case_name")
        claim_num = get("claim_number")
        file_num = self.get_file_number()
        adjuster_name = get("adjuster_name")
        trial_date_raw = get("trial_date")
        fact_bg = get("factual_background")
        
        # Format Data
        trial_date = self._format_date_long(trial_date_raw) if trial_date_raw else ""
        
        # Subject
        subject = f"{case_name}; {claim_num}; ({file_num})"

        # Signature Block Extraction from Untitled.msg (Hardcoded snippet)
        signature_html = """
<p class=MsoNormal><a name="_MailAutoSig"><span style='font-family:"Garamond",serif;color:black;'>ANDREI V. SERPIK</span></a></p>
<p class=MsoNormal><i><span style='font-family:"Garamond",serif;color:black;'>Attorney</span></i></p>
<p class=MsoNormal><span style='font-family:"Garamond",serif;color:black;'>&nbsp;</span></p>
<p class=MsoNormal><span style='font-family:"Garamond",serif;'>BORDIN SEMMER LLP</span><br>
<span style='font-family:"Garamond",serif;'>Howard Hughes Center<br></span>
<span style='font-family:"Garamond",serif;color:#045EC2;'>6100 Center Drive, Suite 1100</span><br>
<span style='font-family:"Garamond",serif;'>Los Angeles, California 90045</span><br>
<span style='font-family:"Garamond",serif;color:#0563C1;'>Aserpik@bordinsemmer.com</span></p>
<p class=MsoNormal><span style='font-size:7.5pt;font-family:"Garamond",serif;'>TEL</span><span style='font-family:"Garamond",serif;'>&nbsp;&nbsp;</span><span style='font-family:"Garamond",serif;color:#045EC2;'>323.457.2110</span><br>
<span style='font-size:7.5pt;font-family:"Garamond",serif;'>FAX</span><span style='font-family:"Garamond",serif;'>&nbsp;&nbsp;</span><span style='font-family:"Garamond",serif;color:#045EC2;'>323.457.2120</span><br>
<span style='font-family:"Garamond",serif;color:black;'>Los Angeles&nbsp;|&nbsp;San Diego&nbsp;| Bay Area</span></p>
"""

        # Body Construction (HTML)
        # Font wrapper
        body_html = "<div style='font-family: Calibri; font-size: 11pt; color: #000000;'>"
        
        # Salutation
        first_name = adjuster_name.split()[0] if adjuster_name else "Adjuster"
        body_html += f"<p>{first_name},</p>"
        
        # Opening
        if trial_date:
            body_html += f"<p>Please allow the following to serve as our interim update in the above-referenced matter, which is currently set for a trial to begin on <b>{trial_date}</b>:</p>"
        else:
            body_html += f"<p>Please allow the following to serve as our interim update in the above-referenced matter, which is currently not set for trial:</p>"
            
        # Factual Background
        body_html += f"<p><b>Factual Background:</b> {fact_bg}</p>"
        
        # Procedural Status (Using LLM Output)
        body_html += f"<p><b>Procedural Status:</b> {proc_status_text}</p>"
        
        # New Topic Placeholder
        body_html += f"<p><b>New Topic Name Placeholder:</b> New Topic Paragraph Placeholder</p>"

        # Further Case Handling
        body_html += f"<p><b>Further Case Handling:</b> [Further Case Handling paragraph Placeholder.]</p>"
        
        # Sign-off
        body_html += "<p>Best,</p><p>Andrei</p>"
        
        # Insert Signature
        body_html += signature_html
        
        body_html += "</div>"

        # Recipients
        cc = client_email
        if cc: 
            cc += "; jbordinwosk@bordinsemmer.com"
        else:
            cc = "jbordinwosk@bordinsemmer.com"

        # Open Outlook
        try:
            if not win32com or not pythoncom:
                raise ImportError("pywin32 module not found.")

            pythoncom.CoInitialize()
            outlook = win32com.client.Dispatch('Outlook.Application')
            mail = outlook.CreateItem(0) # 0 = olMailItem
            mail.To = adjuster_email
            mail.CC = cc
            mail.Subject = subject
            # Use HTMLBody to support formatting. Prepend to existing HTMLBody to keep signature.
            mail.HTMLBody = body_html + mail.HTMLBody 
            
            mail.Display()
            
            # Force Foreground
            try:
                inspector = mail.GetInspector
                inspector.Activate()
                
                if win32gui:
                    hwnd = win32gui.FindWindow(None, inspector.Caption)
                    if hwnd:
                        if win32gui.IsIconic(hwnd):
                             win32gui.ShowWindow(hwnd, 9) # SW_RESTORE
                        else:
                             win32gui.ShowWindow(hwnd, 5) # SW_SHOW
                        win32gui.SetForegroundWindow(hwnd)
            except Exception as fe:
                print(f"Could not force foreground: {fe}")

            pythoncom.CoUninitialize()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to generate report: {e}")

    def open_compose(self):
        default_sub = self.get_file_number() or ""
        dlg = ComposeDialog(self, default_subject=default_sub)
        if dlg.exec():
            data = dlg.get_data()
            self.send_email(data['to'], data['subject'], data['body'])

    def open_gemini_chat(self):
        if not self.check_db_init():
            QMessageBox.warning(self, "Error", "No case file loaded.")
            return

        self.status_label.setText("Preparing email context for Gemini...")
        
        try:
            # Use the chronological list of full email data
            emails = self.db.get_chronological_emails()
            if not emails:
                QMessageBox.information(self, "Info", "No emails found to analyze.")
                self.status_label.setText("")
                return
                
            self.status_label.setText("")
            
            # Pass the list, let Dialog build the context string
            self.chat_dlg = EmailChatDialog(self, email_list=emails)
            # Link click now opens in Outlook instead of just displaying in Reading Pane
            self.chat_dlg.open_email.connect(self.open_email_in_outlook)
            self.chat_dlg.show() # Use show() instead of exec() to allow interacting with the main window (Reading Pane)
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to load context: {e}")
            self.status_label.setText("Error loading context.")

    def open_email_in_outlook(self, email_data):
        """Opens the specific email in Outlook using EntryID, with fallback to search."""
        try:
            if not win32com or not pythoncom:
                raise ImportError("pywin32 module not found.")

            entry_id = email_data.get('entry_id')
            subject = email_data.get('subject')
            
            if not entry_id:
                QMessageBox.warning(self, "Error", "Email data is missing EntryID.")
                return

            pythoncom.CoInitialize()
            outlook = win32com.client.Dispatch('Outlook.Application')
            namespace = outlook.GetNamespace("MAPI")
            
            try:
                # 1. Try Direct ID Access
                item = namespace.GetItemFromID(entry_id)
                item.Display()
                
                # Force to foreground
                try:
                    inspector = item.GetInspector
                    inspector.Activate()
                    
                    if win32gui:
                        # Find the window by caption and bring to front
                        hwnd = win32gui.FindWindow(None, inspector.Caption)
                        if hwnd:
                            win32gui.ShowWindow(hwnd, 5) # SW_SHOW
                            win32gui.SetForegroundWindow(hwnd)
                except Exception as fe:
                    print(f"Note: Could not force foreground: {fe}")
                    
            except Exception as e:
                print(f"Direct ID open failed ({e}), attempting fallback search...")
                QMessageBox.warning(self, "Link Error", f"Could not find exact email in Outlook.\nSubject: {subject}")

            pythoncom.CoUninitialize()
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to open in Outlook: {e}")

    def send_email(self, to, subject, body):
        self.status_label.setText("Sending email...")
        self.email_sender = EmailWorker(to, subject, body)
        self.email_sender.finished.connect(lambda: self.status_label.setText("Email Sent."))
        self.email_sender.error.connect(lambda e: QMessageBox.critical(self, "Error", f"Failed to send: {e}"))
        self.email_sender.start()
