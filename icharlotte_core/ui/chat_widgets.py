"""
Chat UI widgets for the enhanced ChatTab.
"""
import os
import re
from datetime import datetime, timedelta
from typing import Optional, List, Callable

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QFrame, QLabel, QPushButton,
    QListWidget, QListWidgetItem, QLineEdit, QPlainTextEdit, QTextBrowser,
    QScrollArea, QSplitter, QToolButton, QMenu, QSizePolicy, QProgressBar,
    QApplication, QInputDialog, QMessageBox, QToolTip
)
from PySide6.QtCore import Qt, Signal, QSize, QTimer, QPoint
from PySide6.QtGui import (
    QFont, QTextCursor, QColor, QPalette, QAction, QKeyEvent,
    QTextCharFormat, QSyntaxHighlighter, QTextDocument
)

try:
    from pygments import highlight
    from pygments.lexers import get_lexer_by_name, guess_lexer
    from pygments.formatters import HtmlFormatter
    PYGMENTS_AVAILABLE = True
except ImportError:
    PYGMENTS_AVAILABLE = False


# Theme definitions
THEMES = {
    "light": {
        "bg": "#ffffff",
        "panel_bg": "#f5f5f5",
        "text": "#202020",
        "text_secondary": "#666666",
        "user_bubble": "#e3f2fd",
        "assistant_bubble": "#f5f5f5",
        "accent": "#2196F3",
        "accent_hover": "#1976D2",
        "border": "#e0e0e0",
        "success": "#4CAF50",
        "warning": "#FF9800",
        "error": "#f44336",
        "code_bg": "#f8f8f8",
        "selected": "#e3f2fd",
        "hover": "#f0f0f0"
    },
    "dark": {
        "bg": "#1e1e1e",
        "panel_bg": "#252526",
        "text": "#e0e0e0",
        "text_secondary": "#a0a0a0",
        "user_bubble": "#2d3748",
        "assistant_bubble": "#374151",
        "accent": "#60a5fa",
        "accent_hover": "#3b82f6",
        "border": "#4a5568",
        "success": "#48bb78",
        "warning": "#ed8936",
        "error": "#fc8181",
        "code_bg": "#2d2d2d",
        "selected": "#3b82f6",
        "hover": "#3a3a3a"
    }
}

def get_theme(theme_name: str = 'light') -> dict:
    """Get theme colors by name."""
    return THEMES.get(theme_name, THEMES['light'])


class ConversationSidebar(QFrame):
    """Sidebar widget for managing conversations."""

    conversation_selected = Signal(str)  # Emits conversation ID
    new_conversation_requested = Signal()
    conversation_renamed = Signal(str, str)  # ID, new name
    conversation_deleted = Signal(str)  # ID

    def __init__(self, parent=None, theme='light'):
        super().__init__(parent)
        self.theme = theme
        self.conversations = []
        self.current_conversation_id = None
        self.setup_ui()
        self.apply_theme()

    def setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(8, 8, 8, 8)
        layout.setSpacing(8)

        # Header with new chat button
        header = QHBoxLayout()
        title = QLabel("Conversations")
        title.setFont(QFont("Segoe UI", 11, QFont.Weight.Bold))
        header.addWidget(title)
        header.addStretch()

        self.new_btn = QPushButton("+")
        self.new_btn.setFixedSize(28, 28)
        self.new_btn.setToolTip("New Conversation")
        self.new_btn.clicked.connect(self.new_conversation_requested.emit)
        header.addWidget(self.new_btn)

        layout.addLayout(header)

        # Search box
        self.search_box = QLineEdit()
        self.search_box.setPlaceholderText("Search conversations...")
        self.search_box.textChanged.connect(self.filter_conversations)
        layout.addWidget(self.search_box)

        # Conversation list
        self.conv_list = QListWidget()
        self.conv_list.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.conv_list.customContextMenuRequested.connect(self.show_context_menu)
        self.conv_list.itemClicked.connect(self.on_item_clicked)
        self.conv_list.itemDoubleClicked.connect(self.on_item_double_clicked)
        layout.addWidget(self.conv_list)

    def apply_theme(self):
        colors = get_theme(self.theme)
        self.setStyleSheet(f"""
            ConversationSidebar {{
                background-color: {colors['panel_bg']};
                border-right: 1px solid {colors['border']};
            }}
            QLabel {{
                color: {colors['text']};
            }}
            QLineEdit {{
                background-color: {colors['bg']};
                color: {colors['text']};
                border: 1px solid {colors['border']};
                border-radius: 4px;
                padding: 6px;
            }}
            QListWidget {{
                background-color: transparent;
                border: none;
                color: {colors['text']};
            }}
            QListWidget::item {{
                padding: 8px;
                border-radius: 4px;
                margin: 2px 0;
            }}
            QListWidget::item:hover {{
                background-color: {colors['hover']};
            }}
            QListWidget::item:selected {{
                background-color: {colors['selected']};
            }}
            QPushButton {{
                background-color: {colors['accent']};
                color: white;
                border: none;
                border-radius: 4px;
                font-weight: bold;
            }}
            QPushButton:hover {{
                background-color: {colors['accent_hover']};
            }}
        """)

    def set_conversations(self, conversations: list):
        """Set the list of conversations to display."""
        self.conversations = conversations
        self.refresh_list()

    def refresh_list(self):
        """Refresh the conversation list display."""
        self.conv_list.clear()
        search_text = self.search_box.text().lower()

        # Group conversations by date
        today = datetime.now().date()
        yesterday = today - timedelta(days=1)
        this_week = today - timedelta(days=7)

        groups = {
            'Today': [],
            'Yesterday': [],
            'This Week': [],
            'Older': []
        }

        for conv in self.conversations:
            name = conv.get('name', 'Untitled')
            conv_id = conv.get('id', '')
            updated = conv.get('updated_at', '')

            # Filter by search
            if search_text and search_text not in name.lower():
                continue

            # Parse date
            try:
                dt = datetime.fromisoformat(updated.replace('Z', '+00:00'))
                conv_date = dt.date()
            except:
                conv_date = None

            # Categorize
            if conv_date == today:
                groups['Today'].append(conv)
            elif conv_date == yesterday:
                groups['Yesterday'].append(conv)
            elif conv_date and conv_date >= this_week:
                groups['This Week'].append(conv)
            else:
                groups['Older'].append(conv)

        # Add to list with group headers
        for group_name, convs in groups.items():
            if not convs:
                continue

            # Group header
            header = QListWidgetItem(group_name)
            header.setFlags(Qt.ItemFlag.NoItemFlags)
            header.setFont(QFont("Segoe UI", 9, QFont.Weight.Bold))
            self.conv_list.addItem(header)

            # Conversations in group
            for conv in convs:
                item = QListWidgetItem(f"  {conv.get('name', 'Untitled')}")
                item.setData(Qt.ItemDataRole.UserRole, conv.get('id'))
                if conv.get('id') == self.current_conversation_id:
                    item.setSelected(True)
                self.conv_list.addItem(item)

    def filter_conversations(self, text: str):
        """Filter conversations by search text."""
        self.refresh_list()

    def on_item_clicked(self, item: QListWidgetItem):
        """Handle conversation item click."""
        conv_id = item.data(Qt.ItemDataRole.UserRole)
        if conv_id:
            self.current_conversation_id = conv_id
            self.conversation_selected.emit(conv_id)

    def on_item_double_clicked(self, item: QListWidgetItem):
        """Handle double-click to rename."""
        conv_id = item.data(Qt.ItemDataRole.UserRole)
        if conv_id:
            current_name = item.text().strip()
            new_name, ok = QInputDialog.getText(
                self, "Rename Conversation",
                "Enter new name:",
                text=current_name
            )
            if ok and new_name.strip():
                self.conversation_renamed.emit(conv_id, new_name.strip())

    def show_context_menu(self, pos):
        """Show context menu for conversation item."""
        item = self.conv_list.itemAt(pos)
        if not item:
            return
        conv_id = item.data(Qt.ItemDataRole.UserRole)
        if not conv_id:
            return

        menu = QMenu(self)

        rename_action = QAction("Rename", self)
        rename_action.triggered.connect(lambda: self.on_item_double_clicked(item))
        menu.addAction(rename_action)

        delete_action = QAction("Delete", self)
        delete_action.triggered.connect(lambda: self.confirm_delete(conv_id))
        menu.addAction(delete_action)

        menu.exec(self.conv_list.mapToGlobal(pos))

    def confirm_delete(self, conv_id: str):
        """Confirm and delete a conversation."""
        reply = QMessageBox.question(
            self, "Delete Conversation",
            "Are you sure you want to delete this conversation?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        if reply == QMessageBox.StandardButton.Yes:
            self.conversation_deleted.emit(conv_id)

    def set_current_conversation(self, conv_id: str):
        """Set the current conversation selection."""
        self.current_conversation_id = conv_id
        self.refresh_list()


class ResizableInputArea(QWidget):
    """Resizable text input area with template dropdown."""

    text_submitted = Signal(str)  # Emits text when Enter pressed
    template_requested = Signal()  # Emits when template button clicked

    def __init__(self, parent=None, theme='light'):
        super().__init__(parent)
        self.theme = theme
        self.min_height = 60
        self.max_height = 200
        self.setup_ui()
        self.apply_theme()

    def setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(4)

        # Toolbar row
        toolbar = QHBoxLayout()
        toolbar.setSpacing(4)

        # Template button
        self.template_btn = QToolButton()
        self.template_btn.setText("Templates")
        self.template_btn.setPopupMode(QToolButton.ToolButtonPopupMode.InstantPopup)
        self.template_btn.setToolTip("Quick prompts")
        self.template_btn.clicked.connect(self.template_requested.emit)
        toolbar.addWidget(self.template_btn)

        toolbar.addStretch()

        # Attachment indicator
        self.attachment_label = QLabel("")
        self.attachment_label.setVisible(False)
        toolbar.addWidget(self.attachment_label)

        layout.addLayout(toolbar)

        # Text input
        self.text_edit = QPlainTextEdit()
        self.text_edit.setMinimumHeight(self.min_height)
        self.text_edit.setMaximumHeight(self.max_height)
        self.text_edit.setPlaceholderText("Type a message... (Enter to send, Shift+Enter for newline)")
        self.text_edit.textChanged.connect(self.auto_resize)
        layout.addWidget(self.text_edit)

    def apply_theme(self):
        colors = get_theme(self.theme)
        self.setStyleSheet(f"""
            QPlainTextEdit {{
                background-color: {colors['bg']};
                color: {colors['text']};
                border: 1px solid {colors['border']};
                border-radius: 8px;
                padding: 8px;
                font-size: 13px;
            }}
            QPlainTextEdit:focus {{
                border-color: {colors['accent']};
            }}
            QToolButton {{
                background-color: transparent;
                color: {colors['text_secondary']};
                border: 1px solid {colors['border']};
                border-radius: 4px;
                padding: 4px 8px;
            }}
            QToolButton:hover {{
                background-color: {colors['hover']};
            }}
            QLabel {{
                color: {colors['text_secondary']};
                font-size: 12px;
            }}
        """)

    def auto_resize(self):
        """Auto-resize based on content."""
        doc = self.text_edit.document()
        height = int(doc.size().height()) + 20
        height = max(self.min_height, min(self.max_height, height))
        self.text_edit.setFixedHeight(height)

    def set_template_menu(self, menu: QMenu):
        """Set the template dropdown menu."""
        self.template_btn.setMenu(menu)
        self.template_btn.setPopupMode(QToolButton.ToolButtonPopupMode.InstantPopup)

    def show_attachments(self, count: int):
        """Show attachment count indicator."""
        if count > 0:
            self.attachment_label.setText(f"{count} file(s) attached")
            self.attachment_label.setVisible(True)
        else:
            self.attachment_label.setVisible(False)

    def toPlainText(self) -> str:
        return self.text_edit.toPlainText()

    def clear(self):
        self.text_edit.clear()

    def setPlaceholderText(self, text: str):
        self.text_edit.setPlaceholderText(text)

    def setAcceptDrops(self, accept: bool):
        self.text_edit.setAcceptDrops(accept)


class ContextIndicator(QWidget):
    """Widget showing context window usage."""

    clicked = Signal()  # Emitted when clicked for details

    def __init__(self, parent=None, theme='light'):
        super().__init__(parent)
        self.theme = theme
        self.total_tokens = 0
        self.context_limit = 128000
        self.setup_ui()
        self.apply_theme()

    def setup_ui(self):
        layout = QHBoxLayout(self)
        layout.setContentsMargins(8, 4, 8, 4)
        layout.setSpacing(8)

        self.label = QLabel("Context:")
        layout.addWidget(self.label)

        self.progress = QProgressBar()
        self.progress.setFixedHeight(12)
        self.progress.setTextVisible(False)
        layout.addWidget(self.progress, 1)

        self.count_label = QLabel("0 / 128k")
        layout.addWidget(self.count_label)

        self.setCursor(Qt.CursorShape.PointingHandCursor)

    def apply_theme(self):
        colors = get_theme(self.theme)
        self.setStyleSheet(f"""
            ContextIndicator {{
                background-color: {colors['panel_bg']};
                border: 1px solid {colors['border']};
                border-radius: 4px;
            }}
            QLabel {{
                color: {colors['text_secondary']};
                font-size: 11px;
            }}
            QProgressBar {{
                background-color: {colors['bg']};
                border: none;
                border-radius: 4px;
            }}
            QProgressBar::chunk {{
                background-color: {colors['accent']};
                border-radius: 4px;
            }}
        """)

    def update_usage(self, total_tokens: int, context_limit: int):
        """Update the context usage display."""
        self.total_tokens = total_tokens
        self.context_limit = context_limit

        percentage = (total_tokens / context_limit * 100) if context_limit > 0 else 0
        self.progress.setValue(int(percentage))

        # Format token counts
        def format_tokens(n):
            if n >= 1000000:
                return f"{n/1000000:.1f}M"
            elif n >= 1000:
                return f"{n/1000:.1f}k"
            return str(n)

        self.count_label.setText(f"~{format_tokens(total_tokens)} / {format_tokens(context_limit)}")

        # Update color based on usage
        colors = get_theme(self.theme)
        if percentage >= 95:
            chunk_color = colors['error']
        elif percentage >= 80:
            chunk_color = colors['warning']
        else:
            chunk_color = colors['accent']

        self.progress.setStyleSheet(f"""
            QProgressBar {{
                background-color: {colors['bg']};
                border: none;
                border-radius: 4px;
            }}
            QProgressBar::chunk {{
                background-color: {chunk_color};
                border-radius: 4px;
            }}
        """)

    def mousePressEvent(self, event):
        self.clicked.emit()
        super().mousePressEvent(event)


class CodeBlockWidget(QFrame):
    """Widget for displaying code with syntax highlighting and copy button."""

    def __init__(self, code: str, language: str = '', parent=None, theme='light'):
        super().__init__(parent)
        self.code = code
        self.language = language
        self.theme = theme
        self.setup_ui()

    def setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        # Header with language and copy button
        header = QFrame()
        header_layout = QHBoxLayout(header)
        header_layout.setContentsMargins(8, 4, 8, 4)

        lang_label = QLabel(self.language or "Code")
        lang_label.setStyleSheet("font-size: 11px; color: #888;")
        header_layout.addWidget(lang_label)

        header_layout.addStretch()

        copy_btn = QPushButton("Copy")
        copy_btn.setFixedSize(50, 22)
        copy_btn.clicked.connect(self.copy_code)
        copy_btn.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                border: none;
                border-radius: 3px;
                font-size: 11px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
        """)
        header_layout.addWidget(copy_btn)

        layout.addWidget(header)

        # Code content
        self.code_browser = QTextBrowser()
        self.code_browser.setReadOnly(True)
        self.code_browser.setOpenExternalLinks(False)

        # Apply syntax highlighting if available
        if PYGMENTS_AVAILABLE and self.language:
            try:
                lexer = get_lexer_by_name(self.language, stripall=True)
                formatter = HtmlFormatter(style='monokai' if self.theme == 'dark' else 'default')
                highlighted = highlight(self.code, lexer, formatter)
                css = formatter.get_style_defs('.highlight')
                self.code_browser.setHtml(f"<style>{css}</style>{highlighted}")
            except:
                self.code_browser.setPlainText(self.code)
        else:
            self.code_browser.setPlainText(self.code)

        colors = get_theme(self.theme)
        self.code_browser.setStyleSheet(f"""
            QTextBrowser {{
                background-color: {colors['code_bg']};
                color: {colors['text']};
                border: none;
                padding: 8px;
                font-family: Consolas, monospace;
                font-size: 12px;
            }}
        """)

        layout.addWidget(self.code_browser)

        # Frame styling
        colors = get_theme(self.theme)
        header.setStyleSheet(f"""
            QFrame {{
                background-color: {colors['border']};
                border-top-left-radius: 6px;
                border-top-right-radius: 6px;
            }}
        """)
        self.setStyleSheet(f"""
            CodeBlockWidget {{
                border: 1px solid {colors['border']};
                border-radius: 6px;
            }}
        """)

    def copy_code(self):
        """Copy code to clipboard."""
        clipboard = QApplication.clipboard()
        clipboard.setText(self.code)
        QToolTip.showText(self.mapToGlobal(QPoint(0, 0)), "Copied!", self, self.rect(), 1500)


class MessageWidget(QFrame):
    """Widget for displaying a chat message with actions."""

    edit_requested = Signal(str)  # Message ID
    regenerate_requested = Signal(str)  # Message ID
    delete_requested = Signal(str)  # Message ID
    copy_requested = Signal(str)  # Message ID
    pin_toggled = Signal(str)  # Message ID

    def __init__(self, message: dict, parent=None, theme='light'):
        super().__init__(parent)
        self.message = message
        self.theme = theme
        self.is_editing = False
        self.actions_visible = False
        self.setup_ui()
        self.apply_theme()

    def setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 8, 12, 8)
        layout.setSpacing(4)

        # Header row
        header = QHBoxLayout()

        role = self.message.get('role', 'user')
        role_label = QLabel("You" if role == 'user' else "Assistant")
        role_label.setFont(QFont("Segoe UI", 10, QFont.Weight.Bold))
        header.addWidget(role_label)

        # Timestamp
        timestamp = self.message.get('timestamp', '')
        if timestamp:
            try:
                dt = datetime.fromisoformat(timestamp.replace('Z', '+00:00'))
                time_str = dt.strftime('%I:%M %p')
                time_label = QLabel(time_str)
                time_label.setStyleSheet("color: #888; font-size: 11px;")
                header.addWidget(time_label)
            except:
                pass

        # Pin indicator
        if self.message.get('pinned'):
            pin_label = QLabel("pinned")
            pin_label.setStyleSheet("color: #FF9800; font-size: 10px; font-weight: bold;")
            header.addWidget(pin_label)

        header.addStretch()

        # Action buttons (hidden by default)
        self.actions_widget = QWidget()
        actions_layout = QHBoxLayout(self.actions_widget)
        actions_layout.setContentsMargins(0, 0, 0, 0)
        actions_layout.setSpacing(4)

        self.copy_btn = QPushButton("Copy")
        self.copy_btn.setFixedSize(45, 22)
        self.copy_btn.clicked.connect(lambda: self.copy_requested.emit(self.message.get('id', '')))
        actions_layout.addWidget(self.copy_btn)

        if role == 'user':
            self.edit_btn = QPushButton("Edit")
            self.edit_btn.setFixedSize(40, 22)
            self.edit_btn.clicked.connect(lambda: self.edit_requested.emit(self.message.get('id', '')))
            actions_layout.addWidget(self.edit_btn)

        if role == 'assistant':
            self.regen_btn = QPushButton("Regen")
            self.regen_btn.setFixedSize(50, 22)
            self.regen_btn.clicked.connect(lambda: self.regenerate_requested.emit(self.message.get('id', '')))
            actions_layout.addWidget(self.regen_btn)

        self.pin_btn = QPushButton("Unpin" if self.message.get('pinned') else "Pin")
        self.pin_btn.setFixedSize(45, 22)
        self.pin_btn.clicked.connect(lambda: self.pin_toggled.emit(self.message.get('id', '')))
        actions_layout.addWidget(self.pin_btn)

        self.delete_btn = QPushButton("Del")
        self.delete_btn.setFixedSize(35, 22)
        self.delete_btn.clicked.connect(lambda: self.delete_requested.emit(self.message.get('id', '')))
        actions_layout.addWidget(self.delete_btn)

        self.actions_widget.setVisible(False)
        header.addWidget(self.actions_widget)

        layout.addLayout(header)

        # Content
        self.content_label = QLabel()
        self.content_label.setWordWrap(True)
        self.content_label.setTextFormat(Qt.TextFormat.RichText)
        self.content_label.setOpenExternalLinks(True)
        self.content_label.setTextInteractionFlags(
            Qt.TextInteractionFlag.TextSelectableByMouse |
            Qt.TextInteractionFlag.LinksAccessibleByMouse
        )
        self.set_content(self.message.get('content', ''))
        layout.addWidget(self.content_label)

        # Attachments (collapsed by default)
        attachments = self.message.get('attachments', [])
        if attachments:
            self.attachments_frame = QFrame()
            att_layout = QVBoxLayout(self.attachments_frame)
            att_layout.setContentsMargins(8, 4, 8, 4)

            att_header = QPushButton(f"Attachments ({len(attachments)})")
            att_header.setCheckable(True)
            att_header.clicked.connect(self.toggle_attachments)
            att_layout.addWidget(att_header)

            self.att_content = QWidget()
            att_content_layout = QVBoxLayout(self.att_content)
            for att in attachments:
                att_label = QLabel(f"  - {att.get('name', 'Unknown')}")
                att_label.setStyleSheet("color: #666; font-size: 11px;")
                att_content_layout.addWidget(att_label)
            self.att_content.setVisible(False)
            att_layout.addWidget(self.att_content)

            layout.addWidget(self.attachments_frame)

        # Edit indicator
        if self.message.get('edited'):
            edit_label = QLabel("(edited)")
            edit_label.setStyleSheet("color: #888; font-size: 10px; font-style: italic;")
            layout.addWidget(edit_label)

    def set_content(self, content: str):
        """Set message content with basic markdown conversion."""
        # Simple markdown to HTML conversion
        html = content

        # Code blocks
        html = re.sub(r'```(\w+)?\n(.*?)\n```', r'<pre><code>\2</code></pre>', html, flags=re.DOTALL)

        # Inline code
        html = re.sub(r'`([^`]+)`', r'<code style="background:#f0f0f0;padding:2px 4px;border-radius:3px;">\1</code>', html)

        # Bold
        html = re.sub(r'\*\*([^*]+)\*\*', r'<b>\1</b>', html)

        # Italic
        html = re.sub(r'\*([^*]+)\*', r'<i>\1</i>', html)

        # Newlines
        html = html.replace('\n', '<br>')

        self.content_label.setText(html)

    def apply_theme(self):
        colors = get_theme(self.theme)
        role = self.message.get('role', 'user')
        bubble_color = colors['user_bubble'] if role == 'user' else colors['assistant_bubble']

        self.setStyleSheet(f"""
            MessageWidget {{
                background-color: {bubble_color};
                border-radius: 12px;
                margin: 4px 0;
            }}
            QLabel {{
                color: {colors['text']};
            }}
            QPushButton {{
                background-color: transparent;
                color: {colors['text_secondary']};
                border: 1px solid {colors['border']};
                border-radius: 3px;
                font-size: 10px;
            }}
            QPushButton:hover {{
                background-color: {colors['hover']};
            }}
        """)

    def toggle_attachments(self, checked: bool):
        """Toggle attachments visibility."""
        if hasattr(self, 'att_content'):
            self.att_content.setVisible(checked)

    def enterEvent(self, event):
        """Show actions on hover."""
        self.actions_widget.setVisible(True)
        super().enterEvent(event)

    def leaveEvent(self, event):
        """Hide actions when not hovering."""
        self.actions_widget.setVisible(False)
        super().leaveEvent(event)


class SearchResultsWidget(QFrame):
    """Widget for displaying search results."""

    result_selected = Signal(str, str)  # conversation_id, message_id (or empty for conversation)

    def __init__(self, parent=None, theme='light'):
        super().__init__(parent)
        self.theme = theme
        self.setup_ui()
        self.apply_theme()

    def setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(8, 8, 8, 8)

        header = QHBoxLayout()
        self.title = QLabel("Search Results")
        self.title.setFont(QFont("Segoe UI", 10, QFont.Weight.Bold))
        header.addWidget(self.title)

        header.addStretch()

        close_btn = QPushButton("X")
        close_btn.setFixedSize(20, 20)
        close_btn.clicked.connect(self.hide)
        header.addWidget(close_btn)

        layout.addLayout(header)

        self.results_list = QListWidget()
        self.results_list.itemClicked.connect(self.on_result_clicked)
        layout.addWidget(self.results_list)

    def apply_theme(self):
        colors = get_theme(self.theme)
        self.setStyleSheet(f"""
            SearchResultsWidget {{
                background-color: {colors['bg']};
                border: 1px solid {colors['border']};
                border-radius: 8px;
            }}
            QLabel {{
                color: {colors['text']};
            }}
            QListWidget {{
                background-color: transparent;
                border: none;
                color: {colors['text']};
            }}
            QListWidget::item {{
                padding: 8px;
                border-bottom: 1px solid {colors['border']};
            }}
            QListWidget::item:hover {{
                background-color: {colors['hover']};
            }}
        """)

    def set_results(self, results: list):
        """Set search results to display."""
        self.results_list.clear()
        self.title.setText(f"Search Results ({len(results)})")

        for result in results:
            text = result.get('match_text', '')
            conv_name = result.get('conversation_name', '')
            result_type = result.get('type', 'message')

            if result_type == 'conversation':
                display = f"[Conversation] {conv_name}"
            else:
                role = result.get('role', 'user')
                display = f"[{conv_name}] ({role}): {text[:100]}..."

            item = QListWidgetItem(display)
            item.setData(Qt.ItemDataRole.UserRole, {
                'conversation_id': result.get('conversation_id'),
                'message_id': result.get('message_id', '')
            })
            self.results_list.addItem(item)

        self.show()

    def on_result_clicked(self, item: QListWidgetItem):
        """Handle result click."""
        data = item.data(Qt.ItemDataRole.UserRole)
        if data:
            self.result_selected.emit(
                data.get('conversation_id', ''),
                data.get('message_id', '')
            )
