"""
Diff Viewer Dialog for iCharlotte

Displays side-by-side or unified diffs when consolidating discovery responses
or other documents. Allows user to accept or reject changes.
"""

import difflib
from typing import Optional, Tuple, List
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QTextEdit, QSplitter, QFrame, QRadioButton, QButtonGroup,
    QDialogButtonBox, QWidget, QScrollArea
)
from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QFont, QColor, QTextCharFormat, QTextCursor, QSyntaxHighlighter


class DiffHighlighter(QSyntaxHighlighter):
    """Syntax highlighter for diff output."""

    def __init__(self, parent=None):
        super().__init__(parent)

        # Define formats
        self.addition_format = QTextCharFormat()
        self.addition_format.setBackground(QColor(200, 255, 200))  # Light green
        self.addition_format.setForeground(QColor(0, 100, 0))       # Dark green

        self.deletion_format = QTextCharFormat()
        self.deletion_format.setBackground(QColor(255, 200, 200))  # Light red
        self.deletion_format.setForeground(QColor(139, 0, 0))       # Dark red

        self.header_format = QTextCharFormat()
        self.header_format.setForeground(QColor(0, 0, 139))        # Dark blue
        self.header_format.setFontWeight(QFont.Weight.Bold)

        self.context_format = QTextCharFormat()
        self.context_format.setForeground(QColor(100, 100, 100))   # Gray

    def highlightBlock(self, text: str):
        """Apply highlighting to a block of text."""
        if not text:
            return

        if text.startswith('+') and not text.startswith('+++'):
            self.setFormat(0, len(text), self.addition_format)
        elif text.startswith('-') and not text.startswith('---'):
            self.setFormat(0, len(text), self.deletion_format)
        elif text.startswith('@@'):
            self.setFormat(0, len(text), self.header_format)
        elif text.startswith('---') or text.startswith('+++'):
            self.setFormat(0, len(text), self.header_format)


class DiffViewer(QDialog):
    """
    Dialog for viewing and accepting/rejecting diffs.

    Shows differences between old and new versions of text content,
    with options for side-by-side or unified view.
    """

    # Signals
    accepted_changes = Signal()  # User accepted the new version
    rejected_changes = Signal()  # User rejected the changes

    def __init__(
        self,
        old_text: str,
        new_text: str,
        title: str = "Review Changes",
        old_label: str = "Previous Version",
        new_label: str = "New Version",
        parent=None
    ):
        super().__init__(parent)
        self.old_text = old_text
        self.new_text = new_text
        self.result_accepted = False

        self.setWindowTitle(title)
        self.setMinimumSize(1000, 700)
        self.resize(1200, 800)

        self._setup_ui(old_label, new_label)
        self._compute_diff()

    def _setup_ui(self, old_label: str, new_label: str):
        """Set up the dialog UI."""
        layout = QVBoxLayout(self)

        # View mode toggle
        mode_layout = QHBoxLayout()
        mode_layout.addWidget(QLabel("View Mode:"))

        self.mode_group = QButtonGroup(self)

        self.side_by_side_btn = QRadioButton("Side by Side")
        self.side_by_side_btn.setChecked(True)
        self.mode_group.addButton(self.side_by_side_btn, 0)
        mode_layout.addWidget(self.side_by_side_btn)

        self.unified_btn = QRadioButton("Unified")
        self.mode_group.addButton(self.unified_btn, 1)
        mode_layout.addWidget(self.unified_btn)

        mode_layout.addStretch()

        # Statistics
        self.stats_label = QLabel()
        self.stats_label.setStyleSheet("color: gray; font-size: 11px;")
        mode_layout.addWidget(self.stats_label)

        layout.addLayout(mode_layout)

        # Connect mode change
        self.mode_group.buttonClicked.connect(self._on_mode_changed)

        # Stacked views
        self.views_container = QWidget()
        views_layout = QVBoxLayout(self.views_container)
        views_layout.setContentsMargins(0, 0, 0, 0)

        # Side-by-side view
        self.side_by_side_widget = QSplitter(Qt.Orientation.Horizontal)

        # Left panel (old)
        left_container = QFrame()
        left_layout = QVBoxLayout(left_container)
        left_layout.setContentsMargins(0, 0, 0, 0)

        left_header = QLabel(f"<b>{old_label}</b>")
        left_header.setStyleSheet("background-color: #fff0f0; padding: 5px;")
        left_layout.addWidget(left_header)

        self.left_text = QTextEdit()
        self.left_text.setReadOnly(True)
        self.left_text.setFont(QFont("Consolas", 10))
        self.left_text.setStyleSheet("background-color: #fffafa;")
        left_layout.addWidget(self.left_text)

        self.side_by_side_widget.addWidget(left_container)

        # Right panel (new)
        right_container = QFrame()
        right_layout = QVBoxLayout(right_container)
        right_layout.setContentsMargins(0, 0, 0, 0)

        right_header = QLabel(f"<b>{new_label}</b>")
        right_header.setStyleSheet("background-color: #f0fff0; padding: 5px;")
        right_layout.addWidget(right_header)

        self.right_text = QTextEdit()
        self.right_text.setReadOnly(True)
        self.right_text.setFont(QFont("Consolas", 10))
        self.right_text.setStyleSheet("background-color: #f0fff0;")
        right_layout.addWidget(self.right_text)

        self.side_by_side_widget.addWidget(right_container)

        views_layout.addWidget(self.side_by_side_widget)

        # Unified view
        self.unified_widget = QFrame()
        unified_layout = QVBoxLayout(self.unified_widget)
        unified_layout.setContentsMargins(0, 0, 0, 0)

        unified_header = QLabel("<b>Unified Diff</b>")
        unified_header.setStyleSheet("background-color: #f0f0ff; padding: 5px;")
        unified_layout.addWidget(unified_header)

        self.unified_text = QTextEdit()
        self.unified_text.setReadOnly(True)
        self.unified_text.setFont(QFont("Consolas", 10))
        unified_layout.addWidget(self.unified_text)

        self.unified_highlighter = DiffHighlighter(self.unified_text.document())

        views_layout.addWidget(self.unified_widget)
        self.unified_widget.hide()

        layout.addWidget(self.views_container)

        # Button box
        button_layout = QHBoxLayout()

        self.accept_btn = QPushButton("Accept Changes")
        self.accept_btn.setStyleSheet(
            "background-color: #4CAF50; color: white; font-weight: bold; "
            "padding: 8px 20px; font-size: 12px;"
        )
        self.accept_btn.clicked.connect(self._on_accept)
        button_layout.addWidget(self.accept_btn)

        self.reject_btn = QPushButton("Keep Original")
        self.reject_btn.setStyleSheet(
            "background-color: #f44336; color: white; font-weight: bold; "
            "padding: 8px 20px; font-size: 12px;"
        )
        self.reject_btn.clicked.connect(self._on_reject)
        button_layout.addWidget(self.reject_btn)

        button_layout.addStretch()

        self.cancel_btn = QPushButton("Cancel")
        self.cancel_btn.clicked.connect(self.reject)
        button_layout.addWidget(self.cancel_btn)

        layout.addLayout(button_layout)

    def _compute_diff(self):
        """Compute and display the diff."""
        old_lines = self.old_text.splitlines(keepends=True)
        new_lines = self.new_text.splitlines(keepends=True)

        # Compute statistics
        differ = difflib.Differ()
        diff_result = list(differ.compare(old_lines, new_lines))

        additions = sum(1 for line in diff_result if line.startswith('+ '))
        deletions = sum(1 for line in diff_result if line.startswith('- '))
        unchanged = sum(1 for line in diff_result if line.startswith('  '))

        self.stats_label.setText(
            f"<span style='color: green;'>+{additions}</span> | "
            f"<span style='color: red;'>-{deletions}</span> | "
            f"<span style='color: gray;'>{unchanged} unchanged</span>"
        )

        # Side-by-side view
        self._populate_side_by_side(old_lines, new_lines)

        # Unified view
        unified = difflib.unified_diff(
            old_lines, new_lines,
            fromfile='Previous',
            tofile='New',
            lineterm=''
        )
        self.unified_text.setPlainText('\n'.join(unified))

    def _populate_side_by_side(self, old_lines: List[str], new_lines: List[str]):
        """Populate the side-by-side view with highlighted differences."""
        # Use SequenceMatcher for alignment
        matcher = difflib.SequenceMatcher(None, old_lines, new_lines)

        left_content = []
        right_content = []

        for tag, i1, i2, j1, j2 in matcher.get_opcodes():
            if tag == 'equal':
                for line in old_lines[i1:i2]:
                    left_content.append(('equal', line.rstrip('\n\r')))
                    right_content.append(('equal', line.rstrip('\n\r')))
            elif tag == 'delete':
                for line in old_lines[i1:i2]:
                    left_content.append(('delete', line.rstrip('\n\r')))
                    right_content.append(('empty', ''))
            elif tag == 'insert':
                for line in new_lines[j1:j2]:
                    left_content.append(('empty', ''))
                    right_content.append(('insert', line.rstrip('\n\r')))
            elif tag == 'replace':
                old_chunk = old_lines[i1:i2]
                new_chunk = new_lines[j1:j2]

                max_len = max(len(old_chunk), len(new_chunk))
                for k in range(max_len):
                    if k < len(old_chunk):
                        left_content.append(('delete', old_chunk[k].rstrip('\n\r')))
                    else:
                        left_content.append(('empty', ''))

                    if k < len(new_chunk):
                        right_content.append(('insert', new_chunk[k].rstrip('\n\r')))
                    else:
                        right_content.append(('empty', ''))

        # Build HTML
        left_html = self._build_html(left_content, 'left')
        right_html = self._build_html(right_content, 'right')

        self.left_text.setHtml(left_html)
        self.right_text.setHtml(right_html)

    def _build_html(self, content: List[Tuple[str, str]], side: str) -> str:
        """Build HTML for side-by-side view."""
        lines = []
        for tag, text in content:
            escaped_text = (
                text.replace('&', '&amp;')
                    .replace('<', '&lt;')
                    .replace('>', '&gt;')
                    .replace(' ', '&nbsp;')
            )

            if tag == 'equal':
                lines.append(f'<div style="font-family: Consolas;">{escaped_text}&nbsp;</div>')
            elif tag == 'delete':
                lines.append(
                    f'<div style="background-color: #ffcccc; font-family: Consolas;">'
                    f'{escaped_text}&nbsp;</div>'
                )
            elif tag == 'insert':
                lines.append(
                    f'<div style="background-color: #ccffcc; font-family: Consolas;">'
                    f'{escaped_text}&nbsp;</div>'
                )
            elif tag == 'empty':
                lines.append(
                    '<div style="background-color: #f0f0f0; font-family: Consolas;">&nbsp;</div>'
                )

        return ''.join(lines)

    def _on_mode_changed(self, button):
        """Handle view mode change."""
        if button == self.side_by_side_btn:
            self.side_by_side_widget.show()
            self.unified_widget.hide()
        else:
            self.side_by_side_widget.hide()
            self.unified_widget.show()

    def _on_accept(self):
        """User accepted the changes."""
        self.result_accepted = True
        self.accepted_changes.emit()
        self.accept()

    def _on_reject(self):
        """User rejected the changes."""
        self.result_accepted = False
        self.rejected_changes.emit()
        self.reject()

    def get_result(self) -> Tuple[bool, str]:
        """
        Get the result of the dialog.

        Returns:
            Tuple of (accepted, text) where accepted is True if user accepted
            changes and text is the selected version.
        """
        if self.result_accepted:
            return (True, self.new_text)
        return (False, self.old_text)


def show_diff_dialog(
    old_text: str,
    new_text: str,
    title: str = "Review Changes",
    old_label: str = "Previous Version",
    new_label: str = "New Version",
    parent=None
) -> Tuple[bool, str]:
    """
    Show a diff dialog and return the result.

    Args:
        old_text: The original text.
        new_text: The new text with changes.
        title: Dialog window title.
        old_label: Label for the old version panel.
        new_label: Label for the new version panel.
        parent: Parent widget.

    Returns:
        Tuple of (accepted, selected_text).
    """
    dialog = DiffViewer(
        old_text=old_text,
        new_text=new_text,
        title=title,
        old_label=old_label,
        new_label=new_label,
        parent=parent
    )

    dialog.exec()
    return dialog.get_result()


def compute_diff_stats(old_text: str, new_text: str) -> dict:
    """
    Compute statistics about differences between two texts.

    Args:
        old_text: The original text.
        new_text: The new text.

    Returns:
        Dictionary with diff statistics.
    """
    old_lines = old_text.splitlines()
    new_lines = new_text.splitlines()

    differ = difflib.Differ()
    diff_result = list(differ.compare(old_lines, new_lines))

    additions = sum(1 for line in diff_result if line.startswith('+ '))
    deletions = sum(1 for line in diff_result if line.startswith('- '))
    changes = sum(1 for line in diff_result if line.startswith('? '))
    unchanged = sum(1 for line in diff_result if line.startswith('  '))

    # Calculate similarity ratio
    matcher = difflib.SequenceMatcher(None, old_text, new_text)
    similarity = matcher.ratio()

    return {
        'additions': additions,
        'deletions': deletions,
        'changes': changes,
        'unchanged': unchanged,
        'total_old_lines': len(old_lines),
        'total_new_lines': len(new_lines),
        'similarity_ratio': similarity,
        'significant_change': similarity < 0.9  # Less than 90% similar
    }


def has_significant_changes(old_text: str, new_text: str, threshold: float = 0.9) -> bool:
    """
    Check if the new text has significant changes from the old text.

    Args:
        old_text: The original text.
        new_text: The new text.
        threshold: Similarity ratio below which changes are significant.

    Returns:
        True if changes are significant (should show diff dialog).
    """
    if old_text == new_text:
        return False

    matcher = difflib.SequenceMatcher(None, old_text, new_text)
    return matcher.ratio() < threshold
