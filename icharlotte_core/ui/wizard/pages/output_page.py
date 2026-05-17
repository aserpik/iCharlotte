"""OutputPage scaffold. Phase 8 replaces the placeholder body with the mammoth
.docx → HTML editor + Save/Open in Word actions.
"""
import os

from PySide6.QtCore import Signal
from PySide6.QtWidgets import (
    QHBoxLayout,
    QLabel,
    QPushButton,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)


class OutputPage(QWidget):
    """Shows the task's output and action buttons.

    Phase 4: minimal scaffold — header with file name + Open in Word + a plain
    text view of the file. Full mammoth-rendered editor + Save round-trip arrive
    in Phase 8.
    """

    rerun_requested = Signal()
    edit_settings_requested = Signal()
    open_in_word_requested = Signal()
    copy_all_requested = Signal()
    save_requested = Signal()

    def __init__(self, parent: QWidget | None = None):
        super().__init__(parent)
        self._output_path: str | None = None

        outer = QVBoxLayout(self)
        outer.setContentsMargins(24, 24, 24, 24)
        outer.setSpacing(12)

        header = QHBoxLayout()
        self.file_label = QLabel("File: —")
        self.file_label.setStyleSheet("font-weight: 600;")
        header.addWidget(self.file_label, 1)
        self.open_in_word_btn = QPushButton("Open in Word")
        self.open_in_word_btn.clicked.connect(self.open_in_word_requested.emit)
        header.addWidget(self.open_in_word_btn)
        outer.addLayout(header)

        self.editor = QTextEdit()
        self.editor.setReadOnly(False)
        outer.addWidget(self.editor, 1)

        btn_row = QHBoxLayout()
        btn_row.addStretch()
        self.copy_all_btn = QPushButton("Copy All")
        self.copy_all_btn.clicked.connect(self.copy_all_requested.emit)
        btn_row.addWidget(self.copy_all_btn)
        self.rerun_btn = QPushButton("Re-run")
        self.rerun_btn.clicked.connect(self.rerun_requested.emit)
        btn_row.addWidget(self.rerun_btn)
        self.edit_settings_btn = QPushButton("Edit Settings & Re-run")
        self.edit_settings_btn.clicked.connect(self.edit_settings_requested.emit)
        btn_row.addWidget(self.edit_settings_btn)
        self.save_btn = QPushButton("Save")
        self.save_btn.setStyleSheet("background-color: #1976D2; color: white; font-weight: 600; padding: 6px 18px;")
        self.save_btn.clicked.connect(self.save_requested.emit)
        btn_row.addWidget(self.save_btn)
        outer.addLayout(btn_row)

    # ---- Public API ----

    @property
    def output_path(self) -> str | None:
        return self._output_path

    def load_output(self, output_path: str) -> None:
        """Phase 4 stub: shows the file name and a placeholder body."""
        self._output_path = output_path
        self.file_label.setText(f"File: {os.path.basename(output_path)}")
        self.editor.setPlainText(
            f"(Phase 4 scaffold) Output file at:\n{output_path}\n\n"
            "Full mammoth-rendered editor arrives in Phase 8."
        )
