"""Dialog for Med Record Extractor — user pastes prose entries for extraction."""

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel,
    QPlainTextEdit, QPushButton, QMessageBox,
)


class MedRecordExtractorDialog(QDialog):
    """Modal dialog with a text area for pasting medical chronology entries."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Med Record Extractor")
        self.setMinimumSize(600, 400)
        self.setWindowFlags(self.windowFlags() | Qt.WindowType.WindowStaysOnTopHint)
        self._result_text = None
        self._setup_ui()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)

        label = QLabel(
            "Paste medical chronology entries below. Each entry should include "
            "a date of treatment and provider/facility name.\n\n"
            "Example: \"On March 22, 2016, Plaintiff was evaluated by "
            "Vikram M Shaker, MD at Adventist Health for a CT of the abdomen...\""
        )
        label.setWordWrap(True)
        layout.addWidget(label)

        self.text_input = QPlainTextEdit()
        self.text_input.setPlaceholderText(
            "Paste entries here (one or more paragraphs)..."
        )
        self.text_input.setMinimumHeight(250)
        layout.addWidget(self.text_input, stretch=1)

        btn_layout = QHBoxLayout()
        btn_layout.addStretch()

        cancel_btn = QPushButton("Cancel")
        cancel_btn.clicked.connect(self.reject)
        btn_layout.addWidget(cancel_btn)

        extract_btn = QPushButton("Extract")
        extract_btn.setDefault(True)
        extract_btn.clicked.connect(self._on_extract)
        btn_layout.addWidget(extract_btn)

        layout.addLayout(btn_layout)

    def _on_extract(self):
        text = self.text_input.toPlainText().strip()
        if not text:
            QMessageBox.warning(self, "Empty", "Please paste at least one entry.")
            return
        self._result_text = text
        self.accept()

    def get_text(self) -> str:
        return self._result_text or ""
