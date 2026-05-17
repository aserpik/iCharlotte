"""StatusPage — progress bar + log + Cancel button while a task is running."""
from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QHBoxLayout,
    QLabel,
    QPlainTextEdit,
    QProgressBar,
    QPushButton,
    QVBoxLayout,
    QWidget,
)


class StatusPage(QWidget):
    """Shows progress + log lines. Emits cancel_requested when Cancel is clicked."""

    cancel_requested = Signal()

    def __init__(self, parent: QWidget | None = None):
        super().__init__(parent)
        outer = QVBoxLayout(self)
        outer.setContentsMargins(24, 24, 24, 24)
        outer.setSpacing(12)

        self.status_label = QLabel("Starting…")
        self.status_label.setStyleSheet("font-size: 13px; font-weight: 600;")
        outer.addWidget(self.status_label)

        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 0)  # indeterminate by default
        outer.addWidget(self.progress_bar)

        self.log_view = QPlainTextEdit()
        self.log_view.setReadOnly(True)
        self.log_view.setStyleSheet("font-family: Consolas, 'Courier New', monospace; font-size: 11px;")
        outer.addWidget(self.log_view, 1)

        btn_row = QHBoxLayout()
        btn_row.addStretch()
        self.cancel_btn = QPushButton("Cancel")
        self.cancel_btn.setStyleSheet(
            "background-color: #d32f2f; color: white; font-weight: 600; padding: 8px 20px; border-radius: 4px;"
        )
        self.cancel_btn.clicked.connect(self._on_cancel)
        btn_row.addWidget(self.cancel_btn)
        outer.addLayout(btn_row)

    def reset(self) -> None:
        self.status_label.setText("Starting…")
        self.progress_bar.setRange(0, 0)
        self.progress_bar.setValue(0)
        self.log_view.clear()
        self.cancel_btn.setEnabled(True)
        self.cancel_btn.setText("Cancel")

    # ---- Slots / public API for the worker connection ----

    def on_status(self, line: str) -> None:
        self.status_label.setText(line)
        self.log_view.appendPlainText(line)
        self.log_view.verticalScrollBar().setValue(self.log_view.verticalScrollBar().maximum())

    def on_progress(self, pct: int) -> None:
        if pct < 0 or pct > 100:
            return
        if self.progress_bar.maximum() == 0:
            self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(pct)

    def _on_cancel(self) -> None:
        self.cancel_btn.setEnabled(False)
        self.cancel_btn.setText("Cancelling…")
        self.cancel_requested.emit()
