"""SettingsPage — pre-run configuration for a task tab.

This is a placeholder for the per-task settings UI; real per-task
settings are defined in follow-up specs. For now it shows:
  - The list of selected input files (with a Remove button per row).
  - A 'Settings for <task title> — to be defined' label.
  - A Proceed button bottom-right.
"""
import os
from typing import List

from PySide6.QtCore import Signal
from PySide6.QtWidgets import (
    QHBoxLayout,
    QLabel,
    QListWidget,
    QListWidgetItem,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from ..registry import TaskSpec


class SettingsPage(QWidget):
    """Configurable inputs + Proceed button. Emits proceed_requested(settings_dict)."""

    proceed_requested = Signal(dict)  # settings dict (placeholder)

    def __init__(self, spec: TaskSpec, files: List[str], parent: QWidget | None = None):
        super().__init__(parent)
        self._spec = spec
        self._files: List[str] = list(files)

        outer = QVBoxLayout(self)
        outer.setContentsMargins(24, 24, 24, 24)
        outer.setSpacing(16)

        # Files section
        files_label = QLabel(self._format_files_label())
        files_label.setStyleSheet("font-size: 13px; font-weight: 600;")
        outer.addWidget(files_label)
        self.files_label = files_label

        self.files_list = QListWidget()
        self.files_list.setMaximumHeight(150)
        self._refresh_files_list()
        outer.addWidget(self.files_list)

        # Placeholder body
        body = QLabel(f"Settings for {spec.title} — to be defined.")
        body.setStyleSheet("color: #666; font-style: italic; padding: 24px;")
        body.setAlignment(body.alignment())
        outer.addWidget(body, 1)

        # Proceed button bottom-right
        btn_row = QHBoxLayout()
        btn_row.addStretch()
        self.proceed_btn = QPushButton("Proceed")
        self.proceed_btn.setStyleSheet(
            "background-color: #1976D2; color: white; font-weight: 600; padding: 8px 24px; border-radius: 4px;"
        )
        self.proceed_btn.clicked.connect(self._on_proceed)
        btn_row.addWidget(self.proceed_btn)
        outer.addLayout(btn_row)

        self._update_proceed_enabled()

    def _format_files_label(self) -> str:
        return f"Files ({len(self._files)})"

    def _refresh_files_list(self) -> None:
        self.files_list.clear()
        for path in self._files:
            display = os.path.basename(path)
            item = QListWidgetItem(display)
            item.setToolTip(path)
            if not os.path.exists(path):
                item.setText(f"{display}  (missing)")
                item.setForeground(item.foreground())  # placeholder; greyed via stylesheet if desired
            self.files_list.addItem(item)
        self.files_label.setText(self._format_files_label())
        self._update_proceed_enabled()

    def _update_proceed_enabled(self) -> None:
        self.proceed_btn.setEnabled(len(self._files) > 0)

    def _on_proceed(self) -> None:
        self.proceed_requested.emit(self.to_dict())

    # ---- Persistence-friendly API ----

    @property
    def files(self) -> list[str]:
        return list(self._files)

    def to_dict(self) -> dict:
        """Placeholder settings dict. Real per-task settings will override."""
        return {}

    def from_dict(self, data: dict) -> None:
        """Placeholder — real subclasses will restore form state."""
        return None
