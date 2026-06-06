"""SettingsPage — pre-run configuration for a task tab.

This is a placeholder for the per-task settings UI; real per-task
settings are defined in follow-up specs. For now it shows:
  - The list of selected input files with Add Files... / Remove buttons.
  - A 'Settings for <task title> — to be defined' label.
  - A Proceed button bottom-right.
"""
import os
from typing import List

from PySide6.QtCore import Signal
from PySide6.QtWidgets import (
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QListWidget,
    QListWidgetItem,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from .. import theme
from ..file_picker import resolve_default_folder
from ..registry import TaskSpec


class SettingsPage(QWidget):
    """Configurable inputs + Proceed button. Emits proceed_requested(settings_dict)."""

    proceed_requested = Signal(dict)  # settings dict (placeholder)

    def __init__(
        self,
        spec: TaskSpec,
        files: List[str],
        case_root: str | None = None,
        parent: QWidget | None = None,
    ):
        super().__init__(parent)
        self._spec = spec
        self._files: List[str] = list(files)
        self._case_root: str | None = case_root

        outer = QVBoxLayout(self)
        outer.setContentsMargins(
            theme.SPACE_XL, theme.SPACE_XL, theme.SPACE_XL, theme.SPACE_XL
        )
        outer.setSpacing(theme.SPACE_LG)

        # One-line instruction so the page is self-explanatory.
        self.instruction_label = theme.helper_text(
            "Choose the documents to include, then click Continue."
        )
        outer.addWidget(self.instruction_label)

        # Files section
        files_label = theme.section_header(self._format_files_label())
        outer.addWidget(files_label)
        self.files_label = files_label

        # Add Files... / Remove button row
        file_btn_row = QHBoxLayout()
        file_btn_row.setSpacing(theme.SPACE_SM)
        self.add_files_btn = theme.secondary_button("Add Files…")
        self.add_files_btn.clicked.connect(self._on_add_files)
        file_btn_row.addWidget(self.add_files_btn)
        self.remove_btn = theme.secondary_button("Remove")
        self.remove_btn.setEnabled(False)
        self.remove_btn.clicked.connect(self._on_remove_files)
        file_btn_row.addWidget(self.remove_btn)
        file_btn_row.addStretch()
        outer.addLayout(file_btn_row)

        self.files_list = QListWidget()
        self.files_list.setMaximumHeight(150)
        self.files_list.setSelectionMode(QListWidget.SelectionMode.ExtendedSelection)
        self.files_list.itemSelectionChanged.connect(self._on_selection_changed)
        outer.addWidget(self.files_list)

        # Empty-state hint shown under the list when no files are selected.
        self.empty_hint = theme.caption("No files yet — click “Add Files…” to get started.")
        self.empty_hint.setStyleSheet(
            f"font-size: {theme.FONT_CAPTION}px; color: {theme.TEXT_FAINT}; font-style: italic;"
        )
        outer.addWidget(self.empty_hint)

        self._refresh_files_list()

        # Placeholder body
        body = QLabel(f"Settings for {spec.title} — to be defined.")
        body.setStyleSheet(
            f"color: {theme.TEXT_MUTED}; font-style: italic; padding: {theme.SPACE_XL}px;"
        )
        outer.addWidget(body, 1)

        # Continue button bottom-right
        btn_row = QHBoxLayout()
        btn_row.addStretch()
        self.proceed_btn = theme.primary_button("Continue")
        self.proceed_btn.clicked.connect(self._on_proceed)
        btn_row.addWidget(self.proceed_btn)
        outer.addLayout(btn_row)

        self._update_proceed_enabled()

    def _format_files_label(self) -> str:
        return f"Files ({len(self._files)})"

    def _refresh_files_list(self) -> None:
        from PySide6.QtGui import QColor

        self.files_list.clear()
        for path in self._files:
            display = os.path.basename(path)
            item = QListWidgetItem(display)
            item.setToolTip(path)
            if not os.path.exists(path):
                item.setText(f"{display}  (missing)")
                item.setForeground(QColor(theme.ERROR))
            self.files_list.addItem(item)
        self.files_label.setText(self._format_files_label())
        if hasattr(self, "empty_hint"):
            self.empty_hint.setVisible(len(self._files) == 0)
        self._update_proceed_enabled()

    def _update_proceed_enabled(self) -> None:
        if not hasattr(self, "proceed_btn"):
            return
        self.proceed_btn.setEnabled(len(self._files) > 0)

    def _on_selection_changed(self) -> None:
        self.remove_btn.setEnabled(len(self.files_list.selectedItems()) > 0)

    def _on_add_files(self) -> None:
        start_dir = self._file_dialog_start_dir()
        paths, _ = QFileDialog.getOpenFileNames(
            self, "Add files", start_dir, "All files (*.*)"
        )
        existing = set(self._files)
        for p in paths:
            if p not in existing:
                self._files.append(p)
                existing.add(p)
        self._refresh_files_list()

    def _file_dialog_start_dir(self) -> str:
        if not self._case_root:
            return ""
        return resolve_default_folder(self._case_root, self._spec.default_folders)

    def _on_remove_files(self) -> None:
        selected_rows = {idx.row() for idx in self.files_list.selectedIndexes()}
        self._files = [p for i, p in enumerate(self._files) if i not in selected_rows]
        self._refresh_files_list()

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

    def attach_worker(self, worker) -> bool:
        """Override to take control of worker signals (e.g., for speculative runs).

        Return True if the settings page is handling the worker; in that case
        TaskTab will skip wiring status/progress/awaiting_input to the status page.
        """
        return False
