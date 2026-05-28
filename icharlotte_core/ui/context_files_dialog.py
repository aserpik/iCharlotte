"""Reusable modal dialog for picking context files across multiple folders.

``QFileDialog.getOpenFileNames`` only allows selecting multiple files within a
single folder. This dialog lets the user accumulate files from different
folders/subfolders: each "Add files…" click opens a normal file browser and
appends the chosen files to a running list (de-duped by normalized path).
"""
from __future__ import annotations

import os
from typing import List, Optional

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QHBoxLayout,
    QListWidget,
    QListWidgetItem,
    QPushButton,
    QVBoxLayout,
    QWidget,
)


class ContextFilesDialog(QDialog):
    """Accumulate-style file picker that works across multiple folders."""

    def __init__(
        self,
        parent: QWidget | None = None,
        *,
        title: str = "Select context files",
        start_dir: str = "",
        file_filter: str = "All files (*.*)",
        initial: Optional[List[str]] = None,
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle(title)
        self.resize(560, 360)
        self._file_filter = file_filter
        self._next_dir = start_dir or ""
        self._paths: List[str] = []
        self._seen: set[str] = set()

        outer = QVBoxLayout(self)
        outer.setContentsMargins(16, 16, 16, 16)
        outer.setSpacing(10)

        btn_row = QHBoxLayout()
        self.add_btn = QPushButton("Add files…")
        self.add_btn.clicked.connect(self._on_add_files)
        btn_row.addWidget(self.add_btn)
        self.remove_btn = QPushButton("Remove selected")
        self.remove_btn.clicked.connect(self._on_remove_selected)
        btn_row.addWidget(self.remove_btn)
        btn_row.addStretch()
        outer.addLayout(btn_row)

        self.list_widget = QListWidget()
        self.list_widget.setSelectionMode(
            QListWidget.SelectionMode.ExtendedSelection
        )
        outer.addWidget(self.list_widget, 1)

        button_box = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok
            | QDialogButtonBox.StandardButton.Cancel
        )
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        outer.addWidget(button_box)

        if initial:
            self._add_paths(list(initial))

    @staticmethod
    def _key(path: str) -> str:
        return os.path.normcase(os.path.abspath(path))

    @staticmethod
    def _folder_hint(path: str) -> str:
        parent = os.path.dirname(path)
        return os.path.basename(parent) or parent

    def _add_paths(self, paths: List[str]) -> None:
        for path in paths:
            if not path:
                continue
            key = self._key(path)
            if key in self._seen:
                continue
            self._seen.add(key)
            self._paths.append(path)
            display = os.path.basename(path)
            hint = self._folder_hint(path)
            if hint:
                display = f"{display}  —  {hint}"
            item = QListWidgetItem(display)
            item.setToolTip(path)
            item.setData(Qt.ItemDataRole.UserRole, path)
            self.list_widget.addItem(item)
            self._next_dir = os.path.dirname(path) or self._next_dir

    def _on_add_files(self) -> None:
        paths, _ = QFileDialog.getOpenFileNames(
            self, "Add files", self._next_dir, self._file_filter
        )
        if paths:
            self._add_paths(list(paths))

    def _on_remove_selected(self) -> None:
        rows = sorted(
            {i.row() for i in self.list_widget.selectedIndexes()}, reverse=True
        )
        for r in rows:
            if 0 <= r < len(self._paths):
                removed = self._paths.pop(r)
                self._seen.discard(self._key(removed))
                self.list_widget.takeItem(r)

    def selected_files(self) -> List[str]:
        return list(self._paths)

    @classmethod
    def get_files(
        cls,
        parent: QWidget | None = None,
        *,
        title: str = "Select context files",
        start_dir: str = "",
        file_filter: str = "All files (*.*)",
        initial: Optional[List[str]] = None,
    ) -> Optional[List[str]]:
        """Show modally. Return accumulated list on OK, ``None`` on Cancel."""
        dlg = cls(
            parent,
            title=title,
            start_dir=start_dir,
            file_filter=file_filter,
            initial=initial,
        )
        if dlg.exec() == QDialog.DialogCode.Accepted:
            return dlg.selected_files()
        return None
