"""Workbench tab for managing oppose_motion style exemplars."""

from __future__ import annotations

import os
import uuid
from datetime import date

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QCheckBox,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QHBoxLayout,
    QHeaderView,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from icharlotte_core.opposition.style_examples import (
    StyleExample,
    StyleExampleRegistry,
)


_COMMON_MOTION_TAGS = [
    "msj",
    "msa",
    "summary judgment",
    "demurrer",
    "motion to compel",
    "motion to compel further",
    "anti-slapp",
    "motion in limine",
    "motion for reconsideration",
    "motion to set aside",
    "motion to continue",
]


class StyleExamplesTab(QWidget):
    """Editor for oppose_motion style exemplars."""

    def __init__(self, *, registry_path: str, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.registry_path = registry_path
        self.registry = StyleExampleRegistry.load(registry_path)

        layout = QVBoxLayout(self)

        self.table = QTableWidget(0, 4, self)
        self.table.setHorizontalHeaderLabels(["Label", "Path", "Motion Types", "Active"])
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        layout.addWidget(self.table, 1)

        button_row = QHBoxLayout()
        self.add_btn = QPushButton("Add Example")
        self.add_btn.clicked.connect(self._on_add_clicked)
        button_row.addWidget(self.add_btn)
        self.remove_btn = QPushButton("Remove Selected")
        self.remove_btn.clicked.connect(self._on_remove_clicked)
        button_row.addWidget(self.remove_btn)
        button_row.addStretch()
        self.save_btn = QPushButton("Save")
        self.save_btn.clicked.connect(self.save)
        button_row.addWidget(self.save_btn)
        layout.addLayout(button_row)

        self._refresh_table()

    # ---- Programmatic API used by tests --------------------------------

    def add_example_programmatic(
        self,
        *,
        label: str,
        path: str,
        motion_types: list[str],
        active: bool = True,
    ) -> str:
        example_id = uuid.uuid4().hex[:8]
        self.registry.add(StyleExample(
            id=example_id,
            label=label,
            path=path,
            motion_types=[t.strip().lower() for t in motion_types if t.strip()],
            active=active,
            added_at=date.today().isoformat(),
        ))
        self._refresh_table()
        return example_id

    def remove_example_programmatic(self, example_id: str) -> bool:
        ok = self.registry.remove(example_id)
        self._refresh_table()
        return ok

    def save(self) -> None:
        self.registry.save()

    # ---- Interactive handlers ------------------------------------------

    def _on_add_clicked(self) -> None:
        dlg = _ExampleEditDialog(parent=self)
        if dlg.exec() == QDialog.DialogCode.Accepted:
            label, path, motion_types, active = dlg.result_fields()
            if path:
                self.add_example_programmatic(
                    label=label or os.path.basename(path),
                    path=path,
                    motion_types=motion_types,
                    active=active,
                )

    def _on_remove_clicked(self) -> None:
        row = self.table.currentRow()
        if row < 0 or row >= len(self.registry.examples):
            return
        example_id = self.registry.examples[row].id
        confirm = QMessageBox.question(
            self,
            "Remove example",
            f"Remove '{self.registry.examples[row].label}'?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if confirm == QMessageBox.StandardButton.Yes:
            self.remove_example_programmatic(example_id)

    # ---- Helpers --------------------------------------------------------

    def _refresh_table(self) -> None:
        self.table.setRowCount(len(self.registry.examples))
        for i, ex in enumerate(self.registry.examples):
            self.table.setItem(i, 0, QTableWidgetItem(ex.label))
            self.table.setItem(i, 1, QTableWidgetItem(ex.path))
            self.table.setItem(i, 2, QTableWidgetItem(", ".join(ex.motion_types)))
            checkbox = QCheckBox()
            checkbox.setChecked(ex.active)
            checkbox.toggled.connect(lambda checked, eid=ex.id: self._on_active_toggled(eid, checked))
            self.table.setCellWidget(i, 3, checkbox)

    def _on_active_toggled(self, example_id: str, checked: bool) -> None:
        self.registry.update(example_id, active=checked)
        self.registry.save()


class _ExampleEditDialog(QDialog):
    def __init__(self, *, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("Add Style Example")
        self.resize(560, 220)
        layout = QVBoxLayout(self)

        self.label_edit = QLineEdit()
        self.label_edit.setPlaceholderText("Short label (e.g., MTC Opp - Discovery Sanctions)")
        layout.addWidget(self.label_edit)

        path_row = QHBoxLayout()
        self.path_edit = QLineEdit()
        self.path_edit.setPlaceholderText("Path to .docx file")
        path_row.addWidget(self.path_edit, 1)
        browse_btn = QPushButton("Browse...")
        browse_btn.clicked.connect(self._on_browse)
        path_row.addWidget(browse_btn)
        layout.addLayout(path_row)

        self.tags_edit = QLineEdit()
        self.tags_edit.setPlaceholderText(
            "Motion-type tags, comma-separated (e.g., motion to compel, discovery). "
            "Leave empty for universal."
        )
        layout.addWidget(self.tags_edit)

        suggested = QLineEdit(", ".join(_COMMON_MOTION_TAGS))
        suggested.setReadOnly(True)
        suggested.setStyleSheet("color: #5f6368;")
        layout.addWidget(suggested)

        self.active_check = QCheckBox("Active")
        self.active_check.setChecked(True)
        layout.addWidget(self.active_check)

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def _on_browse(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Select exemplar .docx",
            "",
            "Word Documents (*.docx)",
        )
        if path:
            self.path_edit.setText(path)

    def result_fields(self) -> tuple[str, str, list[str], bool]:
        tags = [t.strip().lower() for t in self.tags_edit.text().split(",") if t.strip()]
        return (
            self.label_edit.text().strip(),
            self.path_edit.text().strip(),
            tags,
            self.active_check.isChecked(),
        )
