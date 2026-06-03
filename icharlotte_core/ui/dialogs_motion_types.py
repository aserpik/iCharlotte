"""Workbench tab for managing Generate Motion motion types."""

from __future__ import annotations

from PySide6.QtWidgets import (
    QDialog,
    QDialogButtonBox,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from icharlotte_core.motion_generation.config import (
    MotionTypeConfig,
    reload_motion_types,
)
from icharlotte_core.motion_generation.types_registry import MotionTypeRegistry


def _lines(text: str) -> list[str]:
    return [ln.strip() for ln in text.splitlines() if ln.strip()]


class MotionTypesTab(QWidget):
    """Editor for the Generate Motion motion-type registry."""

    def __init__(self, *, registry_path: str, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.registry_path = registry_path
        self.registry = MotionTypeRegistry.load(registry_path)

        layout = QVBoxLayout(self)
        layout.addWidget(QLabel(
            "Motion types used by the Generate Motion task. Edits apply to new "
            "generations after saving."
        ))

        self.table = QTableWidget(0, 4, self)
        self.table.setHorizontalHeaderLabels(
            ["Type ID", "Display Name", "Sections", "Placeholders"]
        )
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        layout.addWidget(self.table, 1)

        button_row = QHBoxLayout()
        self.add_btn = QPushButton("Add Type")
        self.add_btn.clicked.connect(self._on_add_clicked)
        button_row.addWidget(self.add_btn)
        self.edit_btn = QPushButton("Edit Selected")
        self.edit_btn.clicked.connect(self._on_edit_clicked)
        button_row.addWidget(self.edit_btn)
        self.remove_btn = QPushButton("Remove Selected")
        self.remove_btn.clicked.connect(self._on_remove_clicked)
        button_row.addWidget(self.remove_btn)
        self.restore_btn = QPushButton("Restore Defaults")
        self.restore_btn.clicked.connect(self._on_restore_clicked)
        button_row.addWidget(self.restore_btn)
        button_row.addStretch()
        self.save_btn = QPushButton("Save")
        self.save_btn.clicked.connect(self.save)
        button_row.addWidget(self.save_btn)
        layout.addLayout(button_row)

        self._refresh_table()

    # ---- Programmatic API (used by tests + handlers) ------------------- #

    def add_type_programmatic(self, config: MotionTypeConfig) -> None:
        self.registry.add(config)
        self._refresh_table()

    def update_type_programmatic(self, type_id: str, **fields) -> bool:
        ok = self.registry.update(type_id, **fields)
        self._refresh_table()
        return ok

    def remove_type_programmatic(self, type_id: str) -> bool:
        ok = self.registry.remove(type_id)
        self._refresh_table()
        return ok

    def restore_defaults_programmatic(self) -> None:
        self.registry.restore_defaults()
        self._refresh_table()

    def save(self) -> None:
        self.registry.save()
        # Refresh the running app's get_motion_config singleton.
        reload_motion_types()

    # ---- Interactive handlers ------------------------------------------ #

    def _selected_type_id(self) -> str | None:
        row = self.table.currentRow()
        types = self.registry.list_types()
        if row < 0 or row >= len(types):
            return None
        return types[row].type_id

    def _on_add_clicked(self) -> None:
        dlg = _MotionTypeEditDialog(parent=self)
        if dlg.exec() == QDialog.DialogCode.Accepted:
            cfg = dlg.result_config()
            if not cfg.type_id:
                QMessageBox.warning(self, "Type ID required", "Enter a unique type id.")
                return
            if any(t.type_id == cfg.type_id for t in self.registry.list_types()):
                QMessageBox.warning(self, "Duplicate", f"Type id '{cfg.type_id}' already exists.")
                return
            self.add_type_programmatic(cfg)

    def _on_edit_clicked(self) -> None:
        type_id = self._selected_type_id()
        if type_id is None:
            return
        existing = self.registry.get(type_id)
        dlg = _MotionTypeEditDialog(config=existing, parent=self)
        if dlg.exec() == QDialog.DialogCode.Accepted:
            cfg = dlg.result_config()
            self.registry.add(cfg)  # same id → replace
            self._refresh_table()

    def _on_remove_clicked(self) -> None:
        type_id = self._selected_type_id()
        if type_id is None:
            return
        confirm = QMessageBox.question(
            self,
            "Remove type",
            f"Remove motion type '{type_id}'?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if confirm == QMessageBox.StandardButton.Yes:
            self.remove_type_programmatic(type_id)

    def _on_restore_clicked(self) -> None:
        confirm = QMessageBox.question(
            self,
            "Restore defaults",
            "Replace all motion types with the built-in defaults? This discards "
            "your custom types and edits.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if confirm == QMessageBox.StandardButton.Yes:
            self.restore_defaults_programmatic()

    # ---- Helpers -------------------------------------------------------- #

    def _refresh_table(self) -> None:
        types = self.registry.list_types()
        self.table.setRowCount(len(types))
        for i, cfg in enumerate(types):
            self.table.setItem(i, 0, QTableWidgetItem(cfg.type_id))
            self.table.setItem(i, 1, QTableWidgetItem(cfg.display_name))
            self.table.setItem(i, 2, QTableWidgetItem(str(len(cfg.section_plan))))
            self.table.setItem(i, 3, QTableWidgetItem(str(len(cfg.placeholder_attachments))))


class _MotionTypeEditDialog(QDialog):
    def __init__(self, *, config: MotionTypeConfig | None = None, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self._editing = config is not None
        self.setWindowTitle("Edit Motion Type" if self._editing else "Add Motion Type")
        self.resize(640, 640)
        layout = QVBoxLayout(self)

        layout.addWidget(QLabel("Type ID (lowercase, no spaces; cannot change after creation)"))
        self.type_id_edit = QLineEdit(config.type_id if config else "")
        self.type_id_edit.setReadOnly(self._editing)
        layout.addWidget(self.type_id_edit)

        layout.addWidget(QLabel("Display name"))
        self.display_name_edit = QLineEdit(config.display_name if config else "")
        layout.addWidget(self.display_name_edit)

        layout.addWidget(QLabel("Document guidance (shown on the intake page)"))
        self.guidance_edit = QPlainTextEdit(config.target_doc_guidance if config else "")
        self.guidance_edit.setMaximumHeight(60)
        layout.addWidget(self.guidance_edit)

        layout.addWidget(QLabel("Legal standard (grounds the Legal Standard section)"))
        self.legal_edit = QPlainTextEdit(config.legal_standard_hint if config else "")
        self.legal_edit.setMaximumHeight(90)
        layout.addWidget(self.legal_edit)

        layout.addWidget(QLabel("Section plan (one heading per line)"))
        self.sections_edit = QPlainTextEdit("\n".join(config.section_plan) if config else "")
        self.sections_edit.setMaximumHeight(90)
        layout.addWidget(self.sections_edit)

        layout.addWidget(QLabel("Placeholder attachments (one per line)"))
        self.attachments_edit = QPlainTextEdit("\n".join(config.placeholder_attachments) if config else "")
        self.attachments_edit.setMaximumHeight(70)
        layout.addWidget(self.attachments_edit)

        advanced = QGroupBox("Advanced prompts (optional)")
        adv_layout = QVBoxLayout(advanced)
        adv_layout.addWidget(QLabel("Analyzer prompt"))
        self.analyzer_edit = QPlainTextEdit(config.analyzer_prompt if config else "")
        self.analyzer_edit.setMaximumHeight(60)
        adv_layout.addWidget(self.analyzer_edit)
        adv_layout.addWidget(QLabel("Grounds prompt"))
        self.grounds_edit = QPlainTextEdit(config.grounds_prompt if config else "")
        self.grounds_edit.setMaximumHeight(60)
        adv_layout.addWidget(self.grounds_edit)
        layout.addWidget(advanced)

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def result_config(self) -> MotionTypeConfig:
        return MotionTypeConfig(
            type_id=self.type_id_edit.text().strip().lower().replace(" ", "_"),
            display_name=self.display_name_edit.text().strip(),
            target_doc_guidance=self.guidance_edit.toPlainText().strip(),
            legal_standard_hint=self.legal_edit.toPlainText().strip(),
            section_plan=_lines(self.sections_edit.toPlainText()),
            placeholder_attachments=_lines(self.attachments_edit.toPlainText()),
            analyzer_prompt=self.analyzer_edit.toPlainText().strip(),
            grounds_prompt=self.grounds_edit.toPlainText().strip(),
        )
