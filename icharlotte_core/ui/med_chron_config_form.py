"""Configuration form for the multi-analysis Med-Cron task.

Reads the Phase 1 session JSON, presents:
- a checkbox per curated analysis (Rewrite Chronology pre-checked)
- a "Custom analyses" panel with add/remove rows
- a "narrative missing" warning banner when applicable

On commit_user_config(), validates selection and writes user_config back
to the session, flipping phase to ``ready_to_run``.
"""

from pathlib import Path
from typing import Optional

from PySide6.QtWidgets import (
    QCheckBox, QHBoxLayout, QLabel, QLineEdit, QPlainTextEdit,
    QPushButton, QScrollArea, QVBoxLayout, QWidget,
)

from icharlotte_core.med_chron import session_manager


class CustomAnalysisRow(QWidget):
    """One row in the custom-analyses list: label + instruction + remove btn."""

    def __init__(self, parent: QWidget, on_remove):
        super().__init__(parent)
        self._on_remove = on_remove

        layout = QVBoxLayout(self)
        layout.setContentsMargins(4, 4, 4, 4)
        layout.setSpacing(4)

        top = QHBoxLayout()
        top.addWidget(QLabel("Label:"))
        self.label_edit = QLineEdit()
        self.label_edit.setPlaceholderText("Short name (e.g. 'Left-knee mentions')")
        top.addWidget(self.label_edit, 1)
        self.remove_btn = QPushButton("−")
        self.remove_btn.setFixedSize(24, 24)
        self.remove_btn.setStyleSheet("QPushButton { color: #c62828; font-weight: bold; }")
        self.remove_btn.clicked.connect(self._handle_remove)
        top.addWidget(self.remove_btn)
        layout.addLayout(top)

        layout.addWidget(QLabel("Request:"))
        self.instruction_edit = QPlainTextEdit()
        self.instruction_edit.setPlaceholderText("Describe the analysis…")
        self.instruction_edit.setFixedHeight(60)
        layout.addWidget(self.instruction_edit)

        self.setStyleSheet(
            "CustomAnalysisRow { border: 1px solid #ddd; border-radius: 4px; }"
        )

    def _handle_remove(self):
        self._on_remove(self)

    def label(self) -> str:
        return self.label_edit.text().strip()

    def instruction(self) -> str:
        return self.instruction_edit.toPlainText().strip()

    def is_empty(self) -> bool:
        return not self.label() and not self.instruction()


class MedChronConfigForm(QWidget):
    """Pickable list of catalog analyses + custom analysis rows."""

    def __init__(self, session_path, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.session_path = Path(session_path)
        self._session = session_manager.read_session(self.session_path)

        root = QVBoxLayout(self)
        root.setContentsMargins(8, 8, 8, 8)
        root.setSpacing(8)

        header = QLabel(
            f"<b>Analyses to run on {self._session.get('provider_name', '')}</b>"
        )
        root.addWidget(header)

        # Narrative-missing warning banner
        self.narrative_missing_banner = QLabel(
            "⚠ Narrative text not found in this document — "
            "Rewrite Chronology will be skipped."
        )
        self.narrative_missing_banner.setStyleSheet(
            "background-color: #FFF3CD; color: #856404; "
            "padding: 6px; border-radius: 4px;"
        )
        self.narrative_missing_banner.setVisible(
            bool(self._session.get("narrative_missing"))
        )
        root.addWidget(self.narrative_missing_banner)

        # Curated catalog section
        cat_label = QLabel("<b>Curated analyses:</b>")
        root.addWidget(cat_label)

        self.catalog_checkboxes: dict[str, QCheckBox] = {}
        for entry in self._session.get("catalog", []):
            row = QWidget()
            row_layout = QVBoxLayout(row)
            row_layout.setContentsMargins(16, 2, 0, 2)
            row_layout.setSpacing(0)
            cb = QCheckBox(entry.get("title", entry.get("id", "")))
            cb.setChecked(bool(entry.get("default_selected")))
            self.catalog_checkboxes[entry["id"]] = cb
            row_layout.addWidget(cb)
            desc = QLabel(entry.get("description", ""))
            desc.setStyleSheet("color: #666; font-size: 11px; padding-left: 22px;")
            desc.setWordWrap(True)
            row_layout.addWidget(desc)
            root.addWidget(row)

        # Custom analyses section
        custom_header = QHBoxLayout()
        custom_header.addWidget(QLabel("<b>Custom analyses:</b>"))
        custom_header.addStretch()
        root.addLayout(custom_header)

        # Scrollable container for custom rows.
        self._custom_container = QWidget()
        self._custom_container_layout = QVBoxLayout(self._custom_container)
        self._custom_container_layout.setContentsMargins(0, 0, 0, 0)
        self._custom_container_layout.setSpacing(4)
        self._custom_container_layout.addStretch()  # pin rows to top

        scroll = QScrollArea()
        scroll.setWidget(self._custom_container)
        scroll.setWidgetResizable(True)
        scroll.setMinimumHeight(80)
        scroll.setMaximumHeight(220)
        root.addWidget(scroll)

        self.custom_rows: list[CustomAnalysisRow] = []

        add_btn = QPushButton("+ Add custom analysis")
        add_btn.clicked.connect(self.add_custom_row)
        root.addWidget(add_btn)

        # Inline validation error label.
        self._error_label = QLabel("")
        self._error_label.setStyleSheet("color: #c62828; font-style: italic;")
        self._error_label.setVisible(False)
        root.addWidget(self._error_label)

    def add_custom_row(self) -> "CustomAnalysisRow":
        row = CustomAnalysisRow(self._custom_container, self._remove_custom_row)
        # Insert above the stretch at the end.
        idx = self._custom_container_layout.count() - 1
        self._custom_container_layout.insertWidget(idx, row)
        self.custom_rows.append(row)
        return row

    def _remove_custom_row(self, row: "CustomAnalysisRow") -> None:
        if row in self.custom_rows:
            self.custom_rows.remove(row)
        self._custom_container_layout.removeWidget(row)
        row.deleteLater()

    def _selected_catalog_ids(self) -> list[str]:
        return [cid for cid, cb in self.catalog_checkboxes.items() if cb.isChecked()]

    def _validated_custom_rows(self) -> tuple[list[dict], str]:
        """Return (clean_rows, error_msg). Empty rows are silently dropped.

        Partially-filled rows (one of label/instruction missing) are an error.
        """
        clean = []
        for r in self.custom_rows:
            if r.is_empty():
                continue
            lbl, instr = r.label(), r.instruction()
            if not lbl or not instr:
                return [], (
                    "Custom analyses need both a label and an instruction. "
                    "Fill in (or remove) the partially-completed row."
                )
            clean.append({"label": lbl, "instruction": instr})
        return clean, ""

    def commit_user_config(self) -> bool:
        """Validate and write user_config. Returns True on success."""
        self._error_label.setVisible(False)

        selected = self._selected_catalog_ids()
        clean_custom, err = self._validated_custom_rows()
        if err:
            self._error_label.setText(err)
            self._error_label.setVisible(True)
            return False
        if not selected and not clean_custom:
            self._error_label.setText(
                "Select at least one analysis, or add a custom analysis."
            )
            self._error_label.setVisible(True)
            return False

        session_manager.update_user_config(
            self.session_path,
            {
                "selected_catalog_ids": selected,
                "custom_analyses": clean_custom,
            },
        )
        return True
