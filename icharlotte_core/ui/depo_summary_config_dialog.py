"""Modal dialog the user fills out after phase 1 of the deposition agent."""

from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QCheckBox, QComboBox, QDialog, QDialogButtonBox, QHBoxLayout, QLabel,
    QLineEdit, QListWidget, QListWidgetItem, QMessageBox, QPlainTextEdit,
    QSpinBox, QVBoxLayout, QWidget,
)

from icharlotte_core.deposition import session_manager


class _TopicRow(QWidget):
    def __init__(self, title: str, parent=None):
        super().__init__(parent)
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        self.checkbox = QCheckBox()
        self.checkbox.setChecked(True)
        self.title_edit = QLineEdit(title)
        layout.addWidget(self.checkbox)
        layout.addWidget(self.title_edit, 1)


class DepoSummaryConfigDialog(QDialog):
    def __init__(self, session_path, parent=None):
        super().__init__(parent)
        self.session_path = Path(session_path)
        self._session = session_manager.read_session(self.session_path)

        self.setWindowTitle("Configure Deposition Summary")
        self.setModal(True)
        self.setWindowModality(Qt.ApplicationModal)
        self.resize(700, 600)

        root = QVBoxLayout(self)

        header_text = (
            f"Configure summary for <b>{self._session.get('deponent_name', '')}</b> "
            f"({self._session.get('deponent_type', '')}, "
            f"{self._session.get('deposition_date', 'date unknown')})"
        )
        root.addWidget(QLabel(header_text))

        # Topics list with drag-reorder
        root.addWidget(QLabel("Topics (drag to reorder, uncheck to omit, edit text to rename):"))
        self.topics_list = QListWidget()
        self.topics_list.setDragDropMode(QListWidget.InternalMove)
        self.topics_list.setSelectionMode(QListWidget.SingleSelection)
        self.topics_list.setDefaultDropAction(Qt.MoveAction)
        for t in self._session.get("topics", []):
            row = _TopicRow(t.get("title", ""))
            item = QListWidgetItem()
            item.setSizeHint(row.sizeHint())
            self.topics_list.addItem(item)
            self.topics_list.setItemWidget(item, row)
        root.addWidget(self.topics_list, 1)

        # Summary bias row
        bias_row = QHBoxLayout()
        bias_row.addWidget(QLabel("Summary bias:"))
        self.bias_combo = QComboBox()
        for label, value in (
            ("Neutral", "neutral"),
            ("Most favorable to plaintiff", "pro_plaintiff"),
            ("Most favorable to defense", "pro_defense"),
            ("Custom…", "custom"),
        ):
            self.bias_combo.addItem(label, value)
        bias_row.addWidget(self.bias_combo)
        self.bias_custom_edit = QLineEdit()
        self.bias_custom_edit.setPlaceholderText(
            "Describe the editorial lens (e.g., 'Highlight any inconsistencies in injury testimony')"
        )
        self.bias_custom_edit.setVisible(False)
        bias_row.addWidget(self.bias_custom_edit, 1)
        root.addLayout(bias_row)

        self.bias_combo.currentIndexChanged.connect(self._on_bias_combo_changed)

        # Additional topics
        root.addWidget(QLabel("Additional topics (one per line):"))
        self.added_topics_edit = QPlainTextEdit()
        self.added_topics_edit.setPlaceholderText(
            "One topic per line. These are appended after the checked topics above, in order."
        )
        self.added_topics_edit.setFixedHeight(70)
        root.addWidget(self.added_topics_edit)

        # Settings row
        settings_row = QHBoxLayout()
        settings_row.addWidget(QLabel("Bullets per topic:"))
        self.bullets_spinbox = QSpinBox()
        self.bullets_spinbox.setRange(1, 15)
        self.bullets_spinbox.setValue(5)
        settings_row.addWidget(self.bullets_spinbox)

        settings_row.addSpacing(20)
        settings_row.addWidget(QLabel("Deponent label:"))
        self.deponent_label_edit = QLineEdit(self._session.get("deponent_type", ""))
        settings_row.addWidget(self.deponent_label_edit, 1)

        settings_row.addSpacing(20)
        self.cross_check_checkbox = QCheckBox("Run cross-check pass")
        self.cross_check_checkbox.setChecked(True)
        settings_row.addWidget(self.cross_check_checkbox)
        root.addLayout(settings_row)

        # Custom rules
        root.addWidget(QLabel("Custom rules:"))
        self.custom_rules_edit = QPlainTextEdit()
        self.custom_rules_edit.setPlaceholderText(
            "Any extra instructions for the summary (tense, citation style, things to avoid, etc.)."
        )
        self.custom_rules_edit.setFixedHeight(90)
        root.addWidget(self.custom_rules_edit)

        # Buttons
        buttons = QDialogButtonBox()
        buttons.addButton("Cancel", QDialogButtonBox.RejectRole)
        generate_btn = buttons.addButton("Generate Summary", QDialogButtonBox.AcceptRole)
        generate_btn.setDefault(True)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        root.addWidget(buttons)

    def topic_rows_in_order(self):
        """Yield each topic _TopicRow widget in current visual order."""
        for i in range(self.topics_list.count()):
            item = self.topics_list.item(i)
            yield self.topics_list.itemWidget(item)

    def _on_bias_combo_changed(self, _index):
        is_custom = self.bias_combo.currentData() == "custom"
        self.bias_custom_edit.setVisible(is_custom)
        if not is_custom:
            self.bias_custom_edit.clear()

    def accept(self):
        selected_topics = [
            row.title_edit.text().strip()
            for row in self.topic_rows_in_order()
            if row.checkbox.isChecked() and row.title_edit.text().strip()
        ]
        added_topics = [
            line.strip()
            for line in self.added_topics_edit.toPlainText().splitlines()
            if line.strip()
        ]
        if not selected_topics and not added_topics:
            QMessageBox.warning(
                self,
                "No topics selected",
                "Select at least one topic, or add a custom topic, before generating the summary.",
            )
            return
        cfg = {
            "selected_topics": selected_topics,
            "added_topics": added_topics,
            "bullets_per_topic": self.bullets_spinbox.value(),
            "deponent_label": self.deponent_label_edit.text().strip() or "Deponent",
            "custom_rules": self.custom_rules_edit.toPlainText().strip(),
            "cross_check_enabled": self.cross_check_checkbox.isChecked(),
            "bias": self.bias_combo.currentData() or "neutral",
            "bias_custom": (self.bias_custom_edit.text().strip()
                            if self.bias_combo.currentData() == "custom" else ""),
        }
        session_manager.update_user_config(self.session_path, cfg)
        super().accept()
