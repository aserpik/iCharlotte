"""Task-scoped summary browser for Wizard launcher card actions."""
from __future__ import annotations

import os

from PySide6.QtCore import QEvent, Qt, Signal
from PySide6.QtWidgets import (
    QHBoxLayout,
    QLabel,
    QListWidget,
    QListWidgetItem,
    QMessageBox,
    QPushButton,
    QSplitter,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from icharlotte_core.word_validator import validate_summary_docx

from . import theme
from .summary_outputs import (
    SummaryOutputEntry,
    discover_summary_outputs,
    summary_browser_title,
)
from .docx_io import load_docx_as_html, save_qtextdocument_as_docx


class SummaryBrowserTab(QWidget):
    """Browse prior summaries for one Wizard task."""

    open_requested = Signal(str)

    def __init__(
        self,
        case_path: str,
        file_number: str,
        task_id: str,
        parent: QWidget | None = None,
    ):
        super().__init__(parent)
        self.case_path = case_path
        self.file_number = file_number
        self.task_id = task_id
        self.outputs: list[SummaryOutputEntry] = []
        self._current_path = ""
        self._dirty = False
        self._suppress_preview_signal = False
        self._suppress_selection_signal = False
        self.setProperty("summary_browser_task_id", task_id)
        self.setStyleSheet(theme.wizard_stylesheet())
        self._build_ui()
        self.refresh()

    def _build_ui(self) -> None:
        outer = QVBoxLayout(self)
        outer.setContentsMargins(
            theme.SPACE_XL, theme.SPACE_XL, theme.SPACE_XL, theme.SPACE_XL
        )
        outer.setSpacing(theme.SPACE_MD)

        title = theme.page_title(summary_browser_title(self.task_id))
        outer.addWidget(title)

        self.count_label = QLabel("")
        self.count_label.setStyleSheet(f"color: {theme.TEXT_MUTED};")
        outer.addWidget(self.count_label)

        splitter = QSplitter(Qt.Orientation.Horizontal)

        left = QWidget()
        left_layout = QVBoxLayout(left)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(theme.SPACE_SM)
        left_layout.addWidget(theme.section_header("Summaries"))
        self.output_list = QListWidget()
        self.output_list.installEventFilter(self)
        self.output_list.itemSelectionChanged.connect(self._on_selection_changed)
        left_layout.addWidget(self.output_list, 1)
        splitter.addWidget(left)

        right = QWidget()
        right_layout = QVBoxLayout(right)
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_layout.setSpacing(theme.SPACE_SM)
        self.preview_label = theme.section_header("Preview")
        right_layout.addWidget(self.preview_label)
        self.preview = QTextEdit()
        self.preview.setReadOnly(False)
        self.preview.textChanged.connect(self._on_preview_text_changed)
        right_layout.addWidget(self.preview, 1)
        splitter.addWidget(right)
        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 2)

        outer.addWidget(splitter, 1)

        btn_row = QHBoxLayout()
        btn_row.addStretch()
        self.discard_btn = theme.secondary_button("Discard Changes")
        self.discard_btn.clicked.connect(self._on_discard_changes)
        btn_row.addWidget(self.discard_btn)
        self.save_btn = theme.primary_button("Save")
        self.save_btn.clicked.connect(self._on_save)
        btn_row.addWidget(self.save_btn)
        self.open_btn = theme.secondary_button("Open")
        self.open_btn.clicked.connect(self._on_open)
        btn_row.addWidget(self.open_btn)
        self.external_btn = theme.secondary_button("Open in External Editor")
        self.external_btn.clicked.connect(self._on_open_external)
        btn_row.addWidget(self.external_btn)
        outer.addLayout(btn_row)

    def refresh(self) -> None:
        self.outputs = discover_summary_outputs(self.case_path, self.task_id)
        self.output_list.clear()
        for entry in self.outputs:
            item = QListWidgetItem(f"{entry.name}\n{entry.source}")
            item.setData(Qt.ItemDataRole.UserRole, entry.path)
            item.setToolTip(entry.path)
            self.output_list.addItem(item)

        count = len(self.outputs)
        suffix = "" if count == 1 else "s"
        self.count_label.setText(f"{count} summarized document{suffix} found")

        if count:
            self.output_list.setCurrentRow(0)
        else:
            self._current_path = ""
            self.preview_label.setText("Preview")
            self._set_preview_text("No summarized documents were found for this task.")
            self._refresh_buttons()

    def _on_selection_changed(self) -> None:
        if self._suppress_selection_signal:
            return
        items = self.output_list.selectedItems()
        if not items:
            self._current_path = ""
            self.preview_label.setText("Preview")
            self._set_preview_text("")
            self._refresh_buttons()
            return
        path = items[0].data(Qt.ItemDataRole.UserRole)
        if path != self._current_path and self._dirty:
            if not self._confirm_discard_changes():
                self._select_path(self._current_path)
                return
        self._current_path = path or ""
        self.preview_label.setText(os.path.basename(self._current_path) or "Preview")
        self._render_preview(self._current_path)
        self._refresh_buttons()

    def _render_preview(self, path: str) -> None:
        if not path or not os.path.isfile(path):
            self._set_preview_text("The selected summary file was not found.")
            return
        if path.lower().endswith(".docx"):
            try:
                self._set_preview_html(load_docx_as_html(path))
            except Exception as exc:
                self._set_preview_text(f"Could not preview this .docx file:\n{exc}")
            return
        try:
            with open(path, "r", encoding="utf-8", errors="replace") as fh:
                self._set_preview_text(fh.read())
        except OSError as exc:
            self._set_preview_text(f"Could not preview this file:\n{exc}")

    def _set_preview_html(self, html: str) -> None:
        self._suppress_preview_signal = True
        try:
            self.preview.setHtml(html)
        finally:
            self._suppress_preview_signal = False
        self._dirty = False

    def _set_preview_text(self, text: str) -> None:
        self._suppress_preview_signal = True
        try:
            self.preview.setPlainText(text)
        finally:
            self._suppress_preview_signal = False
        self._dirty = False

    def _refresh_buttons(self) -> None:
        enabled = bool(self._current_path) and os.path.isfile(self._current_path)
        self.open_btn.setEnabled(enabled)
        self.external_btn.setEnabled(enabled)
        self.save_btn.setEnabled(enabled and self._dirty)
        self.discard_btn.setEnabled(enabled and self._dirty)

    def _on_preview_text_changed(self) -> None:
        if self._suppress_preview_signal:
            return
        if not self._current_path:
            return
        self._dirty = True
        self._refresh_buttons()

    def _confirm_discard_changes(self) -> bool:
        answer = QMessageBox.question(
            self,
            "Discard changes?",
            "You have unsaved edits. Discard them and continue?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        return answer == QMessageBox.StandardButton.Yes

    def _select_path(self, path: str) -> None:
        self._suppress_selection_signal = True
        try:
            for row in range(self.output_list.count()):
                item = self.output_list.item(row)
                if item.data(Qt.ItemDataRole.UserRole) == path:
                    self.output_list.setCurrentRow(row)
                    return
        finally:
            self._suppress_selection_signal = False

    def _on_save(self) -> None:
        if not self._current_path:
            return
        saved_path = self._current_path
        try:
            save_qtextdocument_as_docx(self.preview.document(), saved_path)
            validation = validate_summary_docx(saved_path)
            if validation.has_errors:
                validation.print_summary()
                QMessageBox.warning(
                    self,
                    "Validation warning",
                    "Saved, but validation found an issue. Check the console output.",
                )
            else:
                QMessageBox.information(self, "Saved", f"Saved:\n{saved_path}")
            self._dirty = False
            self._refresh_buttons()
            self.refresh()
            self._select_path(saved_path)
            if os.path.isfile(saved_path):
                self._current_path = saved_path
                self.preview_label.setText(os.path.basename(saved_path))
                self._render_preview(saved_path)
                self._refresh_buttons()
        except Exception as exc:
            QMessageBox.critical(self, "Save failed", f"Could not save summary:\n{exc}")

    def _on_discard_changes(self) -> None:
        if self._current_path:
            self._render_preview(self._current_path)
            self._refresh_buttons()

    def _delete_selected_summary(self) -> None:
        items = self.output_list.selectedItems()
        if not items:
            return
        path = items[0].data(Qt.ItemDataRole.UserRole)
        if not path or not os.path.isfile(path):
            return
        detail = f"Delete this summary from disk?\n\n{path}"
        if self._dirty and path == self._current_path:
            detail += "\n\nUnsaved edits will be lost."
        answer = QMessageBox.question(
            self,
            "Delete summary?",
            detail,
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if answer != QMessageBox.StandardButton.Yes:
            return
        try:
            os.remove(path)
        except OSError as exc:
            QMessageBox.critical(self, "Delete failed", f"Could not delete summary:\n{exc}")
            return
        self._dirty = False
        self._current_path = ""
        self.refresh()

    def eventFilter(self, obj, event) -> bool:
        if (
            obj is self.output_list
            and event.type() == QEvent.Type.KeyPress
            and event.key() == Qt.Key.Key_Delete
        ):
            self._delete_selected_summary()
            return True
        return super().eventFilter(obj, event)

    def _on_open(self) -> None:
        if self._current_path:
            self.open_requested.emit(self._current_path)

    def _on_open_external(self) -> None:
        if not self._current_path:
            return
        try:
            os.startfile(self._current_path)
        except Exception as exc:
            QMessageBox.critical(
                self,
                "Open failed",
                f"Could not open the summary in an external editor:\n{exc}",
            )
