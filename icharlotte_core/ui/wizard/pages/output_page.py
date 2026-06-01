"""OutputPage — mammoth-rendered editor + action buttons.

Supports both single-output (one docx) and multi-output (N docxs from a
parallel run) modes. When >1 output is loaded, a picker combo box appears
above the editor so the user can switch between produced files. Save and
Open in Word operate on the currently-selected file.
"""
import os
import shutil
from typing import List

from PySide6.QtCore import Signal
from PySide6.QtGui import QGuiApplication
from PySide6.QtWidgets import (
    QComboBox,
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QPushButton,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from .. import theme
from ..docx_io import load_docx_as_html, save_qtextdocument_as_docx


class OutputPage(QWidget):
    rerun_requested = Signal()
    edit_settings_requested = Signal()

    def __init__(self, parent: QWidget | None = None):
        super().__init__(parent)
        self._output_path: str | None = None
        self._output_paths: List[str] = []
        self._dirty = False
        self._suppress_picker_signal = False
        self._save_as_required = False
        self._save_default_dir = ""
        self._save_suggested_filename = ""

        outer = QVBoxLayout(self)
        outer.setContentsMargins(
            theme.SPACE_XL, theme.SPACE_XL, theme.SPACE_XL, theme.SPACE_XL
        )
        outer.setSpacing(theme.SPACE_MD)

        # Picker row (hidden when only one output)
        picker_row = QHBoxLayout()
        picker_row.setSpacing(theme.SPACE_SM)
        self.picker_label = QLabel("Output:")
        self.picker_label.setStyleSheet(f"font-weight: 600; color: {theme.TEXT};")
        picker_row.addWidget(self.picker_label)
        self.output_picker = QComboBox()
        self.output_picker.setMinimumWidth(360)
        self.output_picker.currentIndexChanged.connect(self._on_picker_changed)
        picker_row.addWidget(self.output_picker, 1)
        self.outputs_count_label = QLabel("")
        self.outputs_count_label.setStyleSheet(f"color: {theme.TEXT_MUTED};")
        picker_row.addWidget(self.outputs_count_label)
        self.picker_row_widget = QWidget()
        self.picker_row_widget.setLayout(picker_row)
        self.picker_row_widget.setVisible(False)
        outer.addWidget(self.picker_row_widget)

        # Inline progress hint shown while sibling files are still processing.
        self.progress_hint_label = QLabel("")
        self.progress_hint_label.setStyleSheet(
            f"color: {theme.PRIMARY}; font-style: italic; padding: 2px 4px;"
        )
        self.progress_hint_label.setVisible(False)
        outer.addWidget(self.progress_hint_label)

        # Failure banner shown when a parallel run finished with some files
        # produced and some failed. Lives on the Output page because by the
        # time `failed` fires, the page has already switched away from Status.
        self.failure_banner_label = QLabel("")
        self.failure_banner_label.setWordWrap(True)
        self.failure_banner_label.setStyleSheet(
            f"background-color: {theme.WARNING_BG}; color: {theme.ERROR_HOVER};"
            f" border: 1px solid {theme.WARNING}; border-radius: {theme.RADIUS_SM}px;"
            f" padding: 8px 12px; font-weight: 600;"
        )
        self.failure_banner_label.setVisible(False)
        outer.addWidget(self.failure_banner_label)

        header = QHBoxLayout()
        self.file_label = QLabel("File: —")
        self.file_label.setStyleSheet(f"font-weight: 600; color: {theme.TEXT};")
        header.addWidget(self.file_label, 1)
        self.open_in_word_btn = theme.secondary_button("Open in Word")
        self.open_in_word_btn.clicked.connect(self._on_open_in_word)
        header.addWidget(self.open_in_word_btn)
        outer.addLayout(header)

        self.editor = QTextEdit()
        self.editor.setReadOnly(False)
        self.editor.textChanged.connect(self._on_text_changed)
        outer.addWidget(self.editor, 1)

        btn_row = QHBoxLayout()
        btn_row.addStretch()
        self.copy_all_btn = theme.secondary_button("Copy All")
        self.copy_all_btn.clicked.connect(self._on_copy_all)
        btn_row.addWidget(self.copy_all_btn)
        self.rerun_btn = theme.secondary_button("Re-run")
        self.rerun_btn.clicked.connect(self.rerun_requested.emit)
        btn_row.addWidget(self.rerun_btn)
        self.edit_settings_btn = theme.secondary_button("Edit Settings & Re-run")
        self.edit_settings_btn.clicked.connect(self.edit_settings_requested.emit)
        btn_row.addWidget(self.edit_settings_btn)
        self.save_btn = theme.primary_button("Save")
        self.save_btn.clicked.connect(self._on_save)
        btn_row.addWidget(self.save_btn)
        outer.addLayout(btn_row)

        self._refresh_save_enabled()

    # ---- Public API ----

    @property
    def output_path(self) -> str | None:
        return self._output_path

    @property
    def output_paths(self) -> List[str]:
        return list(self._output_paths)

    @property
    def is_dirty(self) -> bool:
        return self._dirty

    def load_output(self, output_path: str) -> None:
        """Single-output convenience — equivalent to load_outputs([output_path])."""
        self.load_outputs([output_path] if output_path else [])

    def load_outputs(self, output_paths: List[str]) -> None:
        """Load one or more outputs. If >1, show a picker; first one is rendered."""
        paths = [p for p in (output_paths or []) if p]
        self._output_paths = paths

        # Update picker visibility & contents
        self._suppress_picker_signal = True
        try:
            self.output_picker.clear()
            for p in paths:
                self.output_picker.addItem(os.path.basename(p), p)
        finally:
            self._suppress_picker_signal = False

        if len(paths) > 1:
            self.picker_row_widget.setVisible(True)
            self.outputs_count_label.setText(f"{len(paths)} files produced")
            # Render the first; user can switch via combo.
            self.output_picker.setCurrentIndex(0)
            self._render_path(paths[0])
        elif len(paths) == 1:
            self.picker_row_widget.setVisible(False)
            self.outputs_count_label.setText("")
            self._render_path(paths[0])
        else:
            self.picker_row_widget.setVisible(False)
            self.outputs_count_label.setText("")
            self._output_path = None
            self.file_label.setText("File: —")
            self.editor.clear()
            self._dirty = False
            self._refresh_save_enabled()

    def set_save_as_defaults(
        self,
        default_dir: str,
        suggested_filename: str,
        required: bool = True,
    ) -> None:
        """Configure the Save button to prompt for a final destination."""
        self._save_default_dir = default_dir or ""
        self._save_suggested_filename = suggested_filename or ""
        self._save_as_required = required
        self._refresh_save_enabled()

    def append_output(self, output_path: str) -> None:
        """Add a freshly-produced output to the picker without changing the
        currently-displayed file. If this is the first output the page has
        seen, behave like ``load_outputs([path])``."""
        if not output_path:
            return
        if not self._output_paths:
            self.load_outputs([output_path])
            return
        if output_path in self._output_paths:
            return  # already added
        self._output_paths.append(output_path)
        self._suppress_picker_signal = True
        try:
            self.output_picker.addItem(os.path.basename(output_path), output_path)
        finally:
            self._suppress_picker_signal = False
        self.picker_row_widget.setVisible(True)
        self.outputs_count_label.setText(f"{len(self._output_paths)} files produced")

    def set_progress_hint(self, text: str) -> None:
        """Show an inline 'still processing N more…' style hint. Pass '' to clear."""
        if text:
            self.progress_hint_label.setText(text)
            self.progress_hint_label.setVisible(True)
        else:
            self.progress_hint_label.clear()
            self.progress_hint_label.setVisible(False)

    def set_failure_banner(self, text: str) -> None:
        """Show a red error banner above the editor. Pass '' to clear.

        Used by TaskTab when a parallel run finishes with partial success —
        the user is already on the Output page (because file_completed
        switched to it), so the failure must be communicated here.
        """
        if text:
            self.failure_banner_label.setText(text)
            self.failure_banner_label.setVisible(True)
        else:
            self.failure_banner_label.clear()
            self.failure_banner_label.setVisible(False)

    def failure_banner_text(self) -> str:
        return self.failure_banner_label.text()

    def _render_path(self, output_path: str) -> None:
        """Load the given .docx into the editor."""
        self._output_path = output_path
        self.file_label.setText(f"File: {os.path.basename(output_path)}")
        if os.path.isfile(output_path) and output_path.lower().endswith(".docx"):
            try:
                html = load_docx_as_html(output_path)
                self.editor.blockSignals(True)
                try:
                    self.editor.setHtml(html)
                finally:
                    self.editor.blockSignals(False)
            except Exception as e:
                self.editor.setPlainText(f"(Failed to render {output_path}:\n{e})")
        else:
            self.editor.setPlainText(f"(File not found or not a .docx: {output_path})")
        self._dirty = False
        self._refresh_save_enabled()

    def _on_picker_changed(self, index: int) -> None:
        if self._suppress_picker_signal or index < 0:
            return
        new_path = self.output_picker.itemData(index)
        if not new_path or new_path == self._output_path:
            return
        # Warn before discarding unsaved edits.
        if self._dirty:
            ans = QMessageBox.question(
                self,
                "Discard changes?",
                "You have unsaved changes. Discard them and switch files?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            )
            if ans != QMessageBox.StandardButton.Yes:
                # Revert combo selection without re-triggering.
                self._suppress_picker_signal = True
                try:
                    prev_idx = next(
                        (i for i in range(self.output_picker.count())
                         if self.output_picker.itemData(i) == self._output_path),
                        0,
                    )
                    self.output_picker.setCurrentIndex(prev_idx)
                finally:
                    self._suppress_picker_signal = False
                return
        self._render_path(new_path)

    # ---- Internals ----

    def _on_text_changed(self) -> None:
        self._dirty = True
        self._refresh_save_enabled()

    def _refresh_save_enabled(self) -> None:
        self.save_btn.setEnabled(self._output_path is not None)

    def _on_save(self) -> None:
        if self._output_path is None:
            return
        if self._save_as_required:
            self._on_save_as()
            return
        if not self._dirty:
            QMessageBox.information(
                self,
                "Already saved",
                f"This output is already saved:\n{self._output_path}",
            )
            return
        try:
            save_qtextdocument_as_docx(self.editor.document(), self._output_path)
            self._dirty = False
            self._refresh_save_enabled()
            QMessageBox.information(
                self,
                "Saved",
                f"Saved:\n{self._output_path}",
            )
        except Exception as e:
            QMessageBox.critical(self, "Save failed", f"Could not save .docx:\n{e}")

    def _on_save_as(self) -> None:
        if self._output_path is None:
            return
        default_dir = self._save_default_dir or os.path.dirname(self._output_path)
        suggested_name = (
            self._save_suggested_filename or os.path.basename(self._output_path)
        )
        try:
            if default_dir:
                os.makedirs(default_dir, exist_ok=True)
        except Exception:
            default_dir = os.path.dirname(self._output_path)
        suggested_path = os.path.join(default_dir, suggested_name)
        target_path, _ = QFileDialog.getSaveFileName(
            self,
            "Save Discovery Response",
            suggested_path,
            "Word Documents (*.docx);;All files (*.*)",
        )
        if not target_path:
            return
        if not target_path.lower().endswith(".docx"):
            target_path += ".docx"
        try:
            if self._dirty:
                save_qtextdocument_as_docx(self.editor.document(), target_path)
            elif os.path.abspath(self._output_path) != os.path.abspath(target_path):
                shutil.copyfile(self._output_path, target_path)
            self._output_path = target_path
            self._save_default_dir = os.path.dirname(target_path)
            self._save_suggested_filename = os.path.basename(target_path)
            self.file_label.setText(f"File: {os.path.basename(target_path)}")
            self._dirty = False
            self._refresh_save_enabled()
            QMessageBox.information(
                self,
                "Saved",
                f"Saved:\n{target_path}",
            )
        except Exception as e:
            QMessageBox.critical(self, "Save failed", f"Could not save .docx:\n{e}")

    def _on_copy_all(self) -> None:
        QGuiApplication.clipboard().setText(self.editor.toPlainText())

    def _on_open_in_word(self) -> None:
        if self._output_path is None:
            return
        if self._dirty:
            ans = QMessageBox.question(
                self,
                "Save first?",
                "You have unsaved changes. Save before opening in Word?",
            )
            if ans == QMessageBox.StandardButton.Yes:
                self._on_save()
        try:
            os.startfile(self._output_path)  # Windows
        except Exception as e:
            QMessageBox.critical(self, "Open failed", f"Could not open in Word:\n{e}")
