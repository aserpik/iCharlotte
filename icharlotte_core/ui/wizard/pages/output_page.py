"""OutputPage — mammoth-rendered editor + action buttons."""
import os

from PySide6.QtCore import Signal
from PySide6.QtGui import QGuiApplication
from PySide6.QtWidgets import (
    QHBoxLayout,
    QLabel,
    QPushButton,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from ..docx_io import load_docx_as_html, save_qtextdocument_as_docx


class OutputPage(QWidget):
    rerun_requested = Signal()
    edit_settings_requested = Signal()

    def __init__(self, parent: QWidget | None = None):
        super().__init__(parent)
        self._output_path: str | None = None
        self._dirty = False

        outer = QVBoxLayout(self)
        outer.setContentsMargins(24, 24, 24, 24)
        outer.setSpacing(12)

        header = QHBoxLayout()
        self.file_label = QLabel("File: —")
        self.file_label.setStyleSheet("font-weight: 600;")
        header.addWidget(self.file_label, 1)
        self.open_in_word_btn = QPushButton("Open in Word")
        self.open_in_word_btn.clicked.connect(self._on_open_in_word)
        header.addWidget(self.open_in_word_btn)
        outer.addLayout(header)

        self.editor = QTextEdit()
        self.editor.setReadOnly(False)
        self.editor.textChanged.connect(self._on_text_changed)
        outer.addWidget(self.editor, 1)

        btn_row = QHBoxLayout()
        btn_row.addStretch()
        self.copy_all_btn = QPushButton("Copy All")
        self.copy_all_btn.clicked.connect(self._on_copy_all)
        btn_row.addWidget(self.copy_all_btn)
        self.rerun_btn = QPushButton("Re-run")
        self.rerun_btn.clicked.connect(self.rerun_requested.emit)
        btn_row.addWidget(self.rerun_btn)
        self.edit_settings_btn = QPushButton("Edit Settings & Re-run")
        self.edit_settings_btn.clicked.connect(self.edit_settings_requested.emit)
        btn_row.addWidget(self.edit_settings_btn)
        self.save_btn = QPushButton("Save")
        self.save_btn.setStyleSheet(
            "background-color: #1976D2; color: white; font-weight: 600; padding: 6px 18px;"
        )
        self.save_btn.clicked.connect(self._on_save)
        btn_row.addWidget(self.save_btn)
        outer.addLayout(btn_row)

        self._refresh_save_enabled()

    # ---- Public API ----

    @property
    def output_path(self) -> str | None:
        return self._output_path

    @property
    def is_dirty(self) -> bool:
        return self._dirty

    def load_output(self, output_path: str) -> None:
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

    # ---- Internals ----

    def _on_text_changed(self) -> None:
        self._dirty = True
        self._refresh_save_enabled()

    def _refresh_save_enabled(self) -> None:
        self.save_btn.setEnabled(self._dirty and self._output_path is not None)

    def _on_save(self) -> None:
        if self._output_path is None:
            return
        try:
            save_qtextdocument_as_docx(self.editor.document(), self._output_path)
            self._dirty = False
            self._refresh_save_enabled()
        except Exception as e:
            from PySide6.QtWidgets import QMessageBox
            QMessageBox.critical(self, "Save failed", f"Could not save .docx:\n{e}")

    def _on_copy_all(self) -> None:
        QGuiApplication.clipboard().setText(self.editor.toPlainText())

    def _on_open_in_word(self) -> None:
        if self._output_path is None:
            return
        if self._dirty:
            from PySide6.QtWidgets import QMessageBox
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
            from PySide6.QtWidgets import QMessageBox
            QMessageBox.critical(self, "Open failed", f"Could not open in Word:\n{e}")
