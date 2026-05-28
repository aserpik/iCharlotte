"""Custom output page for Depo Prep — adds a markdown view above the .docx editor."""
from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QSplitter, QTextBrowser, QWidget

from .output_page import OutputPage


class DepoPrepOutputPage(OutputPage):
    def __init__(self, parent: QWidget | None = None):
        super().__init__(parent)

        # Insert a QTextBrowser ABOVE the editor by repacking via a splitter.
        outer = self.layout()

        self.md_viewer = QTextBrowser()
        self.md_viewer.setOpenExternalLinks(True)

        splitter = QSplitter(Qt.Orientation.Vertical)
        # Move the existing editor into the splitter.
        splitter.addWidget(self.md_viewer)
        splitter.addWidget(self.editor)
        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 1)

        # Find the position where the editor used to live and replace with the splitter.
        editor_idx = None
        for i in range(outer.count()):
            item = outer.itemAt(i)
            if item is not None and item.widget() is self.editor:
                editor_idx = i
                break
        if editor_idx is not None:
            outer.takeAt(editor_idx)
        outer.insertWidget(editor_idx if editor_idx is not None else 0, splitter, 1)

    def _render_path(self, output_path: str) -> None:
        # Render docx via base class behaviour.
        super()._render_path(output_path)
        md_path = Path(output_path).with_suffix(".md")
        if md_path.exists():
            try:
                self.md_viewer.setMarkdown(md_path.read_text(encoding="utf-8"))
            except Exception:
                self.md_viewer.setPlainText(md_path.read_text(encoding="utf-8"))
        else:
            self.md_viewer.clear()
