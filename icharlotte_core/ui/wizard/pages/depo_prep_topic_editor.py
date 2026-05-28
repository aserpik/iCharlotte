"""TopicEditor - drag-reorder + checkable + add/remove topic list for Depo Prep.

Uses native Qt.ItemFlag flags; NEVER setItemWidget (MEMORY.md
qlistwidget_setitemwidget_drag.md). The visible row text encodes both title and
strategic note via two lines; on edit, we re-parse.
"""
from __future__ import annotations

from typing import List

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QAbstractItemView, QHBoxLayout, QInputDialog, QListWidget, QListWidgetItem,
    QPushButton, QVBoxLayout, QWidget,
)


def _format_item_text(title: str, strategic_note: str) -> str:
    if strategic_note:
        return f"{title}\n    Strategic: {strategic_note}"
    return title


def _parse_item_text(text: str) -> tuple:
    if "\n    Strategic: " in text:
        title, _, rest = text.partition("\n    Strategic: ")
        return title.strip(), rest.strip()
    return text.strip(), ""


class TopicEditor(QWidget):
    topics_changed = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(4)

        self._list = QListWidget()
        self._list.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self._list.setDragDropMode(QAbstractItemView.DragDropMode.InternalMove)
        self._list.setDefaultDropAction(Qt.DropAction.MoveAction)
        self._list.itemChanged.connect(lambda *_: self.topics_changed.emit())
        self._list.model().rowsMoved.connect(lambda *_: self.topics_changed.emit())
        outer.addWidget(self._list)

        btns = QHBoxLayout()
        self.add_btn = QPushButton("+ Add custom topic")
        self.add_btn.clicked.connect(self._on_add_clicked)
        btns.addWidget(self.add_btn)
        self.remove_btn = QPushButton("Remove selected")
        self.remove_btn.clicked.connect(self._on_remove_clicked)
        btns.addWidget(self.remove_btn)
        btns.addStretch()
        outer.addLayout(btns)

        # Internal: track per-row metadata (id, lawyer_added, refs) keyed by row index.
        self._meta: List[dict] = []

    def set_topics(self, topics: List[dict]) -> None:
        self._list.blockSignals(True)
        try:
            self._list.clear()
            self._meta = []
            for t in topics:
                item = QListWidgetItem(_format_item_text(t.get("title", ""),
                                                          t.get("strategic_note", "")))
                item.setFlags(
                    Qt.ItemFlag.ItemIsSelectable
                    | Qt.ItemFlag.ItemIsEnabled
                    | Qt.ItemFlag.ItemIsUserCheckable
                    | Qt.ItemFlag.ItemIsEditable
                    | Qt.ItemFlag.ItemIsDragEnabled
                )
                item.setCheckState(Qt.CheckState.Checked if t.get("default_checked", True)
                                   else Qt.CheckState.Unchecked)
                self._list.addItem(item)
                self._meta.append({
                    "id": t.get("id"),
                    "lawyer_added": bool(t.get("lawyer_added", False)),
                    "relevant_digest_refs": list(t.get("relevant_digest_refs", [])),
                })
        finally:
            self._list.blockSignals(False)

    def get_topics(self) -> List[dict]:
        out = []
        for row in range(self._list.count()):
            item = self._list.item(row)
            title, strat = _parse_item_text(item.text())
            meta = self._meta[row] if row < len(self._meta) else {}
            out.append({
                "id": meta.get("id") or f"t{row+1:02d}",
                "title": title,
                "strategic_note": strat,
                "relevant_digest_refs": meta.get("relevant_digest_refs", []),
                "default_checked": item.checkState() == Qt.CheckState.Checked,
                "lawyer_added": bool(meta.get("lawyer_added", False)),
            })
        return out

    def set_checked(self, row: int, checked: bool) -> None:
        item = self._list.item(row)
        if item:
            item.setCheckState(Qt.CheckState.Checked if checked else Qt.CheckState.Unchecked)

    def add_topic(self, *, title: str, strategic_note: str) -> None:
        new_id = f"t{(len(self._meta) + 1):02d}"
        item = QListWidgetItem(_format_item_text(title, strategic_note))
        item.setFlags(
            Qt.ItemFlag.ItemIsSelectable
            | Qt.ItemFlag.ItemIsEnabled
            | Qt.ItemFlag.ItemIsUserCheckable
            | Qt.ItemFlag.ItemIsEditable
            | Qt.ItemFlag.ItemIsDragEnabled
        )
        item.setCheckState(Qt.CheckState.Checked)
        self._list.addItem(item)
        self._meta.append({"id": new_id, "lawyer_added": True, "relevant_digest_refs": []})
        self.topics_changed.emit()

    def remove_topic_at(self, row: int) -> None:
        if 0 <= row < self._list.count():
            self._list.takeItem(row)
            if row < len(self._meta):
                self._meta.pop(row)
            self.topics_changed.emit()

    def _on_add_clicked(self) -> None:
        title, ok = QInputDialog.getText(self, "New topic", "Topic title:")
        if not ok or not title.strip():
            return
        note, _ = QInputDialog.getText(self, "Strategic note",
                                        "1-2 sentence strategic note (optional):")
        self.add_topic(title=title.strip(), strategic_note=(note or "").strip())

    def _on_remove_clicked(self) -> None:
        rows = sorted({i.row() for i in self._list.selectedIndexes()}, reverse=True)
        for r in rows:
            self.remove_topic_at(r)
