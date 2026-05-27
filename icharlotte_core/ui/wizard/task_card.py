"""TaskCard — clickable card on the Wizard tab."""
from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QMouseEvent
from PySide6.QtWidgets import QFrame, QHBoxLayout, QLabel, QVBoxLayout

from .registry import TaskSpec


_CARD_STYLE = """
TaskCard {
    background-color: #ffffff;
    border: 1px solid #e0e0e0;
    border-radius: 10px;
}
TaskCard:hover {
    border-color: #1976D2;
    background-color: #fafcff;
}
"""

_ICON_TILE_STYLE = """
QLabel#icon_tile {
    background-color: #fff7e6;
    border-radius: 8px;
    font-size: 22px;
    qproperty-alignment: AlignCenter;
}
"""


class TaskCard(QFrame):
    """A single card representing a task. Clicking emits `clicked(task_id)`."""

    clicked = Signal(str)  # task_id

    def __init__(self, spec: TaskSpec, parent=None):
        super().__init__(parent)
        self._spec = spec
        self.setObjectName("TaskCard")
        self.setStyleSheet(_CARD_STYLE + _ICON_TILE_STYLE)
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setFixedSize(280, 140)

        outer = QVBoxLayout(self)
        outer.setContentsMargins(16, 14, 16, 14)
        outer.setSpacing(8)

        header = QHBoxLayout()
        header.setSpacing(10)

        self.icon_tile = QLabel(spec.icon_glyph)
        self.icon_tile.setObjectName("icon_tile")
        self.icon_tile.setFixedSize(36, 36)
        header.addWidget(self.icon_tile)

        self.title_label = QLabel(spec.title)
        self.title_label.setStyleSheet("font-size: 14px; font-weight: 600; color: #1a1a1a;")
        self.title_label.setWordWrap(True)
        header.addWidget(self.title_label, 1)

        outer.addLayout(header)

        self.description_label = QLabel(spec.description)
        self.description_label.setWordWrap(True)
        self.description_label.setStyleSheet("font-size: 12px; color: #666;")
        outer.addWidget(self.description_label, 1)

    @property
    def task_id(self) -> str:
        return self._spec.task_id

    def mousePressEvent(self, event: QMouseEvent) -> None:
        if event.button() == Qt.MouseButton.LeftButton:
            self.clicked.emit(self._spec.task_id)
        super().mousePressEvent(event)
