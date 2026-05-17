"""WizardTab — header + grid of TaskCards. Recent Tasks added in Phase 7."""
from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QGridLayout,
    QLabel,
    QScrollArea,
    QVBoxLayout,
    QWidget,
)

from .registry import list_tasks
from .task_card import TaskCard


_CARDS_PER_ROW = 3


class WizardTab(QWidget):
    """The 'What would you like to do?' card grid tab."""

    task_requested = Signal(str)  # task_id

    def __init__(self, parent: QWidget | None = None):
        super().__init__(parent)
        self.cards: list[TaskCard] = []
        self._build_ui()

    def _build_ui(self) -> None:
        outer = QVBoxLayout(self)
        outer.setContentsMargins(40, 32, 40, 32)
        outer.setSpacing(24)

        header = QLabel("What would you like to do?")
        header.setStyleSheet("font-size: 22px; font-weight: 400; color: #1a1a1a;")
        outer.addWidget(header)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QScrollArea.Shape.NoFrame)

        container = QWidget()
        grid = QGridLayout(container)
        grid.setContentsMargins(0, 0, 0, 0)
        grid.setHorizontalSpacing(20)
        grid.setVerticalSpacing(20)
        grid.setAlignment(Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)

        for idx, spec in enumerate(list_tasks()):
            card = TaskCard(spec, parent=container)
            card.clicked.connect(self.task_requested.emit)
            row, col = divmod(idx, _CARDS_PER_ROW)
            grid.addWidget(card, row, col)
            self.cards.append(card)

        scroll.setWidget(container)
        outer.addWidget(scroll, 1)
