"""WizardTab — header + grid of TaskCards + Recent Tasks list."""
from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QScrollArea,
    QVBoxLayout,
    QWidget,
)

from .registry import list_tasks
from .task_card import TaskCard


_CARDS_PER_ROW = 3


class WizardTab(QWidget):
    task_requested = Signal(str)            # task_id
    reopen_requested = Signal(dict)         # recent-tasks entry

    def __init__(self, parent=None):
        super().__init__(parent)
        self.cards: list[TaskCard] = []
        self._recent_layout: QVBoxLayout | None = None
        self._build_ui()

    def _build_ui(self):
        outer = QVBoxLayout(self)
        outer.setContentsMargins(40, 32, 40, 32)
        outer.setSpacing(24)

        header = QLabel("What would you like to do?")
        header.setStyleSheet("font-size: 22px; font-weight: 400; color: #1a1a1a;")
        outer.addWidget(header)

        # Card grid
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

        # Divider
        line = QFrame()
        line.setFrameShape(QFrame.Shape.HLine)
        line.setStyleSheet("color: #e0e0e0;")
        outer.addWidget(line)

        # Recent Tasks
        recent_label = QLabel("Recent Tasks")
        recent_label.setStyleSheet("font-size: 14px; font-weight: 600; color: #333;")
        outer.addWidget(recent_label)

        self._recent_layout = QVBoxLayout()
        self._recent_layout.setSpacing(6)
        outer.addLayout(self._recent_layout)
        self._render_recent_empty_state()

    def _render_recent_empty_state(self):
        self._clear_recent_layout()
        empty = QLabel("No completed tasks for this case yet.")
        empty.setStyleSheet("color: #999; font-style: italic;")
        self._recent_layout.addWidget(empty)

    def _clear_recent_layout(self):
        if self._recent_layout is None:
            return
        while self._recent_layout.count():
            item = self._recent_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

    def refresh_recent_tasks(self, entries: list[dict]):
        """Update the Recent Tasks list."""
        if self._recent_layout is None:
            return
        if not entries:
            self._render_recent_empty_state()
            return
        self._clear_recent_layout()
        for entry in entries:
            row = self._build_recent_row(entry)
            self._recent_layout.addWidget(row)

    def _build_recent_row(self, entry: dict) -> QWidget:
        w = QFrame()
        w.setStyleSheet("QFrame { border-bottom: 1px solid #f0f0f0; padding: 4px 0; }")
        h = QHBoxLayout(w)
        h.setContentsMargins(0, 0, 0, 0)
        h.setSpacing(12)

        title = entry.get("title", entry.get("task_id", "Unknown"))
        ts = entry.get("completed_at", "")
        label = QLabel(f"• {title}  —  {ts}")
        label.setStyleSheet("font-size: 12px; color: #333;")
        h.addWidget(label, 1)

        out_path = entry.get("output_path") or ""
        if out_path:
            label.setToolTip(out_path)

        btn = QPushButton("Reopen")
        btn.setFixedHeight(26)
        btn.setStyleSheet("padding: 0 12px;")
        btn.clicked.connect(lambda _=False, e=entry: self.reopen_requested.emit(e))
        h.addWidget(btn)

        return w
