"""WizardTab — header + grid of TaskCards + Recent Tasks list."""
from PySide6.QtCore import QSettings, Qt, Signal
from PySide6.QtWidgets import (
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QScrollArea,
    QSplitter,
    QVBoxLayout,
    QWidget,
)

from . import theme
from .registry import list_tasks
from .task_card import TaskCard


_CARDS_PER_ROW = 3
_SPLITTER_SETTINGS_KEY = "wizard_tab/recent_splitter_sizes"


class WizardTab(QWidget):
    task_requested = Signal(str)            # task_id
    reopen_requested = Signal(dict)         # recent-tasks entry

    def __init__(self, parent=None):
        super().__init__(parent)
        self.cards: list[TaskCard] = []
        self._recent_layout: QVBoxLayout | None = None
        self._settings = QSettings("iCharlotte", "iCharlotte")
        self._build_ui()

    def _build_ui(self):
        self.setStyleSheet(theme.wizard_stylesheet())

        outer = QVBoxLayout(self)
        outer.setContentsMargins(40, 32, 40, 32)
        outer.setSpacing(theme.SPACE_XL)

        header = QLabel("What would you like to do?")
        header.setStyleSheet(
            f"font-size: {theme.FONT_H1}px; font-weight: 600; color: {theme.TEXT};"
        )
        outer.addWidget(header)

        subtitle = theme.helper_text(
            "Pick a task to get started. Each one walks you through Settings → Running → Output."
        )
        outer.addWidget(subtitle)

        # Vertical splitter: cards on top, recent section on bottom
        self._splitter = QSplitter(Qt.Orientation.Vertical)
        self._splitter.setChildrenCollapsible(False)
        self._splitter.setHandleWidth(6)

        # Card grid (top of splitter)
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
        self._splitter.addWidget(scroll)

        # Bottom section: toggle button + (collapsible) recent container
        bottom = QWidget()
        bottom_layout = QVBoxLayout(bottom)
        bottom_layout.setContentsMargins(0, 0, 0, 0)
        bottom_layout.setSpacing(8)

        self._recent_toggle_btn = QPushButton("Show Recent Tasks")
        self._recent_toggle_btn.setStyleSheet(
            f"QPushButton {{ font-size: {theme.FONT_H3}px; color: {theme.TEXT_MUTED};"
            f" background: transparent; border: none; padding: 4px 0; text-align: left; }}"
            f" QPushButton:hover {{ color: {theme.TEXT}; }}"
        )
        self._recent_toggle_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self._recent_toggle_btn.clicked.connect(self._on_toggle_recent)
        bottom_layout.addWidget(
            self._recent_toggle_btn, alignment=Qt.AlignmentFlag.AlignLeft
        )

        self._recent_container = QWidget()
        recent_outer = QVBoxLayout(self._recent_container)
        recent_outer.setContentsMargins(0, 0, 0, 0)
        recent_outer.setSpacing(8)

        recent_label = theme.section_header("Recent Tasks")
        recent_outer.addWidget(recent_label)

        self._recent_layout = QVBoxLayout()
        self._recent_layout.setSpacing(6)
        recent_outer.addLayout(self._recent_layout)

        self._recent_container.setVisible(False)
        bottom_layout.addWidget(self._recent_container, 1)

        self._splitter.addWidget(bottom)
        self._splitter.setStretchFactor(0, 1)
        self._splitter.setStretchFactor(1, 0)
        self._splitter.splitterMoved.connect(self._save_splitter_sizes)
        outer.addWidget(self._splitter, 1)
        self._render_recent_empty_state()

    def _on_toggle_recent(self):
        visible = not self._recent_container.isVisible()
        self._recent_container.setVisible(visible)
        self._recent_toggle_btn.setText(
            "Hide Recent Tasks" if visible else "Show Recent Tasks"
        )
        if visible:
            self._apply_expanded_splitter_sizes()

    def _apply_expanded_splitter_sizes(self):
        """Give the recent section meaningful space when expanded — restore saved sizes if any."""
        saved = self._settings.value(_SPLITTER_SETTINGS_KEY)
        if saved:
            try:
                sizes = [int(x) for x in saved]
                if len(sizes) == 2 and sizes[1] > 0:
                    self._splitter.setSizes(sizes)
                    return
            except (TypeError, ValueError):
                pass
        sizes = self._splitter.sizes()
        total = sum(sizes) or self.height() or 800
        recent_h = max(220, min(320, total // 3))
        self._splitter.setSizes([max(200, total - recent_h), recent_h])

    def _save_splitter_sizes(self, *_args):
        if self._recent_container.isVisible():
            self._settings.setValue(_SPLITTER_SETTINGS_KEY, self._splitter.sizes())

    def _render_recent_empty_state(self):
        self._clear_recent_layout()
        empty = QLabel("No completed tasks for this case yet.")
        empty.setStyleSheet(f"color: {theme.TEXT_FAINT}; font-style: italic;")
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
        w.setStyleSheet(
            f"QFrame {{ border-bottom: 1px solid {theme.BORDER_LIGHT}; padding: 4px 0; }}"
        )
        h = QHBoxLayout(w)
        h.setContentsMargins(0, 0, 0, 0)
        h.setSpacing(theme.SPACE_MD)

        title = entry.get("title", entry.get("task_id", "Unknown"))
        ts = entry.get("completed_at", "")
        label = QLabel(f"• {title}  —  {ts}")
        label.setStyleSheet(f"font-size: {theme.FONT_BODY}px; color: {theme.TEXT_BODY};")
        h.addWidget(label, 1)

        out_path = entry.get("output_path") or ""
        if out_path:
            label.setToolTip(out_path)

        btn = theme.secondary_button("Reopen")
        btn.setFixedHeight(28)
        btn.clicked.connect(lambda _=False, e=entry: self.reopen_requested.emit(e))
        h.addWidget(btn)

        return w
