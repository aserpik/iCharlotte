"""WizardTab — header + search + grouped TaskCards + Recent Tasks list."""
from PySide6.QtCore import QEvent, QObject, QSettings, Qt, Signal
from PySide6.QtWidgets import (
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QPushButton,
    QScrollArea,
    QSplitter,
    QVBoxLayout,
    QWidget,
)

from . import theme
from .registry import CATEGORY_ORDER, filter_tasks, list_tasks
from .task_card import TaskCard


_CARDS_PER_ROW = 3
_SPLITTER_SETTINGS_KEY = "wizard_tab/recent_splitter_sizes"


class WizardTab(QWidget):
    task_requested = Signal(str)            # task_id
    reopen_requested = Signal(dict)         # recent-tasks entry
    card_action_requested = Signal(str)     # card_action_id (corner button)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.cards: dict[str, TaskCard] = {}
        self._cat_sections: dict[str, QWidget] = {}
        self._cat_headers: dict[str, QLabel] = {}
        self._cat_grids: dict[str, QGridLayout] = {}
        self._visible_ids: set[str] = set()
        self._visible_cats: list[str] = []
        self._recent_layout: QVBoxLayout | None = None
        self._settings = QSettings("iCharlotte", "iCharlotte")
        self._build_ui()
        self._apply_filter("")

    # ---- Build ---------------------------------------------------------- #

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

        # Vertical splitter: search + cards on top, recent section on bottom
        self._splitter = QSplitter(Qt.Orientation.Vertical)
        self._splitter.setChildrenCollapsible(False)
        self._splitter.setHandleWidth(6)

        # Top of splitter: search box (fixed) + scrollable grouped sections
        top = QWidget()
        top_layout = QVBoxLayout(top)
        top_layout.setContentsMargins(0, 0, 0, 0)
        top_layout.setSpacing(theme.SPACE_MD)

        self.search_box = QLineEdit()
        self.search_box.setPlaceholderText("Search tasks…")
        self.search_box.setClearButtonEnabled(True)
        self.search_box.textChanged.connect(self._apply_filter)
        self.search_box.installEventFilter(self)
        top_layout.addWidget(self.search_box)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QScrollArea.Shape.NoFrame)
        container = QWidget()
        self._sections_layout = QVBoxLayout(container)
        self._sections_layout.setContentsMargins(0, 0, 0, 0)
        self._sections_layout.setSpacing(theme.SPACE_LG)
        self._sections_layout.setAlignment(
            Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft
        )

        # Build one card per task (created once, reflowed on filter).
        for spec in list_tasks():
            card = TaskCard(spec)
            card.clicked.connect(self.task_requested.emit)
            card.action_requested.connect(self.card_action_requested.emit)
            self.cards[spec.task_id] = card

        # Build one section (header + grid) per category, in order.
        for category in CATEGORY_ORDER:
            section = QWidget()
            sec_layout = QVBoxLayout(section)
            sec_layout.setContentsMargins(0, 0, 0, 0)
            sec_layout.setSpacing(theme.SPACE_SM)

            head = QLabel(category)
            head.setStyleSheet(
                f"font-size: {theme.FONT_H3}px; font-weight: 700;"
                f" letter-spacing: 0.06em; color: {theme.TEXT_MUTED};"
            )
            sec_layout.addWidget(head)

            grid_host = QWidget()
            grid = QGridLayout(grid_host)
            grid.setContentsMargins(0, 0, 0, 0)
            grid.setHorizontalSpacing(20)
            grid.setVerticalSpacing(20)
            grid.setAlignment(Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
            sec_layout.addWidget(grid_host)

            self._cat_sections[category] = section
            self._cat_headers[category] = head
            self._cat_grids[category] = grid
            self._sections_layout.addWidget(section)

        # Empty-state label, shown when a search matches nothing.
        self._empty_label = QLabel("")
        self._empty_label.setStyleSheet(
            f"color: {theme.TEXT_FAINT}; font-style: italic;"
        )
        self._empty_label.setVisible(False)
        self._sections_layout.addWidget(self._empty_label)
        self._sections_layout.addStretch(1)

        scroll.setWidget(container)
        top_layout.addWidget(scroll, 1)
        self._splitter.addWidget(top)

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

    # ---- Search / grouping --------------------------------------------- #

    def _clear_grid(self, grid: QGridLayout) -> None:
        """Remove all items from a grid layout without deleting the widgets."""
        while grid.count():
            grid.takeAt(0)

    def _apply_filter(self, query: str) -> None:
        """Reflow the grouped sections to show only tasks matching `query`."""
        grouped = filter_tasks(list_tasks(), query)
        self._visible_ids = set()
        self._visible_cats = []

        # Hide every card first; matched ones get re-shown below.
        for card in self.cards.values():
            card.hide()

        for category in CATEGORY_ORDER:
            grid = self._cat_grids[category]
            section = self._cat_sections[category]
            self._clear_grid(grid)
            matched = grouped.get(category, [])
            if not matched:
                section.setVisible(False)
                continue
            for idx, spec in enumerate(matched):
                card = self.cards[spec.task_id]
                grid.addWidget(card, idx // _CARDS_PER_ROW, idx % _CARDS_PER_ROW)
                card.show()
                self._visible_ids.add(spec.task_id)
            self._cat_headers[category].setText(f"{category}  ·  {len(matched)}")
            section.setVisible(True)
            self._visible_cats.append(category)

        if not self._visible_cats:
            q = (query or "").strip()
            self._empty_label.setText(f'No tasks match "{q}".')
            self._empty_label.setVisible(True)
        else:
            self._empty_label.setVisible(False)

    def visible_task_ids(self) -> set[str]:
        """Task ids currently rendered (for tests / introspection)."""
        return set(self._visible_ids)

    def visible_categories(self) -> list[str]:
        """Category names currently rendered, in display order."""
        return list(self._visible_cats)

    def category_header_text(self, category: str) -> str:
        return self._cat_headers[category].text()

    def eventFilter(self, obj: QObject, event: QEvent) -> bool:
        if obj is self.search_box and event.type() == QEvent.Type.KeyPress:
            if event.key() == Qt.Key.Key_Escape:
                self.search_box.clear()
                return True
        return super().eventFilter(obj, event)

    def showEvent(self, event) -> None:
        super().showEvent(event)
        # Autofocus the search box when the launcher becomes visible.
        self.search_box.setFocus()

    # ---- Recent tasks (unchanged behavior) ----------------------------- #

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
