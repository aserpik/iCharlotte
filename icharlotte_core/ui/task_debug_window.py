"""Floating task debug console for recorder events."""

from __future__ import annotations

import json
import os
from collections.abc import Callable, Iterable

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QGridLayout,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMainWindow,
    QPlainTextEdit,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from icharlotte_core import task_debug
from icharlotte_core.task_debug import TaskDebugEvent


class TaskDebugRowDetailsDialog(QDialog):
    """Expanded read-only view of every table column for one debug event."""

    def __init__(
        self,
        headers: list[str],
        values: list[str],
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle("Task Debug Row Details")
        self.resize(900, 700)

        content = QWidget(self)
        grid = QGridLayout(content)
        grid.setContentsMargins(12, 12, 12, 12)
        grid.setHorizontalSpacing(14)
        grid.setVerticalSpacing(10)

        for row, (header, value) in enumerate(zip(headers, values)):
            label = QLabel(header)
            label.setAlignment(Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
            label.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
            editor = QPlainTextEdit()
            editor.setObjectName(f"detail_{row}")
            editor.setReadOnly(True)
            editor.setPlainText(value)
            editor.setLineWrapMode(QPlainTextEdit.LineWrapMode.WidgetWidth)
            editor.setMinimumHeight(self._height_for_value(value))
            editor.setSizePolicy(
                QSizePolicy.Policy.Expanding,
                QSizePolicy.Policy.MinimumExpanding,
            )
            grid.addWidget(label, row, 0)
            grid.addWidget(editor, row, 1)
        grid.setColumnStretch(1, 1)

        scroll = QScrollArea(self)
        scroll.setWidgetResizable(True)
        scroll.setWidget(content)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Close)
        buttons.rejected.connect(self.reject)

        layout = QVBoxLayout(self)
        layout.addWidget(scroll)
        layout.addWidget(buttons)

    @staticmethod
    def _height_for_value(value: str) -> int:
        line_count = max(1, value.count("\n") + 1)
        if len(value) > 500:
            return 180
        if len(value) > 160 or line_count > 2:
            return 110
        return 54


class TaskDebugWindow(QMainWindow):
    """Reusable console for inspecting task-debug recorder events."""

    COL_TIME = 0
    COL_TASK = 1
    COL_SOURCE = 2
    COL_LEVEL = 3
    COL_PHASE = 4
    COL_MESSAGE = 5
    COL_ELAPSED = 6
    COL_DETAILS = 7

    HEADERS = [
        "Time",
        "Task",
        "Source",
        "Level",
        "Phase",
        "Message",
        "Elapsed ms",
        "Details",
    ]

    INITIAL_COLUMN_WIDTHS = {
        COL_TIME: 220,
        COL_TASK: 190,
        COL_SOURCE: 190,
        COL_LEVEL: 80,
        COL_PHASE: 130,
        COL_MESSAGE: 360,
        COL_ELAPSED: 95,
        COL_DETAILS: 420,
    }

    ALL_TASKS = "All tasks"
    ALL_SOURCES = "All sources"
    ALL_LEVELS = "All levels"

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("Task Debug Console")
        self.resize(1100, 620)
        self._events: list[TaskDebugEvent] = list(task_debug.get_events())
        self._visible_events: list[TaskDebugEvent] = []
        self._updating_filters = False

        self.task_filter = QComboBox()
        self.source_filter = QComboBox()
        self.level_filter = QComboBox()
        self.search_input = QLineEdit()
        self.pause_autoscroll_check = QCheckBox("Pause autoscroll")
        self.clear_btn = QPushButton("Clear")
        self.copy_btn = QPushButton("Copy")
        self.open_folder_btn = QPushButton("Open Folder")
        self.table = QTableWidget(0, len(self.HEADERS))

        self._build_ui()
        self._connect_signals()
        self._refresh_filter_options()
        self._rebuild_table(scroll_to_bottom=False)

        task_debug.get_bridge().event_emitted.connect(self._on_event_emitted)

    def _build_ui(self) -> None:
        self.task_filter.setObjectName("task_filter")
        self.source_filter.setObjectName("source_filter")
        self.level_filter.setObjectName("level_filter")
        self.search_input.setObjectName("search_input")
        self.pause_autoscroll_check.setObjectName("pause_autoscroll_check")
        self.clear_btn.setObjectName("clear_btn")
        self.copy_btn.setObjectName("copy_btn")
        self.open_folder_btn.setObjectName("open_folder_btn")
        self.table.setObjectName("table")

        self.search_input.setPlaceholderText("Search message/details")
        self.table.setHorizontalHeaderLabels(self.HEADERS)
        self.table.setAlternatingRowColors(True)
        self.table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.table.setSortingEnabled(False)
        self.table.verticalHeader().setVisible(False)
        header = self.table.horizontalHeader()
        header.setStretchLastSection(False)
        header.setMinimumSectionSize(48)
        for col in range(len(self.HEADERS)):
            header.setSectionResizeMode(col, QHeaderView.ResizeMode.Interactive)
            self.table.setColumnWidth(col, self.INITIAL_COLUMN_WIDTHS.get(col, 120))

        filters = QHBoxLayout()
        filters.addWidget(self.task_filter, 2)
        filters.addWidget(self.source_filter, 2)
        filters.addWidget(self.level_filter, 1)
        filters.addWidget(self.search_input, 3)
        filters.addWidget(self.pause_autoscroll_check)
        filters.addWidget(self.copy_btn)
        filters.addWidget(self.clear_btn)
        filters.addWidget(self.open_folder_btn)

        layout = QVBoxLayout()
        layout.addLayout(filters)
        layout.addWidget(self.table)

        central = QWidget(self)
        central.setLayout(layout)
        self.setCentralWidget(central)

    def _connect_signals(self) -> None:
        self.task_filter.currentIndexChanged.connect(self._on_filter_changed)
        self.source_filter.currentIndexChanged.connect(self._on_filter_changed)
        self.level_filter.currentIndexChanged.connect(self._on_filter_changed)
        self.search_input.textChanged.connect(self._on_filter_changed)
        self.clear_btn.clicked.connect(self._clear_events)
        self.copy_btn.clicked.connect(self._copy_visible_rows)
        self.open_folder_btn.clicked.connect(self._open_trace_folder)
        self.table.cellDoubleClicked.connect(self._show_row_details)

    def _on_filter_changed(self) -> None:
        if self._updating_filters:
            return
        self._rebuild_table(scroll_to_bottom=False)

    def _on_event_emitted(self, event: object) -> None:
        if not isinstance(event, TaskDebugEvent):
            return
        # Mirror the recorder's bounded buffer so an open console cannot grow
        # indefinitely during long sessions.
        self._events = list(task_debug.get_events())
        self._refresh_filter_options()
        self._rebuild_table(scroll_to_bottom=not self.pause_autoscroll_check.isChecked())

    def _clear_events(self) -> None:
        task_debug.clear_events()
        self._events.clear()
        self._refresh_filter_options()
        self._rebuild_table(scroll_to_bottom=False)

    def _copy_visible_rows(self) -> None:
        rows: list[str] = []
        for row in range(self.table.rowCount()):
            cells = []
            for col in range(self.table.columnCount()):
                item = self.table.item(row, col)
                cells.append(item.text() if item is not None else "")
            rows.append("\t".join(cells))
        QApplication.clipboard().setText("\n".join(rows))

    def _open_trace_folder(self) -> None:
        startfile = getattr(os, "startfile", None)
        if startfile is None:
            return
        try:
            startfile(str(task_debug.get_trace_dir()))
        except Exception:
            return

    def _refresh_filter_options(self) -> None:
        self._updating_filters = True
        try:
            self._set_combo_options(
                self.task_filter,
                [(self.ALL_TASKS, None)]
                + [
                    (self._task_label(event), self._task_key(event))
                    for event in self._unique_events_by_key(
                        self._events, lambda event: self._task_key(event)
                    )
                ],
            )
            self._set_combo_options(
                self.source_filter,
                [(self.ALL_SOURCES, None)]
                + [
                    (source, source)
                    for source in sorted({e.source for e in self._events if e.source})
                ],
            )
            self._set_combo_options(
                self.level_filter,
                [(self.ALL_LEVELS, None)]
                + [
                    (level, level)
                    for level in sorted({e.level for e in self._events if e.level})
                ],
            )
        finally:
            self._updating_filters = False

    def _set_combo_options(
        self,
        combo: QComboBox,
        options: list[tuple[str, str | None]],
    ) -> None:
        previous_data = combo.currentData()
        previous_text = combo.currentText()
        combo.blockSignals(True)
        try:
            combo.clear()
            for text, data in options:
                combo.addItem(text, data)
            index = -1
            if previous_data is not None:
                index = combo.findData(previous_data)
            if index < 0 and previous_text:
                index = combo.findText(previous_text)
            combo.setCurrentIndex(index if index >= 0 else 0)
        finally:
            combo.blockSignals(False)

    def _rebuild_table(self, *, scroll_to_bottom: bool) -> None:
        visible = [event for event in self._events if self._matches_filters(event)]
        self._visible_events = visible
        self.table.setRowCount(0)
        for event in visible:
            self._append_row(event)
        if scroll_to_bottom:
            self.table.scrollToBottom()

    def _append_row(self, event: TaskDebugEvent) -> None:
        row = self.table.rowCount()
        self.table.insertRow(row)
        values = self._event_values(event)
        for col, value in enumerate(values):
            item = QTableWidgetItem(value)
            if col == self.COL_LEVEL:
                item.setData(Qt.ItemDataRole.UserRole, event.level)
            self.table.setItem(row, col, item)

    def _show_row_details(self, row: int, _column: int) -> None:
        dialog = self._build_row_details_dialog(row)
        if dialog is None:
            return
        dialog.exec()

    def _build_row_details_dialog(self, row: int) -> TaskDebugRowDetailsDialog | None:
        if row < 0 or row >= len(self._visible_events):
            return None
        return TaskDebugRowDetailsDialog(
            self.HEADERS,
            self._event_values(self._visible_events[row]),
            parent=self,
        )

    def _event_values(self, event: TaskDebugEvent) -> list[str]:
        return [
            event.timestamp,
            self._task_label(event),
            event.source,
            event.level,
            event.phase,
            event.message,
            str(event.elapsed_ms),
            self._details_text(event),
        ]

    def _matches_filters(self, event: TaskDebugEvent) -> bool:
        task_key = self.task_filter.currentData()
        if task_key is not None and self._task_key(event) != task_key:
            return False
        source = self.source_filter.currentData()
        if source is not None and event.source != source:
            return False
        level = self.level_filter.currentData()
        if level is not None and event.level != level:
            return False
        query = self.search_input.text().strip().casefold()
        if query:
            haystack = f"{event.message}\n{self._details_text(event)}".casefold()
            if query not in haystack:
                return False
        return True

    def _details_text(self, event: TaskDebugEvent) -> str:
        if not event.details:
            return ""
        try:
            return json.dumps(event.details, ensure_ascii=False, sort_keys=True)
        except Exception:
            return str(event.details)

    def _task_label(self, event: TaskDebugEvent) -> str:
        title = event.task_title.strip()
        task_id = event.task_id.strip()
        if title and task_id and title != task_id:
            return f"{title} ({task_id})"
        return title or task_id or "(untitled task)"

    def _task_key(self, event: TaskDebugEvent) -> str:
        return f"{event.task_title}\0{event.task_id}"

    def _unique_events_by_key(
        self,
        events: Iterable[TaskDebugEvent],
        key_func: Callable[[TaskDebugEvent], str],
    ) -> list[TaskDebugEvent]:
        seen: set[str] = set()
        unique: list[TaskDebugEvent] = []
        for event in events:
            key = key_func(event)
            if key in seen:
                continue
            seen.add(key)
            unique.append(event)
        return unique
