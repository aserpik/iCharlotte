"""Shared scaffold for Wizard task tabs: a persistent header with a step
indicator above an inner page stack.

`WizardTaskContainer` lets the existing task containers (TaskTab,
InProcessTaskTab, OpposeMotionTaskTab) keep their `QStackedWidget`-style API
(`addWidget`, `setCurrentIndex`, `currentIndex`, ...) while gaining a consistent
header and Settings → Running → Output progress indicator for free. The methods
are proxied to an inner QStackedWidget, so call sites and tests that do
`tab.setCurrentIndex(PAGE_OUTPUT)` / `tab.currentIndex()` keep working.
"""
from __future__ import annotations

from typing import List, Optional

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QFrame,
    QHBoxLayout,
    QLabel,
    QStackedWidget,
    QVBoxLayout,
    QWidget,
)

from . import theme


_DEFAULT_STEPS = ["Settings", "Running", "Output"]


class TaskHeader(QWidget):
    """Icon + title + description, plus a numbered step indicator.

    `set_step(i)` marks step *i* active and any earlier steps as done.
    """

    def __init__(self, spec, steps: Optional[List[str]] = None, parent=None):
        super().__init__(parent)
        self._steps = list(steps) if steps else list(_DEFAULT_STEPS)
        self._current = 0

        outer = QVBoxLayout(self)
        outer.setContentsMargins(
            theme.SPACE_XL, theme.SPACE_LG, theme.SPACE_XL, theme.SPACE_MD
        )
        outer.setSpacing(theme.SPACE_MD)

        # --- Title row: icon tile + title/description ---
        title_row = QHBoxLayout()
        title_row.setSpacing(theme.SPACE_MD)

        icon = QLabel(getattr(spec, "icon_glyph", "") or "")
        icon.setFixedSize(40, 40)
        icon.setAlignment(Qt.AlignmentFlag.AlignCenter)
        icon.setStyleSheet(
            f"background-color: {theme.PRIMARY_SUBTLE}; border-radius: {theme.RADIUS_MD}px;"
            f" font-size: 22px;"
        )
        title_row.addWidget(icon, 0, Qt.AlignmentFlag.AlignTop)

        text_col = QVBoxLayout()
        text_col.setSpacing(2)
        title = QLabel(getattr(spec, "title", "") or "")
        title.setStyleSheet(
            f"font-size: {theme.FONT_H1 - 4}px; font-weight: 600; color: {theme.TEXT};"
        )
        text_col.addWidget(title)
        desc = QLabel(getattr(spec, "description", "") or "")
        desc.setWordWrap(True)
        desc.setStyleSheet(f"font-size: {theme.FONT_BODY}px; color: {theme.TEXT_MUTED};")
        text_col.addWidget(desc)
        title_row.addLayout(text_col, 1)
        outer.addLayout(title_row)

        # --- Step indicator row ---
        steps_row = QHBoxLayout()
        steps_row.setSpacing(theme.SPACE_SM)
        self._step_labels: List[QLabel] = []
        self._connectors: List[QLabel] = []
        for i, name in enumerate(self._steps):
            lbl = QLabel(f"{i + 1}. {name}")
            self._step_labels.append(lbl)
            steps_row.addWidget(lbl)
            if i < len(self._steps) - 1:
                conn = QLabel("——")
                self._connectors.append(conn)
                steps_row.addWidget(conn)
        steps_row.addStretch(1)
        outer.addLayout(steps_row)

        # Bottom divider
        line = QFrame()
        line.setFrameShape(QFrame.Shape.HLine)
        line.setStyleSheet(f"color: {theme.BORDER_LIGHT}; background-color: {theme.BORDER_LIGHT};")
        line.setFixedHeight(1)
        outer.addWidget(line)

        self._restyle()

    def set_step(self, index: int) -> None:
        if not self._step_labels:
            return
        self._current = max(0, min(index, len(self._step_labels) - 1))
        self._restyle()

    def _restyle(self) -> None:
        for i, lbl in enumerate(self._step_labels):
            if i == self._current:
                color, weight = theme.PRIMARY, "700"
            elif i < self._current:
                color, weight = theme.SUCCESS, "600"
            else:
                color, weight = theme.TEXT_FAINT, "500"
            lbl.setStyleSheet(
                f"font-size: {theme.FONT_H3}px; color: {color}; font-weight: {weight};"
            )
        for conn in self._connectors:
            conn.setStyleSheet(f"color: {theme.BORDER}; font-size: {theme.FONT_H3}px;")


class WizardTaskContainer(QWidget):
    """A header + inner QStackedWidget, proxying the stacked-widget API.

    Subclasses call ``super().__init__(spec, ...)`` and then use ``addWidget`` /
    ``setCurrentIndex`` / ``currentIndex`` exactly as if they subclassed
    QStackedWidget directly.
    """

    def __init__(self, spec, steps: Optional[List[str]] = None, parent=None):
        super().__init__(parent)
        self._spec = spec
        self.setStyleSheet(theme.wizard_stylesheet())

        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        self.header = TaskHeader(spec, steps=steps)
        outer.addWidget(self.header)

        self._stack = QStackedWidget()
        outer.addWidget(self._stack, 1)
        self._stack.currentChanged.connect(self.header.set_step)

    # ---- Proxied QStackedWidget API ----

    def addWidget(self, w: QWidget) -> int:
        return self._stack.addWidget(w)

    def insertWidget(self, index: int, w: QWidget) -> int:
        return self._stack.insertWidget(index, w)

    def removeWidget(self, w: QWidget) -> None:
        self._stack.removeWidget(w)

    def setCurrentIndex(self, index: int) -> None:
        self._stack.setCurrentIndex(index)

    def currentIndex(self) -> int:
        return self._stack.currentIndex()

    def setCurrentWidget(self, w: QWidget) -> None:
        self._stack.setCurrentWidget(w)

    def currentWidget(self) -> QWidget:
        return self._stack.currentWidget()

    def widget(self, index: int) -> QWidget:
        return self._stack.widget(index)

    def indexOf(self, w: QWidget) -> int:
        return self._stack.indexOf(w)

    def count(self) -> int:
        return self._stack.count()
