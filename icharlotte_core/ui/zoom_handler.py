"""
Ctrl+Scroll zoom for all tabs.

Intercepts Ctrl+Wheel events at the application level and applies
proportional scaling (font-size, padding, margin) to native Qt widgets.
Skips QWebEngineView, PdfViewerWidget, and FilePreviewWidget so their
own zoom behaviour is preserved.

Ctrl+0 resets the current tab to default zoom.
"""

import re
from PySide6.QtCore import QObject, QEvent, Qt
from PySide6.QtWidgets import QApplication, QHeaderView, QTabWidget, QTableWidget, QWidget


# Widget class names whose subtrees we must NOT zoom.
_SKIP_CLASS_NAMES = {"QWebEngineView", "PdfViewerWidget", "FilePreviewWidget"}

ZOOM_MIN = 0.5
ZOOM_MAX = 3.0
ZOOM_STEP = 0.1
ZOOM_DEFAULT = 1.0

# Regex for CSS properties whose px values should scale with zoom.
_SCALABLE_RE = re.compile(
    r"(font-size|padding(?:-[a-z]+)?|margin(?:-[a-z]+)?)\s*:\s*([^;]+)",
    re.IGNORECASE,
)
_PX_VAL_RE = re.compile(r"(\d+(?:\.\d+)?)px")


def _scale_stylesheet(original: str, zoom: float) -> str:
    """Return *original* with font-size/padding/margin px values scaled."""
    def _scale_prop(m: re.Match) -> str:
        prop = m.group(1)
        scaled_value = _PX_VAL_RE.sub(
            lambda v: f"{round(float(v.group(1)) * zoom, 1)}px",
            m.group(2),
        )
        return f"{prop}: {scaled_value}"
    return _SCALABLE_RE.sub(_scale_prop, original)


class ZoomEventFilter(QObject):
    """Application-level event filter that adds Ctrl+Scroll zoom to every tab."""

    def __init__(self, tab_widget: QTabWidget, parent: QObject | None = None):
        super().__init__(parent)
        self._tab_widget = tab_widget
        self._tab_zoom: dict[int, float] = {}
        self._base_font_size: float = max(QApplication.font().pointSizeF(), 9.0)
        # tab_index -> [(widget_ref, original_stylesheet), ...]
        self._original_styles: dict[int, list[tuple[QWidget, str]]] = {}
        # tab_index -> [(table, default_row_h, [col_widths], [resize_modes]), ...]
        self._original_tables: dict[int, list[tuple[QTableWidget, int, list[int], list[int]]]] = {}

    # ------------------------------------------------------------------
    # Public helpers
    # ------------------------------------------------------------------

    def zoom_level(self, tab_index: int) -> float:
        return self._tab_zoom.get(tab_index, ZOOM_DEFAULT)

    def reset_zoom(self, tab_index: int) -> None:
        """Reset a tab to default zoom."""
        self._tab_zoom[tab_index] = ZOOM_DEFAULT
        self._apply_zoom(tab_index)

    # ------------------------------------------------------------------
    # Event filter
    # ------------------------------------------------------------------

    def eventFilter(self, obj: QObject, event: QEvent) -> bool:  # noqa: N802
        if event.type() == QEvent.Type.Wheel:
            return self._handle_wheel(event)
        if event.type() == QEvent.Type.KeyPress:
            return self._handle_key(event)
        return False

    # ------------------------------------------------------------------
    # Wheel zoom
    # ------------------------------------------------------------------

    def _handle_wheel(self, event) -> bool:
        if not (event.modifiers() & Qt.KeyboardModifier.ControlModifier):
            return False

        widget = QApplication.widgetAt(event.globalPosition().toPoint())
        if widget is None:
            return False

        if self._inside_skip_widget(widget):
            return False

        tab_index = self._tab_index_for_widget(widget)
        if tab_index is None:
            return False

        delta = event.angleDelta().y()
        if delta == 0:
            return False

        current = self._tab_zoom.get(tab_index, ZOOM_DEFAULT)
        new_zoom = round(current + (ZOOM_STEP if delta > 0 else -ZOOM_STEP), 2)
        new_zoom = max(ZOOM_MIN, min(ZOOM_MAX, new_zoom))

        if new_zoom == current:
            return True  # at limit, swallow event

        self._tab_zoom[tab_index] = new_zoom
        self._apply_zoom(tab_index)
        return True

    # ------------------------------------------------------------------
    # Ctrl+0 reset
    # ------------------------------------------------------------------

    def _handle_key(self, event) -> bool:
        if not (event.modifiers() & Qt.KeyboardModifier.ControlModifier):
            return False
        if event.key() != Qt.Key.Key_0:
            return False
        tab_index = self._tab_widget.currentIndex()
        if self._tab_zoom.get(tab_index, ZOOM_DEFAULT) != ZOOM_DEFAULT:
            self.reset_zoom(tab_index)
            return True
        return False

    # ------------------------------------------------------------------
    # Internal helpers
    # ------------------------------------------------------------------

    @staticmethod
    def _inside_skip_widget(widget: QWidget) -> bool:
        """Walk the parent chain; return True if inside a skip-listed class."""
        w = widget
        while w is not None:
            if type(w).__name__ in _SKIP_CLASS_NAMES:
                return True
            w = w.parentWidget()
        return False

    def _tab_index_for_widget(self, widget: QWidget) -> int | None:
        """Find the tab index that contains *widget*."""
        w = widget
        while w is not None:
            for i in range(self._tab_widget.count()):
                if self._tab_widget.widget(i) is w:
                    return i
            w = w.parentWidget()
        return None

    # ------------------------------------------------------------------
    # Snapshot & restore original stylesheets
    # ------------------------------------------------------------------

    def _snapshot_styles(self, tab_index: int) -> None:
        """Capture original stylesheets of all widgets with scalable properties."""
        tab_root = self._tab_widget.widget(tab_index)
        if tab_root is None:
            return
        originals: list[tuple[QWidget, str]] = []

        # Tab root itself (strip any prior zoom injection)
        root_ss = re.sub(r"/\* zoom \*/.*?/\* /zoom \*/\s*", "", tab_root.styleSheet())
        if root_ss.strip() and _SCALABLE_RE.search(root_ss):
            originals.append((tab_root, root_ss))

        # All child widgets, skipping those inside PDF/web viewers
        for child in tab_root.findChildren(QWidget):
            if self._inside_skip_widget(child):
                continue
            ss = child.styleSheet()
            if ss and _SCALABLE_RE.search(ss):
                originals.append((child, ss))

        self._original_styles[tab_index] = originals

    def _restore_originals(self, tab_index: int) -> None:
        """Restore every widget to its pre-zoom stylesheet."""
        for widget, original_ss in self._original_styles.get(tab_index, []):
            try:
                widget.setStyleSheet(original_ss)
            except RuntimeError:
                continue  # widget was destroyed
        self._original_styles.pop(tab_index, None)

    # ------------------------------------------------------------------
    # Snapshot & restore table dimensions
    # ------------------------------------------------------------------

    def _snapshot_tables(self, tab_index: int) -> None:
        """Capture original row heights and column widths for all tables in the tab."""
        tab_root = self._tab_widget.widget(tab_index)
        if tab_root is None:
            return
        tables: list[tuple[QTableWidget, int, list[int], list[int]]] = []
        for table in tab_root.findChildren(QTableWidget):
            if self._inside_skip_widget(table):
                continue
            header = table.horizontalHeader()
            row_h = table.verticalHeader().defaultSectionSize()
            col_widths = [table.columnWidth(c) for c in range(table.columnCount())]
            resize_modes = [header.sectionResizeMode(c) for c in range(table.columnCount())]
            tables.append((table, row_h, col_widths, resize_modes))
        self._original_tables[tab_index] = tables

    def _scale_tables(self, tab_index: int, zoom: float) -> None:
        """Scale row heights and interactive column widths for all tables."""
        for table, base_row_h, base_widths, resize_modes in self._original_tables.get(tab_index, []):
            try:
                table.verticalHeader().setDefaultSectionSize(round(base_row_h * zoom))
                for c, (base_w, mode) in enumerate(zip(base_widths, resize_modes)):
                    # Only scale columns with fixed/interactive widths;
                    # leave Stretch and ResizeToContents columns alone.
                    if mode == QHeaderView.ResizeMode.Interactive:
                        table.setColumnWidth(c, round(base_w * zoom))
            except RuntimeError:
                continue  # table was destroyed

    def _restore_tables(self, tab_index: int) -> None:
        """Restore original table dimensions."""
        for table, base_row_h, base_widths, resize_modes in self._original_tables.get(tab_index, []):
            try:
                table.verticalHeader().setDefaultSectionSize(base_row_h)
                for c, (base_w, mode) in enumerate(zip(base_widths, resize_modes)):
                    if mode == QHeaderView.ResizeMode.Interactive:
                        table.setColumnWidth(c, base_w)
            except RuntimeError:
                continue
        self._original_tables.pop(tab_index, None)

    # ------------------------------------------------------------------
    # Apply zoom
    # ------------------------------------------------------------------

    def _apply_zoom(self, tab_index: int) -> None:
        tab_root = self._tab_widget.widget(tab_index)
        if tab_root is None:
            return

        zoom = self._tab_zoom.get(tab_index, ZOOM_DEFAULT)

        # Reset path
        if zoom == ZOOM_DEFAULT:
            self._restore_originals(tab_index)
            self._restore_tables(tab_index)
            self._clear_root_zoom(tab_root)
            return

        # Snapshot on first zoom for this tab
        if tab_index not in self._original_styles:
            self._snapshot_styles(tab_index)
        if tab_index not in self._original_tables:
            self._snapshot_tables(tab_index)

        # Scale every child widget that had hardcoded sizes
        root_in_originals = False
        for widget, original_ss in self._original_styles.get(tab_index, []):
            try:
                scaled = _scale_stylesheet(original_ss, zoom)
                if widget is tab_root:
                    root_in_originals = True
                    # Root gets both the font cascade AND its scaled stylesheet
                    font_pt = round(self._base_font_size * zoom, 1)
                    widget.setStyleSheet(
                        f"/* zoom */ font-size: {font_pt}pt; /* /zoom */ " + scaled
                    )
                else:
                    widget.setStyleSheet(scaled)
            except RuntimeError:
                continue  # widget was destroyed

        # If root had no scalable styles, still inject font cascade
        if not root_in_originals:
            self._inject_root_zoom(tab_root, zoom)

        # Scale table row heights and column widths
        self._scale_tables(tab_index, zoom)

    def _inject_root_zoom(self, tab_root: QWidget, zoom: float) -> None:
        """Inject a font-size rule on the tab root for widgets without explicit sizes."""
        font_pt = round(self._base_font_size * zoom, 1)
        rule = f"/* zoom */ font-size: {font_pt}pt; /* /zoom */"
        existing = tab_root.styleSheet()
        if "/* zoom */" in existing:
            updated = re.sub(r"/\* zoom \*/.*?/\* /zoom \*/", rule, existing)
            tab_root.setStyleSheet(updated)
        else:
            tab_root.setStyleSheet(rule + " " + existing)

    @staticmethod
    def _clear_root_zoom(tab_root: QWidget) -> None:
        """Remove the zoom font-size rule from the tab root."""
        existing = tab_root.styleSheet()
        if "/* zoom */" in existing:
            cleaned = re.sub(r"/\* zoom \*/.*?/\* /zoom \*/\s*", "", existing)
            tab_root.setStyleSheet(cleaned)
