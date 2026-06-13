"""Structured chronology viewer for Med Record Extractor."""
from __future__ import annotations

import html
from dataclasses import dataclass
import os
import re

from PySide6.QtCore import (
    QEvent,
    QItemSelectionModel,
    QRect,
    QSettings,
    QSize,
    Qt,
    QTimer,
    Signal,
)
from PySide6.QtGui import (
    QAbstractTextDocumentLayout,
    QBrush,
    QColor,
    QShortcut,
    QTextDocument,
)
from PySide6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QFrame,
    QFileDialog,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QListView,
    QListWidget,
    QListWidgetItem,
    QMessageBox,
    QPushButton,
    QSplitter,
    QStyle,
    QStyleOptionViewItem,
    QTableWidget,
    QTableWidgetItem,
    QStyledItemDelegate,
    QTabWidget,
    QToolButton,
    QVBoxLayout,
    QWidget,
)

from icharlotte_core.med_record_chronology import (
    ChronologyDocument,
    MatchResult,
    SelectableChronologyRow,
    SelectionState,
    SynopsisParagraph,
    match_synopsis_to_rows,
    parse_chronology_document,
)
from icharlotte_core.med_record_extractor import _build_file_index, _lookup_file
from icharlotte_core.ui.wizard import theme
from icharlotte_core.ui.wizard.file_picker import (
    find_medical_summary_folder,
    resolve_default_folder,
)


TABLE_COLUMN_WIDTHS_KEY = "wizard/med_record_extractor/chronology_table_column_widths"
DEFAULT_TABLE_COLUMN_WIDTHS = [56, 96, 240, 260, 520, 180]
SELECTION_HIGHLIGHT_COLOR = "#fff3a3"
SELECTION_HIGHLIGHT_BORDER_COLOR = "#d1a500"
SEARCH_HIGHLIGHT_COLOR = "#ffeb3b"
ACTIVE_SEARCH_HIGHLIGHT_COLOR = "#ffb74d"
_SELECTION_HIGHLIGHT_BRUSH = QBrush(QColor(SELECTION_HIGHLIGHT_COLOR))


@dataclass(frozen=True)
class _SearchMatch:
    tab_index: int
    row: int
    column: int = 0


def _normalize_search_text(text: str) -> str:
    return " ".join(str(text or "").casefold().split())


def _html_preserve_text(text: str) -> str:
    return html.escape(text).replace("\n", "<br/>")


def _search_ranges(
    text: str,
    query: str,
    *,
    highlight_terms: bool,
) -> list[tuple[int, int]]:
    text = str(text or "")
    query = _normalize_search_text(query)
    terms = query.split()
    if not text or not terms:
        return []

    phrase_pattern = r"\s+".join(re.escape(term) for term in terms)
    ranges = [
        (match.start(), match.end())
        for match in re.finditer(phrase_pattern, text, flags=re.IGNORECASE)
    ]
    if ranges or not highlight_terms:
        return ranges

    term_ranges: list[tuple[int, int]] = []
    for term in terms:
        term_ranges.extend(
            (match.start(), match.end())
            for match in re.finditer(re.escape(term), text, flags=re.IGNORECASE)
        )
    return _merge_ranges(term_ranges)


def _merge_ranges(ranges: list[tuple[int, int]]) -> list[tuple[int, int]]:
    merged: list[tuple[int, int]] = []
    for start, end in sorted(ranges):
        if start >= end:
            continue
        if merged and start <= merged[-1][1]:
            merged[-1] = (merged[-1][0], max(merged[-1][1], end))
        else:
            merged.append((start, end))
    return merged


def search_highlight_html(
    text: str,
    query: str,
    *,
    active: bool = False,
    highlight_terms: bool = False,
) -> str:
    """Return escaped HTML with visible find-result spans around query matches."""

    text = str(text or "")
    ranges = _search_ranges(text, query, highlight_terms=highlight_terms)
    if not ranges:
        return _html_preserve_text(text)

    background = ACTIVE_SEARCH_HIGHLIGHT_COLOR if active else SEARCH_HIGHLIGHT_COLOR
    parts: list[str] = []
    cursor = 0
    for start, end in ranges:
        parts.append(_html_preserve_text(text[cursor:start]))
        parts.append(
            f'<span style="background-color: {background}; color: #111111;">'
            f"{_html_preserve_text(text[start:end])}</span>"
        )
        cursor = end
    parts.append(_html_preserve_text(text[cursor:]))
    return "".join(parts)


def _paint_highlighted_item(
    delegate: QStyledItemDelegate,
    painter,
    option,
    index,
    highlighted_html: str,
) -> None:
    view_option = QStyleOptionViewItem(option)
    delegate.initStyleOption(view_option, index)
    view_option.text = ""

    widget = view_option.widget
    style = widget.style() if widget is not None else QApplication.style()
    style.drawControl(
        QStyle.ControlElement.CE_ItemViewItem,
        view_option,
        painter,
        widget,
    )

    text_rect = style.subElementRect(
        QStyle.SubElement.SE_ItemViewItemText,
        view_option,
        widget,
    )
    if not text_rect.isValid():
        return

    document = QTextDocument()
    document.setDefaultFont(view_option.font)
    document.setDocumentMargin(0)
    document.setTextWidth(text_rect.width())
    document.setHtml(highlighted_html)

    painter.save()
    painter.translate(text_rect.topLeft())
    context = QAbstractTextDocumentLayout.PaintContext()
    context.palette = view_option.palette
    document.documentLayout().draw(painter, context)
    painter.restore()


def _brief_synopsis_stylesheet() -> str:
    check_path = theme._asset_path("checkmark.svg")
    return f"""
    QListWidget {{
        border: 1px solid {theme.BORDER};
        border-radius: {theme.RADIUS_SM}px;
        background-color: #FFFFFF;
        selection-background-color: {SELECTION_HIGHLIGHT_COLOR};
        selection-color: {theme.TEXT};
    }}
    QListWidget::item {{
        padding: 4px 6px;
        border-radius: 3px;
    }}
    QListWidget::item:selected {{
        background-color: {SELECTION_HIGHLIGHT_COLOR};
        color: {theme.TEXT};
    }}
    QListWidget::item:checked {{
        background-color: {SELECTION_HIGHLIGHT_COLOR};
        color: {theme.TEXT};
    }}
    QListWidget::indicator {{
        width: 16px;
        height: 16px;
        border: 1px solid {theme.BORDER};
        border-radius: 3px;
        background-color: #FFFFFF;
    }}
    QListWidget::indicator:hover {{
        border-color: {SELECTION_HIGHLIGHT_BORDER_COLOR};
    }}
    QListWidget::indicator:checked {{
        background-color: {SELECTION_HIGHLIGHT_COLOR};
        border-color: {SELECTION_HIGHLIGHT_BORDER_COLOR};
        image: url("{check_path}");
    }}
    QListWidget::indicator:checked:hover {{
        background-color: {SELECTION_HIGHLIGHT_COLOR};
    }}
    """


class _BriefSynopsisDelegate(QStyledItemDelegate):
    """Paint checked synopsis entries across the full row before default content."""

    def paint(self, painter, option, index) -> None:
        if index.data(Qt.ItemDataRole.CheckStateRole) == Qt.CheckState.Checked:
            painter.save()
            painter.fillRect(option.rect, QColor(SELECTION_HIGHLIGHT_COLOR))
            painter.restore()
        panel = self.parent()
        if (
            isinstance(panel, BriefSynopsisPanel)
            and index.row() in panel.find_highlight_rows
            and panel.find_highlight_query
        ):
            _paint_highlighted_item(
                self,
                painter,
                option,
                index,
                search_highlight_html(
                    str(index.data(Qt.ItemDataRole.DisplayRole) or ""),
                    panel.find_highlight_query,
                    active=index.row() == panel.active_find_highlight_row,
                ),
            )
            return
        super().paint(painter, option, index)


class _ChronologyTableDelegate(QStyledItemDelegate):
    """Paint visible find-result marks inside chronology table cells."""

    def paint(self, painter, option, index) -> None:
        table = self.parent()
        if (
            isinstance(table, ChronologyTablePanel)
            and index.column() > 0
            and index.row() in table.find_highlight_rows
            and table.find_highlight_query
        ):
            highlighted_html = search_highlight_html(
                str(index.data(Qt.ItemDataRole.DisplayRole) or ""),
                table.find_highlight_query,
                active=index.row() == table.active_find_highlight_row,
                highlight_terms=True,
            )
            if "background-color:" in highlighted_html:
                _paint_highlighted_item(
                    self,
                    painter,
                    option,
                    index,
                    highlighted_html,
                )
                return
        super().paint(painter, option, index)


class BriefSynopsisPanel(QListWidget):
    """Checkable list of synopsis paragraphs from the selected chronology."""

    paragraph_toggled = Signal(str, bool)
    paragraph_open_requested = Signal(str)

    def __init__(self, parent: QWidget | None = None):
        super().__init__(parent)
        self._items_by_id: dict[str, QListWidgetItem] = {}
        self.find_highlight_query = ""
        self.find_highlight_rows: set[int] = set()
        self.active_find_highlight_row = -1
        self.setSelectionMode(QAbstractItemView.SelectionMode.MultiSelection)
        self.setWordWrap(True)
        self.setTextElideMode(Qt.TextElideMode.ElideNone)
        self.setResizeMode(QListView.ResizeMode.Adjust)
        self.setUniformItemSizes(False)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.setItemDelegate(_BriefSynopsisDelegate(self))
        self.setStyleSheet(_brief_synopsis_stylesheet())
        self.itemChanged.connect(self._on_item_changed)

    def load_paragraphs(self, paragraphs: list[SynopsisParagraph]) -> None:
        self.blockSignals(True)
        self.clear()
        self._items_by_id.clear()
        self.set_find_highlight("", set(), -1)
        for paragraph in paragraphs:
            item = QListWidgetItem(paragraph.text)
            item.setData(Qt.ItemDataRole.UserRole, paragraph.id)
            item.setFlags(
                item.flags()
                | Qt.ItemFlag.ItemIsUserCheckable
                | Qt.ItemFlag.ItemIsEnabled
            )
            item.setCheckState(Qt.CheckState.Unchecked)
            if paragraph.warning:
                item.setToolTip(paragraph.warning)
            self.addItem(item)
            self._items_by_id[paragraph.id] = item
        self.blockSignals(False)
        self._refresh_item_sizes()

    def set_find_highlight(
        self,
        query: str,
        rows: set[int],
        active_row: int,
    ) -> None:
        self.find_highlight_query = query
        self.find_highlight_rows = set(rows)
        self.active_find_highlight_row = active_row
        self.viewport().update()

    def set_paragraph_checked(self, paragraph_id: str, checked: bool) -> None:
        item = self._items_by_id[paragraph_id]
        item.setCheckState(Qt.CheckState.Checked if checked else Qt.CheckState.Unchecked)
        self._apply_item_highlight(item)

    def mark_warning(self, paragraph_id: str, message: str) -> None:
        item = self._items_by_id.get(paragraph_id)
        if item is not None:
            item.setToolTip(message)

    def mousePressEvent(self, event) -> None:
        if event.button() != Qt.MouseButton.LeftButton:
            super().mousePressEvent(event)
            return
        item = self.itemAt(event.position().toPoint())
        if item is not None:
            checked = item.checkState() == Qt.CheckState.Checked
            item.setCheckState(Qt.CheckState.Unchecked if checked else Qt.CheckState.Checked)
            event.accept()
            return
        super().mousePressEvent(event)

    def mouseDoubleClickEvent(self, event) -> None:
        if event.button() != Qt.MouseButton.LeftButton:
            super().mouseDoubleClickEvent(event)
            return
        item = self.itemAt(event.position().toPoint())
        if item is not None:
            paragraph_id = item.data(Qt.ItemDataRole.UserRole)
            self.paragraph_open_requested.emit(paragraph_id)
            event.accept()
            return
        super().mouseDoubleClickEvent(event)

    def resizeEvent(self, event) -> None:
        super().resizeEvent(event)
        self._refresh_item_sizes()

    def _on_item_changed(self, item: QListWidgetItem) -> None:
        self._apply_item_highlight(item)
        paragraph_id = item.data(Qt.ItemDataRole.UserRole)
        self.paragraph_toggled.emit(
            paragraph_id,
            item.checkState() == Qt.CheckState.Checked,
        )

    def _refresh_item_sizes(self) -> None:
        text_width = max(120, self.viewport().width() - 48)
        for index in range(self.count()):
            item = self.item(index)
            item.setSizeHint(self._size_hint_for_text(item.text(), text_width))

    def _size_hint_for_text(self, text: str, width: int) -> QSize:
        flags = Qt.TextFlag.TextWordWrap | Qt.TextFlag.TextExpandTabs
        rect = self.fontMetrics().boundingRect(
            QRect(0, 0, width, 100000),
            flags,
            text,
        )
        return QSize(width, max(self.fontMetrics().height(), rect.height()) + 16)

    @staticmethod
    def _apply_item_highlight(item: QListWidgetItem) -> None:
        if item.checkState() == Qt.CheckState.Checked:
            item.setBackground(_SELECTION_HIGHLIGHT_BRUSH)
            item.setSelected(True)
        else:
            item.setBackground(QBrush())
            item.setSelected(False)


class ChronologyTablePanel(QTableWidget):
    """Checkable chronology table rows."""

    row_toggled = Signal(str, bool)
    row_open_requested = Signal(str)

    def __init__(self, parent: QWidget | None = None):
        super().__init__(parent)
        self._rows_by_id: dict[str, int] = {}
        self._ids_by_row: dict[int, str] = {}
        self._extractable_by_id: dict[str, bool] = {}
        self.find_highlight_query = ""
        self.find_highlight_rows: set[int] = set()
        self.active_find_highlight_row = -1
        self.setColumnCount(6)
        self.setHorizontalHeaderLabels([
            "Select",
            "DATE",
            "PAGE NO",
            "PROVIDER",
            "DESCRIPTION",
            "Red Flags/Comments",
        ])
        self.setWordWrap(True)
        self.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.setSelectionMode(QAbstractItemView.SelectionMode.NoSelection)
        self.setAlternatingRowColors(True)
        self.setItemDelegate(_ChronologyTableDelegate(self))
        self.verticalHeader().setVisible(False)
        self.horizontalHeader().setStretchLastSection(False)
        for column in range(self.columnCount()):
            self.horizontalHeader().setSectionResizeMode(
                column,
                QHeaderView.ResizeMode.Interactive,
            )
        self.restore_column_widths()
        self._resize_commit_timer = QTimer(self)
        self._resize_commit_timer.setSingleShot(True)
        self._resize_commit_timer.setInterval(150)
        self._resize_commit_timer.timeout.connect(self._commit_column_resize)
        self.horizontalHeader().sectionResized.connect(self._on_section_resized)
        self.cellChanged.connect(self._on_cell_changed)

    def count(self) -> int:
        return self.rowCount()

    def restore_column_widths(self) -> None:
        widths = _coerce_column_widths(
            QSettings("iCharlotte", "iCharlotte").value(TABLE_COLUMN_WIDTHS_KEY),
            self.columnCount(),
        )
        if widths is None:
            widths = DEFAULT_TABLE_COLUMN_WIDTHS[: self.columnCount()]
        for column, width in enumerate(widths):
            self.setColumnWidth(column, width)
        self.resizeRowsToContents()

    def save_column_widths(self) -> None:
        widths = [self.columnWidth(column) for column in range(self.columnCount())]
        settings = QSettings("iCharlotte", "iCharlotte")
        settings.setValue(TABLE_COLUMN_WIDTHS_KEY, widths)
        settings.sync()

    def _on_section_resized(self, logical_index: int, old_size: int, new_size: int) -> None:
        self._resize_commit_timer.start()

    def _commit_column_resize(self) -> None:
        self.save_column_widths()
        self.resizeRowsToContents()

    def load_rows(self, rows: list[SelectableChronologyRow]) -> None:
        self.blockSignals(True)
        self.setRowCount(len(rows))
        self._rows_by_id.clear()
        self._ids_by_row.clear()
        self._extractable_by_id.clear()
        self.set_find_highlight("", set(), -1)
        for index, row in enumerate(rows):
            self._rows_by_id[row.id] = index
            self._ids_by_row[index] = row.id
            self._extractable_by_id[row.id] = row.extractable
            self.setItem(index, 0, self._check_item(row))
            self.setItem(index, 1, self._text_item(row.date))
            self.setItem(index, 2, self._text_item(row.page_no))
            self.setItem(index, 3, self._text_item(row.provider))
            self.setItem(index, 4, self._text_item(row.description))
            self.setItem(index, 5, self._text_item(row.flags))
            if row.warning:
                for column in range(self.columnCount()):
                    item = self.item(index, column)
                    if item is not None:
                        item.setToolTip(row.warning)
            self._apply_row_highlight(index, False)
        self.resizeRowsToContents()
        self.blockSignals(False)

    def set_find_highlight(
        self,
        query: str,
        rows: set[int],
        active_row: int,
    ) -> None:
        self.find_highlight_query = query
        self.find_highlight_rows = set(rows)
        self.active_find_highlight_row = active_row
        self.viewport().update()

    def set_row_checked(self, row_id: str, checked: bool, *, emit: bool = True) -> None:
        row_index = self._rows_by_id[row_id]
        item = self.item(row_index, 0)
        if item is None:
            return
        if checked and not self._extractable_by_id.get(row_id, False):
            checked = False
        if not emit:
            self.blockSignals(True)
        item.setCheckState(Qt.CheckState.Checked if checked else Qt.CheckState.Unchecked)
        self._apply_row_highlight(row_index, checked)
        if not emit:
            self.blockSignals(False)

    def is_row_checked(self, row_id: str) -> bool:
        item = self.item(self._rows_by_id[row_id], 0)
        return item is not None and item.checkState() == Qt.CheckState.Checked

    def mousePressEvent(self, event) -> None:
        if event.button() != Qt.MouseButton.LeftButton:
            super().mousePressEvent(event)
            return
        row_index = self.rowAt(event.position().toPoint().y())
        row_id = self._ids_by_row.get(row_index)
        if row_id is not None:
            if self._extractable_by_id.get(row_id, False):
                self.set_row_checked(row_id, not self.is_row_checked(row_id))
            event.accept()
            return
        super().mousePressEvent(event)

    def mouseDoubleClickEvent(self, event) -> None:
        if event.button() != Qt.MouseButton.LeftButton:
            super().mouseDoubleClickEvent(event)
            return
        row_index = self.rowAt(event.position().toPoint().y())
        row_id = self._ids_by_row.get(row_index)
        if row_id is not None:
            self.row_open_requested.emit(row_id)
            event.accept()
            return
        super().mouseDoubleClickEvent(event)

    def _on_cell_changed(self, row: int, column: int) -> None:
        if column != 0:
            return
        item = self.item(row, 0)
        if item is None:
            return
        row_id = item.data(Qt.ItemDataRole.UserRole)
        checked = item.checkState() == Qt.CheckState.Checked
        self._apply_row_highlight(row, checked)
        self.row_toggled.emit(row_id, checked)

    @staticmethod
    def _check_item(row: SelectableChronologyRow) -> QTableWidgetItem:
        item = QTableWidgetItem("")
        item.setData(Qt.ItemDataRole.UserRole, row.id)
        flags = Qt.ItemFlag.ItemIsEnabled
        if row.extractable:
            flags |= Qt.ItemFlag.ItemIsUserCheckable
        item.setFlags(flags)
        item.setCheckState(Qt.CheckState.Unchecked)
        return item

    @staticmethod
    def _text_item(text: str) -> QTableWidgetItem:
        item = QTableWidgetItem(text)
        item.setFlags(Qt.ItemFlag.ItemIsEnabled)
        return item

    def _apply_row_highlight(self, row: int, checked: bool) -> None:
        brush = _SELECTION_HIGHLIGHT_BRUSH if checked else QBrush()
        for column in range(self.columnCount()):
            item = self.item(row, column)
            if item is not None:
                item.setBackground(brush)


class MedChronologySelectionPage(QWidget):
    """Review a generated chronology and choose rows for record extraction."""

    run_requested = Signal(dict)

    def __init__(
        self,
        case_path: str,
        file_number: str,
        chronology_path: str,
        parent: QWidget | None = None,
    ):
        super().__init__(parent)
        self.case_path = case_path
        self.file_number = file_number
        self.chronology_path = os.path.normpath(chronology_path) if chronology_path else ""
        self.document = self._load_document(self.chronology_path)
        self.selection = SelectionState()
        self._paragraphs = {paragraph.id: paragraph for paragraph in self.document.synopsis_paragraphs}
        self._rows = {row.id: row for row in self.document.rows}
        self._paragraph_match_status: dict[str, str] = {}
        self._file_index: dict[str, str] | None = None
        self.pdf_preview_viewer: QWidget | None = None
        self._pdf_preview_collapsed = False
        self._pdf_preview_width = 420
        self._pdf_preview_target_page = 1
        self._pdf_preview_navigation_attempts = 0
        self._find_matches: list[_SearchMatch] = []
        self._find_index = -1

        self._setup_ui()
        self._refresh_extract_state()
        self._sync_chronology_visibility()

    def set_paragraph_checked(self, paragraph_id: str, checked: bool) -> None:
        self.synopsis_panel.set_paragraph_checked(paragraph_id, checked)

    def set_row_checked(self, row_id: str, checked: bool) -> None:
        self.table_panel.set_row_checked(row_id, checked)

    def is_row_checked(self, row_id: str) -> bool:
        return self.table_panel.is_row_checked(row_id)

    def to_dict(self) -> dict:
        return {
            "chronology_path": self.chronology_path,
            "selected_paragraph_ids": sorted(self.selection.selected_paragraph_ids),
            "selected_row_ids": self.selection.selected_row_ids(),
            "selected_row_sources": self.selection.selected_row_sources(),
        }

    def from_dict(self, data: dict) -> None:
        chronology_path = str(data.get("chronology_path") or "")
        if chronology_path and self.case_path and not os.path.isabs(chronology_path):
            chronology_path = os.path.join(self.case_path, chronology_path)
        if chronology_path and os.path.normpath(chronology_path) != self.chronology_path:
            self.load_chronology(chronology_path)
        if not self.chronology_path:
            return

        self._clear_selection()

        for paragraph_id in data.get("selected_paragraph_ids", []):
            if paragraph_id in self._paragraphs:
                self.set_paragraph_checked(paragraph_id, True)

        row_sources = data.get("selected_row_sources")
        if isinstance(row_sources, dict):
            for row_id, sources in row_sources.items():
                row = self._rows.get(str(row_id))
                if row is None or not row.extractable:
                    continue
                for source in _saved_sources(sources):
                    self.selection.select_row(row.id, source=source)
        else:
            for row_id in data.get("selected_row_ids", []):
                row = self._rows.get(str(row_id))
                if row is not None and row.extractable:
                    self.selection.select_row(row.id, source="manual")

        self._sync_table_checks()
        self._refresh_match_status()
        self._refresh_extract_state()

    def load_chronology(self, chronology_path: str) -> None:
        chronology_path = os.path.normpath(chronology_path)
        self.chronology_path = chronology_path
        self.document = self._load_document(chronology_path)
        self.selection = SelectionState()
        self._paragraphs = {
            paragraph.id: paragraph for paragraph in self.document.synopsis_paragraphs
        }
        self._rows = {row.id: row for row in self.document.rows}
        self._paragraph_match_status.clear()
        self._file_index = None
        self._find_matches = []
        self._find_index = -1
        if self.pdf_preview_viewer is not None and hasattr(self.pdf_preview_viewer, "clear"):
            self.pdf_preview_viewer.clear()
        self.table_panel.load_rows(self.document.rows)
        self.synopsis_panel.load_paragraphs(self.document.synopsis_paragraphs)
        self.tab_widget.setCurrentIndex(0)
        self.find_bar.setVisible(False)
        self.message_label.setText(_status_message(self.document))
        self.message_label.setVisible(bool(self.message_label.text()))
        self._set_match_status("")
        self.pdf_preview_status_label.setText(
            "Double-click a chronology entry to preview its source PDF."
        )
        self.pdf_preview_placeholder.setVisible(True)
        self._refresh_extract_state()
        self._sync_chronology_visibility()

    def _clear_selection(self) -> None:
        self.selection.clear()
        self._paragraph_match_status.clear()
        self.synopsis_panel.blockSignals(True)
        for paragraph_id in self._paragraphs:
            self.synopsis_panel.set_paragraph_checked(paragraph_id, False)
        self.synopsis_panel.blockSignals(False)
        for row_id in self._rows:
            self.table_panel.set_row_checked(row_id, False, emit=False)

    def _load_document(self, chronology_path: str) -> ChronologyDocument:
        if not chronology_path:
            return ChronologyDocument(source_path="")
        try:
            return parse_chronology_document(chronology_path)
        except Exception as exc:
            return ChronologyDocument(
                source_path=os.path.normpath(chronology_path),
                blocking_errors=[f"Could not parse chronology document: {exc}"],
            )

    def _setup_ui(self) -> None:
        self.setStyleSheet(theme.wizard_stylesheet())

        layout = QVBoxLayout(self)
        layout.setContentsMargins(
            theme.SPACE_SM,
            theme.SPACE_SM,
            theme.SPACE_SM,
            theme.SPACE_SM,
        )
        layout.setSpacing(theme.SPACE_SM)

        self.message_label = theme.error_text(_status_message(self.document))
        self.warning_label = self.message_label
        self.message_label.setVisible(bool(self.message_label.text()))
        layout.addWidget(self.message_label)

        self.match_status_label = theme.caption("")
        self.match_status_label.setWordWrap(True)
        self.match_status_label.setVisible(False)
        layout.addWidget(self.match_status_label)

        self.preview_splitter = QSplitter(Qt.Orientation.Horizontal)
        layout.addWidget(self.preview_splitter, 1)

        self.empty_state = self._build_empty_state()
        layout.addWidget(self.empty_state, 1)

        left_pane = QWidget()
        left_layout = QVBoxLayout(left_pane)
        left_layout.setContentsMargins(0, 0, theme.SPACE_SM, 0)
        left_layout.setSpacing(theme.SPACE_SM)

        self.tab_widget = QTabWidget()
        self.find_bar = self._build_find_bar()
        self.find_bar.setVisible(False)
        left_layout.addWidget(self.find_bar)

        self.tab_widget.addTab(self._build_synopsis_pane(), "Brief Synopsis")
        self.tab_widget.addTab(self._build_table_pane(), "Chronology Rows")
        self.tab_widget.currentChanged.connect(self._on_find_scope_changed)
        left_layout.addWidget(self.tab_widget, 1)

        controls = QHBoxLayout()
        controls.setContentsMargins(0, 0, 0, 0)
        self.selected_count_label = QLabel("")
        controls.addWidget(self.selected_count_label)
        controls.addStretch(1)

        self.pdf_preview_toggle_btn = theme.secondary_button("Hide Preview")
        self.pdf_preview_toggle_btn.setToolTip("Show or hide the PDF preview panel")
        self.pdf_preview_toggle_btn.clicked.connect(self.toggle_pdf_preview_collapse)
        controls.addWidget(self.pdf_preview_toggle_btn)

        self.open_original_btn = theme.secondary_button("Open Original")
        self.open_original_btn.clicked.connect(self._open_original)
        controls.addWidget(self.open_original_btn)

        self.extract_btn = theme.primary_button("Extract")
        self.extract_btn.clicked.connect(self._emit_run)
        controls.addWidget(self.extract_btn)
        left_layout.addLayout(controls)

        self.preview_splitter.addWidget(left_pane)
        self.pdf_preview_panel = self._build_pdf_preview_panel()
        self.preview_splitter.addWidget(self.pdf_preview_panel)
        self.preview_splitter.setSizes([900, self._pdf_preview_width])
        self.preview_splitter.setCollapsible(1, True)
        self.find_shortcut = QShortcut("Ctrl+F", self)
        self.find_shortcut.setContext(Qt.ShortcutContext.WidgetWithChildrenShortcut)
        self.find_shortcut.activated.connect(self._show_find_bar)

    def _build_empty_state(self) -> QWidget:
        panel = QFrame()
        panel.setFrameShape(QFrame.Shape.NoFrame)
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addStretch(1)

        self.select_chronology_btn = theme.primary_button(
            "Please select medical chronology"
        )
        self.select_chronology_btn.clicked.connect(self._select_chronology_file)
        layout.addWidget(
            self.select_chronology_btn,
            alignment=Qt.AlignmentFlag.AlignCenter,
        )

        layout.addStretch(1)
        return panel

    def _select_chronology_file(self) -> None:
        start_dir = find_medical_summary_folder(self.case_path) or resolve_default_folder(
            self.case_path,
            ["RECORDS"],
        )
        chronology_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select medical chronology summary",
            start_dir,
            "Word Documents (*.docx)",
        )
        if chronology_path:
            self.load_chronology(chronology_path)

    def _sync_chronology_visibility(self) -> None:
        has_chronology = bool(self.chronology_path)
        self.empty_state.setVisible(not has_chronology)
        self.preview_splitter.setVisible(has_chronology)
        self.message_label.setVisible(
            has_chronology and bool(self.message_label.text())
        )

    def _build_find_bar(self) -> QWidget:
        bar = QFrame()
        bar.setObjectName("medRecordFindBar")
        bar.setFrameShape(QFrame.Shape.NoFrame)
        bar.setStyleSheet(
            f"QFrame#medRecordFindBar {{ background: #FFFFFF; border: 1px solid {theme.BORDER_LIGHT};"
            f" border-radius: {theme.RADIUS_SM}px; }}"
        )

        layout = QHBoxLayout(bar)
        layout.setContentsMargins(8, 6, 8, 6)
        layout.setSpacing(6)

        self.find_input = QLineEdit()
        self.find_input.setPlaceholderText("Find...")
        self.find_input.setClearButtonEnabled(True)
        self.find_input.textChanged.connect(self._on_find_text_changed)
        self.find_input.installEventFilter(self)
        layout.addWidget(self.find_input, 1)

        self.find_status_label = theme.caption("")
        self.find_status_label.setMinimumWidth(54)
        layout.addWidget(self.find_status_label)

        self.find_previous_btn = QToolButton()
        self.find_previous_btn.setText("<")
        self.find_previous_btn.setToolTip("Previous match")
        self.find_previous_btn.setFixedSize(28, 24)
        self.find_previous_btn.clicked.connect(
            lambda: self._activate_next_find_match(backward=True)
        )
        layout.addWidget(self.find_previous_btn)

        self.find_next_btn = QToolButton()
        self.find_next_btn.setText(">")
        self.find_next_btn.setToolTip("Next match")
        self.find_next_btn.setFixedSize(28, 24)
        self.find_next_btn.clicked.connect(self._activate_next_find_match)
        layout.addWidget(self.find_next_btn)

        self.find_close_btn = QToolButton()
        self.find_close_btn.setText("x")
        self.find_close_btn.setToolTip("Close find")
        self.find_close_btn.setFixedSize(24, 24)
        self.find_close_btn.clicked.connect(self._hide_find_bar)
        layout.addWidget(self.find_close_btn)

        self._set_find_navigation_enabled(False)
        return bar

    def _build_pdf_preview_panel(self) -> QWidget:
        panel = QFrame()
        panel.setObjectName("medRecordPdfPreviewPanel")
        panel.setFrameShape(QFrame.Shape.NoFrame)
        panel.setStyleSheet(
            f"QFrame#medRecordPdfPreviewPanel {{ border-left: 1px solid {theme.BORDER_LIGHT}; }}"
        )

        layout = QVBoxLayout(panel)
        layout.setContentsMargins(theme.SPACE_SM, 0, 0, 0)
        layout.setSpacing(theme.SPACE_SM)

        header = QHBoxLayout()
        header.setContentsMargins(0, 0, 0, 0)
        header.addWidget(theme.section_header("PDF Preview"))
        header.addStretch(1)

        self.collapse_pdf_preview_btn = QToolButton()
        self.collapse_pdf_preview_btn.setText("▶")
        self.collapse_pdf_preview_btn.setToolTip("Collapse PDF preview")
        self.collapse_pdf_preview_btn.setFixedSize(24, 24)
        self.collapse_pdf_preview_btn.setStyleSheet(
            "QToolButton { border: none; font-size: 12px; }"
        )
        self.collapse_pdf_preview_btn.clicked.connect(self.toggle_pdf_preview_collapse)
        header.addWidget(self.collapse_pdf_preview_btn)
        layout.addLayout(header)

        self.pdf_preview_status_label = theme.caption(
            "Double-click a chronology entry to preview its source PDF."
        )
        self.pdf_preview_status_label.setWordWrap(True)
        layout.addWidget(self.pdf_preview_status_label)

        self.pdf_preview_placeholder = QLabel("No PDF selected")
        self.pdf_preview_placeholder.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.pdf_preview_placeholder.setWordWrap(True)
        self.pdf_preview_placeholder.setStyleSheet(
            f"color: {theme.TEXT_FAINT}; border: 1px solid {theme.BORDER_LIGHT};"
            f" border-radius: {theme.RADIUS_SM}px; background-color: #FFFFFF;"
        )
        layout.addWidget(self.pdf_preview_placeholder, 1)

        self.pdf_preview_body_layout = layout
        return panel

    def _build_synopsis_pane(self) -> QWidget:
        pane = QFrame()
        pane.setFrameShape(QFrame.Shape.NoFrame)
        layout = QVBoxLayout(pane)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        self.synopsis_panel = BriefSynopsisPanel()
        self.synopsis_panel.load_paragraphs(self.document.synopsis_paragraphs)
        self.synopsis_panel.paragraph_toggled.connect(self._on_paragraph_toggled)
        self.synopsis_panel.paragraph_open_requested.connect(self._open_pdf_for_paragraph)
        layout.addWidget(self.synopsis_panel, 1)
        return pane

    def _build_table_pane(self) -> QWidget:
        pane = QFrame()
        pane.setFrameShape(QFrame.Shape.NoFrame)
        layout = QVBoxLayout(pane)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        self.table_panel = ChronologyTablePanel()
        self.table_panel.load_rows(self.document.rows)
        self.table_panel.row_toggled.connect(self._on_row_toggled)
        self.table_panel.row_open_requested.connect(self._open_pdf_for_row_id)
        layout.addWidget(self.table_panel, 1)
        return pane

    def eventFilter(self, obj, event) -> bool:
        if (
            obj is getattr(self, "find_input", None)
            and event.type() == QEvent.Type.KeyPress
        ):
            if event.key() == Qt.Key.Key_Escape:
                self._hide_find_bar()
                return True
            if event.key() in (Qt.Key.Key_Return, Qt.Key.Key_Enter):
                is_shift = bool(
                    event.modifiers() & Qt.KeyboardModifier.ShiftModifier
                )
                self._activate_next_find_match(backward=is_shift)
                return True
        return super().eventFilter(obj, event)

    def keyPressEvent(self, event) -> None:
        if (
            event.key() == Qt.Key.Key_F
            and event.modifiers() & Qt.KeyboardModifier.ControlModifier
        ):
            self._show_find_bar()
            event.accept()
            return
        super().keyPressEvent(event)

    def _show_find_bar(self) -> None:
        self.find_bar.setVisible(True)
        self.window().activateWindow()
        self.find_input.setFocus(Qt.FocusReason.ShortcutFocusReason)
        self.find_input.selectAll()
        self._refresh_find_matches(reset=True)

    def _hide_find_bar(self) -> None:
        self.find_bar.setVisible(False)
        self._clear_find_highlights()
        self.tab_widget.setFocus()

    def _on_find_text_changed(self, _query: str) -> None:
        self._refresh_find_matches(reset=True)

    def _on_find_scope_changed(self, _index: int) -> None:
        if self.find_bar.isVisible():
            self._refresh_find_matches(reset=True)

    def _refresh_find_matches(self, *, reset: bool) -> None:
        query = _normalize_search_text(self.find_input.text())
        self._find_matches = self._collect_find_matches(query)
        if not query:
            self._find_index = -1
            self.find_status_label.setText("")
            self._set_find_navigation_enabled(False)
            self._clear_find_highlights()
            return
        if not self._find_matches:
            self._find_index = -1
            self.find_status_label.setText("No matches")
            self._set_find_navigation_enabled(False)
            self._clear_find_highlights()
            return
        if (
            reset
            or self._find_index < 0
            or self._find_index >= len(self._find_matches)
        ):
            self._find_index = 0
        self._activate_find_match(self._find_matches[self._find_index])
        self._update_find_status()

    def _collect_find_matches(self, query: str) -> list[_SearchMatch]:
        if not query:
            return []
        if self.tab_widget.currentIndex() == 0:
            matches: list[_SearchMatch] = []
            for index in range(self.synopsis_panel.count()):
                item_text = self.synopsis_panel.item(index).text()
                if query in _normalize_search_text(item_text):
                    matches.append(_SearchMatch(tab_index=0, row=index))
            return matches

        matches: list[_SearchMatch] = []
        for row in range(self.table_panel.rowCount()):
            matching_column = -1
            row_parts: list[str] = []
            for column in range(1, self.table_panel.columnCount()):
                item = self.table_panel.item(row, column)
                text = item.text() if item is not None else ""
                row_parts.append(text)
                if matching_column < 0 and query in _normalize_search_text(text):
                    matching_column = column
            if query in _normalize_search_text("\n".join(row_parts)):
                matches.append(
                    _SearchMatch(
                        tab_index=1,
                        row=row,
                        column=matching_column if matching_column >= 0 else 1,
                    )
                )
        return matches

    def _activate_next_find_match(self, *, backward: bool = False) -> None:
        if not self.find_bar.isVisible():
            self._show_find_bar()
            return
        if not self._find_matches:
            self._refresh_find_matches(reset=True)
            return
        step = -1 if backward else 1
        self._find_index = (self._find_index + step) % len(self._find_matches)
        self._activate_find_match(self._find_matches[self._find_index])
        self._update_find_status()

    def _activate_find_match(self, match: _SearchMatch) -> None:
        self._apply_find_highlights(match)
        selection_flags = QItemSelectionModel.SelectionFlag.NoUpdate
        if match.tab_index == 0:
            model_index = self.synopsis_panel.model().index(match.row, 0)
            self.synopsis_panel.selectionModel().setCurrentIndex(
                model_index,
                selection_flags,
            )
            self.synopsis_panel.scrollTo(
                model_index,
                QAbstractItemView.ScrollHint.PositionAtCenter,
            )
            return

        model_index = self.table_panel.model().index(match.row, match.column)
        self.table_panel.selectionModel().setCurrentIndex(
            model_index,
            selection_flags,
        )
        self.table_panel.scrollTo(
            model_index,
            QAbstractItemView.ScrollHint.PositionAtCenter,
        )

    def _apply_find_highlights(self, active_match: _SearchMatch) -> None:
        query = _normalize_search_text(self.find_input.text())
        if not query:
            self._clear_find_highlights()
            return
        if active_match.tab_index == 0:
            self.synopsis_panel.set_find_highlight(
                query,
                {match.row for match in self._find_matches if match.tab_index == 0},
                active_match.row,
            )
            self.table_panel.set_find_highlight("", set(), -1)
            return
        self.synopsis_panel.set_find_highlight("", set(), -1)
        self.table_panel.set_find_highlight(
            query,
            {match.row for match in self._find_matches if match.tab_index == 1},
            active_match.row,
        )

    def _clear_find_highlights(self) -> None:
        self.synopsis_panel.set_find_highlight("", set(), -1)
        self.table_panel.set_find_highlight("", set(), -1)

    def _update_find_status(self) -> None:
        self.find_status_label.setText(
            f"{self._find_index + 1} of {len(self._find_matches)}"
        )
        self._set_find_navigation_enabled(len(self._find_matches) > 1)

    def _set_find_navigation_enabled(self, enabled: bool) -> None:
        self.find_previous_btn.setEnabled(enabled)
        self.find_next_btn.setEnabled(enabled)

    def _on_paragraph_toggled(self, paragraph_id: str, checked: bool) -> None:
        if checked:
            self.selection.select_paragraph(paragraph_id)
            result = match_synopsis_to_rows(self._paragraphs[paragraph_id], self.document.rows)
            self._apply_match(paragraph_id, result)
        else:
            self.selection.deselect_paragraph(paragraph_id)
            self._paragraph_match_status.pop(paragraph_id, None)

        self._sync_table_checks()
        self._refresh_match_status()
        self._refresh_extract_state()

    def _apply_match(self, paragraph_id: str, result: MatchResult) -> None:
        if result.status == "confident":
            rejected: list[str] = []
            for row_id in result.row_ids:
                row = self._rows.get(row_id)
                if row is None:
                    continue
                if not row.extractable:
                    rejected.append(_row_label(row))
                    continue
                self.selection.select_row(row_id, source=paragraph_id)
            if rejected:
                reason = "Matched row is not extractable: " + "; ".join(rejected)
                self._paragraph_match_status[paragraph_id] = reason
                self.synopsis_panel.mark_warning(paragraph_id, reason)
            else:
                self._paragraph_match_status.pop(paragraph_id, None)
            return

        reason = result.reason or "No confident chronology row match."
        self.synopsis_panel.mark_warning(paragraph_id, reason)
        self._paragraph_match_status[paragraph_id] = (
            f"No row auto-selected for selected synopsis: {reason}"
        )

    def _on_row_toggled(self, row_id: str, checked: bool) -> None:
        if checked:
            self.selection.select_row(row_id, source="manual")
        else:
            self.selection.deselect_row(row_id, source="manual")
        self._sync_table_checks()
        self._refresh_extract_state()

    def _sync_table_checks(self) -> None:
        for row_id in self._rows:
            self.table_panel.set_row_checked(
                row_id,
                self.selection.is_row_selected(row_id),
                emit=False,
            )

    def _set_match_status(self, message: str) -> None:
        self.match_status_label.setText(message)
        self.match_status_label.setVisible(bool(message))

    def _refresh_match_status(self) -> None:
        for paragraph in self.document.synopsis_paragraphs:
            if paragraph.id in self.selection.selected_paragraph_ids:
                message = self._paragraph_match_status.get(paragraph.id)
                if message:
                    self._set_match_status(message)
                    return
        self._set_match_status("")

    def _selected_rows(self) -> list[SelectableChronologyRow]:
        return [
            self._rows[row_id]
            for row_id in self.selection.selected_row_ids()
            if row_id in self._rows and self._rows[row_id].extractable
        ]

    def _refresh_extract_state(self) -> None:
        count = len(self._selected_rows())
        noun = "row" if count == 1 else "rows"
        self.selected_count_label.setText(f"{count} {noun} selected")
        self.extract_btn.setEnabled(count > 0 and not self.document.blocking_errors)

    def _emit_run(self) -> None:
        rows = self._selected_rows()
        if not rows or self.document.blocking_errors:
            QMessageBox.information(
                self,
                "No extractable rows",
                "Select at least one extractable chronology row.",
            )
            return

        self.run_requested.emit({
            "chronology_path": self.chronology_path,
            "selected_rows": rows,
        })

    def _open_pdf_for_paragraph(self, paragraph_id: str) -> None:
        paragraph = self._paragraphs.get(paragraph_id)
        if paragraph is None:
            return
        result = match_synopsis_to_rows(paragraph, self.document.rows)
        if result.status != "confident" or not result.row_ids:
            reason = result.reason or "No confident chronology row match."
            QMessageBox.warning(
                self,
                "Could not open PDF",
                f"Could not identify one source PDF for this synopsis entry: {reason}",
            )
            return
        self._open_pdf_for_row_id(result.row_ids[0])

    def _open_pdf_for_row_id(self, row_id: str) -> None:
        row = self._rows.get(row_id)
        if row is not None:
            self._open_pdf_for_row(row)

    def _open_pdf_for_row(self, row: SelectableChronologyRow) -> None:
        if not row.record_filename or row.page_start <= 0:
            QMessageBox.warning(
                self,
                "Could not open PDF",
                row.warning or f"Could not parse source record/pages from: {row.page_no}",
            )
            return
        pdf_path = _lookup_file(self._get_file_index(), row.record_filename)
        if not pdf_path:
            pdf_path = _lookup_file(
                self._get_file_index(refresh=True),
                row.record_filename,
            )
        if not pdf_path:
            QMessageBox.warning(
                self,
                "PDF not found",
                f"Could not find source PDF: {row.record_filename}",
            )
            return
        self._show_pdf_preview(pdf_path, row.page_start)

    def _get_file_index(self, *, refresh: bool = False) -> dict[str, str]:
        if refresh or self._file_index is None:
            self._file_index = _build_file_index(self.case_path)
        return self._file_index

    def _show_pdf_preview(self, pdf_path: str, page_number: int) -> None:
        page_number = max(1, int(page_number or 1))
        if self._pdf_preview_collapsed:
            self.toggle_pdf_preview_collapse()

        viewer = self._ensure_pdf_preview_viewer()
        self._pdf_preview_target_page = page_number
        self._pdf_preview_navigation_attempts = 0
        self.pdf_preview_status_label.setText(
            f"{os.path.basename(pdf_path)} - page {page_number}"
        )
        self.pdf_preview_placeholder.setVisible(False)
        viewer.load_pdf(pdf_path)
        self._go_to_pdf_preview_target_page()

    def _ensure_pdf_preview_viewer(self) -> QWidget:
        if self.pdf_preview_viewer is None:
            from icharlotte_core.ui.pdf_viewer_widget import PdfViewerWidget

            self.pdf_preview_viewer = PdfViewerWidget()
            self.pdf_preview_body_layout.addWidget(self.pdf_preview_viewer, 1)
        return self.pdf_preview_viewer

    def _go_to_pdf_preview_target_page(self) -> None:
        if self.pdf_preview_viewer is None:
            return
        self._pdf_preview_navigation_attempts += 1
        self.pdf_preview_viewer.go_to_page(self._pdf_preview_target_page)
        if (
            self.pdf_preview_viewer.get_current_page() != self._pdf_preview_target_page
            and self._pdf_preview_navigation_attempts < 10
        ):
            QTimer.singleShot(500, self._go_to_pdf_preview_target_page)

    def toggle_pdf_preview_collapse(self) -> None:
        sizes = self.preview_splitter.sizes()
        if self._pdf_preview_collapsed:
            self.pdf_preview_panel.setVisible(True)
            total_width = max(sum(sizes), self.width(), 1)
            preview_width = min(self._pdf_preview_width, max(280, total_width // 2))
            self.preview_splitter.setSizes([
                max(320, total_width - preview_width),
                preview_width,
            ])
            self.pdf_preview_toggle_btn.setText("Hide Preview")
            self.collapse_pdf_preview_btn.setText("▶")
            self.collapse_pdf_preview_btn.setToolTip("Collapse PDF preview")
            self._pdf_preview_collapsed = False
            return

        if len(sizes) > 1 and sizes[1] > 0:
            self._pdf_preview_width = sizes[1]
        self.pdf_preview_panel.setVisible(False)
        self.preview_splitter.setSizes([max(sum(sizes), 1), 0])
        self.pdf_preview_toggle_btn.setText("Show Preview")
        self._pdf_preview_collapsed = True

    def _open_original(self) -> None:
        os.startfile(self.chronology_path)


def _row_label(row: SelectableChronologyRow) -> str:
    return f"{row.date} {row.provider}".strip()


def _coerce_column_widths(value, expected_count: int) -> list[int] | None:
    if value is None:
        return None
    if isinstance(value, str):
        parts = [part.strip() for part in value.split(",")]
    elif isinstance(value, (list, tuple)):
        parts = list(value)
    else:
        return None
    widths: list[int] = []
    for part in parts:
        try:
            width = int(part)
        except (TypeError, ValueError):
            return None
        if width < 24:
            return None
        widths.append(width)
    if len(widths) != expected_count:
        return None
    return widths


def _saved_sources(value) -> list[str]:
    if isinstance(value, str):
        return [value]
    if isinstance(value, (list, tuple, set)):
        return [str(source) for source in value if source]
    return []


def _status_message(document: ChronologyDocument) -> str:
    messages = [*document.blocking_errors, *document.warnings]
    return "\n".join(message for message in messages if message)
