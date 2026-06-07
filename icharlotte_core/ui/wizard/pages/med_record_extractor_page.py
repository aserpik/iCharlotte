"""Structured chronology viewer for Med Record Extractor."""
from __future__ import annotations

import os

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QAbstractItemView,
    QFrame,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QListWidget,
    QListWidgetItem,
    QMessageBox,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QTabWidget,
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
from icharlotte_core.ui.wizard import theme


class BriefSynopsisPanel(QListWidget):
    """Checkable list of synopsis paragraphs from the selected chronology."""

    paragraph_toggled = Signal(str, bool)

    def __init__(self, parent: QWidget | None = None):
        super().__init__(parent)
        self._items_by_id: dict[str, QListWidgetItem] = {}
        self.setSelectionMode(QAbstractItemView.SelectionMode.NoSelection)
        self.itemChanged.connect(self._on_item_changed)

    def load_paragraphs(self, paragraphs: list[SynopsisParagraph]) -> None:
        self.blockSignals(True)
        self.clear()
        self._items_by_id.clear()
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

    def set_paragraph_checked(self, paragraph_id: str, checked: bool) -> None:
        item = self._items_by_id[paragraph_id]
        item.setCheckState(Qt.CheckState.Checked if checked else Qt.CheckState.Unchecked)

    def mark_warning(self, paragraph_id: str, message: str) -> None:
        item = self._items_by_id.get(paragraph_id)
        if item is not None:
            item.setToolTip(message)

    def mousePressEvent(self, event) -> None:
        item = self.itemAt(event.position().toPoint())
        if item is not None:
            checked = item.checkState() == Qt.CheckState.Checked
            item.setCheckState(Qt.CheckState.Unchecked if checked else Qt.CheckState.Checked)
            event.accept()
            return
        super().mousePressEvent(event)

    def _on_item_changed(self, item: QListWidgetItem) -> None:
        paragraph_id = item.data(Qt.ItemDataRole.UserRole)
        self.paragraph_toggled.emit(
            paragraph_id,
            item.checkState() == Qt.CheckState.Checked,
        )


class ChronologyTablePanel(QTableWidget):
    """Checkable chronology table rows."""

    row_toggled = Signal(str, bool)

    def __init__(self, parent: QWidget | None = None):
        super().__init__(parent)
        self._rows_by_id: dict[str, int] = {}
        self._ids_by_row: dict[int, str] = {}
        self._extractable_by_id: dict[str, bool] = {}
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
        self.verticalHeader().setVisible(False)
        self.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        self.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        self.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        self.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        self.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeMode.Stretch)
        self.horizontalHeader().setSectionResizeMode(5, QHeaderView.ResizeMode.Stretch)
        self.cellChanged.connect(self._on_cell_changed)

    def count(self) -> int:
        return self.rowCount()

    def load_rows(self, rows: list[SelectableChronologyRow]) -> None:
        self.blockSignals(True)
        self.setRowCount(len(rows))
        self._rows_by_id.clear()
        self._ids_by_row.clear()
        self._extractable_by_id.clear()
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
        self.resizeRowsToContents()
        self.blockSignals(False)

    def set_row_checked(self, row_id: str, checked: bool, *, emit: bool = True) -> None:
        item = self.item(self._rows_by_id[row_id], 0)
        if item is None:
            return
        if checked and not self._extractable_by_id.get(row_id, False):
            checked = False
        if not emit:
            self.blockSignals(True)
        item.setCheckState(Qt.CheckState.Checked if checked else Qt.CheckState.Unchecked)
        if not emit:
            self.blockSignals(False)

    def is_row_checked(self, row_id: str) -> bool:
        item = self.item(self._rows_by_id[row_id], 0)
        return item is not None and item.checkState() == Qt.CheckState.Checked

    def mousePressEvent(self, event) -> None:
        row_index = self.rowAt(event.position().toPoint().y())
        row_id = self._ids_by_row.get(row_index)
        if row_id is not None:
            if self._extractable_by_id.get(row_id, False):
                self.set_row_checked(row_id, not self.is_row_checked(row_id))
            event.accept()
            return
        super().mousePressEvent(event)

    def _on_cell_changed(self, row: int, column: int) -> None:
        if column != 0:
            return
        item = self.item(row, 0)
        if item is None:
            return
        row_id = item.data(Qt.ItemDataRole.UserRole)
        self.row_toggled.emit(row_id, item.checkState() == Qt.CheckState.Checked)

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
        self.chronology_path = chronology_path
        self.document = self._load_document(chronology_path)
        self.selection = SelectionState()
        self._paragraphs = {paragraph.id: paragraph for paragraph in self.document.synopsis_paragraphs}
        self._rows = {row.id: row for row in self.document.rows}
        self._paragraph_match_status: dict[str, str] = {}

        self._setup_ui()
        self._refresh_extract_state()

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
            theme.SPACE_XL,
            theme.SPACE_XL,
            theme.SPACE_XL,
            theme.SPACE_XL,
        )
        layout.setSpacing(theme.SPACE_MD)

        layout.addWidget(theme.page_title("Select Medical Chronology Records"))
        layout.addWidget(theme.helper_text(
            "Choose synopsis paragraphs or individual chronology rows to extract source records."
        ))

        self.message_label = theme.error_text(_status_message(self.document))
        self.warning_label = self.message_label
        self.message_label.setVisible(bool(self.message_label.text()))
        layout.addWidget(self.message_label)

        self.match_status_label = theme.caption("")
        self.match_status_label.setWordWrap(True)
        self.match_status_label.setVisible(False)
        layout.addWidget(self.match_status_label)

        self.tab_widget = QTabWidget()
        self.tab_widget.addTab(self._build_synopsis_pane(), "Brief Synopsis")
        self.tab_widget.addTab(self._build_table_pane(), "Chronology Rows")
        layout.addWidget(self.tab_widget, 1)

        controls = QHBoxLayout()
        controls.setContentsMargins(0, 0, 0, 0)
        self.selected_count_label = QLabel("")
        controls.addWidget(self.selected_count_label)
        controls.addStretch(1)

        self.open_original_btn = theme.secondary_button("Open Original")
        self.open_original_btn.clicked.connect(self._open_original)
        controls.addWidget(self.open_original_btn)

        self.extract_btn = theme.primary_button("Extract")
        self.extract_btn.clicked.connect(self._emit_run)
        controls.addWidget(self.extract_btn)
        layout.addLayout(controls)

    def _build_synopsis_pane(self) -> QWidget:
        pane = QFrame()
        pane.setFrameShape(QFrame.Shape.NoFrame)
        layout = QVBoxLayout(pane)
        layout.setContentsMargins(0, 0, theme.SPACE_SM, 0)
        layout.setSpacing(theme.SPACE_SM)
        layout.addWidget(theme.section_header("Brief Synopsis"))

        self.synopsis_panel = BriefSynopsisPanel()
        self.synopsis_panel.load_paragraphs(self.document.synopsis_paragraphs)
        self.synopsis_panel.paragraph_toggled.connect(self._on_paragraph_toggled)
        layout.addWidget(self.synopsis_panel, 1)
        return pane

    def _build_table_pane(self) -> QWidget:
        pane = QFrame()
        pane.setFrameShape(QFrame.Shape.NoFrame)
        layout = QVBoxLayout(pane)
        layout.setContentsMargins(theme.SPACE_SM, 0, 0, 0)
        layout.setSpacing(theme.SPACE_SM)
        layout.addWidget(theme.section_header("Chronology Rows"))

        self.table_panel = ChronologyTablePanel()
        self.table_panel.load_rows(self.document.rows)
        self.table_panel.row_toggled.connect(self._on_row_toggled)
        layout.addWidget(self.table_panel, 1)
        return pane

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

    def _open_original(self) -> None:
        os.startfile(self.chronology_path)


def _row_label(row: SelectableChronologyRow) -> str:
    return f"{row.date} {row.provider}".strip()


def _saved_sources(value) -> list[str]:
    if isinstance(value, str):
        return [value]
    if isinstance(value, (list, tuple, set)):
        return [str(source) for source in value if source]
    return []


def _status_message(document: ChronologyDocument) -> str:
    messages = [*document.blocking_errors, *document.warnings]
    return "\n".join(message for message in messages if message)
