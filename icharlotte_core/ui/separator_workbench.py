"""SeparatorWorkbench — reusable single-PDF document-separation workbench.

Extracted from IndexTab so both Advanced Mode (IndexTab) and Wizard Mode
(SeparateTaskTab) share one implementation. Owns the editable document table,
PDF preview, Mark Start/End range marking, sensitivity slider + Re-analyze,
add/delete rows, and Process (split + merge into PULLED-<source>/).

Decoupled from the main window: Re-analyze emits ``reanalyze_requested(int)``;
the host wires it to whatever re-runs the analysis. Processing emits
``processing_complete(dict)`` with {"created": [...], "errors": [...],
"output_folder": str} so the host decides how to surface the result.
"""
import os

import pypdf
from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QAction
from PySide6.QtWidgets import (
    QHBoxLayout, QHeaderView, QInputDialog, QLabel, QLineEdit, QMenu,
    QMessageBox, QPushButton, QSlider, QSplitter, QTableWidget,
    QTableWidgetItem, QVBoxLayout, QWidget,
)

from ..utils import sanitize_filename, format_date_to_mm_dd_yyyy
from .pdf_viewer_widget import PdfViewerWidget


class SeparatorWorkbench(QWidget):
    reanalyze_requested = Signal(int)        # sensitivity 1..3
    processing_complete = Signal(dict)       # {"created", "errors", "output_folder"}

    def __init__(self, parent=None):
        super().__init__(parent)
        self.current_pdf_path = None
        self.marked_start_page = None
        self._setup_ui()

    def _setup_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        self.splitter = QSplitter(Qt.Orientation.Horizontal)
        root.addWidget(self.splitter)

        # ---- Left: table + controls ----
        middle_widget = QWidget()
        middle_layout = QVBoxLayout(middle_widget)
        middle_layout.setContentsMargins(0, 0, 0, 0)

        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Search by Date or Title...")
        self.search_input.textChanged.connect(self.filter_documents)
        middle_layout.addWidget(self.search_input)

        sens_layout = QHBoxLayout()
        sens_layout.addWidget(QLabel("Separation:"))
        broad = QLabel("Broad"); broad.setStyleSheet("color:#666;font-size:11px;")
        sens_layout.addWidget(broad)
        self.sensitivity_slider = QSlider(Qt.Orientation.Horizontal)
        self.sensitivity_slider.setMinimum(1)
        self.sensitivity_slider.setMaximum(3)
        self.sensitivity_slider.setValue(2)
        self.sensitivity_slider.setTickPosition(QSlider.TickPosition.TicksBelow)
        self.sensitivity_slider.setTickInterval(1)
        self.sensitivity_slider.setPageStep(1)
        self.sensitivity_slider.setFixedWidth(100)
        sens_layout.addWidget(self.sensitivity_slider)
        fine = QLabel("Fine"); fine.setStyleSheet("color:#666;font-size:11px;")
        sens_layout.addWidget(fine)
        self.reanalyze_btn = QPushButton("Re-analyze")
        self.reanalyze_btn.setStyleSheet(
            "background-color:#2196F3;color:white;font-weight:bold;padding:6px 12px;")
        self.reanalyze_btn.clicked.connect(self.on_reanalyze_clicked)
        sens_layout.addWidget(self.reanalyze_btn)
        sens_layout.addStretch()
        middle_layout.addLayout(sens_layout)

        self.doc_table = QTableWidget()
        self.doc_table.setColumnCount(6)
        self.doc_table.setHorizontalHeaderLabels(
            ["Sep.", "Merge Group", "ID", "Date", "Pages", "Title"])
        hdr = self.doc_table.horizontalHeader()
        hdr.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        hdr.setSectionResizeMode(1, QHeaderView.ResizeMode.Interactive)
        self.doc_table.setColumnWidth(1, 100)
        hdr.setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        hdr.setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        hdr.setSectionResizeMode(4, QHeaderView.ResizeMode.ResizeToContents)
        hdr.setSectionResizeMode(5, QHeaderView.ResizeMode.Stretch)
        self.doc_table.setSortingEnabled(True)
        self.doc_table.cellClicked.connect(self.on_doc_clicked)
        self.doc_table.cellDoubleClicked.connect(self.on_doc_double_clicked)
        self.doc_table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.doc_table.customContextMenuRequested.connect(self.show_context_menu)
        middle_layout.addWidget(self.doc_table)

        btn_layout = QHBoxLayout()
        self.add_doc_btn = QPushButton("Add Document")
        self.add_doc_btn.setStyleSheet(
            "background-color:#FF9800;color:white;font-weight:bold;padding:10px;")
        self.add_doc_btn.clicked.connect(self.add_document_row)
        btn_layout.addWidget(self.add_doc_btn)
        self.delete_doc_btn = QPushButton("Delete Selected")
        self.delete_doc_btn.setStyleSheet(
            "background-color:#f44336;color:white;font-weight:bold;padding:10px;")
        self.delete_doc_btn.clicked.connect(self.delete_selected_rows)
        btn_layout.addWidget(self.delete_doc_btn)
        btn_layout.addStretch()
        self.process_btn = QPushButton("Process Documents")
        self.process_btn.setStyleSheet(
            "background-color:#4CAF50;color:white;font-weight:bold;padding:10px;")
        self.process_btn.clicked.connect(self.process_documents)
        btn_layout.addWidget(self.process_btn)
        middle_layout.addLayout(btn_layout)
        self.splitter.addWidget(middle_widget)

        # ---- Right: PDF preview + mark controls ----
        preview_widget = QWidget()
        preview_layout = QVBoxLayout(preview_widget)
        preview_layout.setContentsMargins(0, 0, 0, 0)
        self.pdf_viewer = PdfViewerWidget()
        preview_layout.addWidget(self.pdf_viewer)
        mark_layout = QHBoxLayout()
        self.mark_status = QLabel("Range: Not set")
        self.mark_status.setStyleSheet("font-weight:bold;padding:5px;")
        mark_layout.addWidget(self.mark_status)
        self.mark_start_btn = QPushButton("Mark Start")
        self.mark_start_btn.setStyleSheet("background-color:#2196F3;color:white;padding:8px;")
        self.mark_start_btn.clicked.connect(self.mark_start_page)
        mark_layout.addWidget(self.mark_start_btn)
        self.mark_end_btn = QPushButton("Mark End && Add")
        self.mark_end_btn.setStyleSheet("background-color:#4CAF50;color:white;padding:8px;")
        self.mark_end_btn.clicked.connect(self.mark_end_and_add)
        mark_layout.addWidget(self.mark_end_btn)
        self.clear_mark_btn = QPushButton("Clear")
        self.clear_mark_btn.clicked.connect(self.clear_marked_range)
        mark_layout.addWidget(self.clear_mark_btn)
        preview_layout.addLayout(mark_layout)
        self.splitter.addWidget(preview_widget)
        self.splitter.setSizes([500, 500])

    # ---- Public API ----

    def load_docs(self, pdf_path, docs):
        self.current_pdf_path = pdf_path
        if os.path.exists(pdf_path):
            self.pdf_viewer.load_pdf(pdf_path)
        self.clear_marked_range()
        self.doc_table.setSortingEnabled(False)
        self.doc_table.setRowCount(0)
        for doc in docs:
            self._add_doc_to_table(doc)
        self.doc_table.setSortingEnabled(True)
        self.filter_documents(self.search_input.text())

    def set_busy(self, busy):
        self.reanalyze_btn.setEnabled(not busy)
        self.sensitivity_slider.setEnabled(not busy)

    def on_reanalyze_clicked(self):
        if not self.current_pdf_path:
            return
        self.set_busy(True)
        self.reanalyze_requested.emit(self.sensitivity_slider.value())

    # ==== Moved verbatim from IndexTab (tabs.py) ====

    def filter_documents(self, text):
        text = text.lower()
        for row in range(self.doc_table.rowCount()):
            date_item = self.doc_table.item(row, 3)
            title_widget = self.doc_table.cellWidget(row, 5)  # Title is now a QLineEdit

            date_text = date_item.text().lower() if date_item else ""
            title_text = title_widget.text().lower() if isinstance(title_widget, QLineEdit) else ""

            if text in date_text or text in title_text:
                self.doc_table.setRowHidden(row, False)
            else:
                self.doc_table.setRowHidden(row, True)

    def _add_doc_to_table(self, doc, check_sep=False):
        """Helper to add a document row to the table. Used both for loading and adding new docs."""
        from .tabs import DateTableWidgetItem
        row = self.doc_table.rowCount()
        self.doc_table.insertRow(row)

        # Col 0: Sep. Checkbox
        chk_item = QTableWidgetItem()
        chk_item.setFlags(Qt.ItemFlag.ItemIsUserCheckable | Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable)
        chk_item.setCheckState(Qt.CheckState.Checked if check_sep else Qt.CheckState.Unchecked)
        self.doc_table.setItem(row, 0, chk_item)

        # Col 1: Merge Group (LineEdit)
        merge_edit = QLineEdit()
        merge_edit.setPlaceholderText("")
        self.doc_table.setCellWidget(row, 1, merge_edit)

        # Col 2: ID (read-only)
        id_item = QTableWidgetItem(str(doc.get('id', '')))
        id_item.setFlags(id_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
        self.doc_table.setItem(row, 2, id_item)

        # Col 3: Date (read-only display, but stored)
        date_val = doc.get('date', '')
        formatted_date = format_date_to_mm_dd_yyyy(date_val)
        date_item = DateTableWidgetItem(formatted_date)
        date_item.setFlags(date_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
        self.doc_table.setItem(row, 3, date_item)

        # Col 4: Pages (EDITABLE - use LineEdit for better UX)
        start = doc.get('start', '')
        end = doc.get('end', '')
        pages = f"{start}-{end}" if start != end else str(start)
        pages_edit = QLineEdit(pages)
        pages_edit.setPlaceholderText("e.g., 5-7")
        pages_edit.setToolTip("Edit page range (format: start-end or single page)")
        self.doc_table.setCellWidget(row, 4, pages_edit)

        # Col 5: Title (EDITABLE - use LineEdit)
        title_edit = QLineEdit(str(doc.get('title', '')))
        title_edit.setPlaceholderText("Document title")
        title_edit.setToolTip("Edit document title")
        title_edit.setCursorPosition(0)  # Show beginning of text, not end
        self.doc_table.setCellWidget(row, 5, title_edit)

        return row

    def delete_selected_rows(self):
        """Delete the selected document rows."""
        selected_rows = set()
        for range_ in self.doc_table.selectedRanges():
            for r in range(range_.topRow(), range_.bottomRow() + 1):
                selected_rows.add(r)

        if not selected_rows:
            QMessageBox.information(self, "Info", "No rows selected.")
            return

        reply = QMessageBox.question(
            self, "Confirm Delete",
            f"Delete {len(selected_rows)} selected document(s)?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )

        if reply == QMessageBox.StandardButton.Yes:
            # Delete in reverse order to maintain correct indices
            for row in sorted(selected_rows, reverse=True):
                self.doc_table.removeRow(row)

    def show_context_menu(self, position):
        menu = QMenu()

        add_doc_action = QAction("Add New Document", self)
        add_doc_action.triggered.connect(self.add_document_row)
        menu.addAction(add_doc_action)

        delete_action = QAction("Delete Selected Row(s)", self)
        delete_action.triggered.connect(self.delete_selected_rows)
        menu.addAction(delete_action)

        menu.addSeparator()

        set_group_action = QAction("Set Merge Group for Selected", self)
        set_group_action.triggered.connect(self.set_merge_group_batch)
        menu.addAction(set_group_action)

        clear_group_action = QAction("Clear Merge Group for Selected", self)
        clear_group_action.triggered.connect(self.clear_merge_group_batch)
        menu.addAction(clear_group_action)

        menu.exec(self.doc_table.viewport().mapToGlobal(position))

    def set_merge_group_batch(self):
        selected_rows = set()
        for range_ in self.doc_table.selectedRanges():
            for r in range(range_.topRow(), range_.bottomRow() + 1):
                if not self.doc_table.isRowHidden(r):
                    selected_rows.add(r)

        if not selected_rows:
            return

        group_name, ok = QInputDialog.getText(self, "Set Merge Group", "Enter Merge Group Name:")
        if ok:
            for row in selected_rows:
                widget = self.doc_table.cellWidget(row, 1)
                if isinstance(widget, QLineEdit):
                    widget.setText(group_name)

    def clear_merge_group_batch(self):
        selected_rows = set()
        for range_ in self.doc_table.selectedRanges():
            for r in range(range_.topRow(), range_.bottomRow() + 1):
                if not self.doc_table.isRowHidden(r):
                    selected_rows.add(r)

        for row in selected_rows:
            widget = self.doc_table.cellWidget(row, 1)
            if isinstance(widget, QLineEdit):
                widget.setText("")

    def _parse_pages(self, pages_str):
        """Parse a pages string like '5-7' or '8' into (start, end) tuple."""
        pages_str = pages_str.strip()
        if '-' in pages_str:
            parts = pages_str.split('-')
            try:
                start = int(parts[0].strip())
                end = int(parts[1].strip())
                return start, end
            except (ValueError, IndexError):
                return None, None
        else:
            try:
                page = int(pages_str)
                return page, page
            except ValueError:
                return None, None

    def on_doc_clicked(self, row, column):
        """Single click: Navigate to first page of document."""
        if not hasattr(self, 'pdf_viewer'):
            return
        pages_widget = self.doc_table.cellWidget(row, 4)
        if pages_widget:
            start, _ = self._parse_pages(pages_widget.text())
            if start:
                self.pdf_viewer.go_to_page(start)

    def on_doc_double_clicked(self, row, column):
        """Double click: Navigate to last page of document."""
        if not hasattr(self, 'pdf_viewer'):
            return
        pages_widget = self.doc_table.cellWidget(row, 4)
        if pages_widget:
            _, end = self._parse_pages(pages_widget.text())
            if end:
                self.pdf_viewer.go_to_page(end)

    def mark_start_page(self):
        """Mark current page as start of new document range."""
        if not hasattr(self, 'pdf_viewer'):
            return
        page = self.pdf_viewer.get_current_page()
        if page:
            self.marked_start_page = page
            self.mark_status.setText(f"Range: {page} - ?")
            self.mark_status.setStyleSheet("font-weight: bold; padding: 5px; background-color: #FFF3E0;")

    def mark_end_and_add(self):
        """Mark current page as end and add new document."""
        if not hasattr(self, 'pdf_viewer'):
            return

        if not self.marked_start_page:
            QMessageBox.warning(self, "Warning", "Please mark a start page first.")
            return

        end_page = self.pdf_viewer.get_current_page()
        if not end_page:
            QMessageBox.warning(self, "Warning", "Could not get current page.")
            return

        # Swap if end is before start
        start_page = self.marked_start_page
        if end_page < start_page:
            start_page, end_page = end_page, start_page

        # Create new document with marked range
        new_doc = {
            'id': str(self._get_next_doc_id()),
            'title': 'New Document',
            'date': '',
            'start': start_page,
            'end': end_page
        }

        self.doc_table.setSortingEnabled(False)
        row = self._add_doc_to_table(new_doc, check_sep=True)
        self.doc_table.setSortingEnabled(True)
        self.doc_table.selectRow(row)
        self.doc_table.scrollToItem(self.doc_table.item(row, 0))

        # Focus on title for immediate editing
        title_widget = self.doc_table.cellWidget(row, 5)
        if title_widget:
            title_widget.setFocus()
            title_widget.selectAll()

        # Reset marked range
        self.clear_marked_range()

    def clear_marked_range(self):
        """Clear the marked page range."""
        self.marked_start_page = None
        if hasattr(self, 'mark_status'):
            self.mark_status.setText("Range: Not set")
            self.mark_status.setStyleSheet("font-weight: bold; padding: 5px;")

    def _get_next_doc_id(self):
        """Get next available document ID."""
        max_id = 0
        for row in range(self.doc_table.rowCount()):
            id_item = self.doc_table.item(row, 2)
            if id_item:
                try:
                    max_id = max(max_id, int(id_item.text()))
                except ValueError:
                    pass
        return max_id + 1

    def _get_doc_from_row(self, row):
        """Extract document data from a table row, reading from the editable widgets."""
        doc_id = self.doc_table.item(row, 2).text() if self.doc_table.item(row, 2) else ""

        # Get pages from widget
        pages_widget = self.doc_table.cellWidget(row, 4)
        pages_str = pages_widget.text() if isinstance(pages_widget, QLineEdit) else ""
        start, end = self._parse_pages(pages_str)

        # Get title from widget
        title_widget = self.doc_table.cellWidget(row, 5)
        title = title_widget.text() if isinstance(title_widget, QLineEdit) else ""

        # Get date from item
        date_item = self.doc_table.item(row, 3)
        date = date_item.text() if date_item else ""

        return {
            'id': doc_id,
            'title': title,
            'date': date,
            'start': start,
            'end': end
        }

    # ==== Custom (decoupled) versions ====

    def add_document_row(self):
        if not self.current_pdf_path:
            QMessageBox.warning(self, "Warning", "No PDF loaded.")
            return
        new_doc = {"id": str(self._get_next_doc_id()), "title": "New Document",
                   "date": "", "start": "", "end": ""}
        self.doc_table.setSortingEnabled(False)
        row = self._add_doc_to_table(new_doc, check_sep=True)
        self.doc_table.setSortingEnabled(True)
        self.doc_table.selectRow(row)
        self.doc_table.scrollToItem(self.doc_table.item(row, 0))
        pages_widget = self.doc_table.cellWidget(row, 4)
        if pages_widget:
            pages_widget.setFocus()
            pages_widget.selectAll()

    def process_documents(self):
        pdf_path = self.current_pdf_path
        if not pdf_path or not os.path.exists(pdf_path):
            QMessageBox.warning(self, "Error", f"Source PDF not found: {pdf_path}")
            return
        if self.doc_table.rowCount() == 0:
            QMessageBox.information(self, "Info", "No documents in the table.")
            return

        separate_tasks, merge_groups, validation_errors = [], {}, []
        for row in range(self.doc_table.rowCount()):
            if self.doc_table.isRowHidden(row):
                continue
            doc_obj = self._get_doc_from_row(row)
            is_sep = self.doc_table.item(row, 0).checkState() == Qt.CheckState.Checked
            mw = self.doc_table.cellWidget(row, 1)
            group_name = mw.text().strip() if isinstance(mw, QLineEdit) else ""
            if not is_sep and not group_name:
                continue
            if doc_obj["start"] is None or doc_obj["end"] is None:
                validation_errors.append(f"Row {row + 1} (ID {doc_obj['id']}): Invalid page range")
                continue
            if not doc_obj["title"].strip():
                validation_errors.append(f"Row {row + 1} (ID {doc_obj['id']}): Title is empty")
                continue
            if is_sep:
                separate_tasks.append(doc_obj)
            if group_name:
                merge_groups.setdefault(group_name, []).append(doc_obj)

        if validation_errors:
            QMessageBox.warning(self, "Validation Errors",
                                "The following rows have issues:\n\n" + "\n".join(validation_errors))
        if not separate_tasks and not merge_groups:
            QMessageBox.information(self, "Info",
                                    "No actions selected (Check 'Sep.' or enter a 'Merge Group').")
            return

        base_dir = os.path.dirname(pdf_path)
        source_name = os.path.splitext(os.path.basename(pdf_path))[0]
        output_folder = os.path.join(base_dir, f"PULLED-{source_name}")
        os.makedirs(output_folder, exist_ok=True)

        created, errors = [], list(validation_errors)
        try:
            reader = pypdf.PdfReader(pdf_path)
            for doc in separate_tasks:
                try:
                    writer = pypdf.PdfWriter()
                    start = max(0, int(doc["start"]) - 1)
                    end = min(int(doc["end"]), len(reader.pages))
                    for i in range(start, end):
                        writer.add_page(reader.pages[i])
                    safe = sanitize_filename(doc["title"])[:50]
                    out_path = os.path.join(output_folder, f"{doc['id']} - {safe}.pdf")
                    with open(out_path, "wb") as f:
                        writer.write(f)
                    created.append(os.path.basename(out_path))
                except Exception as e:
                    errors.append(f"Failed to separate '{doc['title']}': {e}")
            for group_name, group_docs in merge_groups.items():
                try:
                    writer = pypdf.PdfWriter()
                    for doc in group_docs:
                        start = max(0, int(doc["start"]) - 1)
                        end = min(int(doc["end"]), len(reader.pages))
                        for i in range(start, end):
                            writer.add_page(reader.pages[i])
                    safe = sanitize_filename(group_name)
                    if not safe.lower().endswith(".pdf"):
                        safe += ".pdf"
                    out_path = os.path.join(output_folder, safe)
                    with open(out_path, "wb") as f:
                        writer.write(f)
                    created.append(f"{os.path.basename(out_path)} ({len(group_docs)} docs)")
                except Exception as e:
                    errors.append(f"Failed to merge group '{group_name}': {e}")
        except Exception as e:
            QMessageBox.critical(self, "Critical Error", f"Processing failed: {e}")
            return

        self.processing_complete.emit(
            {"created": created, "errors": errors, "output_folder": output_folder})
