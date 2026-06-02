# Separate → Wizard Mode Task Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a full-parity "Separate Documents" task to Wizard Mode that analyzes a PDF, lets the user review/edit identified sub-documents in an interactive workbench, and splits/merges them — reusing a shared workbench widget extracted from advanced mode's `IndexTab`.

**Architecture:** Extract `IndexTab`'s single-PDF workbench (editable table + PDF preview + Mark Start/End + sensitivity/Re-analyze + Process) into a reusable `SeparatorWorkbench(QWidget)`. Advanced mode's `IndexTab` becomes a thin host (left PDF list + embedded workbench). Wizard mode adds a custom `SeparateTaskTab(WizardTaskContainer)` (Settings → Status → Workbench), backed by a `SeparateAnalysisWorker(QThread)` that runs `Scripts/separate.py --headless` and embeds the same `SeparatorWorkbench`.

**Tech Stack:** Python 3.x, PySide6, pytest + pytest-qt, pypdf, python-docx, PyMuPDF (fitz).

---

## Reference facts (verified in codebase)

- `SeparatorWorkbench` host helpers live in `icharlotte_core.utils`: `sanitize_filename`, `format_date_to_mm_dd_yyyy`, `log_event`. `pypdf` imported at module level in tabs.py. `PdfViewerWidget` from `icharlotte_core.ui.pdf_viewer_widget` (API: `load_pdf(path)`, `go_to_page(page_num)`, `get_current_page()`).
- Advanced-mode workbench methods currently in `IndexTab` (icharlotte_core/ui/tabs.py): `setup_ui` (workbench portion 2393–2534), `filter_documents` (2775), `_add_doc_to_table` (2814), `on_reanalyze_clicked` (2860), `add_document_row` (2871), `delete_selected_rows` (2914), `save_table_to_index` (2936 — IndexTab-only, stays), `show_context_menu` (2971), `set_merge_group_batch` (3000), `clear_merge_group_batch` (3017), `_parse_pages` (3029), `on_doc_clicked` (3047), `on_doc_double_clicked` (3057), `mark_start_page` (3067), `mark_end_and_add` (3077), `clear_marked_range` (3120), `_get_next_doc_id` (3127), `_get_doc_from_row` (3139), `process_documents` (3164–3313).
- Wizard dispatch: `iCharlotte._open_task_tab` (iCharlotte.py:1510) calls `builder = getattr(in_process_task_tab, in_process_builder_name); task_tab = builder(spec=, case_path=, file_number=, parent=)`; a `None` return aborts silently. `task_completed` signal is connected and the tab is added.
- `build_respond_to_discovery_tab` (in_process_task_tab.py:382) is the model for a builder that opens its own `QFileDialog` and uses `resolve_default_folder(case_path, spec.default_folders)` from `icharlotte_core.ui.wizard.file_picker`.
- `OpposeMotionTaskTab` (oppose_motion_page.py:1280) is the model for a custom `WizardTaskContainer` with its own worker QThread + analysis worker.
- `StatusPage` API: `reset()`, `on_status(str)`, `progress_bar.setRange(0,0)` (indeterminate), `cancel_requested` signal, `cancel_btn`.
- `WizardTaskContainer.__init__(spec, steps=None, parent=None)`; proxies `addWidget`/`setCurrentIndex`/`currentIndex`. Pass `steps=["Settings", "Analyzing", "Review & Split"]`.
- theme helpers: `theme.primary_button(text)`, `theme.secondary_button(text)`, `theme.page_title(text)`, `theme.helper_text(text)`, `theme.section_header(text)`, `theme.SPACE_XL`, `theme.SPACE_MD`.
- Offline docx-validation model: `validate_discovery_response_docx` (word_validator.py:1189). `ValidationResult(context=...)`, `Finding(severity, rule, message, location=None)`. severities: "ERROR"/"WARN"/"INFO"/"PASS".
- Tests use `pytest.importorskip("pytestqt")` then the `qtbot` fixture. Existing wizard tests: `tests/test_wizard/test_wizard_tab.py`.

---

## Task 1: Add offline index-docx validation

**Files:**
- Modify: `icharlotte_core/word_validator.py` (add function near `validate_discovery_response_docx`, ~line 1189)
- Modify: `Scripts/separate.py` (`create_index_word`, ~line 610)
- Test: `tests/test_word_validator_index.py` (create)

- [ ] **Step 1: Write the failing test**

Create `tests/test_word_validator_index.py`:

```python
"""Tests for validate_index_docx (offline separator-index validation)."""
import pytest

pytest.importorskip("docx")
from docx import Document

from icharlotte_core.word_validator import validate_index_docx


def _make_index(path, n_docs):
    doc = Document()
    doc.add_paragraph("INDEX OF DOCUMENTS - test")
    table = doc.add_table(rows=1, cols=4)
    hdr = table.rows[0].cells
    for i, h in enumerate(["Id", "Document Title", "Document Date", "Page Ranges"]):
        hdr[i].text = h
    for i in range(n_docs):
        cells = table.add_row().cells
        cells[0].text = str(i + 1)
        cells[1].text = f"Doc {i + 1}"
    doc.save(str(path))


def test_missing_file_is_error(tmp_path):
    result = validate_index_docx(str(tmp_path / "nope.docx"), expected_doc_count=3)
    assert result.has_errors


def test_valid_index_passes(tmp_path):
    p = tmp_path / "Index_test.docx"
    _make_index(p, 3)
    result = validate_index_docx(str(p), expected_doc_count=3)
    assert not result.has_errors


def test_row_count_mismatch_warns(tmp_path):
    p = tmp_path / "Index_test.docx"
    _make_index(p, 2)
    result = validate_index_docx(str(p), expected_doc_count=5)
    assert result.has_warnings
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_word_validator_index.py -v`
Expected: FAIL with `ImportError: cannot import name 'validate_index_docx'`.

- [ ] **Step 3: Implement `validate_index_docx`**

In `icharlotte_core/word_validator.py`, add directly above `def validate_discovery_response_docx`:

```python
def validate_index_docx(doc_path: str, expected_doc_count: Optional[int] = None) -> ValidationResult:
    """Lightweight offline validation for the separator INDEX .docx.

    Verifies the file exists, opens with python-docx, contains a table with a
    header row plus at least one data row, and (optionally) that the number of
    data rows matches the number of identified documents.
    """
    from docx import Document

    result = ValidationResult(context=f"Separator index: {os.path.basename(doc_path)}")

    if not os.path.exists(doc_path):
        result.findings.append(Finding("ERROR", "file", f"File not found: {doc_path}"))
        return result

    try:
        doc = Document(doc_path)
    except Exception as e:
        result.findings.append(Finding("ERROR", "open", f"Could not open .docx: {e}"))
        return result

    if not doc.tables:
        result.findings.append(Finding("ERROR", "table", "Index has no table"))
        return result

    table = doc.tables[0]
    data_rows = max(0, len(table.rows) - 1)  # minus header
    if data_rows == 0:
        result.findings.append(Finding("ERROR", "rows", "Index table has no data rows"))
        return result

    if expected_doc_count is not None and data_rows != expected_doc_count:
        result.findings.append(Finding(
            "WARN", "row_count",
            f"Index data rows ({data_rows}) != identified documents ({expected_doc_count})",
            expected=expected_doc_count, actual=data_rows,
        ))
    else:
        result.findings.append(Finding("PASS", "row_count", f"{data_rows} document rows"))

    return result
```

- [ ] **Step 4: Run test to verify it passes**

Run: `python -m pytest tests/test_word_validator_index.py -v`
Expected: 3 passed.

- [ ] **Step 5: Wire validation into separate.py**

In `Scripts/separate.py`, inside `create_index_word`, replace the success block (currently):

```python
        doc.save(docx_path)
        logger.info(f"Index saved to: {docx_path}")
        print(f"\n[SUCCESS] Index created: {docx_filename} in INDEXES folder.")
```

with:

```python
        doc.save(docx_path)
        logger.info(f"Index saved to: {docx_path}")
        print(f"\n[SUCCESS] Index created: {docx_filename} in INDEXES folder.")

        # Mandatory Word-output validation (CLAUDE.md). Non-fatal: log only.
        try:
            from icharlotte_core.word_validator import validate_index_docx
            vres = validate_index_docx(docx_path, expected_doc_count=len(docs))
            if vres.has_errors:
                logger.error(f"Index validation FAILED: {docx_path}")
                vres.print_summary()
            elif vres.has_warnings:
                logger.warning(f"Index validation warnings for: {docx_path}")
        except Exception as ve:
            logger.warning(f"Index validation skipped: {ve}")
```

- [ ] **Step 6: Run test to confirm no regression + commit**

Run: `python -m pytest tests/test_word_validator_index.py tests/test_separate_model_config.py -v`
Expected: all pass (separate model-config tests unaffected).

```bash
git add icharlotte_core/word_validator.py Scripts/separate.py tests/test_word_validator_index.py
git commit -m "feat(separate): validate generated index .docx offline"
```

---

## Task 2: Extract `SeparatorWorkbench` widget

**Files:**
- Create: `icharlotte_core/ui/separator_workbench.py`
- Test: `tests/test_separator_workbench.py` (create)

This widget is the self-contained workbench (everything right of IndexTab's PDF list). It does NOT touch `GEMINI_DATA_DIR` persistence (that stays in IndexTab). It emits `reanalyze_requested(int)` instead of calling the main window, and emits `processing_complete(dict)` after split/merge.

- [ ] **Step 1: Write the failing test**

Create `tests/test_separator_workbench.py`:

```python
"""Tests for the reusable SeparatorWorkbench widget."""
import pytest

pytest.importorskip("pytestqt")
pytest.importorskip("pypdf")
import pypdf

from PySide6.QtCore import Qt

from icharlotte_core.ui.separator_workbench import SeparatorWorkbench


def _make_pdf(path, n_pages):
    writer = pypdf.PdfWriter()
    for _ in range(n_pages):
        writer.add_blank_page(width=200, height=200)
    with open(path, "wb") as f:
        writer.write(f)


def test_load_docs_populates_table(qtbot):
    wb = SeparatorWorkbench()
    qtbot.addWidget(wb)
    docs = [
        {"id": "1", "title": "Complaint", "date": "2023-01-01", "start": 1, "end": 2},
        {"id": "2", "title": "Exhibit A", "date": "2023-01-02", "start": 3, "end": 4},
    ]
    wb.load_docs("C:/nonexistent.pdf", docs)
    assert wb.doc_table.rowCount() == 2


def test_reanalyze_emits_sensitivity(qtbot):
    wb = SeparatorWorkbench()
    qtbot.addWidget(wb)
    wb.load_docs("C:/nonexistent.pdf", [])
    wb.sensitivity_slider.setValue(3)
    with qtbot.waitSignal(wb.reanalyze_requested, timeout=500) as blocker:
        wb.reanalyze_btn.click()
    assert blocker.args[0] == 3


def test_process_splits_checked_rows(qtbot, tmp_path):
    pdf = tmp_path / "src.pdf"
    _make_pdf(pdf, 6)
    wb = SeparatorWorkbench()
    qtbot.addWidget(wb)
    docs = [
        {"id": "1", "title": "First", "date": "", "start": 1, "end": 3},
        {"id": "2", "title": "Second", "date": "", "start": 4, "end": 6},
    ]
    wb.load_docs(str(pdf), docs)
    # Check "Sep." on both rows.
    for row in range(wb.doc_table.rowCount()):
        wb.doc_table.item(row, 0).setCheckState(Qt.CheckState.Checked)
    captured = {}
    wb.processing_complete.connect(lambda summary: captured.update(summary))
    wb.process_documents()
    out_dir = tmp_path / "PULLED-src"
    assert out_dir.is_dir()
    assert len(list(out_dir.glob("*.pdf"))) == 2
    assert len(captured.get("created", [])) == 2


def test_process_merges_group(qtbot, tmp_path):
    pdf = tmp_path / "src.pdf"
    _make_pdf(pdf, 6)
    wb = SeparatorWorkbench()
    qtbot.addWidget(wb)
    docs = [
        {"id": "1", "title": "A", "date": "", "start": 1, "end": 2},
        {"id": "2", "title": "B", "date": "", "start": 3, "end": 4},
    ]
    wb.load_docs(str(pdf), docs)
    for row in range(wb.doc_table.rowCount()):
        wb.doc_table.cellWidget(row, 1).setText("Combined")
    wb.process_documents()
    out_dir = tmp_path / "PULLED-src"
    assert (out_dir / "Combined.pdf").exists()
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_separator_workbench.py -v`
Expected: FAIL with `ModuleNotFoundError: No module named 'icharlotte_core.ui.separator_workbench'`.

- [ ] **Step 3: Create `separator_workbench.py`**

Create `icharlotte_core/ui/separator_workbench.py`. The table/preview/mark methods are MOVED VERBATIM from `IndexTab` (see "Reference facts" for current line ranges) with three behavior changes only:
1. `on_reanalyze_clicked` emits `reanalyze_requested(sensitivity)` instead of calling `self.window().run_separator_path(...)`.
2. `process_documents` reads `self.current_pdf_path` (set by `load_docs`) instead of `self.pdf_list.currentItem()`, and emits `processing_complete(summary_dict)` instead of showing a `QMessageBox`.
3. `add_document_row` checks `self.current_pdf_path` instead of `self.pdf_list.currentItem()`.

```python
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
    QMessageBox, QPushButton, QSlider, QTableWidget, QTableWidgetItem,
    QVBoxLayout, QWidget,
)

from ..utils import sanitize_filename, format_date_to_mm_dd_yyyy
from .pdf_viewer_widget import PdfViewerWidget
# DateTableWidgetItem lives in tabs.py; import lazily inside methods to avoid
# a circular import (tabs.py imports many UI widgets).


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

        from PySide6.QtWidgets import QSplitter
        self.splitter = QSplitter(Qt.Orientation.Horizontal)
        root.addWidget(self.splitter)

        # ---- Left: table + controls ----
        middle_widget = QWidget()
        middle_layout = QVBoxLayout(middle_widget)
        middle_layout.setContentsMargins(0, 0, 0, 0)

        # Search bar
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Search by Date or Title...")
        self.search_input.textChanged.connect(self.filter_documents)
        middle_layout.addWidget(self.search_input)

        # Sensitivity slider + Re-analyze
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

        # Document table
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

    # ---- Re-analyze (decoupled) ----

    def on_reanalyze_clicked(self):
        if not self.current_pdf_path:
            return
        self.set_busy(True)
        self.reanalyze_requested.emit(self.sensitivity_slider.value())

    # ---- Methods moved verbatim from IndexTab ----
    # filter_documents, _add_doc_to_table, delete_selected_rows,
    # show_context_menu, set_merge_group_batch, clear_merge_group_batch,
    # _parse_pages, on_doc_clicked, on_doc_double_clicked, mark_start_page,
    # mark_end_and_add, clear_marked_range, _get_next_doc_id, _get_doc_from_row
    # -- copy each verbatim from icharlotte_core/ui/tabs.py (line ranges in
    #    "Reference facts"). They reference only self.doc_table / self.pdf_viewer
    #    / self.marked_start_page / self.mark_status, all defined above.
    #    _add_doc_to_table needs DateTableWidgetItem: add this import at the top
    #    of the method body: `from .tabs import DateTableWidgetItem`.

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

    # ---- Process (emits signal instead of QMessageBox) ----

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
```

When copying the verbatim methods, paste them in place of the comment block. Do NOT include `save_table_to_index` (IndexTab-only) or `on_pdf_selected` (IndexTab-only).

- [ ] **Step 4: Run tests to verify they pass**

Run: `python -m pytest tests/test_separator_workbench.py -v`
Expected: 4 passed.

- [ ] **Step 5: Commit**

```bash
git add icharlotte_core/ui/separator_workbench.py tests/test_separator_workbench.py
git commit -m "feat(separate): extract reusable SeparatorWorkbench widget"
```

---

## Task 3: Refactor `IndexTab` to host `SeparatorWorkbench`

**Files:**
- Modify: `icharlotte_core/ui/tabs.py` (`IndexTab`, lines ~2338–3313)
- Test: manual + existing suite

`IndexTab` keeps: left PDF list, `index_data`/`save_data`/`load_data`/`add_pdf` persistence, and `save_table_to_index`. It embeds one `SeparatorWorkbench` in place of the middle+right panels, and delegates table operations to it.

- [ ] **Step 1: Replace the middle+right UI construction with the workbench**

In `IndexTab.setup_ui`, the PDF list (left) stays. Replace everything from `# Middle: Index Table + Controls` (line ~2393) through the end of `setup_ui` (line 2537, the `setSizes`/`setCollapsible`) with:

```python
        # Right of the PDF list: the shared workbench.
        from .separator_workbench import SeparatorWorkbench
        self.workbench = SeparatorWorkbench()
        self.workbench.reanalyze_requested.connect(self._on_workbench_reanalyze)
        self.workbench.processing_complete.connect(self._on_workbench_processing_complete)
        self.doc_splitter.addWidget(self.workbench)

        self.doc_splitter.setSizes([200, 900])
        self.doc_splitter.setCollapsible(0, True)
```

- [ ] **Step 2: Repoint IndexTab references to the workbench, add handlers**

Replace `IndexTab.on_pdf_selected` body with:

```python
    def on_pdf_selected(self, current, previous):
        if not current:
            return
        path = current.text()
        self.current_pdf_path = path
        docs = self.index_data.get(path, [])
        self.workbench.load_docs(path, docs)
```

Replace `IndexTab.add_pdf`'s trailing "Re-enable sensitivity controls" block:

```python
        # Re-enable sensitivity controls
        if hasattr(self, 'reanalyze_btn'):
            self.reanalyze_btn.setEnabled(True)
            self.sensitivity_slider.setEnabled(True)
```

with:

```python
        if hasattr(self, 'workbench'):
            self.workbench.set_busy(False)
```

Add two new handler methods and update `save_table_to_index` to read from the workbench table. Insert after `add_pdf`:

```python
    def _on_workbench_reanalyze(self, sensitivity):
        if not self.current_pdf_path:
            return
        main_window = self.window()
        if hasattr(main_window, 'run_separator_path'):
            main_window.run_separator_path(self.current_pdf_path, sensitivity=sensitivity)

    def _on_workbench_processing_complete(self, summary):
        created = summary.get("created", [])
        errors = summary.get("errors", [])
        msg = f"Processed {len(created)} item(s).\n\nFiles Created:\n" + "\n".join(created[:10])
        if len(created) > 10:
            msg += f"\n...and {len(created) - 10} more."
        if errors:
            msg += "\n\nErrors:\n" + "\n".join(errors)
            QMessageBox.warning(self, "Result with Errors", msg)
        else:
            QMessageBox.information(self, "Success", msg)
```

Replace `save_table_to_index` so it reads from `self.workbench`:

```python
    def save_table_to_index(self):
        if not self.current_pdf_path:
            QMessageBox.warning(self, "Warning", "No PDF selected.")
            return
        wb = self.workbench
        new_docs = []
        for row in range(wb.doc_table.rowCount()):
            doc_obj = wb._get_doc_from_row(row)
            if doc_obj['start'] is None or doc_obj['end'] is None:
                QMessageBox.warning(self, "Validation Error",
                    f"Row {row + 1} (ID {doc_obj['id']}): Invalid page range. Cannot save.")
                return
            new_docs.append(doc_obj)
        self.index_data[self.current_pdf_path] = new_docs
        self.save_data()
        QMessageBox.information(self, "Success", f"Saved {len(new_docs)} document(s) to index.")
```

- [ ] **Step 3: Delete the now-duplicated methods from IndexTab**

Delete these methods from `IndexTab` (they now live in `SeparatorWorkbench`): `filter_documents`, `_add_doc_to_table`, `on_reanalyze_clicked`, `add_document_row`, `delete_selected_rows`, `show_context_menu`, `set_merge_group_batch`, `clear_merge_group_batch`, `_parse_pages`, `on_doc_clicked`, `on_doc_double_clicked`, `mark_start_page`, `mark_end_and_add`, `clear_marked_range`, `_get_next_doc_id`, `_get_doc_from_row`, and `process_documents`. Keep: `load_data`, `save_data`, `add_pdf`, `on_pdf_selected`, `save_table_to_index`, `toggle_pdf_list_collapse`, and the new handlers. Remove now-unused widget refs (`self.doc_table`, `self.pdf_viewer`, `self.search_input`, etc.) — they're accessed via `self.workbench`.

- [ ] **Step 4: Check for other references to removed IndexTab attributes**

Run: `python -m pytest tests/ -k "index or separat" -v` and grep for external callers:

Run: `git grep -n "index_tab\.\(doc_table\|pdf_viewer\|process_documents\|sensitivity_slider\|reanalyze_btn\)" -- "*.py"`
Expected: no hits outside tabs.py. If `iCharlotte.py`'s `run_separator_path` (line 2192) references `self.index_tab.reanalyze_btn` / `self.index_tab.sensitivity_slider`, update those to `self.index_tab.workbench.set_busy(False)`.

Apply this edit to `iCharlotte.py` `run_separator_path` `on_finished` (lines 2191-2194):

```python
            # Re-enable sensitivity controls even on failure
            if hasattr(self, 'index_tab') and hasattr(self.index_tab, 'workbench'):
                self.index_tab.workbench.set_busy(False)
```

- [ ] **Step 5: Verify advanced mode still imports + tabs construct**

Run: `python -c "import icharlotte_core.ui.tabs as t; print('ok')"`
Expected: `ok` (no import error).

Run: `python -m pytest tests/ -q`
Expected: no NEW failures vs. baseline (note any pre-existing failures before this task).

- [ ] **Step 6: Manual smoke (advanced mode)**

Launch the app, switch to Advanced Mode, open a case, run Separate on a PDF, confirm the table fills, edit a title, check Sep. on a row, click Process, confirm a file lands in `PULLED-<source>/`. (Per CLAUDE.md: test after changing a feature.)

Run: `python iCharlotte.py`

- [ ] **Step 7: Commit**

```bash
git add icharlotte_core/ui/tabs.py iCharlotte.py
git commit -m "refactor(separate): IndexTab hosts shared SeparatorWorkbench"
```

---

## Task 4: Register the wizard task (registry + routing + builder shim)

**Files:**
- Modify: `icharlotte_core/ui/wizard/registry.py`
- Modify: `icharlotte_core/ui/wizard/task_routing.py`
- Modify: `icharlotte_core/ui/wizard/in_process_task_tab.py`
- Test: `tests/test_wizard/test_separate_routing.py` (create)

- [ ] **Step 1: Write the failing test**

Create `tests/test_wizard/test_separate_routing.py`:

```python
"""Registry + routing wiring for the Separate wizard task."""
from icharlotte_core.ui.wizard.registry import get_task, TASK_REGISTRY
from icharlotte_core.ui.wizard.task_routing import (
    get_in_process_task_builder_name,
    is_in_process_task,
    requires_initial_file_picker,
)


def test_separate_in_registry():
    spec = get_task("separate")
    assert spec.title == "Separate Documents"
    assert spec.script_name == "separate.py"


def test_separate_is_in_process_no_picker():
    assert is_in_process_task("separate")
    assert get_in_process_task_builder_name("separate") == "build_separate_tab"
    assert requires_initial_file_picker("separate") is False


def test_builder_attribute_exists():
    from icharlotte_core.ui.wizard import in_process_task_tab
    assert hasattr(in_process_task_tab, "build_separate_tab")
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_wizard/test_separate_routing.py -v`
Expected: FAIL with `KeyError: 'separate'`.

- [ ] **Step 3: Add the registry entry**

In `icharlotte_core/ui/wizard/registry.py`, add inside `TASK_REGISTRY` (after `"med_record_extractor"` is fine):

```python
    "separate": TaskSpec(
        task_id="separate",
        title="Separate Documents",
        description="Split a combined PDF into individually-named documents using AI.",
        icon_glyph="\U0001F4D1",  # 📑
        script_name="separate.py",
        default_folders=[],
    ),
```

- [ ] **Step 4: Register the routing builder**

In `icharlotte_core/ui/wizard/task_routing.py`, add to `_IN_PROCESS_TASK_BUILDERS`:

```python
_IN_PROCESS_TASK_BUILDERS = {
    "subpoena_tracker": "build_subpoena_tab",
    "respond_to_discovery": "build_respond_to_discovery_tab",
    "oppose_motion": "build_oppose_motion_tab",
    "separate": "build_separate_tab",
}
```

- [ ] **Step 5: Add the builder shim**

In `icharlotte_core/ui/wizard/in_process_task_tab.py`, add at the end of the file (mirrors how `build_oppose_motion_tab` re-exports from its page module):

```python
def build_separate_tab(
    spec,
    case_path: str,
    file_number: str,
    parent: QWidget | None,
):
    from icharlotte_core.ui.wizard.pages.separate_page import (
        build_separate_tab as _build,
    )

    return _build(
        spec=spec,
        case_path=case_path,
        file_number=file_number,
        parent=parent,
    )
```

> Note: this imports `separate_page`, created in Task 6. Step 6 below tolerates that by testing only the registry/routing pieces that don't import the page. The full builder test runs in Task 6.

- [ ] **Step 6: Run tests to verify they pass**

Run: `python -m pytest tests/test_wizard/test_separate_routing.py::test_separate_in_registry tests/test_wizard/test_separate_routing.py::test_separate_is_in_process_no_picker -v`
Expected: 2 passed. (`test_builder_attribute_exists` passes once Task 6 creates `separate_page.py`; it will pass now too because the shim is defined at module top level and only imports lazily.)

Run the third test as well: `python -m pytest tests/test_wizard/test_separate_routing.py -v`
Expected: 3 passed (the shim function exists even though its body's import resolves later).

- [ ] **Step 7: Commit**

```bash
git add icharlotte_core/ui/wizard/registry.py icharlotte_core/ui/wizard/task_routing.py icharlotte_core/ui/wizard/in_process_task_tab.py tests/test_wizard/test_separate_routing.py
git commit -m "feat(separate): register Separate Documents wizard task"
```

---

## Task 5: `SeparateAnalysisWorker` (runs separate.py --headless)

**Files:**
- Create: `icharlotte_core/ui/wizard/pages/separate_page.py` (worker only, this task)
- Test: `tests/test_wizard/test_separate_worker.py` (create)

The worker runs `python Scripts/separate.py --headless --sensitivity N <pdf>` as a subprocess (UTF-8, `errors="replace"`), parses `JSON_MAP: <path>` from stdout, loads the doc map, emits `finished_analysis(True, docs)`.

- [ ] **Step 1: Write the failing test**

Create `tests/test_wizard/test_separate_worker.py`:

```python
"""Tests for SeparateAnalysisWorker output parsing (no real subprocess)."""
import json

import pytest

pytest.importorskip("pytestqt")

from icharlotte_core.ui.wizard.pages.separate_page import SeparateAnalysisWorker


def test_parse_json_map_from_stdout(tmp_path):
    docs = [{"id": "1", "title": "X", "date": "", "start": 1, "end": 2}]
    map_path = tmp_path / "map.json"
    map_path.write_text(json.dumps(docs), encoding="utf-8")
    stdout = f"some log line\nJSON_MAP: {map_path}\nmore\n"
    parsed = SeparateAnalysisWorker._parse_docs(stdout)
    assert parsed == docs


def test_parse_missing_marker_returns_none():
    assert SeparateAnalysisWorker._parse_docs("no marker here") is None
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_wizard/test_separate_worker.py -v`
Expected: FAIL with `ModuleNotFoundError` (separate_page not created).

- [ ] **Step 3: Create `separate_page.py` with the worker**

Create `icharlotte_core/ui/wizard/pages/separate_page.py`:

```python
"""Wizard Mode "Separate Documents" task.

Settings (pick sensitivity, Analyze) → Status (analyzing) → Workbench
(embedded SeparatorWorkbench for review + split/merge). Backed by
SeparateAnalysisWorker, which runs Scripts/separate.py --headless to produce
the document map (and the Word index, as a side effect, exactly like Advanced
Mode).
"""
import json
import os
import re
import subprocess
import sys

from PySide6.QtCore import QThread, Signal

from icharlotte_core.config import SCRIPTS_DIR


class SeparateAnalysisWorker(QThread):
    progress = Signal(str)
    finished_analysis = Signal(bool, object)  # (success, list[dict] | error str)

    def __init__(self, pdf_path: str, sensitivity: int, parent=None):
        super().__init__(parent)
        self.pdf_path = pdf_path
        self.sensitivity = sensitivity

    @staticmethod
    def _parse_docs(stdout: str):
        match = re.search(r"JSON_MAP:\s*(.+)", stdout)
        if not match:
            return None
        json_path = match.group(1).strip()
        try:
            with open(json_path, "r", encoding="utf-8") as f:
                docs = json.load(f)
        except Exception:
            return None
        try:
            os.remove(json_path)
        except OSError:
            pass
        return docs

    def run(self):
        try:
            script_path = os.path.join(SCRIPTS_DIR, "separate.py")
            self.progress.emit(f"Analyzing {os.path.basename(self.pdf_path)}…")
            proc = subprocess.run(
                [sys.executable, script_path, "--headless",
                 "--sensitivity", str(self.sensitivity), self.pdf_path],
                capture_output=True, text=True, encoding="utf-8", errors="replace",
            )
            for line in (proc.stdout or "").splitlines():
                if line.strip():
                    self.progress.emit(line.strip())
            if proc.returncode != 0:
                self.finished_analysis.emit(
                    False, f"Analysis failed (exit {proc.returncode}). {proc.stderr[-500:]}")
                return
            docs = self._parse_docs(proc.stdout or "")
            if docs is None:
                self.finished_analysis.emit(
                    False, "Analysis completed but no document map was produced.")
                return
            self.finished_analysis.emit(True, docs)
        except Exception as e:
            self.finished_analysis.emit(False, str(e))
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `python -m pytest tests/test_wizard/test_separate_worker.py -v`
Expected: 2 passed.

- [ ] **Step 5: Commit**

```bash
git add icharlotte_core/ui/wizard/pages/separate_page.py tests/test_wizard/test_separate_worker.py
git commit -m "feat(separate): SeparateAnalysisWorker runs separate.py headless"
```

---

## Task 6: Settings page, Workbench page, TaskTab, and builder

**Files:**
- Modify: `icharlotte_core/ui/wizard/pages/separate_page.py` (add pages, tab, builder)
- Test: `tests/test_wizard/test_separate_task.py` (create)

- [ ] **Step 1: Write the failing test**

Create `tests/test_wizard/test_separate_task.py`:

```python
"""Tests for the Separate wizard task tab (settings + workbench wiring)."""
import pytest

pytest.importorskip("pytestqt")
pytest.importorskip("pypdf")

from icharlotte_core.ui.wizard.registry import get_task
from icharlotte_core.ui.wizard.pages.separate_page import (
    SeparateSettingsPage,
    SeparateTaskTab,
    PAGE_SETTINGS,
    PAGE_STATUS,
    PAGE_WORKBENCH,
)


def test_settings_emits_analyze_with_sensitivity(qtbot):
    page = SeparateSettingsPage(pdf_path="C:/x.pdf")
    qtbot.addWidget(page)
    page.sensitivity_slider.setValue(1)
    with qtbot.waitSignal(page.analyze_requested, timeout=500) as blocker:
        page.analyze_btn.click()
    assert blocker.args[0] == 1


def test_tab_starts_on_settings(qtbot):
    spec = get_task("separate")
    tab = SeparateTaskTab(spec, case_path="C:/case", file_number="1234.001",
                          pdf_path="C:/x.pdf")
    qtbot.addWidget(tab)
    assert tab.currentIndex() == PAGE_SETTINGS


def test_analysis_success_loads_workbench(qtbot, tmp_path):
    import pypdf
    pdf = tmp_path / "src.pdf"
    w = pypdf.PdfWriter()
    for _ in range(4):
        w.add_blank_page(width=200, height=200)
    with open(pdf, "wb") as f:
        w.write(f)

    spec = get_task("separate")
    tab = SeparateTaskTab(spec, case_path=str(tmp_path), file_number="1234.001",
                          pdf_path=str(pdf))
    qtbot.addWidget(tab)
    docs = [{"id": "1", "title": "A", "date": "", "start": 1, "end": 4}]
    # Simulate worker completion directly (no real subprocess).
    tab._on_analysis_finished(True, docs)
    assert tab.currentIndex() == PAGE_WORKBENCH
    assert tab.workbench.doc_table.rowCount() == 1
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_wizard/test_separate_task.py -v`
Expected: FAIL with `ImportError: cannot import name 'SeparateSettingsPage'`.

- [ ] **Step 3: Append pages, tab, and builder to `separate_page.py`**

Append to `icharlotte_core/ui/wizard/pages/separate_page.py`:

```python
import os as _os  # already imported above; safe noop if linter flags, remove dup

from PySide6.QtWidgets import (
    QFileDialog, QHBoxLayout, QLabel, QMessageBox, QPushButton, QSlider,
    QVBoxLayout, QWidget,
)
from PySide6.QtCore import Qt

from icharlotte_core.ui.wizard import theme
from icharlotte_core.ui.wizard.pages.status_page import StatusPage
from icharlotte_core.ui.wizard.task_scaffold import WizardTaskContainer
from icharlotte_core.ui.separator_workbench import SeparatorWorkbench
from icharlotte_core.ui.wizard.file_picker import resolve_default_folder

PAGE_SETTINGS = 0
PAGE_STATUS = 1
PAGE_WORKBENCH = 2


class SeparateSettingsPage(QWidget):
    analyze_requested = Signal(int)  # sensitivity

    def __init__(self, pdf_path: str, parent=None):
        super().__init__(parent)
        self.pdf_path = pdf_path
        layout = QVBoxLayout(self)
        layout.setContentsMargins(theme.SPACE_XL, theme.SPACE_XL, theme.SPACE_XL, theme.SPACE_XL)
        layout.setSpacing(theme.SPACE_MD)

        layout.addWidget(theme.page_title("Separate Documents"))
        layout.addWidget(theme.helper_text(
            "iCharlotte will scan the PDF, identify the distinct documents inside it, "
            "and let you review, rename, split, and merge them. An index is also saved "
            "to NOTES/AI OUTPUT/INDEXES."))
        layout.addWidget(QLabel(f"File: {os.path.basename(pdf_path)}"))

        sens_row = QHBoxLayout()
        sens_row.addWidget(QLabel("Separation sensitivity:"))
        broad = QLabel("Broad"); broad.setStyleSheet("color:#666;font-size:11px;")
        sens_row.addWidget(broad)
        self.sensitivity_slider = QSlider(Qt.Orientation.Horizontal)
        self.sensitivity_slider.setMinimum(1)
        self.sensitivity_slider.setMaximum(3)
        self.sensitivity_slider.setValue(2)
        self.sensitivity_slider.setTickPosition(QSlider.TickPosition.TicksBelow)
        self.sensitivity_slider.setTickInterval(1)
        self.sensitivity_slider.setPageStep(1)
        self.sensitivity_slider.setFixedWidth(120)
        sens_row.addWidget(self.sensitivity_slider)
        fine = QLabel("Fine"); fine.setStyleSheet("color:#666;font-size:11px;")
        sens_row.addWidget(fine)
        sens_row.addStretch()
        layout.addLayout(sens_row)

        layout.addStretch(1)
        btn_row = QHBoxLayout()
        btn_row.addStretch()
        self.analyze_btn = theme.primary_button("Analyze")
        self.analyze_btn.clicked.connect(
            lambda: self.analyze_requested.emit(self.sensitivity_slider.value()))
        btn_row.addWidget(self.analyze_btn)
        layout.addLayout(btn_row)

    def to_dict(self) -> dict:
        return {"sensitivity": self.sensitivity_slider.value()}

    def from_dict(self, data: dict) -> None:
        if "sensitivity" in data:
            self.sensitivity_slider.setValue(int(data["sensitivity"]))


class SeparateWorkbenchPage(QWidget):
    """Hosts the SeparatorWorkbench plus a result banner + Open Folder button."""

    def __init__(self, parent=None):
        super().__init__(parent)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(theme.SPACE_MD, theme.SPACE_MD, theme.SPACE_MD, theme.SPACE_MD)
        layout.setSpacing(theme.SPACE_SM if hasattr(theme, "SPACE_SM") else 6)

        self.banner = QLabel("")
        self.banner.setWordWrap(True)
        self.banner.setVisible(False)
        layout.addWidget(self.banner)

        self.workbench = SeparatorWorkbench()
        layout.addWidget(self.workbench, 1)

        btn_row = QHBoxLayout()
        self.open_folder_btn = theme.secondary_button("Open Output Folder")
        self.open_folder_btn.setVisible(False)
        self.open_folder_btn.clicked.connect(self._open_folder)
        btn_row.addWidget(self.open_folder_btn)
        btn_row.addStretch()
        layout.addLayout(btn_row)

        self._output_folder = ""
        self.workbench.processing_complete.connect(self._on_processing_complete)

    def _on_processing_complete(self, summary: dict):
        created = summary.get("created", [])
        errors = summary.get("errors", [])
        self._output_folder = summary.get("output_folder", "")
        text = f"✓ Created {len(created)} file(s) in {os.path.basename(self._output_folder)}."
        if errors:
            text += f"  ⚠ {len(errors)} error(s): " + "; ".join(errors[:3])
        self.banner.setText(text)
        self.banner.setVisible(True)
        self.open_folder_btn.setVisible(bool(self._output_folder))

    def _open_folder(self):
        if self._output_folder and os.path.isdir(self._output_folder):
            try:
                os.startfile(self._output_folder)  # Windows
            except Exception as e:
                QMessageBox.critical(self, "Open failed", f"Could not open folder:\n{e}")


class SeparateTaskTab(WizardTaskContainer):
    task_completed = Signal(dict)

    def __init__(self, spec, case_path: str, file_number: str, pdf_path: str, parent=None):
        super().__init__(spec, steps=["Settings", "Analyzing", "Review & Split"], parent=parent)
        self._case_path = case_path
        self._file_number = file_number
        self._pdf_path = pdf_path
        self._worker = None

        self.settings_page = SeparateSettingsPage(pdf_path)
        self.status_page = StatusPage()
        self.workbench_page = SeparateWorkbenchPage()
        self.workbench = self.workbench_page.workbench  # convenience for tests/wiring

        self.addWidget(self.settings_page)    # PAGE_SETTINGS
        self.addWidget(self.status_page)      # PAGE_STATUS
        self.addWidget(self.workbench_page)   # PAGE_WORKBENCH

        self.settings_page.analyze_requested.connect(self._start_analysis)
        self.status_page.cancel_requested.connect(self._on_cancel)
        self.workbench.reanalyze_requested.connect(self._start_analysis)
        self.workbench.processing_complete.connect(self._on_processing_complete)

    @property
    def spec(self):
        return self._spec

    @property
    def files(self) -> list:
        return [self._pdf_path]

    def _start_analysis(self, sensitivity: int):
        if self._worker is not None and self._worker.isRunning():
            return
        self.status_page.reset()
        self.status_page.progress_bar.setRange(0, 0)
        self.status_page.on_status("Analyzing…")
        self.setCurrentIndex(PAGE_STATUS)
        worker = SeparateAnalysisWorker(self._pdf_path, sensitivity, parent=None)
        worker.progress.connect(self.status_page.on_status)
        worker.finished_analysis.connect(self._on_analysis_finished)
        worker.finished.connect(worker.deleteLater)
        self._worker = worker
        worker.start()

    def _on_analysis_finished(self, success: bool, payload: object):
        self._worker = None
        if not success:
            self.status_page.on_status(f"FAILED: {payload}")
            self.status_page.cancel_btn.setText("Back to Settings")
            self.status_page.cancel_btn.setEnabled(True)
            try:
                self.status_page.cancel_btn.clicked.disconnect()
            except RuntimeError:
                pass
            self.status_page.cancel_btn.clicked.connect(
                lambda: self.setCurrentIndex(PAGE_SETTINGS))
            return
        docs = payload if isinstance(payload, list) else []
        self.workbench.set_busy(False)
        self.workbench.load_docs(self._pdf_path, docs)
        self.setCurrentIndex(PAGE_WORKBENCH)

    def _on_cancel(self):
        # The analysis subprocess can't be cooperatively cancelled; just bounce
        # back to settings if it hasn't produced anything yet.
        self.setCurrentIndex(PAGE_SETTINGS)

    def _on_processing_complete(self, summary: dict):
        from datetime import datetime
        self.task_completed.emit({
            "task_id": self._spec.task_id,
            "title": self._spec.title,
            "files": [self._pdf_path],
            "settings": self.settings_page.to_dict(),
            "output_path": summary.get("output_folder", ""),
            "completed_at": datetime.now().isoformat(timespec="seconds"),
        })

    def closeEvent(self, event):
        if self._worker is not None and self._worker.isRunning():
            QMessageBox.information(
                self, "Task running",
                "Analysis is still running. Wait for it to finish before closing this tab.")
            event.ignore()
            return
        super().closeEvent(event)


def build_separate_tab(spec, case_path: str, file_number: str, parent=None):
    start_dir = resolve_default_folder(case_path, spec.default_folders)
    pdf_path, _ = QFileDialog.getOpenFileName(
        parent, "Select a PDF to separate", start_dir, "PDF files (*.pdf)")
    if not pdf_path:
        return None
    return SeparateTaskTab(
        spec=spec, case_path=case_path, file_number=file_number,
        pdf_path=pdf_path, parent=parent)
```

> Remove the duplicate `import os as _os` line if your editor flags it — `os`, `Signal` are already imported at the top of the file from Task 5. Consolidate imports so there is exactly one import of each.

- [ ] **Step 4: Run tests to verify they pass**

Run: `python -m pytest tests/test_wizard/test_separate_task.py tests/test_wizard/test_separate_routing.py -v`
Expected: all pass (including `test_builder_attribute_exists` now fully resolvable).

- [ ] **Step 5: Commit**

```bash
git add icharlotte_core/ui/wizard/pages/separate_page.py tests/test_wizard/test_separate_task.py
git commit -m "feat(separate): SeparateTaskTab settings + workbench page + builder"
```

---

## Task 7: Recent-task reopen + final integration

**Files:**
- Modify: `iCharlotte.py` (`_on_reopen_recent_task` ~line 1185; `_restore_task_tabs_for_case` ~line 1260)
- Test: manual

The existing reopen/restore paths special-case only `build_oppose_motion_tab` and otherwise build a generic `TaskTab` — which is wrong for Separate (an in-process custom tab). Make those paths re-open Separate via its picker-based builder, matching the design's "reopen re-picks the PDF" decision.

- [ ] **Step 1: Guard reopen for in-process custom tabs**

In `iCharlotte.py._on_reopen_recent_task`, the block at line ~1185 handles oppose_motion. Add a sibling branch for any in-process builder (so Separate uses its builder rather than `TaskTab`). After the existing oppose_motion `if` block, add:

```python
        builder_name = get_in_process_task_builder_name(task_id)
        if builder_name and builder_name != "build_oppose_motion_tab":
            from icharlotte_core.ui.wizard import in_process_task_tab
            builder = getattr(in_process_task_tab, builder_name)
            task_tab = builder(
                spec=spec, case_path=self.case_path,
                file_number=self.file_number, parent=self)
            if task_tab is None:
                return
            task_tab.setProperty("wizard_task_id", spec.task_id)
            task_tab.setProperty("wizard_instance_suffix", suffix)
            task_tab.task_completed.connect(self._on_task_completed)
            new_index = self.tabs.addTab(task_tab, title)
            self.tabs.setCurrentIndex(new_index)
            self._hide_fixed_close_buttons()
            return
```

Make the same guard in `_restore_task_tabs_for_case` (line ~1260): after the oppose_motion branch, before the generic `TaskTab` construction, add an equivalent branch that calls the builder for other in-process tasks (so a restored Separate tab re-opens its picker rather than constructing an empty `TaskTab`).

```python
            builder_name = get_in_process_task_builder_name(task_id)
            if builder_name and builder_name != "build_oppose_motion_tab":
                # In-process custom tabs (Separate) re-pick their source on
                # restore; skip silently if the user cancels.
                from icharlotte_core.ui.wizard import in_process_task_tab
                builder = getattr(in_process_task_tab, builder_name)
                task_tab = builder(
                    spec=spec, case_path=self.case_path,
                    file_number=self.file_number, parent=self)
                if task_tab is None:
                    continue
                task_tab.setProperty("wizard_task_id", task_id)
                task_tab.task_completed.connect(self._on_task_completed)
                self.tabs.addTab(task_tab, entry.get("title") or spec.title)
                continue
```

> Verify `spec`, `suffix`, `title`, and `task_id` are in scope at each insertion point; if `spec` isn't yet defined in `_restore_task_tabs_for_case`, add `spec = get_task(task_id)` above the branch (it's already imported).

- [ ] **Step 2: Verify imports + full suite**

Run: `python -c "import iCharlotte" 2>&1 | tail -5`
Expected: no import/syntax error (a Qt "no display" style message is fine; a `SyntaxError`/`ImportError` is not).

Run: `python -m pytest tests/test_wizard/ -v`
Expected: all separate-task tests pass; no new failures elsewhere.

- [ ] **Step 3: Manual end-to-end (wizard mode)** — MANDATORY per CLAUDE.md

Run: `python iCharlotte.py`
Then: open a case → Wizard Mode → click the **Separate Documents** card → pick a multi-document PDF → set sensitivity → **Analyze** → confirm the workbench fills with identified docs and the PDF preview loads → edit a title, mark a custom range with Mark Start/Mark End, check a couple **Sep.** rows and set a **Merge Group** on two rows → **Process Documents** → confirm:
  - result banner appears + **Open Output Folder** works,
  - `PULLED-<source>/` contains the split files and the merged file,
  - `NOTES/AI OUTPUT/INDEXES/Index_<source>.docx` exists.
Also re-run **Re-analyze** from the workbench at a different sensitivity and confirm the table reloads.

- [ ] **Step 4: Update memory + commit**

Add a one-line index entry to `MEMORY.md` topic index (e.g. `- separate_wizard_task.md — Separate ported to wizard; shared SeparatorWorkbench extracted from IndexTab; SeparateAnalysisWorker runs separate.py --headless; reopen re-picks PDF`). Create that topic file with the key gotchas discovered during implementation.

```bash
git add iCharlotte.py
git commit -m "feat(separate): wizard reopen/restore via in-process builder"
```

---

## Self-Review

**Spec coverage:**
- Behavior/flow (spec §1) → Tasks 5, 6 (settings→status→workbench).
- Extract shared workbench (spec §2A) → Task 2.
- IndexTab thin host (spec §2B) → Task 3.
- SeparateTaskTab + worker (spec §2C) → Tasks 5, 6.
- Registry & routing (spec §2D) → Task 4.
- Output locations unchanged (spec §2E) → preserved by reusing separate.py + workbench process logic (Tasks 2, 5).
- Word validation (spec §2F) → Task 1.
- Testing (spec §3) → tests in Tasks 1–6 + manual in Tasks 3, 7.
- Known limitation: reopen re-picks PDF (spec §5) → Task 7.

**Placeholder scan:** No TBD/TODO; every code step shows complete code. The one "move verbatim" instruction (Task 2 Step 3) names exact methods + line ranges and lists the only three behavior changes — acceptable because rewriting 250 correct lines verbatim adds transcription risk; the methods are self-contained and reference only widgets defined in `_setup_ui`.

**Type consistency:** `reanalyze_requested(int)`, `processing_complete(dict)`, `analyze_requested(int)`, `finished_analysis(bool, object)` used consistently across Tasks 2/5/6. Page constants `PAGE_SETTINGS/PAGE_STATUS/PAGE_WORKBENCH` defined once (Task 6) and used in tests. `SeparatorWorkbench.load_docs/set_busy/process_documents` signatures match call sites in IndexTab (Task 3) and SeparateTaskTab (Task 6). Builder signature `(spec, case_path, file_number, parent)` matches the dispatch call in iCharlotte.py.
