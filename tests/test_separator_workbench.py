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
    for row in range(wb.doc_table.rowCount()):
        wb.doc_table.item(row, 0).setCheckState(Qt.CheckState.Checked)
    captured = {}
    wb.processing_complete.connect(lambda summary: captured.update(summary))
    wb.process_documents()
    out_dir = tmp_path / "PULLED-src"
    assert out_dir.is_dir()
    assert len(list(out_dir.glob("*.pdf"))) == 2
    assert len(captured.get("created", [])) == 2


def test_split_page_ranges_are_correct(qtbot, tmp_path):
    """Guard the 1-based->0-based page math: 1-3 yields 3 pages, 4-6 yields 3."""
    pdf = tmp_path / "src.pdf"
    _make_pdf(pdf, 6)
    wb = SeparatorWorkbench()
    qtbot.addWidget(wb)
    docs = [
        {"id": "1", "title": "First", "date": "", "start": 1, "end": 3},
        {"id": "2", "title": "Second", "date": "", "start": 4, "end": 6},
    ]
    wb.load_docs(str(pdf), docs)
    for row in range(wb.doc_table.rowCount()):
        wb.doc_table.item(row, 0).setCheckState(Qt.CheckState.Checked)
    wb.process_documents()
    out_dir = tmp_path / "PULLED-src"
    first = pypdf.PdfReader(str(out_dir / "1 - First.pdf"))
    second = pypdf.PdfReader(str(out_dir / "2 - Second.pdf"))
    assert len(first.pages) == 3
    assert len(second.pages) == 3


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
