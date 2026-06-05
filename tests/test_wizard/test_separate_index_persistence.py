"""Wizard Separate task persists its document map to the shared index store."""
import os

import pytest

pytest.importorskip("pytestqt")
pytest.importorskip("pypdf")

from icharlotte_core import config
from icharlotte_core import case_index_store as store
from icharlotte_core.ui.wizard.registry import get_task
from icharlotte_core.ui.wizard.pages import separate_page
from icharlotte_core.ui.wizard.pages.separate_page import SeparateTaskTab


def _make_pdf(tmp_path, pages=4):
    import pypdf
    pdf = tmp_path / "src.pdf"
    w = pypdf.PdfWriter()
    for _ in range(pages):
        w.add_blank_page(width=200, height=200)
    with open(pdf, "wb") as f:
        w.write(f)
    return str(pdf)


def test_docs_from_workbench_reads_rows():
    from types import SimpleNamespace
    rows = [{"id": "1", "title": "A"}, {"id": "2", "title": "B"}]
    wb = SimpleNamespace(
        doc_table=SimpleNamespace(rowCount=lambda: len(rows)),
        _get_doc_from_row=lambda r: rows[r],
    )
    assert separate_page._docs_from_workbench(wb) == rows


def test_analysis_persists_to_index(qtbot, tmp_path, monkeypatch):
    monkeypatch.setattr(config, "GEMINI_DATA_DIR", str(tmp_path))
    pdf = _make_pdf(tmp_path)
    spec = get_task("separate")
    tab = SeparateTaskTab(spec, case_path=str(tmp_path),
                          file_number="1234.001", pdf_path=pdf)
    qtbot.addWidget(tab)
    docs = [{"id": "1", "title": "A", "date": "", "start": 1, "end": 4}]
    tab._on_analysis_finished(True, docs)
    assert store.load_index("1234.001") == {pdf: docs}


def test_processing_persists_workbench_rows(qtbot, tmp_path, monkeypatch):
    monkeypatch.setattr(config, "GEMINI_DATA_DIR", str(tmp_path))
    pdf = _make_pdf(tmp_path)
    spec = get_task("separate")
    tab = SeparateTaskTab(spec, case_path=str(tmp_path),
                          file_number="1234.001", pdf_path=pdf)
    qtbot.addWidget(tab)
    docs = [{"id": "1", "title": "A", "date": "", "start": 1, "end": 4}]
    tab._on_analysis_finished(True, docs)
    tab._on_processing_complete({"output_folder": str(tmp_path)})
    saved = store.load_index("1234.001")[pdf]
    assert len(saved) == 1
    assert saved[0]["id"] == "1"
    assert saved[0]["title"] == "A"
    assert saved[0]["start"] == 1
    assert saved[0]["end"] == 4


def test_no_file_number_writes_nothing(qtbot, tmp_path, monkeypatch):
    monkeypatch.setattr(config, "GEMINI_DATA_DIR", str(tmp_path))
    pdf = _make_pdf(tmp_path)
    spec = get_task("separate")
    tab = SeparateTaskTab(spec, case_path=str(tmp_path),
                          file_number="", pdf_path=pdf)
    qtbot.addWidget(tab)
    tab._on_analysis_finished(True, [{"id": "1", "title": "A", "date": "", "start": 1, "end": 4}])
    assert not os.path.exists(store.index_path(""))
