from pathlib import Path

import pytest

pytest.importorskip("pytestqt")
pytest.importorskip("PySide6")


def test_output_page_loads_md_alongside_docx(qtbot, tmp_path):
    from icharlotte_core.ui.wizard.pages.depo_prep_output_page import DepoPrepOutputPage
    from Scripts.depo_prep_lib.render_docx import render_outline_docx
    from Scripts.depo_prep_lib.render_md import render_outline_md

    outline = {"deponent_name": "Jane Doe", "deponent_role": "Plaintiff",
               "topics": [{"topic_id": "t01", "title": "T1", "strategic_note": "s",
                            "questions": [{"n": 1, "text": "Q1"}]}],
               "coverage_gaps": []}
    docx = tmp_path / "outline.docx"
    md = tmp_path / "outline.md"
    render_outline_docx(outline=outline, output_path=docx)
    render_outline_md(outline=outline, output_path=md)

    page = DepoPrepOutputPage()
    qtbot.addWidget(page)
    page.load_output(str(docx))
    md_text = page.md_viewer.toPlainText()
    assert "Jane Doe" in md_text
    assert "T1" in md_text


def test_output_page_falls_back_to_docx_only_when_md_missing(qtbot, tmp_path):
    from icharlotte_core.ui.wizard.pages.depo_prep_output_page import DepoPrepOutputPage
    from Scripts.depo_prep_lib.render_docx import render_outline_docx
    outline = {"deponent_name": "X", "deponent_role": "",
               "topics": [{"topic_id": "t01", "title": "T", "strategic_note": "",
                            "questions": [{"n": 1, "text": "Q"}]}], "coverage_gaps": []}
    docx = tmp_path / "outline.docx"
    render_outline_docx(outline=outline, output_path=docx)
    # No outline.md alongside.
    page = DepoPrepOutputPage()
    qtbot.addWidget(page)
    page.load_output(str(docx))
    # Doesn't crash; md_viewer empty or hidden.
    assert page.md_viewer is not None
