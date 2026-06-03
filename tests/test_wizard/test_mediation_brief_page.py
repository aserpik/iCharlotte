"""Tests for the Mediation Brief Wizard task."""
import os

import pytest


# ---- Registry / routing (pure logic, no Qt) ----

def test_mediation_brief_registered():
    from icharlotte_core.ui.wizard.registry import TASK_REGISTRY

    spec = TASK_REGISTRY["mediation_brief"]
    assert spec.title == "Mediation Brief"
    assert spec.category == "Motions & Drafting"
    assert spec.script_name == ""
    assert "mediation" in spec.keywords


def test_mediation_brief_has_valid_category():
    from icharlotte_core.ui.wizard.registry import CATEGORY_ORDER, TASK_REGISTRY

    assert TASK_REGISTRY["mediation_brief"].category in CATEGORY_ORDER


def test_mediation_brief_is_in_process():
    from icharlotte_core.ui.wizard.task_routing import (
        get_in_process_task_builder_name,
        is_in_process_task,
        requires_initial_file_picker,
    )

    assert get_in_process_task_builder_name("mediation_brief") == "build_mediation_brief_tab"
    assert is_in_process_task("mediation_brief") is True
    # In-process tasks own their source selection — no pre-Settings picker.
    assert requires_initial_file_picker("mediation_brief") is False


def test_build_mediation_brief_tab_attribute_exists():
    from icharlotte_core.ui.wizard import in_process_task_tab

    assert hasattr(in_process_task_tab, "build_mediation_brief_tab")


# ---- Document reader ----

pytestqt = pytest.importorskip("pytestqt")  # ensures PySide6/Qt present for module import


def test_read_documents_txt_and_docx(tmp_path):
    from docx import Document

    from icharlotte_core.ui.wizard.pages.mediation_brief_page import _read_documents

    txt = tmp_path / "note.txt"
    txt.write_text("Plain text body", encoding="utf-8")

    docx_path = tmp_path / "summary.docx"
    doc = Document()
    doc.add_paragraph("Intro paragraph text")
    table = doc.add_table(rows=1, cols=2)
    table.rows[0].cells[0].text = "CellA"
    table.rows[0].cells[1].text = "CellB"
    doc.save(str(docx_path))

    content, warnings = _read_documents([str(txt), str(docx_path)])

    assert "Plain text body" in content
    assert "Intro paragraph text" in content
    # Table-aware extraction must surface cell text (no silent data loss).
    assert "CellA" in content and "CellB" in content
    assert warnings == []


def test_read_documents_reports_missing_file(tmp_path):
    from icharlotte_core.ui.wizard.pages.mediation_brief_page import _read_documents

    content, warnings = _read_documents([str(tmp_path / "nope.pdf")])

    assert content == ""
    assert any("not found" in w.lower() for w in warnings)
