"""Tests for the per-row context-documents UI feature in MedChronConfigForm."""

import sys
from pathlib import Path

import pytest

pytest.importorskip("pytestqt")  # NOTE: no underscore — pytest_qt silently skips

PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))


@pytest.fixture(autouse=True)
def _isolated_custom_analyses_store(tmp_path, monkeypatch):
    """Redirect the global custom-analyses JSON to a per-test tmp path so
    tests don't see (or pollute) the developer's real saved analyses."""
    from icharlotte_core.med_chron import custom_analyses_store
    monkeypatch.setattr(
        custom_analyses_store,
        "_STORE_PATH",
        tmp_path / "store" / "med_chron_custom_analyses.json",
    )
    yield


# -----------------------------
# Task 3: sniff_text_layer tests
# -----------------------------

def test_sniff_text_layer_txt_has_text(tmp_path):
    from icharlotte_core.ui.med_chron_config_form import sniff_text_layer
    p = tmp_path / "ok.txt"
    p.write_text("This is a useful status report with content.", encoding="utf-8")
    has_text, _ = sniff_text_layer(str(p))
    assert has_text is True


def test_sniff_text_layer_txt_empty(tmp_path):
    from icharlotte_core.ui.med_chron_config_form import sniff_text_layer
    p = tmp_path / "empty.txt"
    p.write_text("   \n  ", encoding="utf-8")
    has_text, _ = sniff_text_layer(str(p))
    assert has_text is False


def test_sniff_text_layer_docx_has_text(tmp_path):
    from docx import Document
    from icharlotte_core.ui.med_chron_config_form import sniff_text_layer
    p = tmp_path / "doc.docx"
    doc = Document()
    doc.add_paragraph("Some legible paragraph content for the sniff to find.")
    doc.save(str(p))
    has_text, _ = sniff_text_layer(str(p))
    assert has_text is True


def test_sniff_text_layer_docx_empty(tmp_path):
    from docx import Document
    from icharlotte_core.ui.med_chron_config_form import sniff_text_layer
    p = tmp_path / "blank.docx"
    Document().save(str(p))
    has_text, _ = sniff_text_layer(str(p))
    assert has_text is False


def test_sniff_text_layer_unreadable_returns_false(tmp_path):
    from icharlotte_core.ui.med_chron_config_form import sniff_text_layer
    # File that does not exist.
    has_text, reason = sniff_text_layer(str(tmp_path / "ghost.pdf"))
    assert has_text is False
    assert reason  # non-empty reason string
