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


def test_sniff_text_layer_docx_with_table_content_only(tmp_path):
    """python-docx's doc.paragraphs skips tables — sniff must also sample
    table cells so docs whose content lives in tables (legal chronological
    summaries, intake forms) aren't falsely flagged as 'no text layer'."""
    from docx import Document
    from icharlotte_core.ui.med_chron_config_form import sniff_text_layer
    p = tmp_path / "tables_only.docx"
    doc = Document()
    table = doc.add_table(rows=2, cols=2)
    table.cell(0, 0).text = "Date"
    table.cell(0, 1).text = "Provider"
    table.cell(1, 0).text = "2024-02-01"
    table.cell(1, 1).text = "Acme PT"
    doc.save(str(p))
    has_text, _ = sniff_text_layer(str(p))
    assert has_text is True


def test_sniff_text_layer_pdf_with_text_layer(tmp_path):
    """A PDF whose first page has > 200 chars of extractable text returns
    (True, '')."""
    from pypdf import PdfWriter
    from icharlotte_core.ui.med_chron_config_form import sniff_text_layer

    try:
        from reportlab.pdfgen import canvas
        from reportlab.lib.pagesizes import letter
    except ImportError:
        pytest.skip("reportlab not installed — needed to build a text-layer PDF")

    p = tmp_path / "with_text.pdf"
    c = canvas.Canvas(str(p), pagesize=letter)
    # > 200 chars of extractable text. Spread across a few lines.
    long_line = "The defense theory rests on plaintiff's pre-existing degenerative changes documented across multiple imaging studies."
    y = 750
    for _ in range(3):
        c.drawString(72, y, long_line)
        y -= 20
    c.save()

    has_text, _ = sniff_text_layer(str(p))
    assert has_text is True


def test_sniff_text_layer_pdf_below_threshold(tmp_path):
    """A PDF with very little text (under 200 chars) returns (False, reason)."""
    from icharlotte_core.ui.med_chron_config_form import sniff_text_layer

    try:
        from reportlab.pdfgen import canvas
        from reportlab.lib.pagesizes import letter
    except ImportError:
        pytest.skip("reportlab not installed — needed to build a text-layer PDF")

    p = tmp_path / "tiny_text.pdf"
    c = canvas.Canvas(str(p), pagesize=letter)
    c.drawString(72, 750, "Tiny.")  # only 5 chars
    c.save()

    has_text, reason = sniff_text_layer(str(p))
    assert has_text is False
    assert reason  # non-empty reason


def test_sniff_text_layer_unsupported_extension(tmp_path):
    """An .rtf or .png is reported as unsupported with a descriptive reason."""
    from icharlotte_core.ui.med_chron_config_form import sniff_text_layer
    p = tmp_path / "image.png"
    p.write_bytes(b"\x89PNG\r\n")
    has_text, reason = sniff_text_layer(str(p))
    assert has_text is False
    assert ".png" in reason


# -----------------------------
# Task 4: ContextDropTextEdit
# -----------------------------

def test_context_drop_textedit_accepts_supported_file_urls(qtbot, tmp_path):
    from PySide6.QtCore import QMimeData, QPoint, Qt, QUrl
    from PySide6.QtGui import QDragEnterEvent
    from icharlotte_core.ui.med_chron_config_form import ContextDropTextEdit

    p1 = tmp_path / "a.pdf"
    p1.write_bytes(b"%PDF-1.4 stub")
    p2 = tmp_path / "b.docx"
    p2.write_bytes(b"docx stub")

    w = ContextDropTextEdit()
    qtbot.addWidget(w)

    received: list[list[str]] = []
    w.files_dropped.connect(lambda paths: received.append(paths))

    mime = QMimeData()
    mime.setUrls([QUrl.fromLocalFile(str(p1)), QUrl.fromLocalFile(str(p2))])
    event = QDragEnterEvent(QPoint(10, 10), Qt.DropAction.CopyAction, mime,
                            Qt.MouseButton.LeftButton, Qt.KeyboardModifier.NoModifier)
    w.dragEnterEvent(event)
    assert event.isAccepted()


def test_context_drop_textedit_rejects_unsupported_file_urls(qtbot, tmp_path):
    from PySide6.QtCore import QMimeData, QPoint, Qt, QUrl
    from PySide6.QtGui import QDragEnterEvent
    from icharlotte_core.ui.med_chron_config_form import ContextDropTextEdit

    p = tmp_path / "bad.png"
    p.write_bytes(b"\x89PNG\r\n")

    w = ContextDropTextEdit()
    qtbot.addWidget(w)

    mime = QMimeData()
    mime.setUrls([QUrl.fromLocalFile(str(p))])
    event = QDragEnterEvent(QPoint(10, 10), Qt.DropAction.CopyAction, mime,
                            Qt.MouseButton.LeftButton, Qt.KeyboardModifier.NoModifier)
    w.dragEnterEvent(event)
    # Not accepted as a "file drop", so the base class falls through; since
    # the base QPlainTextEdit doesn't know how to handle a PNG URL either,
    # the event is not accepted.
    assert not event.isAccepted()


def test_context_drop_textedit_emits_files_dropped(qtbot, tmp_path):
    from PySide6.QtCore import QMimeData, QPoint, Qt, QUrl
    from PySide6.QtGui import QDropEvent
    from icharlotte_core.ui.med_chron_config_form import ContextDropTextEdit

    p = tmp_path / "good.txt"
    p.write_text("hi", encoding="utf-8")

    w = ContextDropTextEdit()
    qtbot.addWidget(w)

    received: list[list[str]] = []
    w.files_dropped.connect(lambda paths: received.append(paths))

    mime = QMimeData()
    mime.setUrls([QUrl.fromLocalFile(str(p))])
    event = QDropEvent(QPoint(10, 10), Qt.DropAction.CopyAction, mime,
                       Qt.MouseButton.LeftButton, Qt.KeyboardModifier.NoModifier)
    w.dropEvent(event)
    assert received == [[str(p)]]


def test_context_drop_textedit_plaintext_drop_still_works(qtbot):
    """Dropping plain text (not file URLs) should fall through to the base
    QPlainTextEdit so the user can still drop a snippet."""
    from PySide6.QtCore import QMimeData, QPoint, Qt
    from PySide6.QtGui import QDropEvent
    from icharlotte_core.ui.med_chron_config_form import ContextDropTextEdit

    w = ContextDropTextEdit()
    qtbot.addWidget(w)

    received: list[list[str]] = []
    w.files_dropped.connect(lambda paths: received.append(paths))

    mime = QMimeData()
    mime.setText("some snippet")
    event = QDropEvent(QPoint(10, 10), Qt.DropAction.CopyAction, mime,
                       Qt.MouseButton.LeftButton, Qt.KeyboardModifier.NoModifier)
    w.dropEvent(event)
    # files_dropped should NOT fire for plain text
    assert received == []
