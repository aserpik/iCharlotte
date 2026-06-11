import os

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

import pytest

pytest.importorskip("pytestqt")

from docx import Document
from PySide6.QtCore import Qt
from PySide6.QtWidgets import QMessageBox

from icharlotte_core.ui.wizard.summary_browser import SummaryBrowserTab


def _docx(path, text="Original summary"):
    path.parent.mkdir(parents=True, exist_ok=True)
    doc = Document()
    doc.add_paragraph(text)
    doc.save(str(path))
    return str(path)


def test_delete_key_removes_selected_summary_file(qtbot, tmp_path, monkeypatch):
    output = _docx(tmp_path / "NOTES" / "AI OUTPUT" / "AI_OUTPUT.docx")
    monkeypatch.setattr(
        QMessageBox,
        "question",
        staticmethod(lambda *args, **kwargs: QMessageBox.StandardButton.Yes),
    )

    page = SummaryBrowserTab(str(tmp_path), "1234.001", "summarize_documents")
    qtbot.addWidget(page)

    assert page.output_list.count() == 1
    qtbot.keyClick(page.output_list, Qt.Key.Key_Delete)

    assert not os.path.exists(output)
    assert page.output_list.count() == 0
    assert "No summarized documents" in page.preview.toPlainText()


def test_preview_is_editable_and_save_writes_selected_docx(qtbot, tmp_path, monkeypatch):
    output = _docx(tmp_path / "NOTES" / "AI OUTPUT" / "AI_OUTPUT.docx")
    saved = {}

    def fake_save(document, out_path):
        saved["path"] = out_path
        saved["text"] = document.toPlainText()

    monkeypatch.setattr(
        "icharlotte_core.ui.wizard.summary_browser.save_qtextdocument_as_docx",
        fake_save,
    )
    monkeypatch.setattr(
        QMessageBox,
        "information",
        staticmethod(lambda *args, **kwargs: None),
    )

    page = SummaryBrowserTab(str(tmp_path), "1234.001", "summarize_documents")
    qtbot.addWidget(page)

    assert not page.preview.isReadOnly()
    page.preview.setPlainText("Edited summary text")

    assert page.save_btn.isEnabled()
    page.save_btn.click()

    assert saved == {"path": output, "text": "Edited summary text"}
    assert not page.save_btn.isEnabled()
