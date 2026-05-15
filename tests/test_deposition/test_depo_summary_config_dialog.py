"""Tests for the DepoSummaryConfigDialog popup."""

import json
import os
from pathlib import Path
from unittest.mock import patch

import pytest

pytest.importorskip("pytestqt")

from PySide6.QtCore import Qt

from icharlotte_core.deposition import session_manager
from icharlotte_core.ui.depo_summary_config_dialog import DepoSummaryConfigDialog


def _make_session(tmp_path) -> Path:
    session_path = tmp_path / "session.json"
    cached = tmp_path / "session.txt"
    cached.write_text("fake transcript", encoding="utf-8")
    session_manager.write_session(session_path, {
        "version": 1,
        "phase": "awaiting_input",
        "input_path": str(tmp_path / "Smith.pdf"),
        "cached_text_path": str(cached),
        "deponent_name": "John Smith",
        "deposition_date": "January 15, 2024",
        "deponent_type": "Plaintiff",
        "file_number": "3850.084",
        "topics": [
            {"id": 1, "title": "Pre-Accident History", "rank": 1, "discussion_density": "high"},
            {"id": 2, "title": "Mechanism Of Injury", "rank": 2, "discussion_density": "high"},
            {"id": 3, "title": "Damages", "rank": 3, "discussion_density": "medium"},
        ],
        "user_config": None,
    })
    return session_path


def test_dialog_loads_session_and_populates_topics(qtbot, tmp_path):
    session_path = _make_session(tmp_path)
    dlg = DepoSummaryConfigDialog(session_path)
    qtbot.addWidget(dlg)

    titles = [item.text() for item in dlg.topic_rows_in_order()]
    assert titles == ["Pre-Accident History", "Mechanism Of Injury", "Damages"]
    assert all(item.checkState() == Qt.Checked for item in dlg.topic_rows_in_order())
    assert dlg.deponent_label_edit.text() == "Plaintiff"
    assert dlg.bullets_spinbox.value() == 5
    assert dlg.cross_check_checkbox.isChecked()


def test_dialog_accept_writes_user_config_back_to_session(qtbot, tmp_path):
    session_path = _make_session(tmp_path)
    dlg = DepoSummaryConfigDialog(session_path)
    qtbot.addWidget(dlg)

    # Uncheck topic 2, rename topic 1, add a custom topic, change settings.
    items = list(dlg.topic_rows_in_order())
    items[1].setCheckState(Qt.Unchecked)
    items[0].setText("Pre-Accident Lower Back Treatment")
    dlg.added_topics_edit.setPlainText("Communications With Treating Providers\n")
    dlg.bullets_spinbox.setValue(7)
    dlg.deponent_label_edit.setText("Mr. Smith")
    dlg.custom_rules_edit.setPlainText("Use past tense.")
    dlg.cross_check_checkbox.setChecked(False)

    dlg.accept()

    loaded = session_manager.read_session(session_path)
    assert loaded["phase"] == "ready_for_summary"
    cfg = loaded["user_config"]
    assert cfg["selected_topics"] == ["Pre-Accident Lower Back Treatment", "Damages"]
    assert cfg["added_topics"] == ["Communications With Treating Providers"]
    assert cfg["bullets_per_topic"] == 7
    assert cfg["deponent_label"] == "Mr. Smith"
    assert cfg["custom_rules"] == "Use past tense."
    assert cfg["cross_check_enabled"] is False


def test_dialog_cancel_does_not_modify_session(qtbot, tmp_path):
    session_path = _make_session(tmp_path)
    before = session_path.read_text(encoding="utf-8")

    dlg = DepoSummaryConfigDialog(session_path)
    qtbot.addWidget(dlg)
    list(dlg.topic_rows_in_order())[0].setCheckState(Qt.Unchecked)
    dlg.bullets_spinbox.setValue(99)
    dlg.reject()

    after = session_path.read_text(encoding="utf-8")
    assert before == after


def test_dialog_atomic_write_preserves_original_on_failure(qtbot, tmp_path, monkeypatch):
    session_path = _make_session(tmp_path)
    before = session_path.read_text(encoding="utf-8")

    dlg = DepoSummaryConfigDialog(session_path)
    qtbot.addWidget(dlg)
    dlg.bullets_spinbox.setValue(9)

    def boom(*a, **kw):
        raise OSError("simulated")

    monkeypatch.setattr(os, "replace", boom)
    with pytest.raises(OSError):
        dlg.accept()

    # Session file untouched
    after = session_path.read_text(encoding="utf-8")
    assert before == after


def test_dialog_accept_blocks_when_no_topics_selected(qtbot, tmp_path, monkeypatch):
    """If the user unchecks all topics and adds none, accept must not write user_config."""
    from PySide6.QtWidgets import QMessageBox

    session_path = _make_session(tmp_path)
    before = session_path.read_text(encoding="utf-8")

    dlg = DepoSummaryConfigDialog(session_path)
    qtbot.addWidget(dlg)

    for item in dlg.topic_rows_in_order():
        item.setCheckState(Qt.Unchecked)
    dlg.added_topics_edit.setPlainText("")

    # Stub QMessageBox.warning so the test doesn't block on a real dialog.
    warning_called = []
    monkeypatch.setattr(QMessageBox, "warning",
                        lambda *a, **kw: warning_called.append(a) or QMessageBox.Ok)

    dlg.accept()

    assert warning_called, "QMessageBox.warning should have been shown"
    # Session JSON must be unchanged: phase still 'awaiting_input', user_config still None.
    after = session_path.read_text(encoding="utf-8")
    assert before == after


def test_dialog_topic_drag_reorder_changes_selected_topics_order(qtbot, tmp_path):
    """Programmatically reorder topics via QListWidget and verify selected_topics order."""
    session_path = _make_session(tmp_path)
    dlg = DepoSummaryConfigDialog(session_path)
    qtbot.addWidget(dlg)

    # Initial fixture order: Pre-Accident History, Mechanism Of Injury, Damages.
    # Move "Damages" (row 2) to row 0.
    item = dlg.topics_list.takeItem(2)
    dlg.topics_list.insertItem(0, item)

    dlg.accept()

    loaded = session_manager.read_session(session_path)
    cfg = loaded["user_config"]
    assert cfg["selected_topics"] == ["Damages", "Pre-Accident History", "Mechanism Of Injury"]


def test_dialog_bias_combo_defaults_to_neutral_and_writes_neutral_to_session(qtbot, tmp_path):
    session_path = _make_session(tmp_path)
    dlg = DepoSummaryConfigDialog(session_path)
    qtbot.addWidget(dlg)

    assert dlg.bias_combo.currentData() == "neutral"
    assert dlg.bias_custom_edit.isVisible() is False

    dlg.accept()
    cfg = session_manager.read_session(session_path)["user_config"]
    assert cfg["bias"] == "neutral"
    assert cfg["bias_custom"] == ""


def test_dialog_bias_custom_reveals_text_field_and_round_trips(qtbot, tmp_path):
    session_path = _make_session(tmp_path)
    dlg = DepoSummaryConfigDialog(session_path)
    qtbot.addWidget(dlg)
    dlg.show()

    custom_idx = next(
        i for i in range(dlg.bias_combo.count())
        if dlg.bias_combo.itemData(i) == "custom"
    )
    dlg.bias_combo.setCurrentIndex(custom_idx)
    assert dlg.bias_custom_edit.isVisible() is True

    dlg.bias_custom_edit.setText("Highlight inconsistencies in injury testimony.")
    dlg.accept()

    cfg = session_manager.read_session(session_path)["user_config"]
    assert cfg["bias"] == "custom"
    assert cfg["bias_custom"] == "Highlight inconsistencies in injury testimony."


def test_dialog_bias_switching_away_from_custom_clears_custom_field(qtbot, tmp_path):
    session_path = _make_session(tmp_path)
    dlg = DepoSummaryConfigDialog(session_path)
    qtbot.addWidget(dlg)

    custom_idx = next(
        i for i in range(dlg.bias_combo.count())
        if dlg.bias_combo.itemData(i) == "custom"
    )
    pro_def_idx = next(
        i for i in range(dlg.bias_combo.count())
        if dlg.bias_combo.itemData(i) == "pro_defense"
    )
    dlg.bias_combo.setCurrentIndex(custom_idx)
    dlg.bias_custom_edit.setText("Some custom directive")
    dlg.bias_combo.setCurrentIndex(pro_def_idx)
    assert dlg.bias_custom_edit.isVisible() is False
    assert dlg.bias_custom_edit.text() == ""

    dlg.accept()
    cfg = session_manager.read_session(session_path)["user_config"]
    assert cfg["bias"] == "pro_defense"
    assert cfg["bias_custom"] == ""


def test_dialog_context_docs_accept_via_add_button(qtbot, tmp_path):
    """Programmatically exercise the same path the 'Add files…' button does."""
    session_path = _make_session(tmp_path)
    pdf_path = tmp_path / "complaint.pdf"
    pdf_path.write_bytes(b"%PDF-1.4\n")
    docx_path = tmp_path / "med_summary.docx"
    docx_path.write_bytes(b"PK\x03\x04")

    dlg = DepoSummaryConfigDialog(session_path)
    qtbot.addWidget(dlg)
    dlg._append_context_path(pdf_path)
    dlg._append_context_path(docx_path)
    assert dlg.context_docs_list.count() == 2

    dlg.accept()
    cfg = session_manager.read_session(session_path)["user_config"]
    paths = cfg["context_doc_paths"]
    assert any("complaint.pdf" in p for p in paths)
    assert any("med_summary.docx" in p for p in paths)


def test_dialog_context_docs_reject_unsupported_extensions(qtbot, tmp_path):
    """Dropping a .txt file is rejected and surfaces a status message."""
    from PySide6.QtCore import QUrl
    session_path = _make_session(tmp_path)
    txt_path = tmp_path / "notes.txt"
    txt_path.write_text("nope", encoding="utf-8")

    dlg = DepoSummaryConfigDialog(session_path)
    qtbot.addWidget(dlg)
    dlg._add_context_paths_from_urls([QUrl.fromLocalFile(str(txt_path))])

    assert dlg.context_docs_list.count() == 0
    assert "Unsupported" in dlg._ctx_status_label.text()
    assert "notes.txt" in dlg._ctx_status_label.text()


def test_dialog_context_docs_remove_button_drops_path(qtbot, tmp_path):
    session_path = _make_session(tmp_path)
    pdf_a = tmp_path / "a.pdf"
    pdf_a.write_bytes(b"%PDF-1.4\n")
    pdf_b = tmp_path / "b.pdf"
    pdf_b.write_bytes(b"%PDF-1.4\n")

    dlg = DepoSummaryConfigDialog(session_path)
    qtbot.addWidget(dlg)
    dlg._append_context_path(pdf_a)
    dlg._append_context_path(pdf_b)
    dlg._remove_context_path(pdf_a)
    assert dlg.context_docs_list.count() == 1
    assert pdf_a not in dlg._context_doc_paths
    assert pdf_b in dlg._context_doc_paths

    dlg.accept()
    cfg = session_manager.read_session(session_path)["user_config"]
    assert len(cfg["context_doc_paths"]) == 1
    assert "b.pdf" in cfg["context_doc_paths"][0]


def test_dialog_context_docs_drop_via_mime(qtbot, tmp_path):
    """Simulate a drop event built from QMimeData."""
    from PySide6.QtCore import QMimeData, QPoint, QUrl
    from PySide6.QtGui import QDropEvent
    from PySide6.QtCore import Qt as _Qt

    session_path = _make_session(tmp_path)
    pdf_path = tmp_path / "ev.pdf"
    pdf_path.write_bytes(b"%PDF-1.4\n")

    dlg = DepoSummaryConfigDialog(session_path)
    qtbot.addWidget(dlg)

    mime = QMimeData()
    mime.setUrls([QUrl.fromLocalFile(str(pdf_path))])
    event = QDropEvent(
        QPoint(0, 0), _Qt.CopyAction, mime, _Qt.LeftButton, _Qt.NoModifier
    )
    dlg._context_drop(event)
    qtbot.wait(50)  # pump event loop so the deferred QTimer.singleShot(0, ...) fires

    assert dlg.context_docs_list.count() == 1
    assert pdf_path in dlg._context_doc_paths


def test_dialog_context_docs_missing_at_accept_are_silently_dropped(qtbot, tmp_path):
    session_path = _make_session(tmp_path)
    pdf_path = tmp_path / "gone.pdf"
    pdf_path.write_bytes(b"%PDF-1.4\n")

    dlg = DepoSummaryConfigDialog(session_path)
    qtbot.addWidget(dlg)
    dlg._append_context_path(pdf_path)
    pdf_path.unlink()  # remove file after it was added

    dlg.accept()
    cfg = session_manager.read_session(session_path)["user_config"]
    assert cfg["context_doc_paths"] == []


def test_dialog_compute_case_folder_extracts_matter_root(qtbot, tmp_path):
    """Helper returns the case root (matter folder) given the deposition input path."""
    session_path = _make_session(tmp_path)
    # _make_session sets input_path to "tmp_path/Smith.pdf" — no matter folder pattern.
    # Override input_path with a realistic case structure for this test.
    session = session_manager.read_session(session_path)
    session["input_path"] = r"Z:\Shared\Current Clients\3800- NATIONWIDE\3850\084 - Dudash\Depositions\Smith Depo.pdf"
    session_manager.write_session(session_path, session)

    dlg = DepoSummaryConfigDialog(session_path)
    qtbot.addWidget(dlg)
    assert dlg._compute_case_folder().endswith("084 - Dudash")


def test_dialog_compute_case_folder_falls_back_to_parent_when_no_matter_pattern(qtbot, tmp_path):
    session_path = _make_session(tmp_path)
    # The fixture's input_path is just "<tmp>/Smith.pdf" — no folder named with 3 digits.
    dlg = DepoSummaryConfigDialog(session_path)
    qtbot.addWidget(dlg)
    fallback = dlg._compute_case_folder()
    # Should be the tmp_path itself (parent of the fake pdf)
    assert fallback == str(tmp_path)
