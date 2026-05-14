"""Tests for the DepoSummaryConfigDialog popup."""

import json
import os
from pathlib import Path
from unittest.mock import patch

import pytest

pytest.importorskip("pytestqt")

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

    titles = [row.title_edit.text() for row in dlg.topic_rows_in_order()]
    assert titles == ["Pre-Accident History", "Mechanism Of Injury", "Damages"]
    assert all(row.checkbox.isChecked() for row in dlg.topic_rows_in_order())
    assert dlg.deponent_label_edit.text() == "Plaintiff"
    assert dlg.bullets_spinbox.value() == 5
    assert dlg.cross_check_checkbox.isChecked()


def test_dialog_accept_writes_user_config_back_to_session(qtbot, tmp_path):
    session_path = _make_session(tmp_path)
    dlg = DepoSummaryConfigDialog(session_path)
    qtbot.addWidget(dlg)

    # Uncheck topic 2, rename topic 1, add a custom topic, change settings.
    rows = list(dlg.topic_rows_in_order())
    rows[1].checkbox.setChecked(False)
    rows[0].title_edit.setText("Pre-Accident Lower Back Treatment")
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
    list(dlg.topic_rows_in_order())[0].checkbox.setChecked(False)
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

    for row in dlg.topic_rows_in_order():
        row.checkbox.setChecked(False)
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
    # Move "Damages" (row 2) to row 0. takeItem destroys the row widget, so we
    # rebuild it as a fresh _TopicRow to mirror what happens during a real drag.
    from icharlotte_core.ui.depo_summary_config_dialog import _TopicRow
    from PySide6.QtWidgets import QListWidgetItem
    dlg.topics_list.takeItem(2)
    new_row = _TopicRow("Damages")
    new_item = QListWidgetItem()
    new_item.setSizeHint(new_row.sizeHint())
    dlg.topics_list.insertItem(0, new_item)
    dlg.topics_list.setItemWidget(new_item, new_row)

    dlg.accept()

    loaded = session_manager.read_session(session_path)
    cfg = loaded["user_config"]
    assert cfg["selected_topics"] == ["Damages", "Pre-Accident History", "Mechanism Of Injury"]
