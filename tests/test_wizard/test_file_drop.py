"""Tests for wizard settings-page file drag/drop helpers."""

from __future__ import annotations

import pytest

pytest.importorskip("pytestqt")

from PySide6.QtCore import QEvent, QMimeData, QUrl

from icharlotte_core.ui.wizard.file_drop import (
    FileDropTarget,
    local_file_paths_from_mime_data,
)
from icharlotte_core.ui.wizard.pages.settings_page import SettingsPage
from icharlotte_core.ui.wizard.registry import get_task


class _FakeDropEvent:
    def __init__(self, event_type: QEvent.Type, paths: list[str]):
        self._event_type = event_type
        self._mime = QMimeData()
        self._mime.setUrls([QUrl.fromLocalFile(path) for path in paths])
        self.accepted = False
        self.ignored = False

    def type(self):
        return self._event_type

    def mimeData(self):
        return self._mime

    def acceptProposedAction(self):
        self.accepted = True

    def ignore(self):
        self.ignored = True


def test_local_file_paths_from_mime_data_ignores_dirs_and_dedupes(tmp_path):
    first = tmp_path / "one.pdf"
    second = tmp_path / "two.docx"
    first.write_bytes(b"%PDF-1.4\n")
    second.write_text("text", encoding="utf-8")
    folder = tmp_path / "folder"
    folder.mkdir()

    mime = QMimeData()
    mime.setUrls(
        [
            QUrl.fromLocalFile(str(first)),
            QUrl.fromLocalFile(str(folder)),
            QUrl("https://example.test/not-local.pdf"),
            QUrl.fromLocalFile(str(first)),
            QUrl.fromLocalFile(str(second)),
        ]
    )

    assert local_file_paths_from_mime_data(mime) == [str(first), str(second)]


def test_file_drop_target_accepts_drag_and_calls_callback(tmp_path):
    source = tmp_path / "source.pdf"
    source.write_bytes(b"%PDF-1.4\n")
    received: list[list[str]] = []
    target = FileDropTarget(lambda paths: received.append(paths))

    drag_event = _FakeDropEvent(QEvent.Type.DragEnter, [str(source)])
    assert target.eventFilter(None, drag_event) is True
    assert drag_event.accepted is True

    drop_event = _FakeDropEvent(QEvent.Type.Drop, [str(source)])
    assert target.eventFilter(None, drop_event) is True
    assert drop_event.accepted is True
    assert received == [[str(source)]]


def test_generic_settings_page_drop_appends_unique_files(qtbot, tmp_path):
    first = tmp_path / "pleading.pdf"
    second = tmp_path / "records.docx"
    first.write_bytes(b"%PDF-1.4\n")
    second.write_text("text", encoding="utf-8")

    page = SettingsPage(
        get_task("summarize_documents"),
        files=[str(first)],
        case_root=str(tmp_path),
    )
    qtbot.addWidget(page)

    page.add_files([str(first), str(second)])

    assert page.files == [str(first), str(second)]
    assert page.files_label.text() == "Source Queue (2)"
    assert page.proceed_btn.isEnabled() is True


def test_settings_page_file_list_viewport_accepts_drops(qtbot, tmp_path):
    page = SettingsPage(
        get_task("summarize_documents"),
        files=[],
        case_root=str(tmp_path),
    )
    qtbot.addWidget(page)

    assert page.files_list.acceptDrops() is True
    assert page.files_list.viewport().acceptDrops() is True
    assert getattr(page.files_list.viewport(), "_icharlotte_file_drop_handlers", [])
