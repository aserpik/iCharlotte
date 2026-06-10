"""Tests for the generic wizard Settings page."""

import os

import pytest

pytest.importorskip("pytestqt")

from PySide6.QtWidgets import QFrame

from icharlotte_core.ui.wizard.pages import settings_page as settings_page_mod
from icharlotte_core.ui.wizard.pages.settings_page import SettingsPage
from icharlotte_core.ui.wizard.registry import get_task


def test_settings_page_uses_source_review_workbench_layout(qtbot, tmp_path):
    selected = tmp_path / "pleading.pdf"
    selected.write_bytes(b"%PDF-1.4\n")

    page = SettingsPage(
        get_task("summarize_documents"),
        files=[str(selected)],
        case_root=str(tmp_path),
    )
    qtbot.addWidget(page)

    assert page.findChild(QFrame, "sourceReviewWorkbench") is not None
    assert page.findChild(QFrame, "sourceQueuePanel") is not None
    assert page.findChild(QFrame, "taskSetupPanel") is not None
    assert page.empty_hint.objectName() == "sourceDropZoneHint"
    assert page.source_count_value.text() == "1 file"
    assert page.proceed_btn.isEnabled() is True


def test_add_files_defaults_to_task_default_folder(qtbot, tmp_path, monkeypatch):
    target_dir = tmp_path / "DISCOVERY" / "RESPONSES"
    target_dir.mkdir(parents=True)
    selected = target_dir / "responses.pdf"
    selected.write_bytes(b"%PDF-1.4\n")
    captured = {}

    def fake_get_open_file_names(parent, title, start_dir, file_filter):
        captured["title"] = title
        captured["start_dir"] = start_dir
        captured["file_filter"] = file_filter
        return ([str(selected)], "")

    monkeypatch.setattr(
        settings_page_mod.QFileDialog,
        "getOpenFileNames",
        fake_get_open_file_names,
    )

    page = SettingsPage(
        get_task("summarize_discovery"),
        files=[],
        case_root=str(tmp_path),
    )
    qtbot.addWidget(page)

    page._on_add_files()

    assert os.path.normpath(captured["start_dir"]) == os.path.normpath(str(target_dir))
    assert page.files == [str(selected)]


def test_add_files_for_document_summary_defaults_to_case_root(qtbot, tmp_path, monkeypatch):
    selected = tmp_path / "pleading.pdf"
    selected.write_bytes(b"%PDF-1.4\n")
    captured = {}

    def fake_get_open_file_names(parent, title, start_dir, file_filter):
        captured["start_dir"] = start_dir
        return ([str(selected)], "")

    monkeypatch.setattr(
        settings_page_mod.QFileDialog,
        "getOpenFileNames",
        fake_get_open_file_names,
    )

    page = SettingsPage(
        get_task("summarize_documents"),
        files=[],
        case_root=str(tmp_path),
    )
    qtbot.addWidget(page)

    page._on_add_files()

    assert os.path.normpath(captured["start_dir"]) == os.path.normpath(str(tmp_path))
    assert page.files == [str(selected)]
