"""Tests for Summarize Depositions settings-page source selection."""

import os

import pytest

pytest.importorskip("pytestqt")

from icharlotte_core.ui.wizard.pages import deposition_settings_page as depo_page_mod
from icharlotte_core.ui.wizard.pages.deposition_settings_page import (
    DepositionSettingsPage,
)
from icharlotte_core.ui.wizard.registry import get_task


def test_empty_deposition_settings_waits_for_transcript(qtbot, tmp_path):
    page = DepositionSettingsPage(
        get_task("summarize_depositions"),
        files=[],
        case_root=str(tmp_path),
    )
    qtbot.addWidget(page)

    assert "Add a transcript" in page._small_status.text()
    assert page.proceed_btn.isEnabled() is False


def test_add_file_defaults_to_transcripts_folder_and_starts_phase1(
    qtbot, tmp_path, monkeypatch
):
    target_dir = tmp_path / "DISCOVERY" / "TRANSCRIPTS"
    target_dir.mkdir(parents=True)
    selected = target_dir / "smith_depo.pdf"
    selected.write_bytes(b"%PDF-1.4\n")
    captured = {}

    def fake_get_open_file_names(parent, title, start_dir, file_filter):
        captured["title"] = title
        captured["start_dir"] = start_dir
        return ([str(selected)], "")

    monkeypatch.setattr(
        depo_page_mod.QFileDialog,
        "getOpenFileNames",
        fake_get_open_file_names,
    )

    page = DepositionSettingsPage(
        get_task("summarize_depositions"),
        files=[],
        case_root=str(tmp_path),
    )
    qtbot.addWidget(page)
    requested = []
    page.restart_phase1_requested.connect(lambda files: requested.append(list(files)))

    page._on_add_files()

    assert os.path.normpath(captured["start_dir"]) == os.path.normpath(str(target_dir))
    assert page.files == [str(selected)]
    assert requested == [[str(selected)]]


def test_removing_deposition_file_cancels_phase1(qtbot, tmp_path):
    selected = tmp_path / "depo.pdf"
    selected.write_bytes(b"%PDF-1.4\n")
    page = DepositionSettingsPage(
        get_task("summarize_depositions"),
        files=[str(selected)],
        case_root=str(tmp_path),
    )
    qtbot.addWidget(page)
    requested = []
    page.restart_phase1_requested.connect(lambda files: requested.append(list(files)))

    page.files_list.setCurrentRow(0)
    page._on_remove_files()

    assert page.files == []
    assert requested == [[]]
