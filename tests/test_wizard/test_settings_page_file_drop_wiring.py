"""File-add adapters used by custom wizard settings pages."""

from __future__ import annotations

import pytest

pytest.importorskip("pytestqt")

from PySide6.QtCore import QEvent, QMimeData, QUrl

from icharlotte_core.ui.wizard.pages.depo_prep_settings_page import (
    DepoPrepSettingsPage,
)
from icharlotte_core.ui.wizard.pages.deposition_settings_page import (
    DepositionSettingsPage,
)
from icharlotte_core.ui.wizard.pages.generate_motion_page import (
    GenerateMotionSettingsPage,
)
from icharlotte_core.ui.wizard.pages.mediation_brief_page import (
    MediationBriefSettingsPage,
)
from icharlotte_core.ui.wizard.pages.oppose_motion_page import (
    OpposeMotionSettingsPage,
)
from icharlotte_core.ui.wizard.pages.med_chron_settings_page import (
    MedChronSettingsPage,
)
from icharlotte_core.ui.wizard.registry import get_task


class _FakeDropEvent:
    def __init__(self, paths: list[str]):
        self._mime = QMimeData()
        self._mime.setUrls([QUrl.fromLocalFile(path) for path in paths])
        self.accepted = False

    def type(self):
        return QEvent.Type.Drop

    def mimeData(self):
        return self._mime

    def acceptProposedAction(self):
        self.accepted = True

    def ignore(self):
        return None


def _send_drop(widget, paths: list[str]) -> None:
    handlers = getattr(widget, "_icharlotte_file_drop_handlers", [])
    assert handlers, f"{widget!r} has no file drop handler"
    event = _FakeDropEvent(paths)
    assert handlers[0].eventFilter(widget, event) is True
    assert event.accepted is True


def test_deposition_drop_replaces_single_transcript_slot(qtbot, tmp_path):
    first = tmp_path / "first.pdf"
    second = tmp_path / "second.pdf"
    first.write_bytes(b"%PDF-1.4\n")
    second.write_bytes(b"%PDF-1.4\n")
    page = DepositionSettingsPage(
        get_task("summarize_depositions"),
        files=[],
        case_root=str(tmp_path),
    )
    qtbot.addWidget(page)
    restarts: list[list[str]] = []
    page.restart_phase1_requested.connect(lambda files: restarts.append(list(files)))

    page.add_files([str(first), str(second)])

    assert page.files == [str(first)]
    assert restarts == [[str(first)]]


def test_med_chron_drop_replaces_single_chronology_slot(qtbot, tmp_path):
    first = tmp_path / "chronology.docx"
    second = tmp_path / "other.docx"
    first.write_text("chron", encoding="utf-8")
    second.write_text("other", encoding="utf-8")
    page = MedChronSettingsPage(
        get_task("med_chron_analysis"),
        files=[],
        case_root=str(tmp_path),
    )
    qtbot.addWidget(page)
    restarts: list[list[str]] = []
    page.restart_phase1_requested.connect(lambda files: restarts.append(list(files)))

    page.add_files([str(first), str(second)])

    assert page.files == [str(first)]
    assert restarts == [[str(first)]]


def test_depo_prep_drop_targets_route_to_each_source_bucket(qtbot, tmp_path):
    deponent = tmp_path / "deponent.pdf"
    context = tmp_path / "complaint.pdf"
    deponent.write_bytes(b"%PDF-1.4\n")
    context.write_bytes(b"%PDF-1.4\n")
    page = DepoPrepSettingsPage(get_task("depo_prep"), files=[], case_root=str(tmp_path))
    qtbot.addWidget(page)

    _send_drop(page.deponent_files_list, [str(deponent)])
    _send_drop(page.context_files_list, [str(context)])

    settings = page.to_dict()
    assert settings["deponent_sources"] == [str(deponent)]
    assert settings["context_sources"] == [str(context)]


def test_generate_motion_drop_adds_supported_target_files(qtbot, tmp_path):
    supported = tmp_path / "motion_context.pdf"
    unsupported = tmp_path / "program.exe"
    supported.write_bytes(b"%PDF-1.4\n")
    unsupported.write_bytes(b"bin")
    page = GenerateMotionSettingsPage(str(tmp_path), "1234.001")
    qtbot.addWidget(page)

    page.add_target_files([str(supported), str(unsupported), str(supported)])

    assert page.current_target_files() == [str(supported)]


def test_oppose_motion_drop_routes_motion_and_context_files(qtbot, tmp_path):
    motion = tmp_path / "msj.pdf"
    context = tmp_path / "declaration.docx"
    unsupported = tmp_path / "image.png"
    motion.write_bytes(b"%PDF-1.4\n")
    context.write_text("doc", encoding="utf-8")
    unsupported.write_bytes(b"png")
    page = OpposeMotionSettingsPage(str(tmp_path), "1234.001", "", [])
    qtbot.addWidget(page)

    page.set_motion_file(str(motion))
    page.add_context_files([str(context), str(unsupported), str(context)])

    assert page.motion_file == str(motion)
    assert page.context_files == [str(context)]


def test_mediation_brief_drop_adds_source_files(qtbot, tmp_path):
    first = tmp_path / "brief_source.pdf"
    second = tmp_path / "email.msg"
    first.write_bytes(b"%PDF-1.4\n")
    second.write_text("msg", encoding="utf-8")
    page = MediationBriefSettingsPage(str(tmp_path), "1234.001")
    qtbot.addWidget(page)

    page.add_files([str(first), str(second), str(first)])

    assert page.current_files() == [str(first), str(second)]


def test_mediation_brief_drop_sets_caption_template(qtbot, tmp_path):
    caption = tmp_path / "caption.docx"
    ignored = tmp_path / "caption.pdf"
    caption.write_text("docx", encoding="utf-8")
    ignored.write_bytes(b"%PDF-1.4\n")
    page = MediationBriefSettingsPage(str(tmp_path), "1234.001")
    qtbot.addWidget(page)

    page.add_caption_files([str(ignored), str(caption)])

    assert page.caption_edit.text() == str(caption)
