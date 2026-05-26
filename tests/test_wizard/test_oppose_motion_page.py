import pytest

pytest.importorskip("pytestqt")

from unittest.mock import patch

from icharlotte_core.opposition.models import (
    CitationVerification,
    DraftDocument,
    MotionMetadata,
    OutlineNode,
)
from icharlotte_core.ui.wizard.pages.oppose_motion_page import (
    OpposeMotionOutputPage,
    OpposeMotionSettingsPage,
    OpposeMotionTaskTab,
    SETTINGS_PAGE_CONFIRM,
    TASK_PAGE_OUTPUT,
    build_oppose_motion_tab,
)


def test_confirmation_blocks_missing_required_fields(qtbot):
    page = OpposeMotionSettingsPage(
        case_root="/tmp/case",
        file_number="0000.000",
        motion_file="/tmp/motion.pdf",
        context_files=[],
    )
    qtbot.addWidget(page)
    page.set_metadata(MotionMetadata())

    assert page.can_continue_to_outline() is False

    page.set_metadata(
        MotionMetadata(
            motion_type="Motion for Summary Judgment",
            relief_requested="summary judgment",
            principal_arguments=["no duty"],
        )
    )

    assert page.can_continue_to_outline() is True


def test_outline_items_start_checked(qtbot):
    page = OpposeMotionSettingsPage(
        case_root="/tmp/case",
        file_number="0000.000",
        motion_file="/tmp/motion.pdf",
        context_files=[],
    )
    qtbot.addWidget(page)
    page.set_outline(
        [
            OutlineNode(
                id="a",
                text="Main",
                children=[OutlineNode(id="a1", text="Sub")],
            )
        ]
    )

    assert page.outline_tree.topLevelItem(0).checkState(0).value == 2
    assert page.outline_tree.topLevelItem(0).child(0).checkState(0).value == 2


def test_task_tab_starts_on_confirmation_page(qtbot):
    tab = OpposeMotionTaskTab(
        spec=type("Spec", (), {"task_id": "oppose_motion", "title": "Oppose a Motion"})(),
        case_path="/tmp/case",
        file_number="0000.000",
        motion_file="/tmp/motion.pdf",
        context_files=[],
    )
    qtbot.addWidget(tab)

    assert tab.settings_page.currentIndex() == SETTINGS_PAGE_CONFIRM


def test_task_tab_loads_draft_result(qtbot):
    tab = OpposeMotionTaskTab(
        spec=type("Spec", (), {"task_id": "oppose_motion", "title": "Oppose a Motion"})(),
        case_path="/tmp/case",
        file_number="0000.000",
        motion_file="/tmp/motion.pdf",
        context_files=[],
    )
    qtbot.addWidget(tab)
    draft = DraftDocument(
        title="Opposition",
        body_text="Argument text",
        citations=[
            CitationVerification(
                citation_text="69 Cal.2d 108",
                status="verified",
                supporting_passage="ordinary care",
            )
        ],
    )

    tab._on_worker_finished(True, draft)

    assert tab.currentIndex() == TASK_PAGE_OUTPUT
    assert "Argument text" in tab.output_page.editor.toPlainText()
    assert "ordinary care" in tab.output_page.source_drawer.toPlainText()


def test_settings_from_dict_accepts_missing_state_and_refreshes_motion_label(qtbot):
    page = OpposeMotionSettingsPage(
        case_root="/tmp/case",
        file_number="0000.000",
        motion_file="/tmp/old.pdf",
        context_files=[],
    )
    qtbot.addWidget(page)

    page.from_dict(None)
    assert "old.pdf" in page.motion_label.text()

    page.from_dict({"motion_file": "/tmp/new_motion.pdf"})
    assert page.motion_file == "/tmp/new_motion.pdf"
    assert "new_motion.pdf" in page.motion_label.text()


def test_task_tab_stores_last_settings_when_run_is_requested(qtbot, monkeypatch):
    started = []

    class FakeWorker:
        def __init__(self, case_path, file_number, settings, parent=None):
            self.settings = settings
            self.progress = type("Signal", (), {"connect": lambda self, slot: None})()
            self.finished_result = type("Signal", (), {"connect": lambda self, slot: None})()

        def start(self):
            started.append(self.settings)

    monkeypatch.setattr(
        "icharlotte_core.ui.wizard.pages.oppose_motion_page.OpposeMotionWorker",
        FakeWorker,
    )
    tab = OpposeMotionTaskTab(
        spec=type("Spec", (), {"task_id": "oppose_motion", "title": "Oppose a Motion"})(),
        case_path="/tmp/case",
        file_number="0000.000",
        motion_file="/tmp/motion.pdf",
        context_files=[],
    )
    qtbot.addWidget(tab)

    tab._on_run({"motion_file": "/tmp/motion.pdf", "outline": []})

    assert tab._last_settings["motion_file"] == "/tmp/motion.pdf"
    assert started[0]["motion_file"] == "/tmp/motion.pdf"


def test_task_tab_does_not_start_second_worker_while_running(qtbot, monkeypatch):
    started = []

    class FakeWorker:
        def __init__(self, case_path, file_number, settings, parent=None):
            self.settings = settings
            self.progress = type("Signal", (), {"connect": lambda self, slot: None})()
            self.finished_result = type("Signal", (), {"connect": lambda self, slot: None})()

        def start(self):
            started.append(self.settings)

    monkeypatch.setattr(
        "icharlotte_core.ui.wizard.pages.oppose_motion_page.OpposeMotionWorker",
        FakeWorker,
    )
    tab = OpposeMotionTaskTab(
        spec=type("Spec", (), {"task_id": "oppose_motion", "title": "Oppose a Motion"})(),
        case_path="/tmp/case",
        file_number="0000.000",
        motion_file="/tmp/motion.pdf",
        context_files=[],
    )
    qtbot.addWidget(tab)

    tab._on_run({"motion_file": "/tmp/first.pdf"})
    tab._on_run({"motion_file": "/tmp/second.pdf"})

    assert len(started) == 1
    assert started[0]["motion_file"] == "/tmp/first.pdf"


def test_output_page_show_citation_updates_drawer(qtbot):
    page = OpposeMotionOutputPage()
    qtbot.addWidget(page)
    draft = DraftDocument(
        title="Opposition",
        body_text="Text",
        citations=[
            CitationVerification(
                citation_text="69 Cal.2d 108",
                normalized_citation="69 Cal.2d 108",
                status="verified",
                case_name="Rowland v. Christian",
                court="California Supreme Court",
                date="1968-08-08",
                opinion_url="https://www.courtlistener.com/opinion/123/",
                supporting_passage="ordinary care language",
                warning="warning text",
            )
        ],
    )
    page.show_result(draft)
    page.show_citation(0)

    drawer_text = page.source_drawer.toPlainText()
    assert "69 Cal.2d 108" in drawer_text
    assert "Normalized: 69 Cal.2d 108" in drawer_text
    assert "Status: verified" in drawer_text
    assert "Rowland v. Christian" in drawer_text
    assert "California Supreme Court" in drawer_text
    assert "1968-08-08" in drawer_text
    assert "https://www.courtlistener.com/opinion/123/" in drawer_text
    assert "ordinary care language" in drawer_text
    assert "warning text" in drawer_text


def test_output_editor_is_read_only(qtbot):
    page = OpposeMotionOutputPage()
    qtbot.addWidget(page)

    assert page.editor.isReadOnly() is True


def test_save_as_uses_dialog_and_does_not_save_when_cancelled(qtbot, tmp_path):
    page = OpposeMotionOutputPage()
    qtbot.addWidget(page)
    source = tmp_path / "preview.docx"
    source.write_bytes(b"docx-bytes")
    page.show_result(
        DraftDocument(
            title="Opposition",
            body_text="Text",
            preview_path=str(source),
        )
    )

    with patch(
        "icharlotte_core.ui.wizard.pages.oppose_motion_page.QFileDialog.getSaveFileName",
        return_value=("", ""),
    ) as dialog, patch(
        "icharlotte_core.ui.wizard.pages.oppose_motion_page.shutil.copyfile"
    ) as copyfile, patch(
        "icharlotte_core.ui.wizard.pages.oppose_motion_page.QMessageBox.information"
    ) as info:
        page.save_as()

    assert dialog.called
    copyfile.assert_not_called()
    info.assert_not_called()


def test_save_as_defaults_outside_internal_preview_folder(qtbot, tmp_path):
    page = OpposeMotionOutputPage()
    qtbot.addWidget(page)
    preview = (
        tmp_path
        / "NOTES"
        / "AI OUTPUT"
        / ".icharlotte"
        / "wizard_previews"
        / "oppose_motion"
        / "preview.docx"
    )
    preview.parent.mkdir(parents=True)
    preview.write_bytes(b"docx-bytes")
    page.show_result(
        DraftDocument(
            title="Opposition",
            body_text="Text",
            preview_path=str(preview),
        )
    )

    with patch(
        "icharlotte_core.ui.wizard.pages.oppose_motion_page.QFileDialog.getSaveFileName",
        return_value=("", ""),
    ) as dialog:
        page.save_as()

    suggested_path = dialog.call_args.args[2]
    assert ".icharlotte" not in suggested_path
    assert suggested_path.endswith("Opposition.docx")


def test_save_as_appends_docx_and_copies_preview(qtbot, tmp_path):
    page = OpposeMotionOutputPage()
    qtbot.addWidget(page)
    source = tmp_path / "preview.docx"
    source.write_bytes(b"docx-bytes")
    target = tmp_path / "saved-opposition"
    page.show_result(
        DraftDocument(
            title="Opposition",
            body_text="Text",
            preview_path=str(source),
        )
    )

    with patch(
        "icharlotte_core.ui.wizard.pages.oppose_motion_page.QFileDialog.getSaveFileName",
        return_value=(str(target), ""),
    ), patch(
        "icharlotte_core.ui.wizard.pages.oppose_motion_page.QMessageBox.information"
    ) as info:
        page.save_as()

    assert (tmp_path / "saved-opposition.docx").read_bytes() == b"docx-bytes"
    assert info.called


def test_save_as_reports_copy_errors(qtbot, tmp_path):
    page = OpposeMotionOutputPage()
    qtbot.addWidget(page)
    source = tmp_path / "preview.docx"
    source.write_bytes(b"docx-bytes")
    target = tmp_path / "saved.docx"
    page.show_result(
        DraftDocument(
            title="Opposition",
            body_text="Text",
            preview_path=str(source),
        )
    )

    with patch(
        "icharlotte_core.ui.wizard.pages.oppose_motion_page.QFileDialog.getSaveFileName",
        return_value=(str(target), ""),
    ), patch(
        "icharlotte_core.ui.wizard.pages.oppose_motion_page.shutil.copyfile",
        side_effect=OSError("locked"),
    ), patch(
        "icharlotte_core.ui.wizard.pages.oppose_motion_page.QMessageBox.critical"
    ) as critical:
        page.save_as()

    assert critical.called


def test_builder_rejects_cancelled_motion_picker(qtbot):
    spec = type("Spec", (), {"task_id": "oppose_motion", "title": "Oppose a Motion"})()
    with patch(
        "icharlotte_core.ui.wizard.pages.oppose_motion_page.QFileDialog.getOpenFileName",
        return_value=("", ""),
    ):
        tab = build_oppose_motion_tab(spec, "/tmp/case", "0000.000", None)
    assert tab is None


def test_builder_rejects_unsupported_motion_file(qtbot):
    spec = type("Spec", (), {"task_id": "oppose_motion", "title": "Oppose a Motion"})()
    with patch(
        "icharlotte_core.ui.wizard.pages.oppose_motion_page.QFileDialog.getOpenFileName",
        return_value=("/tmp/motion.txt", ""),
    ):
        with patch(
            "icharlotte_core.ui.wizard.pages.oppose_motion_page.QMessageBox.warning"
        ) as warning:
            tab = build_oppose_motion_tab(spec, "/tmp/case", "0000.000", None)
    assert tab is None
    assert warning.called


def test_builder_filters_unsupported_context_files(qtbot, tmp_path):
    spec = type("Spec", (), {"task_id": "oppose_motion", "title": "Oppose a Motion"})()
    motion = tmp_path / "motion.pdf"
    motion.write_bytes(b"")
    good_context = tmp_path / "facts.txt"
    bad_context = tmp_path / "notes.xlsx"
    good_context.write_text("facts")
    bad_context.write_text("spreadsheet")

    with patch(
        "icharlotte_core.ui.wizard.pages.oppose_motion_page.QFileDialog.getOpenFileName",
        return_value=(str(motion), ""),
    ), patch(
        "icharlotte_core.ui.wizard.pages.oppose_motion_page.QFileDialog.getOpenFileNames",
        return_value=([str(good_context), str(bad_context)], ""),
    ):
        tab = build_oppose_motion_tab(spec, str(tmp_path), "0000.000", None)

    qtbot.addWidget(tab)
    assert tab.settings_page.context_files == [str(good_context)]
