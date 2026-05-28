import json
from unittest.mock import MagicMock, patch

import pytest

pytest.importorskip("pytestqt")
pytest.importorskip("PySide6")


@pytest.fixture
def spec():
    from icharlotte_core.ui.wizard.registry import TaskSpec
    return TaskSpec(task_id="depo_prep", title="Depo Prep", description="d",
                    icon_glyph="?", script_name="depo_prep.py",
                    phase1_args=["--phase=analyze"], phase2_flag="--phase=generate")


def test_settings_page_renders_all_controls(qtbot, spec, tmp_path):
    from icharlotte_core.ui.wizard.pages.depo_prep_settings_page import DepoPrepSettingsPage
    page = DepoPrepSettingsPage(spec, files=[], case_root=str(tmp_path))
    qtbot.addWidget(page)
    assert page.deponent_name_combo is not None
    assert page.deponent_role_edit is not None
    assert page.add_deponent_files_btn is not None
    assert page.add_context_files_btn is not None
    assert page.style_combo is not None  # or radio_group
    assert page.free_text_edit is not None
    assert page.analyze_btn is not None
    assert page.flag_strategic.isChecked() is True   # default ON
    assert page.flag_source_facts.isChecked() is True  # default ON
    assert page.flag_impeachment.isChecked() is False
    assert page.flag_objection.isChecked() is False
    # Analyze button disabled until deponent + at least one source.
    assert page.analyze_btn.isEnabled() is False


def test_settings_page_analyze_writes_config_and_emits(qtbot, spec, tmp_path):
    from icharlotte_core.ui.wizard.pages.depo_prep_settings_page import DepoPrepSettingsPage
    page = DepoPrepSettingsPage(spec, files=[], case_root=str(tmp_path))
    qtbot.addWidget(page)

    page.set_deponent_name("Jane Doe")
    page.set_deponent_role("Plaintiff")
    page.add_deponent_files([str(tmp_path / "fake_depo.pdf")])
    # Fake-fake the existence check by creating the file.
    (tmp_path / "fake_depo.pdf").write_bytes(b"")
    assert page.analyze_btn.isEnabled() is True

    captured = []
    page.proceed_requested.connect(lambda d: captured.append(d))
    with qtbot.waitSignal(page.proceed_requested, timeout=1000):
        page._on_analyze_clicked()

    # Settings page persists config to disk and reports its path as the single "file".
    files = page.files
    assert len(files) == 1
    assert files[0].endswith("config.json")
    cfg = json.loads(open(files[0], "r", encoding="utf-8").read())
    assert cfg["deponent_name"] == "Jane Doe"
    assert cfg["deponent_role"] == "Plaintiff"
    assert cfg["style"] in ("discovery", "lockdown", "expert", "friendly")
    assert "case_root" in cfg


def test_settings_page_reveals_topic_editor_on_attach_phase1_complete(qtbot, spec, tmp_path):
    from icharlotte_core.ui.wizard.pages.depo_prep_settings_page import DepoPrepSettingsPage

    # Pre-create session+topics files to simulate Phase 1 output.
    session_dir = tmp_path / "NOTES" / "AI Output" / "Depo Prep - Jane - 2026-05-27 1432"
    session_dir.mkdir(parents=True)
    (session_dir / "session.json").write_text(json.dumps({
        "deponent_name": "Jane Doe", "topics_warning": None,
    }), encoding="utf-8")
    (session_dir / "topics.json").write_text(json.dumps({
        "topics": [{"id": "t01", "title": "T1", "strategic_note": "s",
                    "relevant_digest_refs": [], "default_checked": True, "lawyer_added": False}],
    }), encoding="utf-8")

    page = DepoPrepSettingsPage(spec, files=[], case_root=str(tmp_path))
    qtbot.addWidget(page)
    page.show()
    page._on_phase1_complete(str(session_dir / "session.json"))

    assert page.topic_editor.isVisible() is True
    assert len(page.topic_editor.get_topics()) == 1
    assert page.generate_btn.isEnabled() is True


def test_settings_page_generate_emits_phase2_requested(qtbot, spec, tmp_path):
    from icharlotte_core.ui.wizard.pages.depo_prep_settings_page import DepoPrepSettingsPage

    session_dir = tmp_path / "session"
    session_dir.mkdir()
    (session_dir / "session.json").write_text(json.dumps({"deponent_name": "X"}), encoding="utf-8")
    (session_dir / "topics.json").write_text(json.dumps({"topics": [
        {"id": "t01", "title": "T", "strategic_note": "", "relevant_digest_refs": [],
         "default_checked": True, "lawyer_added": False},
    ]}), encoding="utf-8")

    page = DepoPrepSettingsPage(spec, files=[], case_root=str(tmp_path))
    qtbot.addWidget(page)
    page._on_phase1_complete(str(session_dir / "session.json"))

    captured = []
    page.phase2_requested.connect(lambda p: captured.append(p))
    with qtbot.waitSignal(page.phase2_requested, timeout=1000):
        page._on_generate_clicked()

    # Topic editor state was persisted to topics.json before emit.
    assert captured[0] == str(session_dir / "session.json")
    saved_topics = json.loads((session_dir / "topics.json").read_text(encoding="utf-8"))
    assert isinstance(saved_topics["topics"], list)
