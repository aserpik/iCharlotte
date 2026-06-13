"""Tests for webcompanion.task_defs."""
import json
from pathlib import Path

from webcompanion import task_defs as T


def test_seven_tasks_registered():
    assert set(T.TASKS) == {
        "summarize_documents", "summarize_discovery", "summarize_depositions",
        "depo_prep", "medical_records", "med_chron_analysis", "separate",
    }


def test_phase1_argv_plain():
    task = T.TASKS["summarize_documents"]
    argv = T.build_phase1_argv(task, r"E:\case\doc.pdf")
    assert argv[0].endswith("summarize.py") and argv[-1] == r"E:\case\doc.pdf"
    assert len(argv) == 2


def test_phase1_argv_med_chron_has_prep_flag():
    task = T.TASKS["med_chron_analysis"]
    argv = T.build_phase1_argv(task, r"E:\case\chron.docx")
    assert argv[1] == "--phase=prep"


def test_phase2_argv():
    task = T.TASKS["summarize_depositions"]
    argv = T.build_phase2_argv(task, r"C:\logs\s.json")
    assert argv[1] == "--phase=summary" and argv[2] == r"C:\logs\s.json"
    assert T.build_phase2_argv(T.TASKS["depo_prep"], "x")[1] == "--phase=generate"


def test_write_depo_prep_config_roundtrip():
    cfg = {"deponent_name": "Smith", "style": "lockdown"}
    path = T.write_depo_prep_config(cfg)
    loaded = json.loads(Path(path).read_text(encoding="utf-8"))
    assert loaded["deponent_name"] == "Smith" and Path(path).name == "config.json"


def test_depo_prep_topics_roundtrip(tmp_path):
    session = tmp_path / "session.json"
    session.write_text("{}", encoding="utf-8")
    topics = [{"id": "t01", "title": "Background", "strategic_note": "",
               "relevant_digest_refs": [], "default_checked": True,
               "lawyer_added": False}]
    T.write_depo_prep_topics(str(session), topics)
    assert T.read_depo_prep_topics(str(session)) == topics


def test_script_path_points_into_scripts_dir():
    p = Path(T.script_path("summarize.py"))
    assert p.parent.name == "Scripts" and p.exists()
