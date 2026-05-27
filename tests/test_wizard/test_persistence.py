"""Tests for WizardStatePersistence."""
import json
import os
import pytest

from icharlotte_core.ui.wizard.persistence import WizardStatePersistence


def test_load_missing_file_returns_default(tmp_path):
    p = WizardStatePersistence(str(tmp_path))
    data = p.load()
    assert data["version"] == 1
    assert data["open_tabs"] == []
    assert data["recent_tasks"] == []


def test_save_then_load_roundtrips(tmp_path):
    p = WizardStatePersistence(str(tmp_path))
    p.set_open_tabs([
        {"task_id": "summarize_documents", "instance_suffix": "", "files": ["a.pdf"],
         "settings": {}, "page": "settings", "output_path": None},
    ])
    p.add_recent_task({
        "task_id": "summarize_documents", "title": "Summarize Documents",
        "files": ["a.pdf"], "settings": {},
        "output_path": "NOTES/AI Output/x.docx",
        "completed_at": "2026-05-15T10:42:00",
    })
    p.save()

    p2 = WizardStatePersistence(str(tmp_path))
    data = p2.load()
    assert len(data["open_tabs"]) == 1
    assert data["open_tabs"][0]["task_id"] == "summarize_documents"
    assert len(data["recent_tasks"]) == 1


def test_recent_tasks_capped_at_20(tmp_path):
    p = WizardStatePersistence(str(tmp_path))
    for i in range(25):
        p.add_recent_task({
            "task_id": "summarize_documents", "title": f"Run {i}",
            "files": [], "settings": {}, "output_path": "x.docx",
            "completed_at": f"2026-05-15T{i:02d}:00:00",
        })
    p.save()
    data = WizardStatePersistence(str(tmp_path)).load()
    assert len(data["recent_tasks"]) == 20
    # Newest first.
    titles = [t["title"] for t in data["recent_tasks"]]
    assert titles[0] == "Run 24"
    assert titles[-1] == "Run 5"


def test_atomic_write_uses_tmp_then_rename(tmp_path):
    p = WizardStatePersistence(str(tmp_path))
    p.set_open_tabs([])
    p.save()
    state_file = os.path.join(str(tmp_path), "NOTES", "AI OUTPUT", ".icharlotte", "wizard_state.json")
    assert os.path.exists(state_file)
    # tmp file should not be lingering.
    assert not os.path.exists(state_file + ".tmp")


def test_corrupt_file_falls_back_to_default(tmp_path):
    folder = os.path.join(str(tmp_path), "NOTES", "AI OUTPUT", ".icharlotte")
    os.makedirs(folder)
    with open(os.path.join(folder, "wizard_state.json"), "w") as f:
        f.write("{ not json")
    p = WizardStatePersistence(str(tmp_path))
    data = p.load()
    assert data["open_tabs"] == []
    assert data["recent_tasks"] == []


def test_readme_created_on_first_save(tmp_path):
    p = WizardStatePersistence(str(tmp_path))
    p.set_open_tabs([])
    p.save()
    readme = os.path.join(str(tmp_path), "NOTES", "AI OUTPUT", ".icharlotte", "README.txt")
    assert os.path.exists(readme)


def test_folder_path_under_notes_ai_output(tmp_path):
    """`.icharlotte` lives under NOTES/AI OUTPUT, not at the case root."""
    p = WizardStatePersistence(str(tmp_path))
    p.set_open_tabs([])
    p.save()
    # Folder must NOT be created at case root.
    assert not os.path.exists(os.path.join(str(tmp_path), ".icharlotte"))
    # Folder must be under NOTES/AI OUTPUT.
    assert os.path.isdir(os.path.join(str(tmp_path), "NOTES", "AI OUTPUT", ".icharlotte"))
