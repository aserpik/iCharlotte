"""Generate Motion task is registered in the launcher (spec #7)."""
from icharlotte_core.ui.wizard.registry import TASK_REGISTRY, get_task


def test_generate_motion_is_registered():
    assert "generate_motion" in TASK_REGISTRY


def test_generate_motion_in_motions_category():
    spec = get_task("generate_motion")
    assert spec.category == "Motions"


def test_generate_motion_is_in_process_task():
    # In-process (no subprocess script), like Oppose a Motion.
    assert get_task("generate_motion").script_name == ""


def test_generate_motion_title_and_keywords():
    spec = get_task("generate_motion")
    assert "motion" in spec.title.lower()
    kws = [k.lower() for k in spec.keywords]
    assert any("demurrer" in k for k in kws)
    assert any("compel" in k for k in kws)
