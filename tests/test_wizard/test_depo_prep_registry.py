from icharlotte_core.ui.wizard.registry import TASK_REGISTRY, get_task


def test_depo_prep_task_registered():
    spec = get_task("depo_prep")
    assert spec.title == "Depo Prep"
    assert spec.script_name == "depo_prep.py"
    assert "--phase=analyze" in spec.phase1_args
    assert spec.phase2_flag == "--phase=generate"


def test_depo_prep_settings_page_cls():
    from icharlotte_core.ui.wizard.pages.depo_prep_settings_page import DepoPrepSettingsPage
    spec = get_task("depo_prep")
    assert spec.settings_page_cls is DepoPrepSettingsPage


def test_depo_prep_output_page_cls():
    from icharlotte_core.ui.wizard.pages.depo_prep_output_page import DepoPrepOutputPage
    spec = get_task("depo_prep")
    assert spec.output_page_cls is DepoPrepOutputPage
