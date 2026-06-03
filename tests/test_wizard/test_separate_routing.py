"""Registry + routing wiring for the Separate wizard task."""
from icharlotte_core.ui.wizard.registry import get_task, TASK_REGISTRY
from icharlotte_core.ui.wizard.task_routing import (
    get_in_process_task_builder_name,
    is_in_process_task,
    requires_initial_file_picker,
)


def test_separate_in_registry():
    spec = get_task("separate")
    assert spec.title == "Separate Documents"
    assert spec.script_name == "separate.py"


def test_separate_is_in_process_no_picker():
    assert is_in_process_task("separate")
    assert get_in_process_task_builder_name("separate") == "build_separate_tab"
    assert requires_initial_file_picker("separate") is False


def test_builder_attribute_exists():
    from icharlotte_core.ui.wizard import in_process_task_tab
    assert hasattr(in_process_task_tab, "build_separate_tab")
