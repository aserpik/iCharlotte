"""Tests for the Mediation Brief Wizard task."""
import os

import pytest


# ---- Registry / routing (pure logic, no Qt) ----

def test_mediation_brief_registered():
    from icharlotte_core.ui.wizard.registry import TASK_REGISTRY

    spec = TASK_REGISTRY["mediation_brief"]
    assert spec.title == "Mediation Brief"
    assert spec.category == "Motions"
    assert spec.script_name == ""
    assert "mediation" in spec.keywords


def test_mediation_brief_has_valid_category():
    from icharlotte_core.ui.wizard.registry import CATEGORY_ORDER, TASK_REGISTRY

    assert TASK_REGISTRY["mediation_brief"].category in CATEGORY_ORDER


def test_mediation_brief_is_in_process():
    from icharlotte_core.ui.wizard.task_routing import (
        get_in_process_task_builder_name,
        is_in_process_task,
        requires_initial_file_picker,
    )

    assert get_in_process_task_builder_name("mediation_brief") == "build_mediation_brief_tab"
    assert is_in_process_task("mediation_brief") is True
    # In-process tasks own their source selection — no pre-Settings picker.
    assert requires_initial_file_picker("mediation_brief") is False


def test_build_mediation_brief_tab_attribute_exists():
    from icharlotte_core.ui.wizard import in_process_task_tab

    assert hasattr(in_process_task_tab, "build_mediation_brief_tab")
