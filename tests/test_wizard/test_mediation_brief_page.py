"""Tests for the Mediation Brief Wizard task."""
import os

import pytest


# ---- Registry / routing (pure logic, no Qt) ----

def test_mediation_brief_registered():
    from icharlotte_core.ui.wizard.registry import TASK_REGISTRY

    spec = TASK_REGISTRY["mediation_brief"]
    assert spec.title == "Mediation Brief"
    assert spec.category == "Motions & Drafting"
    assert spec.script_name == ""
    assert "mediation" in spec.keywords


def test_mediation_brief_has_valid_category():
    from icharlotte_core.ui.wizard.registry import CATEGORY_ORDER, TASK_REGISTRY

    assert TASK_REGISTRY["mediation_brief"].category in CATEGORY_ORDER
