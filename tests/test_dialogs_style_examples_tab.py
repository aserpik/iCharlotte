"""Tests for the workbench Style Examples tab."""

from __future__ import annotations

import os

import pytest


@pytest.fixture(autouse=True)
def _skip_if_no_qt():
    pytest.importorskip("PySide6")


def test_tab_loads_existing_examples(qtbot, tmp_path):
    from icharlotte_core.opposition.style_examples import StyleExample, StyleExampleRegistry
    from icharlotte_core.ui.dialogs_style_examples import StyleExamplesTab

    registry_path = str(tmp_path / "style_examples.json")
    reg = StyleExampleRegistry(path=registry_path)
    reg.add(StyleExample(id="1", label="MTC Opp", path="/x.docx", motion_types=["motion to compel"], active=True))
    reg.save()

    tab = StyleExamplesTab(registry_path=registry_path)
    qtbot.addWidget(tab)
    assert tab.table.rowCount() == 1
    assert tab.table.item(0, 0).text() == "MTC Opp"


def test_tab_add_then_save_persists(qtbot, tmp_path):
    from icharlotte_core.ui.dialogs_style_examples import StyleExamplesTab
    from icharlotte_core.opposition.style_examples import StyleExampleRegistry

    registry_path = str(tmp_path / "style_examples.json")
    docx_path = str(tmp_path / "exemplar.docx")
    with open(docx_path, "wb") as f:
        f.write(b"PK")  # not a real docx, doesn't matter for registry shape

    tab = StyleExamplesTab(registry_path=registry_path)
    qtbot.addWidget(tab)
    tab.add_example_programmatic(
        label="MSJ Opp",
        path=docx_path,
        motion_types=["summary judgment"],
        active=True,
    )
    tab.save()

    reloaded = StyleExampleRegistry.load(registry_path)
    assert len(reloaded.examples) == 1
    assert reloaded.examples[0].label == "MSJ Opp"


def test_tab_remove_clears_row(qtbot, tmp_path):
    from icharlotte_core.opposition.style_examples import StyleExample, StyleExampleRegistry
    from icharlotte_core.ui.dialogs_style_examples import StyleExamplesTab

    registry_path = str(tmp_path / "style_examples.json")
    reg = StyleExampleRegistry(path=registry_path)
    reg.add(StyleExample(id="abc", label="A", path="/a", motion_types=[], active=True))
    reg.save()

    tab = StyleExamplesTab(registry_path=registry_path)
    qtbot.addWidget(tab)
    tab.remove_example_programmatic("abc")
    tab.save()

    reloaded = StyleExampleRegistry.load(registry_path)
    assert reloaded.examples == []
