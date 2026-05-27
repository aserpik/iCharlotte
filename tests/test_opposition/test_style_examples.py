"""Tests for the style-examples registry + motion-type matching."""

from __future__ import annotations

import json
import os
import tempfile

import pytest

from icharlotte_core.opposition.style_examples import (
    StyleExample,
    StyleExampleRegistry,
)


@pytest.fixture
def tmp_registry_path():
    with tempfile.TemporaryDirectory() as tmp:
        yield os.path.join(tmp, "style_examples.json")


def test_empty_registry_load(tmp_registry_path):
    reg = StyleExampleRegistry.load(tmp_registry_path)
    assert reg.examples == []


def test_save_then_load_roundtrip(tmp_registry_path):
    reg = StyleExampleRegistry(path=tmp_registry_path)
    reg.add(StyleExample(
        id="ex-1",
        label="MTC Opp",
        path="C:/x/y.docx",
        motion_types=["motion to compel", "discovery"],
        active=True,
    ))
    reg.save()

    loaded = StyleExampleRegistry.load(tmp_registry_path)
    assert len(loaded.examples) == 1
    ex = loaded.examples[0]
    assert ex.label == "MTC Opp"
    assert "motion to compel" in ex.motion_types
    assert ex.active is True


def test_match_by_motion_type_substring(tmp_registry_path):
    reg = StyleExampleRegistry(path=tmp_registry_path)
    reg.add(StyleExample(id="1", label="MTC", path="/x", motion_types=["motion to compel"], active=True))
    reg.add(StyleExample(id="2", label="MSJ", path="/y", motion_types=["summary judgment"], active=True))
    reg.add(StyleExample(id="3", label="Universal", path="/z", motion_types=[], active=True))

    matches = reg.matches_for_motion_type("Motion to Compel Form Interrogatories")
    ids = sorted(m.id for m in matches)
    # MTC matches via substring; Universal always matches (no tags); MSJ doesn't.
    assert ids == ["1", "3"]


def test_inactive_examples_excluded_from_match(tmp_registry_path):
    reg = StyleExampleRegistry(path=tmp_registry_path)
    reg.add(StyleExample(id="1", label="MTC", path="/x", motion_types=["motion to compel"], active=False))
    matches = reg.matches_for_motion_type("Motion to Compel Form Interrogatories")
    assert matches == []


def test_max_matches_caps_at_three(tmp_registry_path):
    reg = StyleExampleRegistry(path=tmp_registry_path)
    for i in range(5):
        reg.add(StyleExample(id=str(i), label=f"e{i}", path=f"/p{i}", motion_types=["motion to compel"], active=True))
    matches = reg.matches_for_motion_type("motion to compel x", max_results=3)
    assert len(matches) == 3


def test_remove_and_update(tmp_registry_path):
    reg = StyleExampleRegistry(path=tmp_registry_path)
    reg.add(StyleExample(id="a", label="A", path="/a", motion_types=[], active=True))
    reg.add(StyleExample(id="b", label="B", path="/b", motion_types=[], active=True))
    reg.update("a", label="A revised", motion_types=["msj"])
    reg.remove("b")

    assert len(reg.examples) == 1
    assert reg.examples[0].label == "A revised"
    assert reg.examples[0].motion_types == ["msj"]
