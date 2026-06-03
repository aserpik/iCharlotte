"""Tests for loading motion-type style samples into the drafter (Workbench feature)."""
from docx import Document

from icharlotte_core.motion_generation.samples import load_exemplars
from icharlotte_core.opposition.style_examples import StyleExample, StyleExampleRegistry


def _make_sample(tmp_path, text):
    doc = Document()
    doc.add_paragraph(text)
    p = tmp_path / "sample.docx"
    doc.save(str(p))
    return str(p)


def _registry_with(tmp_path, *, motion_types, active=True, text="SAMPLE STYLE TEXT"):
    sample = _make_sample(tmp_path, text)
    reg_path = str(tmp_path / "style_examples.json")
    reg = StyleExampleRegistry.load(reg_path)
    reg.add(StyleExample(id="1", label="S", path=sample, motion_types=motion_types, active=active))
    reg.save()
    return reg_path


def test_load_exemplars_matches_by_type(tmp_path):
    reg_path = _registry_with(tmp_path, motion_types=["compel"])
    cache = str(tmp_path / "cache")
    got = load_exemplars("compel", registry_path=reg_path, cache_dir=cache)
    assert any("SAMPLE STYLE TEXT" in t for t in got)


def test_load_exemplars_excludes_other_types(tmp_path):
    reg_path = _registry_with(tmp_path, motion_types=["compel"])
    cache = str(tmp_path / "cache")
    assert load_exemplars("demurrer", registry_path=reg_path, cache_dir=cache) == []


def test_inactive_samples_are_skipped(tmp_path):
    reg_path = _registry_with(tmp_path, motion_types=["compel"], active=False)
    cache = str(tmp_path / "cache")
    assert load_exemplars("compel", registry_path=reg_path, cache_dir=cache) == []


def test_missing_registry_returns_empty(tmp_path):
    got = load_exemplars(
        "compel",
        registry_path=str(tmp_path / "nope.json"),
        cache_dir=str(tmp_path / "cache"),
    )
    assert got == []
