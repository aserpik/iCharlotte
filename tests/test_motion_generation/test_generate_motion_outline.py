"""generate_motion_outline: motion-aware nested outline with graceful fallback."""
from icharlotte_core.motion_generation.analyzer import (
    generate_motion_outline,
    outline_from_config,
)
from icharlotte_core.motion_generation.config import get_motion_config
from icharlotte_core.opposition.models import MotionMetadata


def _md(args, motion="Motion in Limine to Exclude Witnesses"):
    return MotionMetadata(motion_type=motion, relief_requested="Exclude X",
                          principal_arguments=args)


def test_nested_argument_subheadings():
    def fake_llm(system, user):
        return ('{"outline": [{"text": "Introduction"}, '
                '{"text": "Argument", "children": [{"text": "Sub A"}, {"text": "Sub B"}]}, '
                '{"text": "Conclusion"}]}')
    cfg = get_motion_config("generic")
    nodes = generate_motion_outline(cfg, _md(["g1", "g2"]), target_text="facts",
                                    llm_callback=fake_llm)
    arg = [n for n in nodes if n.text == "Argument"]
    assert arg and len(arg[0].children) >= 2
    assert all(c.selected for c in arg[0].children)


def test_outline_enforces_required_spine_and_three_to_four_argument_headings():
    def fake_llm(system, user):
        return ('{"outline": [{"text": "Introduction"}, '
                '{"text": "Legal Standard"}, '
                '{"text": "Argument", "children": [{"text": "Sub A"}]}, '
                '{"text": "Conclusion"}]}')

    cfg = get_motion_config("generic")
    nodes = generate_motion_outline(
        cfg,
        _md([
            "The moving party satisfied the statutory prerequisites",
            "The opposing party's objections lack merit",
            "The requested relief is narrowly tailored",
            "No prejudice will result from granting the motion",
            "A fifth point should not become a fifth subheading",
        ]),
        llm_callback=fake_llm,
    )

    assert [n.text for n in nodes] == [
        "Introduction",
        "Statement of Facts",
        "Argument",
        "Conclusion",
    ]
    argument = next(n for n in nodes if n.text == "Argument")
    assert 3 <= len(argument.children) <= 4
    assert all(child.selected for child in argument.children)
    assert "Legal Standard" not in [n.text for n in nodes]


def test_fence_tolerant():
    def fake_llm(system, user):
        return '```json\n{"outline": [{"text": "Argument", "children": [{"text": "Sub A"}]}]}\n```'
    cfg = get_motion_config("generic")
    nodes = generate_motion_outline(cfg, _md(["g1"]), llm_callback=fake_llm)
    assert any(n.text == "Argument" and n.children for n in nodes)


def test_empty_grounds_falls_back_without_calling_llm():
    calls = {"n": 0}
    def fake_llm(system, user):
        calls["n"] += 1
        return "{}"
    cfg = get_motion_config("generic")
    nodes = generate_motion_outline(cfg, _md([]), llm_callback=fake_llm)
    assert calls["n"] == 0
    assert [n.text for n in nodes] == [n.text for n in outline_from_config(cfg)]


def test_llm_failure_falls_back_to_flat_outline():
    def fake_llm(system, user):
        return "not json at all"
    cfg = get_motion_config("generic")
    nodes = generate_motion_outline(cfg, _md(["g1"]), llm_callback=fake_llm)
    assert [n.text for n in nodes] == [n.text for n in outline_from_config(cfg)]


def test_prompt_contains_motion_and_grounds():
    captured = {}
    def fake_llm(system, user):
        captured["user"] = user
        return '{"outline": [{"text": "Argument", "children": [{"text": "X"}]}]}'
    cfg = get_motion_config("generic")
    generate_motion_outline(cfg, _md(["unique-ground-phrase-xyz"]), llm_callback=fake_llm)
    assert "Motion in Limine to Exclude Witnesses" in captured["user"]
    assert "unique-ground-phrase-xyz" in captured["user"]
