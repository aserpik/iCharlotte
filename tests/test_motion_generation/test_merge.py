"""Tests for merging user-specified arguments/relief with AI-proposed grounds (spec #7 redesign)."""
from icharlotte_core.motion_generation.analyzer import merge_intake_with_analysis
from icharlotte_core.opposition.models import MotionMetadata


def _ai(relief="AI relief", args=("AI ground one", "AI ground two")):
    return MotionMetadata(relief_requested=relief, principal_arguments=list(args))


def test_user_relief_overrides_ai_relief():
    md = merge_intake_with_analysis("My relief", [], _ai(), "Motion to Compel")
    assert md.relief_requested == "My relief"


def test_empty_user_relief_falls_back_to_ai():
    md = merge_intake_with_analysis("  ", [], _ai(relief="AI relief"), "Motion")
    assert md.relief_requested == "AI relief"


def test_user_arguments_come_first_then_ai():
    md = merge_intake_with_analysis(
        "r", ["User ground"], _ai(args=("AI ground",)), "Motion"
    )
    assert md.principal_arguments == ["User ground", "AI ground"]


def test_duplicate_arguments_are_deduped_case_insensitively():
    md = merge_intake_with_analysis(
        "r", ["Boilerplate Objections"], _ai(args=("boilerplate objections", "New one")), "M"
    )
    assert md.principal_arguments == ["Boilerplate Objections", "New one"]


def test_motion_type_name_is_applied():
    md = merge_intake_with_analysis("r", ["g"], _ai(), "Motion for Protective Order")
    assert md.motion_type == "Motion for Protective Order"


def test_blank_user_arguments_are_ignored():
    md = merge_intake_with_analysis("r", ["", "  ", "real"], _ai(args=()), "M")
    assert md.principal_arguments == ["real"]


def test_works_with_no_ai_grounds():
    md = merge_intake_with_analysis("r", ["only user"], _ai(relief="", args=()), "M")
    assert md.principal_arguments == ["only user"]
    assert md.relief_requested == "r"
