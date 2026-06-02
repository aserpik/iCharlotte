"""UI-logic tests for the Generate Motion settings page (spec #7)."""
import pytest

pytest.importorskip("pytestqt")

from icharlotte_core.opposition.models import MotionMetadata  # noqa: E402
from icharlotte_core.ui.wizard.pages.generate_motion_page import (  # noqa: E402
    SETTINGS_PAGE_OUTLINE,
    SETTINGS_PAGE_REVIEW,
    GenerateMotionSettingsPage,
)


def _page(qtbot, type_id="compel"):
    page = GenerateMotionSettingsPage("", "", type_id, ["/case/responses.pdf"])
    qtbot.addWidget(page)
    return page


def test_starts_on_review_page(qtbot):
    page = _page(qtbot)
    assert page.currentIndex() == SETTINGS_PAGE_REVIEW


def test_type_combo_reflects_initial_type(qtbot):
    page = _page(qtbot, "demurrer")
    assert page.current_motion_type_id() == "demurrer"


def test_set_metadata_populates_fields(qtbot):
    page = _page(qtbot)
    page.set_metadata(MotionMetadata(
        relief_requested="Compel further responses",
        principal_arguments=["Boilerplate objections", "No privilege log"],
    ))
    md = page.current_metadata()
    assert md.relief_requested == "Compel further responses"
    assert md.principal_arguments == ["Boilerplate objections", "No privilege log"]
    # motion_type is derived from the selected type's display name.
    assert "Compel" in md.motion_type


def test_cannot_continue_without_relief_and_grounds(qtbot):
    page = _page(qtbot)
    assert page.can_continue_to_outline() is False
    page.set_metadata(MotionMetadata(
        relief_requested="x", principal_arguments=["y"]
    ))
    assert page.can_continue_to_outline() is True


def test_continue_advances_to_outline_when_valid(qtbot):
    page = _page(qtbot)
    page.set_metadata(MotionMetadata(relief_requested="x", principal_arguments=["y"]))
    page._on_continue_to_outline()
    assert page.currentIndex() == SETTINGS_PAGE_OUTLINE


def test_to_dict_and_from_dict_round_trip(qtbot):
    page = _page(qtbot)
    page.set_metadata(MotionMetadata(relief_requested="Relief", principal_arguments=["G1", "G2"]))
    from icharlotte_core.motion_generation.analyzer import outline_from_config
    from icharlotte_core.motion_generation.config import get_motion_config
    page.set_outline(outline_from_config(get_motion_config("compel")))

    data = page.to_dict()
    assert data["motion_type_id"] == "compel"
    assert data["target_files"] == ["/case/responses.pdf"]
    assert data["metadata"]["relief_requested"] == "Relief"
    assert len(data["outline"]) >= 5  # the standard section spine

    page2 = _page(qtbot, "generic")
    page2.from_dict(data)
    assert page2.current_motion_type_id() == "compel"
    assert page2.current_metadata().principal_arguments == ["G1", "G2"]
