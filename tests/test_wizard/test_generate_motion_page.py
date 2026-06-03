"""UI-logic tests for the Generate Motion intake settings page (spec #7 redesign)."""
import pytest

pytest.importorskip("pytestqt")

from icharlotte_core.opposition.models import MotionMetadata  # noqa: E402
from icharlotte_core.ui.wizard.pages.generate_motion_page import (  # noqa: E402
    SETTINGS_PAGE_INTAKE,
    SETTINGS_PAGE_OUTLINE,
    SETTINGS_PAGE_REVIEW,
    GenerateMotionSettingsPage,
    _OTHER,
)


def _page(qtbot):
    page = GenerateMotionSettingsPage("/case", "12345")
    qtbot.addWidget(page)
    return page


def test_starts_on_intake_page(qtbot):
    page = _page(qtbot)
    assert page.currentIndex() == SETTINGS_PAGE_INTAKE


def test_builder_opens_directly_on_intake_with_no_dialog(qtbot):
    # The builder must NOT pop a file/type dialog; it lands on the intake page.
    from icharlotte_core.ui.wizard.registry import get_task
    from icharlotte_core.ui.wizard.pages.generate_motion_page import (
        TASK_PAGE_SETTINGS,
        build_generate_motion_tab,
    )

    tab = build_generate_motion_tab(get_task("generate_motion"), "/case", "123", None)
    qtbot.addWidget(tab)
    assert tab is not None
    assert tab.currentIndex() == TASK_PAGE_SETTINGS
    assert tab.settings_page.currentIndex() == SETTINGS_PAGE_INTAKE


def test_type_combo_includes_other(qtbot):
    page = _page(qtbot)
    assert page.type_combo.findData(_OTHER) >= 0


def test_other_type_reveals_custom_name_and_routes_to_generic(qtbot):
    page = _page(qtbot)
    page.type_combo.setCurrentIndex(page.type_combo.findData(_OTHER))
    assert page.custom_name_edit.isVisibleTo(page)
    page.custom_name_edit.setText("Motion for Protective Order")
    assert page.current_motion_type_id() == "generic"
    assert page.current_motion_type_name() == "Motion for Protective Order"


def test_configured_type_name_from_config(qtbot):
    page = _page(qtbot)
    page.type_combo.setCurrentIndex(page.type_combo.findData("demurrer"))
    assert page.current_motion_type_id() == "demurrer"
    assert "Demurrer" in page.current_motion_type_name()


def test_intake_settings_collects_all_fields(qtbot):
    page = _page(qtbot)
    page.type_combo.setCurrentIndex(page.type_combo.findData("compel"))
    page.files_list.addItem("/case/responses.pdf")
    page.user_relief_edit.setText("Compel further responses")
    page.user_arguments_edit.setPlainText("Boilerplate objections\nNo privilege log")
    s = page.intake_settings()
    assert s["motion_type_id"] == "compel"
    assert s["target_files"] == ["/case/responses.pdf"]
    assert s["user_relief"] == "Compel further responses"
    assert s["user_arguments"] == ["Boilerplate objections", "No privilege log"]


def test_analyze_emits_when_arguments_present(qtbot):
    page = _page(qtbot)
    page.user_arguments_edit.setPlainText("A ground")
    with qtbot.waitSignal(page.analyze_requested, timeout=500) as blocker:
        page._on_analyze_continue()
    assert blocker.args[0]["user_arguments"] == ["A ground"]


def test_set_metadata_then_continue_to_outline(qtbot):
    page = _page(qtbot)
    page.set_metadata(MotionMetadata(
        motion_type="Motion to Compel Further Responses",
        relief_requested="Compel",
        principal_arguments=["g1", "g2"],
    ))
    md = page.current_metadata()
    assert md.relief_requested == "Compel"
    assert md.principal_arguments == ["g1", "g2"]
    page._on_continue_to_outline()
    assert page.currentIndex() == SETTINGS_PAGE_OUTLINE


def test_to_dict_from_dict_round_trip_custom_type(qtbot):
    page = _page(qtbot)
    page.type_combo.setCurrentIndex(page.type_combo.findData(_OTHER))
    page.custom_name_edit.setText("Motion for Protective Order")
    page.files_list.addItem("/case/doc.pdf")
    page.set_metadata(MotionMetadata(relief_requested="Protect", principal_arguments=["G1"]))
    from icharlotte_core.motion_generation.analyzer import outline_from_config
    from icharlotte_core.motion_generation.config import get_motion_config
    page.set_outline(outline_from_config(get_motion_config("generic")))

    data = page.to_dict()
    assert data["motion_type_id"] == "generic"
    assert data["motion_type_name"] == "Motion for Protective Order"
    assert data["target_files"] == ["/case/doc.pdf"]

    page2 = _page(qtbot)
    page2.from_dict(data)
    assert page2.current_motion_type_name() == "Motion for Protective Order"
    assert page2.current_target_files() == ["/case/doc.pdf"]
    assert page2.current_metadata().principal_arguments == ["G1"]
