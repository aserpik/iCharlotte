"""Tests for the Separate wizard task tab (settings + workbench wiring)."""
import pytest

pytest.importorskip("pytestqt")
pytest.importorskip("pypdf")

from icharlotte_core.ui.wizard.registry import get_task
from icharlotte_core.ui.wizard.pages.separate_page import (
    SeparateSettingsPage,
    SeparateTaskTab,
    PAGE_SETTINGS,
    PAGE_STATUS,
    PAGE_WORKBENCH,
)


def test_settings_emits_analyze_with_sensitivity(qtbot):
    page = SeparateSettingsPage(pdf_path="C:/x.pdf")
    qtbot.addWidget(page)
    page.sensitivity_slider.setValue(1)
    with qtbot.waitSignal(page.analyze_requested, timeout=500) as blocker:
        page.analyze_btn.click()
    assert blocker.args[0] == 1


def test_tab_starts_on_settings(qtbot):
    spec = get_task("separate")
    tab = SeparateTaskTab(spec, case_path="C:/case", file_number="1234.001",
                          pdf_path="C:/x.pdf")
    qtbot.addWidget(tab)
    assert tab.currentIndex() == PAGE_SETTINGS


def test_analysis_success_loads_workbench(qtbot, tmp_path):
    import pypdf
    pdf = tmp_path / "src.pdf"
    w = pypdf.PdfWriter()
    for _ in range(4):
        w.add_blank_page(width=200, height=200)
    with open(pdf, "wb") as f:
        w.write(f)

    spec = get_task("separate")
    tab = SeparateTaskTab(spec, case_path=str(tmp_path), file_number="1234.001",
                          pdf_path=str(pdf))
    qtbot.addWidget(tab)
    docs = [{"id": "1", "title": "A", "date": "", "start": 1, "end": 4}]
    tab._on_analysis_finished(True, docs)
    assert tab.currentIndex() == PAGE_WORKBENCH
    assert tab.workbench.doc_table.rowCount() == 1
