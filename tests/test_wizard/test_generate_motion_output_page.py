"""Citation-review surface on the Generate Motion output page (parity with
the oppose output page)."""
from __future__ import annotations

import pytest

pytest.importorskip("PySide6")
pytest.importorskip("pytestqt")

from PySide6.QtCore import QUrl  # noqa: E402

from icharlotte_core.opposition.models import (  # noqa: E402
    CitationVerification,
    DraftDocument,
)
from icharlotte_core.ui.wizard.pages.generate_motion_page import (  # noqa: E402
    GenerateMotionOutputPage,
)


def _draft(citations):
    return DraftDocument(
        title="Motion to Compel",
        body_text="See *Smith v. Jones* (2010) 50 Cal.4th 100 for support.",
        citations=citations,
    )


def test_supported_citation_renders_green(qtbot):
    page = GenerateMotionOutputPage()
    qtbot.addWidget(page)
    page.show_result(_draft([
        CitationVerification(
            citation_text="Smith v. Jones (2010) 50 Cal.4th 100",
            verdict="SUPPORTED", kind="case",
        )
    ]))
    html = page.editor.toHtml().lower()
    assert "#1e8e3e" in html  # green underline color


def test_not_supported_citation_renders_red(qtbot):
    page = GenerateMotionOutputPage()
    qtbot.addWidget(page)
    page.show_result(_draft([
        CitationVerification(
            citation_text="Smith v. Jones (2010) 50 Cal.4th 100",
            verdict="NOT_SUPPORTED", kind="case",
        )
    ]))
    assert "#c5221f" in page.editor.toHtml().lower()  # red


def test_summary_banner_counts_per_verdict(qtbot):
    page = GenerateMotionOutputPage()
    qtbot.addWidget(page)
    page.show_result(DraftDocument(
        title="M", body_text="Body.",
        citations=[
            CitationVerification(citation_text="a", verdict="SUPPORTED"),
            CitationVerification(citation_text="b", verdict="SUPPORTED"),
            CitationVerification(citation_text="c", verdict="PARTIAL"),
            CitationVerification(citation_text="d", verdict="NOT_SUPPORTED"),
        ],
    ))
    banner = page.summary_banner.text().lower()
    assert "supported" in banner
    assert "2" in banner


def test_clicking_anchor_updates_detail_panel(qtbot):
    page = GenerateMotionOutputPage()
    qtbot.addWidget(page)
    page.show_result(DraftDocument(
        title="M",
        body_text="First *A v. B* (2010) 1 Cal.5th 1; then *C v. D* (2011) 2 Cal.5th 2.",
        citations=[
            CitationVerification(citation_text="A v. B (2010) 1 Cal.5th 1",
                                 case_name="A v. B", verdict="SUPPORTED", kind="case"),
            CitationVerification(citation_text="C v. D (2011) 2 Cal.5th 2",
                                 case_name="C v. D", verdict="NOT_SUPPORTED", kind="case"),
        ],
    ))
    page._on_anchor_clicked(QUrl("citation:1"))
    assert "C v. D" in page.detail_panel.header_label.text()
    assert "NOT SUPPORTED" in page.detail_panel.header_label.text()


def test_empty_citations_shows_motion_specific_message(qtbot):
    page = GenerateMotionOutputPage()
    qtbot.addWidget(page)
    page.show_result(DraftDocument(title="M", body_text="Body.", citations=[]))
    assert "motion" in page.detail_panel.body_html.lower()


def test_save_warns_on_red_verdicts(qtbot, monkeypatch, tmp_path):
    from PySide6.QtWidgets import QFileDialog, QMessageBox

    page = GenerateMotionOutputPage()
    qtbot.addWidget(page)
    preview = tmp_path / "preview.docx"
    preview.write_bytes(b"dummy")
    page.show_result(DraftDocument(
        title="M", body_text="b", preview_path=str(preview),
        citations=[CitationVerification(citation_text="x", verdict="NOT_SUPPORTED")],
    ))

    warned = {"yes": False}

    def fake_question(parent, title, text, *args, **kwargs):
        warned["yes"] = True
        return QMessageBox.StandardButton.Cancel  # user cancels

    monkeypatch.setattr(QMessageBox, "question", fake_question)
    monkeypatch.setattr(QFileDialog, "getSaveFileName", lambda *a, **k: ("", ""))

    page.save_as()
    assert warned["yes"]


def test_open_in_word_button_present(qtbot):
    page = GenerateMotionOutputPage()
    qtbot.addWidget(page)
    assert hasattr(page, "open_btn")
    assert page.open_btn.text() == "Open in Word"
