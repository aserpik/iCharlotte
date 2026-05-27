"""Tests for verdict-colored underline rendering in the wizard output page."""

from __future__ import annotations

import pytest


@pytest.fixture(autouse=True)
def _skip_if_no_qt():
    pytest.importorskip("PySide6")


def test_underline_color_for_supported_verdict_is_green():
    from icharlotte_core.opposition.models import CitationVerification, DraftDocument
    from icharlotte_core.ui.wizard.pages.oppose_motion_page import _render_draft_html

    draft = DraftDocument(
        title="Opp",
        body_text="See *Smith v. Jones* (2010) 50 Cal.4th 100 for support.",
        citations=[CitationVerification(
            citation_text="Smith v. Jones (2010) 50 Cal.4th 100",
            verdict="SUPPORTED",
            kind="case",
        )],
    )
    html = _render_draft_html(draft)
    # Find the anchor wrapping this citation; its underline color should encode SUPPORTED.
    assert "#1e8e3e" in html.lower() or "supported" in html.lower()


def test_ampersand_inside_italics_renders_as_single_entity():
    # Regression: pre-escaping inside the italic-marker lambda combined
    # with the outer html.escape() produced "&amp;amp;" in HTML, which
    # rendered as a visible "&amp;" in the QTextBrowser.
    from icharlotte_core.opposition.models import CitationVerification, DraftDocument
    from icharlotte_core.ui.wizard.pages.oppose_motion_page import _render_draft_html

    draft = DraftDocument(
        title="Opp",
        body_text="See *Goldberg & Bagula v. Court* (2006) 137 Cal.App.4th 579.",
        citations=[CitationVerification(
            citation_text="Goldberg & Bagula v. Court (2006) 137 Cal.App.4th 579",
            verdict="SUPPORTED",
            kind="case",
        )],
    )
    html = _render_draft_html(draft)
    # Properly escaped: exactly "&amp;" once, never "&amp;amp;".
    assert "&amp;amp;" not in html
    assert "&amp;" in html


def test_underline_color_for_not_supported_is_red():
    from icharlotte_core.opposition.models import CitationVerification, DraftDocument
    from icharlotte_core.ui.wizard.pages.oppose_motion_page import _render_draft_html

    draft = DraftDocument(
        title="Opp",
        body_text="See *Smith v. Jones* (2010) 50 Cal.4th 100 for support.",
        citations=[CitationVerification(
            citation_text="Smith v. Jones (2010) 50 Cal.4th 100",
            verdict="NOT_SUPPORTED",
            kind="case",
        )],
    )
    html = _render_draft_html(draft)
    assert "#c5221f" in html.lower() or "not_supported" in html.lower()


def test_underline_color_for_partial_is_yellow():
    from icharlotte_core.opposition.models import CitationVerification, DraftDocument
    from icharlotte_core.ui.wizard.pages.oppose_motion_page import _render_draft_html

    draft = DraftDocument(
        title="Opp",
        body_text="See *Smith v. Jones* (2010) 50 Cal.4th 100 for support.",
        citations=[CitationVerification(
            citation_text="Smith v. Jones (2010) 50 Cal.4th 100",
            verdict="PARTIAL",
            kind="case",
        )],
    )
    html = _render_draft_html(draft)
    assert "#f9ab00" in html.lower() or "partial" in html.lower()


def test_unverified_uses_gray_color():
    from icharlotte_core.opposition.models import CitationVerification, DraftDocument
    from icharlotte_core.ui.wizard.pages.oppose_motion_page import _render_draft_html

    draft = DraftDocument(
        title="Opp",
        body_text="See *Smith v. Jones* (2010) 50 Cal.4th 100 for support.",
        citations=[CitationVerification(
            citation_text="Smith v. Jones (2010) 50 Cal.4th 100",
            verdict="UNVERIFIED",
            kind="case",
        )],
    )
    html = _render_draft_html(draft)
    assert "#80868b" in html.lower() or "unverified" in html.lower()


def test_summary_banner_counts_per_verdict(qtbot):
    from icharlotte_core.opposition.models import CitationVerification, DraftDocument
    from icharlotte_core.ui.wizard.pages.oppose_motion_page import OpposeMotionOutputPage

    page = OpposeMotionOutputPage()
    qtbot.addWidget(page)
    page.show_result(DraftDocument(
        title="Opp",
        body_text="Body text.",
        citations=[
            CitationVerification(citation_text="a", verdict="SUPPORTED"),
            CitationVerification(citation_text="b", verdict="SUPPORTED"),
            CitationVerification(citation_text="c", verdict="PARTIAL"),
            CitationVerification(citation_text="d", verdict="NOT_SUPPORTED"),
            CitationVerification(citation_text="e", verdict="UNVERIFIED"),
        ],
    ))
    banner = page.summary_banner.text()
    assert "2" in banner  # SUPPORTED count
    assert "1" in banner  # PARTIAL count
    assert "supported" in banner.lower()
    assert "partial" in banner.lower()


def test_save_warns_on_red_verdicts(qtbot, monkeypatch, tmp_path):
    from icharlotte_core.opposition.models import CitationVerification, DraftDocument
    from icharlotte_core.ui.wizard.pages.oppose_motion_page import OpposeMotionOutputPage
    from PySide6.QtWidgets import QMessageBox

    page = OpposeMotionOutputPage()
    qtbot.addWidget(page)
    preview = tmp_path / "preview.docx"
    preview.write_bytes(b"dummy")
    page.show_result(DraftDocument(
        title="Opp",
        body_text="b",
        preview_path=str(preview),
        citations=[CitationVerification(citation_text="x", verdict="NOT_SUPPORTED")],
    ))

    warned = {"yes": False}

    def fake_question(parent, title, text, *args, **kwargs):
        warned["yes"] = True
        return QMessageBox.StandardButton.Cancel  # user cancels

    monkeypatch.setattr(QMessageBox, "question", fake_question)
    monkeypatch.setattr(
        "icharlotte_core.ui.wizard.pages.oppose_motion_page.QFileDialog.getSaveFileName",
        lambda *a, **k: ("", ""),
    )

    page.save_as()
    assert warned["yes"]


def test_dialog_shows_evidence_quote_for_supported(qtbot):
    from icharlotte_core.opposition.models import CitationVerification
    from icharlotte_core.ui.wizard.pages.oppose_motion_page import CitationDetailDialog

    cv = CitationVerification(
        citation_text="Smith v. Jones (2010) 50 Cal.4th 100",
        case_name="Smith v. Jones",
        verdict="SUPPORTED",
        proposition="Trial courts have discretion.",
        evidence="The court did not abuse its discretion.",
        note="Direct support.",
        opinion_url="https://www.courtlistener.com/opinion/123/",
    )
    dlg = CitationDetailDialog(cv)
    qtbot.addWidget(dlg)
    text = dlg.findChild(type(dlg.body_label)).text() if hasattr(dlg, "body_label") else ""
    # Best-effort: dialog HTML must contain the evidence string somewhere.
    all_html = " ".join(w.text() for w in dlg.findChildren(type(dlg.header)) if hasattr(w, "text"))
    full = all_html + (text or "")
    assert "did not abuse" in full or "did not abuse" in dlg.body_html


def test_dialog_shows_what_case_actually_holds_for_not_supported(qtbot):
    from icharlotte_core.opposition.models import CitationVerification
    from icharlotte_core.ui.wizard.pages.oppose_motion_page import CitationDetailDialog

    cv = CitationVerification(
        citation_text="Sinaiko Healthcare (2007) 148 Cal.App.4th 390",
        case_name="Sinaiko Healthcare",
        verdict="NOT_SUPPORTED",
        proposition="Serving discovery responses moots a motion to compel.",
        evidence="A party who fails to serve timely responses waives objections.",
        note="Sinaiko's holding addresses waiver, not mootness.",
    )
    dlg = CitationDetailDialog(cv)
    qtbot.addWidget(dlg)
    assert "waiver" in dlg.body_html.lower() or "waives" in dlg.body_html.lower()
    assert "not_supported" in dlg.body_html.lower() or "does not hold" in dlg.body_html.lower() or "actually holds" in dlg.body_html.lower()
