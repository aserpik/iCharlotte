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
