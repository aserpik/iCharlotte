"""The shared citation-review toolkit exposes its symbols, and oppose
re-exports them for backward compatibility with existing imports/tests."""
from __future__ import annotations

import pytest

pytest.importorskip("PySide6")


def test_citation_review_exposes_symbols():
    from icharlotte_core.ui.wizard.pages import citation_review as cr

    for name in (
        "CitationReviewOutputPage",
        "CitationDetailPanel",
        "CitationDetailDialog",
        "_render_draft_html",
        "_build_citation_index",
        "_format_inline_html",
        "_color_for_verdict",
        "_citation_header_html",
        "_citation_body_html",
        "_run_find_replacement",
        "_VERDICT_COLORS",
        "_VERDICT_HEADER_COLORS",
        "_VERDICT_LABELS",
    ):
        assert hasattr(cr, name), f"citation_review missing {name}"


def test_oppose_motion_page_reexports_shared_symbols():
    # Existing tests import these names from oppose_motion_page; they must
    # keep resolving after the extraction.
    from icharlotte_core.ui.wizard.pages.oppose_motion_page import (  # noqa: F401
        CitationDetailDialog,
        CitationDetailPanel,
        OpposeMotionOutputPage,
        _render_draft_html,
    )

    from icharlotte_core.ui.wizard.pages.citation_review import (
        CitationReviewOutputPage,
    )

    assert issubclass(OpposeMotionOutputPage, CitationReviewOutputPage)
