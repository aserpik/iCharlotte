# tests/test_firm_briefs/test_panel_provenance_html.py
from icharlotte_core.opposition.models import CitationVerification
from icharlotte_core.ui.wizard.pages.citation_review import _citation_body_html


def test_firm_badge_and_tier_rendered():
    c = CitationVerification(citation_text="Townsend (1998) 61 Cal.App.4th 1431", verdict="SUPPORTED",
                             source="firm", source_brief=r"C:\lib\Oppositions\Motion to Compel\Smith Opp.pdf",
                             firm_verification="local",
                             alternatives=[{"case_name": "Leko v. Cornerstone", "citation": "86 Cal.App.4th 1109"}])
    html = _citation_body_html(c, "SUPPORTED")
    assert "From your brief" in html
    assert "Smith Opp" in html            # basename of source_brief shown
    assert "Leko v. Cornerstone" in html  # alternative listed


def test_unverified_firm_amber_warning():
    c = CitationVerification(citation_text="Smith v. Jones (2024) 999 F.3d 1", verdict="UNVERIFIED",
                             source="firm", source_brief=r"C:\lib\x.pdf", firm_verification="unverified_firm")
    html = _citation_body_html(c, "UNVERIFIED")
    assert "from firm brief" in html.lower()
    assert "not independently verified" in html.lower()


def test_corpus_citation_no_firm_badge():
    c = CitationVerification(citation_text="X v. Y (2000) 1 Cal.5th 1", verdict="SUPPORTED")
    assert "From your brief" not in _citation_body_html(c, "SUPPORTED")
