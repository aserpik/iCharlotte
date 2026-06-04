# tests/test_firm_briefs/test_alt_swap.py
from icharlotte_core.ui.wizard.pages.citation_review import apply_alternative_to_body


def test_swap_replaces_cite_text():
    body = "As held in Townsend v. Superior Court (1998) 61 Cal.App.4th 1431, the duty applies."
    new = apply_alternative_to_body(
        body, old_cite="61 Cal.App.4th 1431",
        alternative={"case_name": "Leko v. Cornerstone Building Inspection Service",
                     "citation": "86 Cal.App.4th 1109", "year": "2001"})
    assert "86 Cal.App.4th 1109" in new
    assert "Leko v. Cornerstone" in new
    assert "61 Cal.App.4th 1431" not in new


def test_swap_noop_when_cite_absent():
    body = "No citation here."
    assert apply_alternative_to_body(body, "1 Cal.5th 1", {"citation": "2 Cal.5th 2"}) == body
