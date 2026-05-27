"""Tests for the citation parser."""

from icharlotte_core.opposition.citation_parser import (
    Citation,
    extract_citations,
)


def test_simple_case_cite_extracted():
    body = "The court held this in *Cottini v. Enloe Medical Center* (2014) 226 Cal.App.4th 401."
    cites = extract_citations(body)
    assert len(cites) == 1
    c = cites[0]
    assert c.kind == "case"
    assert c.case_name == "Cottini v. Enloe Medical Center"
    assert c.year == "2014"
    assert c.reporter_citation == "226 Cal.App.4th 401"


def test_pincite_preserved_in_raw_text_but_stripped_from_normalized():
    body = "See *Cottini v. Enloe Medical Center* (2014) 226 Cal.App.4th 401, 415."
    cites = extract_citations(body)
    assert len(cites) == 1
    c = cites[0]
    assert "415" in c.raw_text
    assert c.normalized.endswith("226 Cal.App.4th 401")
    assert ", 415" not in c.normalized


def test_case_cite_raw_text_strips_italic_markers():
    # The UI converts *Name* into <i>Name</i> in the rendered HTML before
    # wrapping the citation as a clickable anchor. If raw_text kept the
    # asterisks, the regex match against the rendered HTML would fail and
    # the cite would render unclickable. So raw_text must NOT contain *.
    body = "See *Cottini v. Enloe Medical Center* (2014) 226 Cal.App.4th 401."
    cites = extract_citations(body)
    assert len(cites) == 1
    assert "*" not in cites[0].raw_text
    assert "_" not in cites[0].raw_text
    assert cites[0].raw_text.startswith("Cottini")


def test_case_name_with_comma_inc_suffix_captures_full_name():
    # Real-world case names often include a comma before a company
    # suffix like ", Inc." / ", LLC" / ", L.P." — the parser must keep
    # the full name on the left side of "v." rather than starting the
    # match at "Inc. v. ...".
    body = (
        "As the court held in *Sinaiko Healthcare-Consulting, Inc. v. "
        "Pacific Healthcare Consultants* (2007) 148 Cal.App.4th 390, 402, "
        "the trial court has discretion."
    )
    cites = extract_citations(body)
    assert len(cites) == 1
    c = cites[0]
    assert c.case_name == (
        "Sinaiko Healthcare-Consulting, Inc. v. Pacific Healthcare Consultants"
    )
    assert c.year == "2007"
    assert c.reporter_citation == "148 Cal.App.4th 390"
    assert c.raw_text.startswith("Sinaiko Healthcare-Consulting,")


def test_case_name_with_comma_on_right_side():
    # Symmetric coverage: comma+suffix on the right side too.
    body = "See *Smith v. Jones, Inc.* (2010) 50 Cal.4th 100 for support."
    cites = extract_citations(body)
    assert len(cites) == 1
    assert cites[0].case_name == "Smith v. Jones, Inc."


def test_law_firm_case_name_with_ampersand_and_commas():
    # Real-world law-firm case name with multiple commas and an ampersand.
    body = (
        "In *Hecht, Solberg, Robinson, Goldberg & Bagula v. Superior Court* "
        "(2006) 137 Cal.App.4th 579, 591-592, the court held that ..."
    )
    cites = extract_citations(body)
    assert len(cites) == 1
    c = cites[0]
    assert c.case_name == (
        "Hecht, Solberg, Robinson, Goldberg & Bagula v. Superior Court"
    )
    assert c.year == "2006"
    assert c.reporter_citation == "137 Cal.App.4th 579"


def test_case_name_ampersand_only():
    body = "See *Latham & Watkins v. Smith* (2015) 60 Cal.App.4th 100."
    cites = extract_citations(body)
    assert len(cites) == 1
    assert cites[0].case_name == "Latham & Watkins v. Smith"


def test_amp_entity_in_body_is_normalized():
    # If the LLM ever emits "&amp;" instead of "&", the parser must still
    # capture the full case name. The renderer's double-escape bug also
    # produces "&amp;" in HTML — this defensive normalization handles it.
    body = "See *X &amp; Y v. Z* (2010) 50 Cal.4th 100."
    cites = extract_citations(body)
    assert len(cites) == 1
    assert cites[0].case_name == "X & Y v. Z"


def test_vs_period_form_accepted():
    body = "See *Smith vs. Jones* (2010) 50 Cal.4th 100."
    cites = extract_citations(body)
    assert len(cites) == 1
    assert "Smith vs. Jones" == cites[0].case_name


def test_case_name_without_italic_markers():
    body = "The court held this in Cottini v. Enloe Medical Center (2014) 226 Cal.App.4th 401."
    cites = extract_citations(body)
    assert len(cites) == 1
    assert cites[0].case_name == "Cottini v. Enloe Medical Center"


def test_no_case_cite_returns_empty():
    assert extract_citations("This sentence has no citation.") == []


def test_multiple_case_cites_in_one_paragraph():
    body = (
        "Two cases apply. *Smith v. Jones* (2010) 50 Cal.4th 100 "
        "and *Brown v. Davis* (2015) 60 Cal.App.4th 200 both hold this."
    )
    cites = extract_citations(body)
    assert len(cites) == 2
    assert {c.case_name for c in cites} == {"Smith v. Jones", "Brown v. Davis"}


def test_statute_cite_with_section_symbol():
    body = "Plaintiff failed to comply with Code Civ. Proc., § 2024.020."
    cites = extract_citations(body)
    assert len(cites) == 1
    c = cites[0]
    assert c.kind == "statute"
    assert c.law_code == "CCP"
    assert c.section_num == "2024.020"


def test_statute_cite_evidence_code():
    body = "Under Evid. Code § 352 the court may exclude this."
    cites = extract_citations(body)
    assert len(cites) == 1
    c = cites[0]
    assert c.kind == "statute"
    assert c.law_code == "EVID"
    assert c.section_num == "352"


def test_statute_full_name_form():
    body = "The Code of Civil Procedure section 2031.030 governs."
    cites = extract_citations(body)
    assert len(cites) == 1
    c = cites[0]
    assert c.kind == "statute"
    assert c.law_code == "CCP"


def test_rule_of_court_cite():
    body = "Pursuant to California Rules of Court, rule 3.1345."
    cites = extract_citations(body)
    assert len(cites) == 1
    c = cites[0]
    assert c.kind == "rule"
    assert "3.1345" in c.raw_text


def test_mixed_case_statute_rule_in_one_body():
    body = (
        "*Smith v. Jones* (2010) 50 Cal.4th 100 establishes the rule. "
        "Code Civ. Proc., § 2024.020 codifies the deadline. "
        "California Rules of Court, rule 3.1345 controls format."
    )
    cites = extract_citations(body)
    kinds = sorted(c.kind for c in cites)
    assert kinds == ["case", "rule", "statute"]


def test_proposition_is_containing_sentence_plus_prior():
    body = (
        "Discovery cutoffs must be respected. "
        "Trial courts retain discretion to deny untimely motions. "
        "*Cottini v. Enloe Medical Center* (2014) 226 Cal.App.4th 401. "
        "This is the next sentence."
    )
    cites = extract_citations(body)
    assert len(cites) == 1
    p = cites[0].proposition
    # Should include the containing sentence + the prior one, but not the next.
    assert "untimely motions" in p
    assert "Discovery cutoffs" in p
    assert "next sentence" not in p


def test_proposition_for_first_sentence_has_no_prior():
    body = "*Cottini v. Enloe* (2014) 226 Cal.App.4th 401 controls. Other stuff follows."
    cites = extract_citations(body)
    assert len(cites) == 1
    p = cites[0].proposition
    assert "controls" in p
    assert "Other stuff" not in p


def test_multiple_cites_share_sentence_share_proposition():
    body = (
        "Two cases agree. *Smith v. Jones* (2010) 50 Cal.4th 100 and "
        "*Brown v. Davis* (2015) 60 Cal.App.4th 200 both so hold."
    )
    cites = extract_citations(body)
    assert len(cites) == 2
    # Both share the same sentence — propositions identical.
    assert cites[0].proposition == cites[1].proposition
    assert "Two cases agree" in cites[0].proposition
