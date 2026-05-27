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
