from icharlotte_core.firm_briefs.citation_harvest import harvest_cites, HarvestedCite

TEXT = (
    "Plaintiff failed to meet and confer before moving. "
    "A party must engage in a reasonable and good faith effort. "
    "Townsend v. Superior Court (1998) 61 Cal.App.4th 1431, 1438. "
    "The motion is therefore procedurally improper."
)


def test_harvests_case_cite_with_norm_and_proposition():
    cites = harvest_cites(TEXT)
    assert len(cites) == 1
    c = cites[0]
    assert isinstance(c, HarvestedCite)
    assert c.case_name.startswith("Townsend")
    assert c.reporter_citation == "61 Cal.App.4th 1431"
    assert c.year == "1998"
    assert c.norm_cite == "61cal.app.4th1431"
    assert "meet and confer" in c.proposition.lower() or "good faith" in c.proposition.lower()
    assert c.quoted_passage  # non-empty


def test_skips_statutes_in_phase1():
    cites = harvest_cites("See Code Civ. Proc. section 2031.310. Nothing else.")
    assert cites == []
