from icharlotte_core.opposition.models import CitationVerification, RetrievedAuthority
from icharlotte_core.firm_briefs.provenance import attach_firm_provenance


def test_attach_marks_firm_and_alternatives():
    cits = [CitationVerification(citation_text="Townsend v. Superior Court (1998) 61 Cal.App.4th 1431",
                                 normalized_citation="61 Cal.App.4th 1431")]
    pool = [RetrievedAuthority(case_name="Townsend v. Superior Court", citation="61 Cal.App.4th 1431",
                              source="firm", verification="local",
                              source_brief=r"C:\lib\x.pdf",
                              alternatives=[RetrievedAuthority(case_name="Leko v. Cornerstone",
                                                              citation="86 Cal.App.4th 1109", source="corpus")])]
    attach_firm_provenance(cits, pool)
    assert cits[0].source == "firm"
    assert cits[0].firm_verification == "local"
    assert cits[0].source_brief.endswith("x.pdf")
    assert cits[0].alternatives[0]["citation"] == "86 Cal.App.4th 1109"


def test_no_match_leaves_citation_untouched():
    cits = [CitationVerification(citation_text="Other v. Case", normalized_citation="1 Cal.5th 1")]
    attach_firm_provenance(cits, [])
    assert cits[0].source == ""
