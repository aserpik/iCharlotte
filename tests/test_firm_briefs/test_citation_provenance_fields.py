from icharlotte_core.opposition.models import CitationVerification


def test_provenance_fields_default():
    c = CitationVerification()
    assert c.source == ""
    assert c.source_brief == ""
    assert c.firm_verification == ""
    assert c.alternatives == []


def test_provenance_roundtrips_through_dict():
    c = CitationVerification(citation_text="Townsend v. Superior Court (1998) 61 Cal.App.4th 1431",
                             source="firm", source_brief=r"C:\lib\Oppositions\Motion to Compel\x.pdf",
                             firm_verification="local",
                             alternatives=[{"case_name": "Leko v. Cornerstone", "citation": "86 Cal.App.4th 1109"}])
    d = c.to_dict()
    c2 = CitationVerification.from_dict(d)
    assert c2.source == "firm"
    assert c2.source_brief.endswith("x.pdf")
    assert c2.firm_verification == "local"
    assert c2.alternatives[0]["citation"] == "86 Cal.App.4th 1109"
