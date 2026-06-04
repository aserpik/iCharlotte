from icharlotte_core.firm_briefs.profile import extract_headings, compose_profile, profile_from_text

DOC = (
    "NOTICE OF MOTION\n"
    "I. PLAINTIFF FAILED TO MEET AND CONFER\n"
    "Some argument prose here that is not a heading.\n"
    "II. THE DISCOVERY CUTOFF HAS PASSED\n"
    "More prose.\n"
)


def test_extract_headings_picks_caps_lines():
    heads = extract_headings(DOC)
    assert any("MEET AND CONFER" in h for h in heads)
    assert any("DISCOVERY CUTOFF" in h for h in heads)
    assert "Some argument prose here that is not a heading." not in heads


def test_compose_profile_concatenates():
    prof = compose_profile("compel further responses", ["FAILED TO MEET AND CONFER"], ["cutoff passed"])
    assert "compel further responses" in prof
    assert "MEET AND CONFER" in prof
    assert "cutoff passed" in prof


def test_profile_from_text_nonempty():
    assert profile_from_text(DOC).strip()
