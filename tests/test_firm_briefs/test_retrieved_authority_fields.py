from icharlotte_core.opposition.models import RetrievedAuthority


def test_new_provenance_fields_default_safely():
    ra = RetrievedAuthority()
    assert ra.source == "corpus"
    assert ra.verification == "local"
    assert ra.source_brief == ""
    assert ra.alternatives == []


def test_alternatives_are_independent_lists():
    a, b = RetrievedAuthority(), RetrievedAuthority()
    a.alternatives.append(b)
    assert b.alternatives == []  # no shared default
