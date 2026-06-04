from types import SimpleNamespace
from icharlotte_core.firm_briefs.profile import profile_from_metadata


def test_profile_from_metadata_combines_relief_and_arguments():
    meta = SimpleNamespace(
        relief_requested="compel further responses to RFP set one",
        principal_arguments=["Plaintiff failed to meet and confer", "Responses are evasive"],
    )
    prof = profile_from_metadata(meta)
    assert "compel further responses" in prof
    assert "meet and confer" in prof
    assert "evasive" in prof


def test_profile_from_metadata_handles_missing_fields():
    assert profile_from_metadata(SimpleNamespace()) == ""
    assert profile_from_metadata(None) == ""
