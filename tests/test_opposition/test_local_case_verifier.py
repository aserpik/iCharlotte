import json

from icharlotte_core.opposition.citation_parser import Citation
from icharlotte_core.opposition.local_case_verifier import LocalCaseVerifier


class _FakeCorpus:
    def __init__(self, found):
        self._found = found
    def lookup_by_citation(self, cite):
        return self._found


def _llm_supported(_sys, _user):
    return json.dumps({
        "verdict": "SUPPORTED",
        "evidence": "duty of care is established",
        "note": "on point",
    })


def test_not_found_when_citation_absent():
    v = LocalCaseVerifier(corpus=_FakeCorpus(None), llm_callback=_llm_supported)
    c = Citation(kind="case", raw_text="30 Cal. 4th 43", normalized="30 Cal. 4th 43",
                 reporter_citation="30 Cal. 4th 43", proposition="duty exists")
    cv = v.verify(c)
    assert cv.verdict == "NOT_FOUND"


def test_supported_when_text_supports(monkeypatch):
    import icharlotte_core.opposition.local_case_verifier as mod
    monkeypatch.setattr(mod, "get_prompt", lambda *_a, **_k: "{proposition}|{citation_text}|{authority_text}")
    found = {"case_uid": "cap:1", "full_text": "The duty of care is established.",
             "name": "Duty v. Care", "url": "u", "court": "Cal.", "decision_date": "2003-01-01",
             "citation_count": 9, "latest_citing_year": "2019"}
    v = LocalCaseVerifier(corpus=_FakeCorpus(found), llm_callback=_llm_supported)
    c = Citation(kind="case", raw_text="30 Cal. 4th 43", normalized="30 Cal. 4th 43",
                 reporter_citation="30 Cal. 4th 43", proposition="duty exists")
    cv = v.verify(c)
    assert cv.verdict == "SUPPORTED"
    assert cv.cluster_id == "cap:1"
    assert cv.citation_count == 9


def test_verifier_uses_full_text_not_parenthetical_context(monkeypatch):
    import icharlotte_core.opposition.local_case_verifier as mod
    monkeypatch.setattr(mod, "get_prompt", lambda *_a, **_k: "{authority_text}")
    captured = {}

    def llm(_sys, user):
        captured["user"] = user
        return json.dumps({
            "verdict": "NOT_SUPPORTED",
            "evidence": "",
            "note": "parenthetical text is not in the opinion",
        })

    found = {
        "case_uid": "cap:1",
        "full_text": "primary opinion text only",
        "name": "Aguilar v. Atlantic Richfield Co.",
        "url": "u",
        "court": "Cal.",
        "decision_date": "2001-06-14",
    }
    verifier = LocalCaseVerifier(corpus=_FakeCorpus(found), llm_callback=llm)
    citation = Citation(
        kind="case",
        raw_text="25 Cal. 4th 826",
        normalized="25 Cal. 4th 826",
        reporter_citation="25 Cal. 4th 826",
        proposition="summary judgment burden",
    )

    cv = verifier.verify(citation)

    assert "primary opinion text only" in captured["user"]
    assert "summary judgment burden from parenthetical" not in captured["user"]
    assert cv.verdict == "NOT_SUPPORTED"
