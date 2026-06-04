from icharlotte_core.firm_briefs.provider import FirmAuthorityProvider


class FakeIndex:
    def __init__(self, rows): self._rows = rows
    def authority_candidates(self, proposition, *, motion_type, limit=8):
        return self._rows


class FakeCorpus:
    def __init__(self, texts): self._texts = texts  # norm_cite -> (uid, text)
    def lookup_by_citation(self, citation):
        norm = (citation or "").replace(" ", "").lower()
        if norm in self._texts:
            uid, _ = self._texts[norm]
            return {"case_uid": uid, "citation": citation}
        return None
    def get_opinion_text(self, uid):
        for _norm, (u, text) in self._texts.items():
            if u == uid:
                return text
        return None


ROW = {"case_name": "Townsend v. Superior Court", "reporter_cite": "61 Cal.App.4th 1431",
       "year": "1998", "norm_cite": "61cal.app.4th1431",
       "proposition": "meet and confer required", "quoted_passage": "good faith effort",
       "source_brief": "Oppositions/Motion to Compel/x.pdf"}


def test_resolves_via_corpus_local():
    idx = FakeIndex([ROW])
    corpus = FakeCorpus({"61cal.app.4th1431": ("cap:1", "... reasonable and good faith effort ...")})
    prov = FirmAuthorityProvider(idx, corpus, cl_client=None)
    cands = prov.candidates_for("meet and confer", motion_type="compel", side="opposition")
    assert len(cands) == 1
    c = cands[0]
    assert c["source"] == "firm"
    assert c["verification"] == "local"
    assert c["cluster_id"] == "cap:1"
    assert "good faith" in c["text"]
    assert c["source_brief"].endswith("x.pdf")


def test_unverified_when_not_in_corpus_and_no_cl():
    idx = FakeIndex([ROW])
    corpus = FakeCorpus({})  # not found
    prov = FirmAuthorityProvider(idx, corpus, cl_client=None)
    cands = prov.candidates_for("meet and confer", motion_type="compel", side="opposition")
    assert len(cands) == 1
    c = cands[0]
    assert c["verification"] == "unverified_firm"
    assert c["text"] == ""  # no opinion text → handled specially downstream
    assert c["passage"] == "good faith effort"


def test_cl_fallback_verifies():
    idx = FakeIndex([ROW])
    corpus = FakeCorpus({})

    class FakeCL:
        def get_opinion_text(self, cite):
            return "court text mentioning good faith effort"
        def lookup_by_citation(self, cite):
            return {"case_uid": "cl:9"}
    prov = FirmAuthorityProvider(idx, corpus, cl_client=FakeCL())
    cands = prov.candidates_for("meet and confer", motion_type="compel", side="opposition")
    assert cands[0]["verification"] == "courtlistener"
    assert cands[0]["text"]
