from icharlotte_core.opposition.argument_research import research_argument


class StubCorpusClient:
    """Minimal cl_client: returns no corpus search hits (isolate firm path)."""
    def search_opinions(self, q, *, semantic=False, max_results=20, published_only=True):
        return []
    def get_opinion_text(self, uid):
        return ""
    def get_authority_signals(self, uid):
        return {}


class StubProvider:
    def __init__(self, cands): self._c = cands
    def candidates_for(self, proposition, *, motion_type, side, limit=6):
        return self._c


def _llm(_sys, _user):
    # rerank reply selecting the firm candidate with a verbatim passage.
    return '{"selections":[{"id":"cap:1","passage":"good faith effort","supports":"meet and confer required"}]}'


def test_firm_local_candidate_selected_and_tagged():
    firm = [{"cluster_id": "cap:1", "case_name": "Townsend v. Superior Court",
             "citation": "61 Cal.App.4th 1431", "year": "1998",
             "text": "a reasonable and good faith effort", "opinion_url": "",
             "source": "firm", "verification": "local",
             "source_brief": "x.pdf", "passage": "good faith effort",
             "proposition": "meet and confer required"}]
    out = research_argument(
        "Plaintiff failed to meet and confer",
        cl_client=StubCorpusClient(), query_llm=lambda s, u: '{"queries":[]}',
        rerank_llm=_llm, motion_type="compel", side="opposition",
        firm_provider=StubProvider(firm),
    )
    assert len(out) == 1
    assert out[0].source == "firm"
    assert out[0].verification == "local"
    assert out[0].source_brief == "x.pdf"


def test_unverified_firm_cite_appended_flagged():
    firm = [{"cluster_id": "firm:1", "case_name": "Smith v. Jones",
             "citation": "999 F.3d 1", "year": "2024", "text": "",
             "opinion_url": "", "source": "firm", "verification": "unverified_firm",
             "source_brief": "y.pdf", "passage": "federal rule applies",
             "proposition": "federal rule applies"}]
    out = research_argument(
        "Federal standard governs",
        cl_client=StubCorpusClient(), query_llm=lambda s, u: '{"queries":[]}',
        rerank_llm=lambda s, u: '{"selections":[]}', motion_type="compel",
        side="opposition", firm_provider=StubProvider(firm),
    )
    assert len(out) == 1
    assert out[0].verification == "unverified_firm"
    assert out[0].case_name == "Smith v. Jones"
    assert out[0].passage == "federal rule applies"
