import pytest

from icharlotte_core.chat.legal_research import (
    ChatLegalResearchService,
    ChatResearchSettings,
    CourtListenerMode,
    is_current_law_query,
    normalize_settings,
)
from icharlotte_core.legal_research.models import CaseResult


def test_default_settings_are_firm_local_and_courtlistener_fallback():
    settings = ChatResearchSettings.default()

    assert settings.firm_authority is True
    assert settings.local_corpus is True
    assert settings.courtlistener_mode == CourtListenerMode.FALLBACK_CURRENT_LAW


def test_courtlistener_fallback_without_local_sources_becomes_always():
    settings = ChatResearchSettings(
        firm_authority=False,
        local_corpus=False,
        courtlistener_mode=CourtListenerMode.FALLBACK_CURRENT_LAW,
    )

    normalized = normalize_settings(settings)

    assert normalized.firm_authority is False
    assert normalized.local_corpus is False
    assert normalized.courtlistener_mode == CourtListenerMode.ALWAYS_SEARCH


def test_courtlistener_off_stays_off_without_local_sources():
    settings = ChatResearchSettings(
        firm_authority=False,
        local_corpus=False,
        courtlistener_mode=CourtListenerMode.OFF,
    )

    normalized = normalize_settings(settings)

    assert normalized.courtlistener_mode == CourtListenerMode.OFF


def test_settings_from_values_handles_qsettings_strings():
    settings = ChatResearchSettings.from_values(
        firm_authority="false",
        local_corpus="true",
        courtlistener_mode="always_search",
    )

    assert settings.firm_authority is False
    assert settings.local_corpus is True
    assert settings.courtlistener_mode == CourtListenerMode.ALWAYS_SEARCH


def test_unknown_courtlistener_mode_uses_default():
    settings = ChatResearchSettings.from_values(
        firm_authority=True,
        local_corpus=True,
        courtlistener_mode="unknown",
    )

    assert settings.courtlistener_mode == CourtListenerMode.FALLBACK_CURRENT_LAW


def test_settings_from_values_preserves_enum_mode():
    settings = ChatResearchSettings.from_values(
        courtlistener_mode=CourtListenerMode.ALWAYS_SEARCH,
    )

    assert settings.courtlistener_mode == CourtListenerMode.ALWAYS_SEARCH


def test_settings_from_values_preserves_enum_off_without_local_sources():
    settings = ChatResearchSettings.from_values(
        firm_authority=False,
        local_corpus=False,
        courtlistener_mode=CourtListenerMode.OFF,
    )

    assert settings.firm_authority is False
    assert settings.local_corpus is False
    assert settings.courtlistener_mode == CourtListenerMode.OFF


def test_extract_propositions_from_json_response():
    calls = []

    def llm(system_prompt, user_prompt):
        calls.append((system_prompt, user_prompt))
        return '{"propositions":["landlord duty to repair stairs","comparative fault open and obvious condition"]}'

    service = ChatLegalResearchService(llm_callback=llm)

    props = service.extract_propositions(
        user_text="Can we defeat summary judgment on premises liability?",
        context_text="Plaintiff fell on stairs after repeated repair requests.",
    )

    assert props == [
        "landlord duty to repair stairs",
        "comparative fault open and obvious condition",
    ]
    assert "Plaintiff fell on stairs" in calls[0][1]


def test_extract_propositions_accepts_positional_llm_callback():
    service = ChatLegalResearchService(lambda _system, _user: "{}")

    props = service.extract_propositions(
        user_text="What is the California rule for negligent hiring?",
        context_text="",
    )

    assert props == ["What is the California rule for negligent hiring?"]


def test_extract_propositions_falls_back_to_user_text_when_llm_returns_bad_json():
    service = ChatLegalResearchService(llm_callback=lambda _system, _user: "not json")

    props = service.extract_propositions(
        user_text="What is the California rule for negligent hiring?",
        context_text="",
    )

    assert props == ["What is the California rule for negligent hiring?"]


def test_extract_propositions_limits_to_five_items_and_drops_blanks():
    service = ChatLegalResearchService(
        llm_callback=lambda _system, _user: (
            '{"propositions":["a"," ","b","c","d","e","f"]}'
        )
    )

    props = service.extract_propositions(user_text="research", context_text="")

    assert props == ["a", "b", "c", "d", "e"]


@pytest.mark.parametrize(
    "query",
    [
        "Find the most recent California cases on discovery sanctions",
        "What is the current law on arbitration unconscionability?",
        "Are there any new cases about negligent hiring?",
        "Use up to date authority on premises liability",
        "Use current authority on premises liability",
        "Use up-to-date authority on premises liability",
    ],
)
def test_current_law_query_detection_positive(query):
    assert is_current_law_query(query) is True


def test_current_law_query_detection_negative():
    assert is_current_law_query("What is the rule for negligence duty?") is False


class FakeCorpusClient:
    def __init__(self, results=None, text_by_id=None, metadata=None):
        self.results = results or []
        self.text_by_id = text_by_id or {}
        self.metadata = metadata or {
            "source_counts": {"cl": 1},
            "max_decision_date": "2026-01-01",
        }
        self.calls = []
        self.signal_calls = []

    def search_opinions(self, query, *, semantic=False, max_results=15, published_only=True):
        self.calls.append((query, semantic, max_results, published_only))
        return self.results

    def get_opinion_text(self, case_uid):
        return self.text_by_id.get(str(case_uid), "")

    def get_authority_signals(self, case_uid):
        self.signal_calls.append(case_uid)
        return {"citation_count": 7, "latest_citing_year": "2025"}

    def corpus_metadata(self):
        return self.metadata


class FakeFirmProvider:
    def __init__(self, candidates=None):
        self.candidates = candidates or []
        self.calls = []

    def candidates_for(self, proposition, *, motion_type, side, limit=6):
        self.calls.append((proposition, motion_type, side, limit))
        return self.candidates


class FakeFirmProviderWithLiveFallback(FakeFirmProvider):
    def __init__(self, *, cl_client, candidates=None):
        super().__init__(candidates=candidates)
        self.cl_client = cl_client

    def candidates_for(self, proposition, *, motion_type, side, limit=6):
        if self.cl_client is not None:
            self.cl_client.search_opinions(
                proposition,
                semantic=True,
                max_results=limit,
                published_only=True,
            )
        return super().candidates_for(
            proposition,
            motion_type=motion_type,
            side=side,
            limit=limit,
        )


def _case(
    name="Duty v. Care",
    cite="30 Cal. 4th 43",
    uid="cap:1",
    text="The duty rule controls.",
):
    return CaseResult(
        name=name,
        citation=cite,
        date="2020-01-01",
        court="Cal.",
        snippet=text,
        url="https://example.test/case",
        cluster_id=uid,
    )


def _service_for_sources(*, local=None, firm=None, courtlistener=None, token=None):
    return ChatLegalResearchService(
        llm_callback=lambda _system, _user: '{"propositions":["duty rule"]}',
        local_corpus=local,
        firm_provider=firm,
        courtlistener_client=courtlistener,
        courtlistener_token=("token" if courtlistener else "") if token is None else token,
    )


def test_collect_local_only_searches_local_and_not_courtlistener():
    local = FakeCorpusClient(results=[_case()], text_by_id={"cap:1": "The duty rule controls."})
    cl = FakeCorpusClient(
        results=[_case(name="Live v. Case", cite="55 Cal.App.5th 10", uid="cl:1")]
    )
    service = _service_for_sources(local=local, courtlistener=cl)

    candidates, warnings, searches = service.collect_candidates(
        propositions=["duty rule"],
        settings=ChatResearchSettings(
            firm_authority=False,
            local_corpus=True,
            courtlistener_mode=CourtListenerMode.OFF,
        ),
        original_query="duty rule",
    )

    assert len(candidates) == 1
    assert candidates[0].case_name == "Duty v. Care"
    assert candidates[0].citation_count == 7
    assert candidates[0].latest_citing_year == "2025"
    assert len(local.calls) == 2
    assert local.signal_calls == ["cap:1"]
    assert cl.calls == []
    assert warnings == []
    assert any("Local California corpus" in item for item in searches)


def test_collect_firm_only_uses_firm_provider():
    firm = FakeFirmProvider(
        candidates=[
            {
                "cluster_id": "cap:firm",
                "case_name": "Townsend v. Superior Court",
                "citation": "61 Cal.App.4th 1431",
                "year": "1998",
                "text": "The court required a good faith effort.",
                "source": "firm",
                "verification": "local",
                "source_brief": "sample.pdf",
                "passage": "good faith effort",
                "proposition": "meet and confer required",
            }
        ]
    )
    service = _service_for_sources(firm=firm)

    candidates, warnings, searches = service.collect_candidates(
        propositions=["meet and confer required"],
        settings=ChatResearchSettings(
            firm_authority=True,
            local_corpus=False,
            courtlistener_mode=CourtListenerMode.OFF,
        ),
        original_query="meet and confer required",
    )

    assert len(candidates) == 1
    assert candidates[0].sources[0].kind == "firm"
    assert candidates[0].sources[0].reference == "sample.pdf"
    assert warnings == []
    assert searches == ["Firm/sample-motion authority: meet and confer required"]


def test_collect_firm_with_courtlistener_off_disables_provider_live_fallback():
    live_fallback = FakeCorpusClient(
        results=[_case(name="Live v. Case", cite="55 Cal.App.5th 10", uid="cl:1")]
    )
    firm = FakeFirmProviderWithLiveFallback(
        cl_client=live_fallback,
        candidates=[
            {
                "cluster_id": "cap:firm",
                "case_name": "Townsend v. Superior Court",
                "citation": "61 Cal.App.4th 1431",
                "year": "1998",
                "text": "The court required a good faith effort.",
                "source": "firm",
                "verification": "local",
                "source_brief": "sample.pdf",
                "passage": "good faith effort",
                "proposition": "meet and confer required",
            }
        ],
    )
    service = _service_for_sources(firm=firm)

    candidates, warnings, searches = service.collect_candidates(
        propositions=["meet and confer required"],
        settings=ChatResearchSettings(
            firm_authority=True,
            local_corpus=False,
            courtlistener_mode=CourtListenerMode.OFF,
        ),
        original_query="meet and confer required",
    )

    assert len(candidates) == 1
    assert live_fallback.calls == []
    assert firm.cl_client is live_fallback
    assert warnings == []
    assert searches == ["Firm/sample-motion authority: meet and confer required"]


def test_courtlistener_off_never_calls_live_client_even_with_thin_local_results():
    local = FakeCorpusClient(results=[])
    cl = FakeCorpusClient(
        results=[_case(name="Live v. Case", cite="55 Cal.App.5th 10", uid="cl:1")]
    )
    service = _service_for_sources(local=local, courtlistener=cl)

    candidates, warnings, searches = service.collect_candidates(
        propositions=["thin issue"],
        settings=ChatResearchSettings(
            firm_authority=False,
            local_corpus=True,
            courtlistener_mode=CourtListenerMode.OFF,
        ),
        original_query="thin issue",
    )

    assert candidates == []
    assert cl.calls == []
    assert any("Local corpus returned thin results" in warning for warning in warnings)


def test_courtlistener_off_never_calls_live_client_for_current_law_query():
    local = FakeCorpusClient(results=[])
    cl = FakeCorpusClient(
        results=[_case(name="Live v. Case", cite="55 Cal.App.5th 10", uid="cl:1")]
    )
    service = _service_for_sources(local=local, courtlistener=cl)

    service.collect_candidates(
        propositions=["recent discovery sanctions"],
        settings=ChatResearchSettings(
            firm_authority=False,
            local_corpus=True,
            courtlistener_mode=CourtListenerMode.OFF,
        ),
        original_query="Find the most recent California cases on discovery sanctions",
    )

    assert cl.calls == []


def test_courtlistener_fallback_calls_live_when_local_results_are_thin():
    local = FakeCorpusClient(results=[])
    cl = FakeCorpusClient(
        results=[_case(name="Live v. Case", cite="55 Cal.App.5th 10", uid="cl:1")],
        text_by_id={"cl:1": "Live authority supports the rule."},
    )
    service = _service_for_sources(local=local, courtlistener=cl)

    candidates, warnings, searches = service.collect_candidates(
        propositions=["thin issue"],
        settings=ChatResearchSettings(
            firm_authority=False,
            local_corpus=True,
            courtlistener_mode=CourtListenerMode.FALLBACK_CURRENT_LAW,
        ),
        original_query="thin issue",
    )

    assert any(candidate.case_name == "Live v. Case" for candidate in candidates)
    assert cl.calls
    assert any("CourtListener API" in item for item in searches)


def test_courtlistener_fallback_does_not_call_live_when_firm_results_are_not_thin():
    firm = FakeFirmProvider(
        candidates=[
            {
                "cluster_id": "firm:1",
                "case_name": "Alpha v. Rule",
                "citation": "10 Cal.App.5th 1",
                "year": "2020",
                "text": "Alpha supports the rule.",
            },
            {
                "cluster_id": "firm:2",
                "case_name": "Beta v. Rule",
                "citation": "20 Cal.App.5th 2",
                "year": "2021",
                "text": "Beta supports the rule.",
            },
            {
                "cluster_id": "firm:3",
                "case_name": "Gamma v. Rule",
                "citation": "30 Cal.App.5th 3",
                "year": "2022",
                "text": "Gamma supports the rule.",
            },
        ]
    )
    cl = FakeCorpusClient(
        results=[_case(name="Live v. Case", cite="55 Cal.App.5th 10", uid="cl:1")]
    )
    service = _service_for_sources(firm=firm, courtlistener=cl)

    candidates, warnings, searches = service.collect_candidates(
        propositions=["firm-supported issue"],
        settings=ChatResearchSettings(
            firm_authority=True,
            local_corpus=False,
            courtlistener_mode=CourtListenerMode.FALLBACK_CURRENT_LAW,
        ),
        original_query="firm-supported issue",
    )

    assert {candidate.case_name for candidate in candidates} == {
        "Alpha v. Rule",
        "Beta v. Rule",
        "Gamma v. Rule",
    }
    assert cl.calls == []
    assert warnings == []
    assert searches == ["Firm/sample-motion authority: firm-supported issue"]


def test_courtlistener_fallback_calls_live_when_firm_results_are_thin():
    firm = FakeFirmProvider(
        candidates=[
            {
                "cluster_id": "firm:1",
                "case_name": "Alpha v. Rule",
                "citation": "10 Cal.App.5th 1",
                "year": "2020",
                "text": "Alpha supports the rule.",
            }
        ]
    )
    cl = FakeCorpusClient(
        results=[_case(name="Live v. Case", cite="55 Cal.App.5th 10", uid="cl:1")],
        text_by_id={"cl:1": "Live authority supports the rule."},
    )
    service = _service_for_sources(firm=firm, courtlistener=cl)

    candidates, warnings, searches = service.collect_candidates(
        propositions=["thin firm issue"],
        settings=ChatResearchSettings(
            firm_authority=True,
            local_corpus=False,
            courtlistener_mode=CourtListenerMode.FALLBACK_CURRENT_LAW,
        ),
        original_query="thin firm issue",
    )

    assert any(candidate.case_name == "Alpha v. Rule" for candidate in candidates)
    assert any(candidate.case_name == "Live v. Case" for candidate in candidates)
    assert cl.calls
    assert any("Selected non-live sources returned thin results" in warning for warning in warnings)
    assert any("CourtListener API" in item for item in searches)


def test_courtlistener_fallback_counts_deduped_local_candidates_for_thinness():
    local = FakeCorpusClient(
        results=[_case(name="Local v. Case", cite="10 Cal.App.5th 1", uid="cap:local")],
        text_by_id={"cap:local": "Local authority supports the rule."},
    )
    cl = FakeCorpusClient(
        results=[_case(name="Live v. Case", cite="55 Cal.App.5th 10", uid="cl:1")],
        text_by_id={"cl:1": "Live authority supports the rule."},
    )
    service = _service_for_sources(local=local, courtlistener=cl)

    candidates, warnings, searches = service.collect_candidates(
        propositions=["thin issue"],
        settings=ChatResearchSettings(
            firm_authority=False,
            local_corpus=True,
            courtlistener_mode=CourtListenerMode.FALLBACK_CURRENT_LAW,
        ),
        original_query="thin issue",
    )

    assert any(candidate.case_name == "Local v. Case" for candidate in candidates)
    assert any(candidate.case_name == "Live v. Case" for candidate in candidates)
    assert len(local.calls) == 2
    assert cl.calls
    assert any("Local corpus returned thin results" in warning for warning in warnings)
    assert any("CourtListener API" in item for item in searches)


def test_courtlistener_fallback_calls_live_for_stale_local_metadata():
    local = FakeCorpusClient(
        results=[_case()],
        text_by_id={"cap:1": "The duty rule controls."},
        metadata={"source_counts": {"cl": 1}, "max_decision_date": "2020-01-01"},
    )
    cl = FakeCorpusClient(
        results=[_case(name="Live v. Case", cite="55 Cal.App.5th 10", uid="cl:1")]
    )
    service = _service_for_sources(local=local, courtlistener=cl)

    candidates, warnings, searches = service.collect_candidates(
        propositions=["duty rule"],
        settings=ChatResearchSettings(
            firm_authority=False,
            local_corpus=True,
            courtlistener_mode=CourtListenerMode.FALLBACK_CURRENT_LAW,
        ),
        original_query="duty rule",
    )

    assert any(candidate.case_name == "Live v. Case" for candidate in candidates)
    assert cl.calls
    assert any("Local corpus is stale" in warning for warning in warnings)
    assert any("CourtListener API" in item for item in searches)


def test_courtlistener_fallback_calls_live_when_local_metadata_has_no_recent_slice():
    local = FakeCorpusClient(
        results=[_case()],
        text_by_id={"cap:1": "The duty rule controls."},
        metadata={"source_counts": {"cl": 0}, "max_decision_date": "2026-01-01"},
    )
    cl = FakeCorpusClient(
        results=[_case(name="Live v. Case", cite="55 Cal.App.5th 10", uid="cl:1")]
    )
    service = _service_for_sources(local=local, courtlistener=cl)

    candidates, warnings, searches = service.collect_candidates(
        propositions=["duty rule"],
        settings=ChatResearchSettings(
            firm_authority=False,
            local_corpus=True,
            courtlistener_mode=CourtListenerMode.FALLBACK_CURRENT_LAW,
        ),
        original_query="duty rule",
    )

    assert any(candidate.case_name == "Live v. Case" for candidate in candidates)
    assert cl.calls
    assert any("no CourtListener recent slice" in warning for warning in warnings)
    assert any("CourtListener API" in item for item in searches)


def test_courtlistener_fallback_calls_live_for_current_law_query_even_with_local_results():
    local = FakeCorpusClient(results=[_case()], text_by_id={"cap:1": "The duty rule controls."})
    cl = FakeCorpusClient(results=[_case(name="Live v. Case", uid="cl:1")])
    service = _service_for_sources(local=local, courtlistener=cl)

    service.collect_candidates(
        propositions=["recent discovery sanctions"],
        settings=ChatResearchSettings(
            firm_authority=False,
            local_corpus=True,
            courtlistener_mode=CourtListenerMode.FALLBACK_CURRENT_LAW,
        ),
        original_query="Find the most recent California cases on discovery sanctions",
    )

    assert cl.calls


def test_courtlistener_always_search_calls_live_even_when_local_has_results():
    local = FakeCorpusClient(results=[_case()], text_by_id={"cap:1": "The duty rule controls."})
    cl = FakeCorpusClient(
        results=[_case(name="Live v. Case", cite="55 Cal.App.5th 10", uid="cl:1")],
        text_by_id={"cl:1": "Live authority supports the rule."},
    )
    service = _service_for_sources(local=local, courtlistener=cl)

    candidates, warnings, searches = service.collect_candidates(
        propositions=["duty rule"],
        settings=ChatResearchSettings(
            firm_authority=False,
            local_corpus=True,
            courtlistener_mode=CourtListenerMode.ALWAYS_SEARCH,
        ),
        original_query="duty rule",
    )

    assert len(candidates) == 2
    assert cl.calls
    assert warnings == []


def test_courtlistener_selected_without_token_warns_and_skips_live():
    cl = FakeCorpusClient(results=[_case(name="Live v. Case", uid="cl:1")])
    service = _service_for_sources(courtlistener=cl, token="")

    candidates, warnings, searches = service.collect_candidates(
        propositions=["duty rule"],
        settings=ChatResearchSettings(
            firm_authority=False,
            local_corpus=False,
            courtlistener_mode=CourtListenerMode.ALWAYS_SEARCH,
        ),
        original_query="duty rule",
    )

    assert candidates == []
    assert searches == []
    assert cl.calls == []
    assert warnings == ["CourtListener API selected but COURTLISTENER_API_TOKEN is not set."]


def test_courtlistener_selected_without_client_warns_and_skips_live():
    service = ChatLegalResearchService(
        llm_callback=lambda _system, _user: '{"propositions":["duty rule"]}',
        courtlistener_client=None,
        courtlistener_token="token",
    )

    candidates, warnings, searches = service.collect_candidates(
        propositions=["duty rule"],
        settings=ChatResearchSettings(
            firm_authority=False,
            local_corpus=False,
            courtlistener_mode=CourtListenerMode.ALWAYS_SEARCH,
        ),
        original_query="duty rule",
    )

    assert candidates == []
    assert searches == []
    assert warnings == ["CourtListener API selected but COURTLISTENER_API_TOKEN is not set."]
