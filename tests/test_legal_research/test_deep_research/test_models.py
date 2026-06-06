from icharlotte_core.legal_research.deep_research import (
    AuthorityCandidate,
    CourtListenerMode,
    DeepResearchRequest,
    ParentheticalWeightPolicy,
    ResearchSurface,
    ResearchTaskType,
    SourcePolicy,
    TreatmentClassification,
    TreatmentSignal,
    normalize_source_policy,
)
from icharlotte_core.legal_research.deep_research.models import _bool_value


def test_default_request_is_california_fail_closed():
    request = DeepResearchRequest(question="What is the summary judgment standard?")

    assert request.surface == ResearchSurface.CHAT
    assert request.task_type == ResearchTaskType.DISCRETE_QUESTION
    assert request.jurisdiction == "California"
    assert request.fail_closed is True
    assert request.max_questions == 5


def test_source_policy_defaults_to_firm_local_and_courtlistener_fallback():
    policy = SourcePolicy.default()

    assert policy.firm_authority is True
    assert policy.local_corpus is True
    assert policy.courtlistener_mode == CourtListenerMode.FALLBACK_CURRENT_LAW
    assert policy.ca_leginfo is True
    assert policy.ca_courts_recent is True


def test_source_policy_from_values_handles_strings():
    policy = SourcePolicy.from_values(
        firm_authority="false",
        local_corpus="true",
        courtlistener_mode="always_search",
        ca_leginfo="0",
        ca_courts_recent="yes",
    )

    assert policy.firm_authority is False
    assert policy.local_corpus is True
    assert policy.courtlistener_mode == CourtListenerMode.ALWAYS_SEARCH
    assert policy.ca_leginfo is False
    assert policy.ca_courts_recent is True


def test_invalid_boolean_string_uses_default():
    policy = SourcePolicy.from_values(
        firm_authority="flase",
        local_corpus="tru",
        ca_leginfo="wat",
    )

    assert policy.firm_authority is True
    assert policy.local_corpus is True
    assert policy.ca_leginfo is True


def test_invalid_boolean_string_can_default_false():
    assert _bool_value("flase", False) is False


def test_fallback_mode_without_local_sources_becomes_always_search():
    policy = SourcePolicy(
        firm_authority=False,
        local_corpus=False,
        courtlistener_mode=CourtListenerMode.FALLBACK_CURRENT_LAW,
    )

    normalized = normalize_source_policy(policy)

    assert normalized.courtlistener_mode == CourtListenerMode.ALWAYS_SEARCH


def test_courtlistener_off_stays_off_without_local_sources():
    policy = SourcePolicy(
        firm_authority=False,
        local_corpus=False,
        courtlistener_mode=CourtListenerMode.OFF,
    )

    normalized = normalize_source_policy(policy)

    assert normalized.courtlistener_mode == CourtListenerMode.OFF


def test_parenthetical_weight_policy_defaults_to_ten_percent_cap():
    policy = ParentheticalWeightPolicy.default()

    assert policy.max_score_contribution == 0.10
    assert policy.allow_parenthetical_as_sole_support is False
    assert policy.duplicate_similarity_threshold == 0.92


def test_treatment_signal_round_trip_preserves_classification():
    signal = TreatmentSignal(
        signal_id="sig-1",
        described_citation="12 Cal.5th 100",
        citing_case_name="Later v. Case",
        citing_citation="15 Cal.5th 200",
        parenthetical_text="holding that the trial court abused its discretion",
        classification=TreatmentClassification.SUPPORTING,
        confidence=0.82,
    )

    restored = TreatmentSignal.from_dict(signal.to_dict())

    assert restored.signal_id == "sig-1"
    assert restored.classification == TreatmentClassification.SUPPORTING
    assert restored.parenthetical_text.startswith("holding")


def test_authority_candidate_exposes_parenthetical_fields():
    candidate = AuthorityCandidate(
        candidate_id="c1",
        case_name="Smith v. Jones",
        citation="12 Cal.5th 100",
        parenthetical_match_score=0.7,
        treatment_signals=[
            TreatmentSignal(
                signal_id="sig-1",
                described_citation="12 Cal.5th 100",
                parenthetical_text="explaining the rule",
            )
        ],
    )

    data = candidate.to_dict()

    assert data["parenthetical_match_score"] == 0.7
    assert data["treatment_signals"][0]["parenthetical_text"] == "explaining the rule"
