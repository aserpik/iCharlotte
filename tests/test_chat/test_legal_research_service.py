import pytest

from icharlotte_core.chat.legal_research import (
    ChatLegalResearchService,
    ChatResearchSettings,
    CourtListenerMode,
    is_current_law_query,
    normalize_settings,
)


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
