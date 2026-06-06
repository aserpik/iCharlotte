import pytest

from icharlotte_core.chat.legal_research import (
    ChatResearchSettings,
    CourtListenerMode,
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
