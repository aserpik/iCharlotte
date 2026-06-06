import os
from types import SimpleNamespace

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

import pytest
from PySide6.QtCore import QSettings
from PySide6.QtWidgets import QApplication

from icharlotte_core.chat.legal_research import CourtListenerMode


@pytest.fixture(autouse=True)
def isolated_qsettings(tmp_path):
    previous_default_format = QSettings.defaultFormat()
    QSettings.setPath(QSettings.Format.IniFormat, QSettings.Scope.UserScope, str(tmp_path))
    QSettings.setDefaultFormat(QSettings.Format.IniFormat)
    yield
    QSettings("iCharlotte", "iCharlotte").sync()
    QSettings.setDefaultFormat(previous_default_format)


def _app():
    return QApplication.instance() or QApplication([])


def _clear_chat_research_settings():
    settings = QSettings("iCharlotte", "iCharlotte")
    for key in (
        "chat_tab/legal_research_firm_authority",
        "chat_tab/legal_research_local_corpus",
        "chat_tab/legal_research_courtlistener_mode",
    ):
        settings.remove(key)
    settings.sync()


def _make_chat_tab(qtbot, monkeypatch):
    from icharlotte_core.ui import tabs

    monkeypatch.setattr(tabs.ChatTab, "update_models", lambda self, provider: None)
    tab = tabs.ChatTab()
    qtbot.addWidget(tab)
    return tab


def test_chat_research_source_defaults(qtbot, monkeypatch):
    _app()
    _clear_chat_research_settings()

    tab = _make_chat_tab(qtbot, monkeypatch)

    settings = tab._current_chat_research_settings()

    assert settings.firm_authority is True
    assert settings.local_corpus is True
    assert settings.courtlistener_mode == CourtListenerMode.FALLBACK_CURRENT_LAW
    assert tab.research_sources_btn.text() == "Sources: Firm + Local + CL Fallback"


def test_chat_research_source_choices_persist(qtbot, monkeypatch):
    _app()
    _clear_chat_research_settings()

    tab = _make_chat_tab(qtbot, monkeypatch)
    tab.firm_authority_action.setChecked(False)
    tab.local_corpus_action.setChecked(False)
    tab.courtlistener_always_action.setChecked(True)
    tab._on_research_source_changed()

    tab2 = _make_chat_tab(qtbot, monkeypatch)
    settings = tab2._current_chat_research_settings()

    assert settings.firm_authority is False
    assert settings.local_corpus is False
    assert settings.courtlistener_mode == CourtListenerMode.ALWAYS_SEARCH
    assert tab2.courtlistener_always_action.isChecked() is True


def test_chat_research_source_menu_normalizes_fallback_when_no_local_sources(
    qtbot,
    monkeypatch,
):
    _app()
    _clear_chat_research_settings()

    tab = _make_chat_tab(qtbot, monkeypatch)
    tab.firm_authority_action.setChecked(False)
    tab.local_corpus_action.setChecked(False)
    tab.courtlistener_fallback_action.setChecked(True)
    tab._on_research_source_changed()

    settings = tab._current_chat_research_settings()

    assert settings.courtlistener_mode == CourtListenerMode.ALWAYS_SEARCH
    assert tab.courtlistener_always_action.isChecked() is True


def test_run_chat_legal_research_passes_selected_settings(qtbot, monkeypatch):
    _app()
    _clear_chat_research_settings()
    from icharlotte_core.ui import tabs

    captured = {}

    class FakeService:
        @classmethod
        def from_environment(cls, *, llm_callback):
            captured["has_llm_callback"] = callable(llm_callback)
            return cls()

        def research(self, *, user_text, context_text, settings, status_callback):
            captured["user_text"] = user_text
            captured["context_text"] = context_text
            captured["settings"] = settings
            status_callback("fake progress")
            return SimpleNamespace(
                selected_authorities=[],
                get_known_case_names=lambda: [],
                build_augmented_system_prompt=lambda base: base + "\nAUGMENTED",
                format_research_basis_html=lambda: ["<b>Legal Research Basis</b>"],
            )

    monkeypatch.setattr(tabs, "ChatLegalResearchService", FakeService)
    tab = _make_chat_tab(qtbot, monkeypatch)
    tab.firm_authority_action.setChecked(False)
    tab.local_corpus_action.setChecked(False)
    tab.courtlistener_always_action.setChecked(True)
    tab._on_research_source_changed()

    packet = tab._run_chat_legal_research("research this", "context text")

    assert captured["has_llm_callback"] is True
    assert captured["user_text"] == "research this"
    assert captured["context_text"] == "context text"
    assert captured["settings"].firm_authority is False
    assert captured["settings"].local_corpus is False
    assert captured["settings"].courtlistener_mode == CourtListenerMode.ALWAYS_SEARCH
    assert "fake progress" in tab.chat_history.toPlainText()
    assert packet is not None


def test_run_chat_legal_research_fail_closed_restores_buttons(qtbot, monkeypatch):
    _app()
    _clear_chat_research_settings()
    from icharlotte_core.ui import tabs
    from icharlotte_core.chat.legal_research import ChatResearchError

    class FakeService:
        @classmethod
        def from_environment(cls, *, llm_callback):
            return cls()

        def research(self, *, user_text, context_text, settings, status_callback):
            raise ChatResearchError("No verified legal authorities were found.")

    monkeypatch.setattr(tabs, "ChatLegalResearchService", FakeService)
    tab = _make_chat_tab(qtbot, monkeypatch)
    tab.send_btn.setEnabled(False)
    tab.stop_btn.setEnabled(True)

    packet = tab._run_chat_legal_research("research this", "")

    assert packet is None
    assert "No verified legal authorities were found." in tab.chat_history.toPlainText()
    assert tab.send_btn.isEnabled() is True
    assert tab.stop_btn.isEnabled() is False
