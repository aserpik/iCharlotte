import os

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtCore import QSettings
from PySide6.QtWidgets import QApplication

from icharlotte_core.chat.legal_research import CourtListenerMode


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
