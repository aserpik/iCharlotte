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


def test_run_chat_legal_research_llm_callback_uses_model_snapshot_during_status_events(
    qtbot,
    monkeypatch,
):
    _app()
    _clear_chat_research_settings()
    from icharlotte_core import llm
    from icharlotte_core.ui import tabs

    captured = {}

    def fake_generate(**kwargs):
        captured["llm_kwargs"] = kwargs
        return "research synthesis"

    class FakeService:
        @classmethod
        def from_environment(cls, *, llm_callback):
            captured["llm_callback"] = llm_callback
            return cls()

        def research(self, *, user_text, context_text, settings, status_callback):
            status_callback("fake progress")
            captured["llm_callback"]("snapshot system", "snapshot user")
            return SimpleNamespace(
                selected_authorities=[],
                get_known_case_names=lambda: [],
                build_augmented_system_prompt=lambda base: base + "\nAUGMENTED",
                format_research_basis_html=lambda: ["<b>Legal Research Basis</b>"],
            )

    monkeypatch.setattr(tabs, "ChatLegalResearchService", FakeService)
    monkeypatch.setattr(llm.LLMHandler, "generate", fake_generate)
    tab = _make_chat_tab(qtbot, monkeypatch)
    tab.provider_combo.setCurrentText("Gemini")
    tab.model_combo.addItems(["gemini-snapshot", "gpt-live"])
    tab.model_combo.setCurrentText("gemini-snapshot")
    tab.settings = {
        "temperature": 0.6,
        "top_p": 0.8,
        "max_tokens": 2048,
        "thinking_level": "Low",
    }

    original_process_events = QApplication.processEvents

    def mutate_during_events():
        tab.provider_combo.setCurrentText("OpenAI")
        tab.model_combo.setCurrentText("gpt-live")
        tab.settings["max_tokens"] = 999
        tab.settings["thinking_level"] = "High"

    monkeypatch.setattr(QApplication, "processEvents", mutate_during_events)
    try:
        packet = tab._run_chat_legal_research("research this", "context text")
    finally:
        monkeypatch.setattr(QApplication, "processEvents", original_process_events)

    llm_kwargs = captured["llm_kwargs"]
    assert packet is not None
    assert llm_kwargs["provider"] == "Gemini"
    assert llm_kwargs["model"] == "gemini-snapshot"
    assert llm_kwargs["settings"]["max_tokens"] == 2048
    assert llm_kwargs["settings"]["thinking_level"] == "Low"
    assert llm_kwargs["settings"]["stream"] is False
    assert llm_kwargs["settings"]["temperature"] == 0.2


def test_send_message_uses_model_snapshot_for_final_worker_after_research_events(
    qtbot,
    monkeypatch,
):
    _app()
    _clear_chat_research_settings()
    from icharlotte_core.ui import tabs

    worker_args = {}

    class FakeSignal:
        def connect(self, callback):
            pass

    class FakeWorker:
        def __init__(
            self,
            provider,
            model,
            system,
            user,
            files,
            settings,
            history=None,
            media_files=None,
        ):
            worker_args.update(
                provider=provider,
                model=model,
                system=system,
                user=user,
                files=files,
                settings=settings,
                history=history,
                media_files=media_files,
            )
            self.new_token = FakeSignal()
            self.finished = FakeSignal()
            self.error = FakeSignal()

        def start(self):
            worker_args["started"] = True

    class FakePersistence:
        def add_message(self, conversation_id, message):
            pass

    class FakeService:
        @classmethod
        def from_environment(cls, *, llm_callback):
            return cls()

        def research(self, *, user_text, context_text, settings, status_callback):
            status_callback("fake progress")
            return SimpleNamespace(
                selected_authorities=[],
                cases=[],
                statutes=[],
                verification=[],
                get_known_case_names=lambda: [],
                build_augmented_system_prompt=lambda base: base + "\nAUGMENTED",
                format_research_basis_html=lambda: ["<b>Legal Research Basis</b>"],
            )

    monkeypatch.setattr(tabs, "ChatLegalResearchService", FakeService)
    monkeypatch.setattr(tabs, "LLMWorker", FakeWorker)
    tab = _make_chat_tab(qtbot, monkeypatch)
    tab.persistence = FakePersistence()
    tab.current_conversation_id = "conversation-1"
    tab.provider_combo.setCurrentText("Gemini")
    tab.model_combo.addItems(["gemini-snapshot", "gpt-live"])
    tab.model_combo.setCurrentText("gemini-snapshot")
    tab.settings = {
        "temperature": 0.6,
        "top_p": 0.8,
        "max_tokens": 2048,
        "thinking_level": "Low",
    }
    tab.legal_research_check.setChecked(True)
    tab.chat_input.setPlainText("research this")
    tab.read_files_content = lambda: ""
    tab.read_library_content = lambda: ""
    tab.get_attachment_info = lambda: []
    tab._get_checked_audio_files = lambda: []

    original_process_events = QApplication.processEvents

    def mutate_during_events():
        tab.provider_combo.setCurrentText("OpenAI")
        tab.model_combo.setCurrentText("gpt-live")
        tab.settings["max_tokens"] = 999
        tab.settings["thinking_level"] = "High"

    monkeypatch.setattr(QApplication, "processEvents", mutate_during_events)
    try:
        tab.send_message()
    finally:
        monkeypatch.setattr(QApplication, "processEvents", original_process_events)

    assert worker_args["started"] is True
    assert worker_args["provider"] == "Gemini"
    assert worker_args["model"] == "gemini-snapshot"
    assert worker_args["settings"]["max_tokens"] == 2048
    assert worker_args["settings"]["thinking_level"] == "Low"
    assert worker_args["settings"]["stream"] is True
    assert worker_args["media_files"] is None


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
    tab._pending_research = object()

    packet = tab._run_chat_legal_research("research this", "")

    assert packet is None
    assert tab._pending_research is None
    assert "No verified legal authorities were found." in tab.chat_history.toPlainText()
    assert tab.send_btn.isEnabled() is True
    assert tab.stop_btn.isEnabled() is False


def test_on_error_clears_pending_research_packet(qtbot, monkeypatch):
    _app()
    _clear_chat_research_settings()
    tab = _make_chat_tab(qtbot, monkeypatch)
    tab._pending_research = object()

    tab.on_error("LLM failed")

    assert tab._pending_research is None


def _make_research_packet(*, quote="The duty rule controls the negligence analysis."):
    from icharlotte_core.chat.legal_research import (
        ChatResearchPacket,
        ChatResearchSettings,
        ChatResearchSource,
        ChatSelectedAuthority,
    )

    return ChatResearchPacket(
        query="duty rule",
        settings=ChatResearchSettings.default(),
        selected_authorities=[
            ChatSelectedAuthority(
                id="cap:1",
                proposition="duty rule",
                case_name="Duty v. Care",
                citation="30 Cal. 4th 43",
                year="2020",
                reason="It states the governing duty rule.",
                supports="Duty controls negligence.",
                quote=quote,
                sources=[
                    ChatResearchSource(
                        kind="local_corpus",
                        label="Local California corpus",
                    )
                ],
            )
        ],
    )


def test_finalize_response_appends_research_basis_for_packet(qtbot, monkeypatch):
    _app()
    _clear_chat_research_settings()
    tab = _make_chat_tab(qtbot, monkeypatch)
    tab.stream_start_time = 1.0
    tab.stream_start_pos = tab.chat_history.textCursor().position()
    tab._pending_research = _make_research_packet()

    tab.finalize_response("Duty is governed by Duty v. Care (2020) 30 Cal. 4th 43.")

    plain = tab.chat_history.toPlainText()
    assert "Legal Research Basis" in plain
    assert "It states the governing duty rule." in plain
    assert "The duty rule controls the negligence analysis." in plain
    assert tab._pending_research is None


def test_finalize_response_keeps_research_basis_before_separator(qtbot, monkeypatch):
    _app()
    _clear_chat_research_settings()
    tab = _make_chat_tab(qtbot, monkeypatch)
    tab.stream_start_time = 1.0
    tab.stream_start_pos = tab.chat_history.textCursor().position()
    tab._pending_research = _make_research_packet()

    tab.finalize_response("Assistant answer.")

    plain = tab.chat_history.toPlainText()
    assert plain.index("Assistant answer.") < plain.index("Legal Research Basis")
    assert plain.index("Legal Research Basis") < plain.index("-" * 50)
    assert plain.count("-" * 50) == 1


def test_finalize_response_renders_research_basis_html_entities(qtbot, monkeypatch):
    _app()
    _clear_chat_research_settings()
    tab = _make_chat_tab(qtbot, monkeypatch)
    tab.stream_start_time = 1.0
    tab.stream_start_pos = tab.chat_history.textCursor().position()
    quote = 'The court said "quoted" & binding <rule>.'
    tab._pending_research = _make_research_packet(quote=quote)

    tab.finalize_response("Assistant answer.")

    plain = tab.chat_history.toPlainText()
    assert "&quot;" not in plain
    assert quote in plain


def test_finalize_response_persists_research_basis_with_assistant_message(
    qtbot,
    monkeypatch,
):
    _app()
    _clear_chat_research_settings()

    class FakePersistence:
        def __init__(self):
            self.messages = []

        def add_message(self, conversation_id, message):
            self.messages.append((conversation_id, message))

    tab = _make_chat_tab(qtbot, monkeypatch)
    tab.persistence = FakePersistence()
    tab.current_conversation_id = "conversation-1"
    tab.stream_start_time = 1.0
    tab.stream_start_pos = tab.chat_history.textCursor().position()
    quote = 'The court said "quoted" & binding <rule>.'
    tab._pending_research = _make_research_packet(quote=quote)

    tab.finalize_response("Assistant answer.")

    conversation_id, message = tab.persistence.messages[0]
    assert conversation_id == "conversation-1"
    assert "Assistant answer." in message.content
    assert "Legal Research Basis" in message.content
    assert quote in message.content
