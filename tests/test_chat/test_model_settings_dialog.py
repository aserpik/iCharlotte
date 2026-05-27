from types import SimpleNamespace


def _combo_items(combo):
    return [combo.itemText(i) for i in range(combo.count())]


def test_settings_dialog_limits_thinking_levels_for_gpt55_pro(qtbot):
    from icharlotte_core.ui.dialogs import SettingsDialog

    dlg = SettingsDialog(
        {"temperature": 1.0, "top_p": 0.95, "max_tokens": -1, "thinking_level": "None"},
        selected_model="gpt-5.5-pro",
    )
    qtbot.addWidget(dlg)

    assert _combo_items(dlg.thinking_combo) == ["Medium", "High", "xhigh"]
    assert dlg.thinking_combo.currentText() == "Medium"


def test_chat_tab_passes_selected_model_to_settings_dialog(monkeypatch):
    from icharlotte_core.ui import tabs

    captured = {}

    class FakeDialog:
        def __init__(self, settings, parent=None, selected_model=None):
            captured["settings"] = settings
            captured["parent"] = parent
            captured["selected_model"] = selected_model

        def exec(self):
            return True

        def get_settings(self):
            return {"thinking_level": "Medium"}

    class FakeModelCombo:
        def currentText(self):
            return "gpt-5.5-pro-2026-04-23"

    monkeypatch.setattr(tabs, "SettingsDialog", FakeDialog)
    fake_tab = SimpleNamespace(
        settings={"thinking_level": "None"},
        model_combo=FakeModelCombo(),
    )

    tabs.ChatTab.open_settings(fake_tab)

    assert captured["selected_model"] == "gpt-5.5-pro-2026-04-23"
    assert fake_tab.settings == {"thinking_level": "Medium"}
