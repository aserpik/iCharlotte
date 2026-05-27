import importlib
import os
import sys


def _reload_config():
    sys.modules.pop("icharlotte_core.config", None)
    return importlib.import_module("icharlotte_core.config")


def test_project_env_gemini_key_overrides_stale_inherited_env(monkeypatch, tmp_path):
    project_key = "AIzaSy" + ("1" * 33)

    monkeypatch.chdir(tmp_path)
    monkeypatch.setenv("GEMINI_API_KEY", "stale_invalid_key")
    monkeypatch.setenv("GOOGLE_API_KEY", "stale_google_key")
    (tmp_path / ".env").write_text(f"GEMINI_API_KEY={project_key}\n", encoding="utf-8")

    config = _reload_config()

    assert config.API_KEYS["Gemini"] == project_key
    assert os.environ["GEMINI_API_KEY"] == project_key
    assert "GOOGLE_API_KEY" not in os.environ


def test_placeholder_project_env_does_not_replace_inherited_key(monkeypatch, tmp_path):
    inherited_key = "AIzaSy" + ("2" * 33)

    monkeypatch.chdir(tmp_path)
    monkeypatch.setenv("GEMINI_API_KEY", inherited_key)
    (tmp_path / ".env").write_text(
        "GEMINI_API_KEY=your_gemini_api_key_here\n",
        encoding="utf-8",
    )

    config = _reload_config()

    assert config.API_KEYS["Gemini"] == inherited_key
    assert os.environ["GEMINI_API_KEY"] == inherited_key
