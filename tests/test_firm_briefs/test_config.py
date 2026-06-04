# tests/test_firm_briefs/test_config.py
import os
import importlib


def test_firm_briefs_config_present():
    config = importlib.import_module("icharlotte_core.config")
    assert isinstance(config.FIRM_BRIEFS_DATA_DIR, str)
    assert config.FIRM_BRIEFS_DATA_DIR  # non-empty
    assert isinstance(config.FIRM_BRIEFS_ROOTS, list)
    # Default seeds the 5800 library if present in cwd, but the value must be a list.
    assert all(isinstance(p, str) for p in config.FIRM_BRIEFS_ROOTS)


def test_firm_briefs_data_dir_env_override(monkeypatch):
    monkeypatch.setenv("FIRM_BRIEFS_DATA_DIR", os.path.join("X:", "fb"))
    config = importlib.reload(importlib.import_module("icharlotte_core.config"))
    assert config.FIRM_BRIEFS_DATA_DIR == os.path.join("X:", "fb")
