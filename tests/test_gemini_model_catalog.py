import json
from pathlib import Path
import unittest

from icharlotte_core.chat.token_counter import TokenCounter
from icharlotte_core.llm_config import (
    DEFAULT_MODEL_SEQUENCE,
    FAST_MODEL_SEQUENCE,
    GEMINI3_MODEL_SEQUENCE,
)
from icharlotte_core.ui.dialogs import AVAILABLE_MODELS as WORKBENCH_MODELS
from icharlotte_core.word_hotkey import AVAILABLE_MODELS as WORD_MODELS


PROJECT_ROOT = Path(__file__).resolve().parents[1]

CURRENT_GEMINI_MODEL_IDS = {
    "gemini-3.1-pro-preview",
    "gemini-3.5-flash",
    "gemini-3.1-flash-lite",
}

BLOCKED_GEMINI_MODEL_IDS = {
    "gemini-3.1-flash-lite-preview",
    "gemini-3-pro-preview",
    "gemini-3-flash-preview",
    "models/gemini-2.0-flash",
    "gemini-2.5-pro",
    "gemini-2.5-flash",
    "gemini-2.0-flash",
    "gemini-1.5-pro",
    "gemini-1.5-flash",
    "gemini-1.0-pro",
}


def _gemini_models_from_sequence(sequence):
    return [spec.model for spec in sequence if spec.provider == "Gemini"]


def _collect_json_model_ids(node):
    if isinstance(node, dict):
        for key, value in node.items():
            if key == "model" and isinstance(value, str):
                yield value
            else:
                yield from _collect_json_model_ids(value)
    elif isinstance(node, list):
        for item in node:
            yield from _collect_json_model_ids(item)


class GeminiModelCatalogTests(unittest.TestCase):
    def test_default_sequences_use_current_gemini_models(self):
        sequences = [
            DEFAULT_MODEL_SEQUENCE,
            FAST_MODEL_SEQUENCE,
            GEMINI3_MODEL_SEQUENCE,
        ]

        for model in [model for seq in sequences for model in _gemini_models_from_sequence(seq)]:
            self.assertIn(model, CURRENT_GEMINI_MODEL_IDS)

    def test_ui_model_selectors_only_offer_current_gemini_models(self):
        workbench_ids = [model_id for model_id, _ in WORKBENCH_MODELS["Gemini"]]
        word_ids = [model_id for _, provider, model_id in WORD_MODELS if provider == "Gemini"]

        for model in workbench_ids + word_ids:
            self.assertIn(model, CURRENT_GEMINI_MODEL_IDS)

    def test_preferences_file_contains_no_blocked_gemini_ids(self):
        prefs_path = PROJECT_ROOT / "config" / "llm_preferences.json"
        preferences = json.loads(prefs_path.read_text(encoding="utf-8"))
        model_ids = set(_collect_json_model_ids(preferences))

        blocked = sorted(model_ids & BLOCKED_GEMINI_MODEL_IDS)
        self.assertEqual(blocked, [])

    def test_token_counter_knows_current_gemini_models(self):
        # Each supported Gemini model must resolve to its real 1M window —
        # not the 128k ultimate fallback. The lookup is family-pattern based,
        # so we test the resolved value rather than dict membership.
        for model in CURRENT_GEMINI_MODEL_IDS:
            self.assertEqual(
                TokenCounter.get_context_limit(model),
                1_048_576,
                f"{model} should resolve to the Gemini 1M context window",
            )

    def test_runtime_source_has_no_blocked_hard_coded_gemini_ids(self):
        source_roots = [
            PROJECT_ROOT / "icharlotte_core",
            PROJECT_ROOT / "Scripts",
            PROJECT_ROOT / "config",
        ]
        checked_files = []
        matches = []

        for root in source_roots:
            patterns = ("*.py", "*.json")
            for pattern in patterns:
                for path in root.rglob(pattern):
                    if "__pycache__" in path.parts:
                        continue
                    checked_files.append(path)
                    text = path.read_text(encoding="utf-8", errors="ignore")
                    for model in BLOCKED_GEMINI_MODEL_IDS:
                        if model in text:
                            rel_path = path.relative_to(PROJECT_ROOT)
                            matches.append(f"{rel_path}: {model}")

        self.assertGreater(len(checked_files), 0)
        self.assertEqual(matches, [])
