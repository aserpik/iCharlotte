import os
import json
import tempfile
import unittest
from unittest.mock import patch, MagicMock

from icharlotte_core import llm


class MapModelTests(unittest.TestCase):
    def test_thinking_maps_to_codex(self):
        self.assertEqual(llm._map_openai_model_to_codex("gpt-5.2-thinking"), "gpt-5.2-codex")

    def test_instant_maps_to_codex(self):
        self.assertEqual(llm._map_openai_model_to_codex("gpt-5.2-instant"), "gpt-5.2-codex")

    def test_codex_id_passthrough(self):
        self.assertEqual(llm._map_openai_model_to_codex("gpt-5-codex"), "gpt-5-codex")

    def test_unknown_gpt5_maps_to_generic_codex(self):
        self.assertEqual(llm._map_openai_model_to_codex("gpt-5.9-foo"), "gpt-5-codex")

    def test_gpt4o_unsupported(self):
        self.assertIsNone(llm._map_openai_model_to_codex("gpt-4o"))

    def test_o1_unsupported(self):
        self.assertIsNone(llm._map_openai_model_to_codex("o1"))

    def test_empty_unsupported(self):
        self.assertIsNone(llm._map_openai_model_to_codex(""))
        self.assertIsNone(llm._map_openai_model_to_codex(None))


class CodexAvailableTests(unittest.TestCase):
    def test_available_when_on_path_and_logged_in(self):
        with patch.object(llm.shutil, "which", return_value="C:/codex.exe"), \
             patch.object(llm.os.path, "isfile", return_value=True):
            self.assertTrue(llm.codex_available())

    def test_unavailable_when_not_on_path(self):
        with patch.object(llm.shutil, "which", return_value=None), \
             patch.object(llm.os.path, "isfile", return_value=True):
            self.assertFalse(llm.codex_available())

    def test_unavailable_when_not_logged_in(self):
        with patch.object(llm.shutil, "which", return_value="C:/codex.exe"), \
             patch.object(llm.os.path, "isfile", return_value=False):
            self.assertFalse(llm.codex_available())


class SubscriptionEnabledTests(unittest.TestCase):
    def _write_prefs(self, payload):
        fd, path = tempfile.mkstemp(suffix=".json")
        os.close(fd)
        with open(path, "w", encoding="utf-8") as f:
            json.dump(payload, f)
        self.addCleanup(lambda: os.remove(path))
        return path

    def test_enabled_when_flag_true(self):
        path = self._write_prefs({"openai_use_subscription": True})
        with patch.object(llm, "_subscription_prefs_path", return_value=path):
            self.assertTrue(llm.openai_subscription_enabled())

    def test_disabled_when_flag_false(self):
        path = self._write_prefs({"openai_use_subscription": False})
        with patch.object(llm, "_subscription_prefs_path", return_value=path):
            self.assertFalse(llm.openai_subscription_enabled())

    def test_default_true_when_key_missing(self):
        path = self._write_prefs({"version": "2.1"})
        with patch.object(llm, "_subscription_prefs_path", return_value=path):
            self.assertTrue(llm.openai_subscription_enabled())

    def test_default_true_when_file_missing(self):
        with patch.object(llm, "_subscription_prefs_path",
                          return_value="C:/nonexistent/does-not-exist.json"):
            self.assertTrue(llm.openai_subscription_enabled())


if __name__ == "__main__":
    unittest.main()
