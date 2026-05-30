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


if __name__ == "__main__":
    unittest.main()
