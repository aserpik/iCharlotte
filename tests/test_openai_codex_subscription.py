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


if __name__ == "__main__":
    unittest.main()
