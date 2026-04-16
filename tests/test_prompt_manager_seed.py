"""Tests for seed_pipeline_prompts() in PromptManager."""
import os
import shutil
import tempfile
import unittest

from icharlotte_core.prompt_manager import PromptManager


class TestSeedPipelinePrompts(unittest.TestCase):

    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        self.pm = PromptManager(prompts_dir=self.tmp)

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)

    def test_seed_creates_word_assistant_passes(self):
        self.pm.seed_pipeline_prompts()
        expected = [
            "system_prompt", "redline_system_prompt", "email_system_prompt",
            "redline_prefix", "placeholder_instructions",
            "cursor_instructions", "selection_instructions",
        ]
        for pass_name in expected:
            text = self.pm.get_prompt("word_assistant", pass_name)
            self.assertIsNotNone(text, f"word_assistant:{pass_name} missing")
            self.assertTrue(len(text) > 20, f"word_assistant:{pass_name} too short")

    def test_seed_creates_legal_research_passes(self):
        self.pm.seed_pipeline_prompts()
        expected = [
            "query_planning", "query_extraction", "synthesis",
            "verification", "relevance_ranking",
            "research_framing", "citation_instruction",
        ]
        for pass_name in expected:
            text = self.pm.get_prompt("legal_research", pass_name)
            self.assertIsNotNone(text, f"legal_research:{pass_name} missing")
            self.assertTrue(len(text) > 20, f"legal_research:{pass_name} too short")

    def test_seed_creates_mediation_brief_passes(self):
        self.pm.seed_pipeline_prompts()
        for pass_name in ["style_guide", "formatting_rules"]:
            text = self.pm.get_prompt("mediation_brief", pass_name)
            self.assertIsNotNone(text, f"mediation_brief:{pass_name} missing")
            self.assertTrue(len(text) > 20, f"mediation_brief:{pass_name} too short")

    def test_seed_is_idempotent(self):
        self.pm.seed_pipeline_prompts()
        v1_text = self.pm.get_prompt("word_assistant", "system_prompt")
        self.pm.seed_pipeline_prompts()
        v1_text_again = self.pm.get_prompt("word_assistant", "system_prompt")
        self.assertEqual(v1_text, v1_text_again)
        versions = self.pm.list_versions("word_assistant", "system_prompt")
        self.assertEqual(len(versions), 1)

    def test_seed_does_not_overwrite_user_edits(self):
        self.pm.seed_pipeline_prompts()
        self.pm.create_version(
            "word_assistant", "system_prompt",
            "My custom system prompt",
            version="v2", set_as_current=True,
        )
        custom = self.pm.get_prompt("word_assistant", "system_prompt")
        self.assertEqual(custom, "My custom system prompt")
        self.pm.seed_pipeline_prompts()
        after_seed = self.pm.get_prompt("word_assistant", "system_prompt")
        self.assertEqual(after_seed, "My custom system prompt")


if __name__ == "__main__":
    unittest.main()
