"""Unit tests for MediationBriefGenerator.build_refinement_prompts."""
import unittest
from unittest.mock import patch

from icharlotte_core.mediation_brief import MediationBriefGenerator


class TestBuildRefinementPrompts(unittest.TestCase):
    def setUp(self):
        self.gen = MediationBriefGenerator()
        self.sections_dict = {
            "introduction": "Plaintiff Smith sues Defendant Jones for negligence.",
            "statement_of_facts": "On March 1, 2025, the parties met at the crossing.",
            "liability": "Defendant had the right of way under Vehicle Code 21800.",
            "damages": "Plaintiff alleges $50,000 in medical specials.",
            "settlement_position": "Defendant offers $15,000.",
            "conclusion": "For the foregoing reasons, liability is in dispute.",
        }

    @patch.object(MediationBriefGenerator, "get_style_excerpts", return_value={})
    def test_returns_system_and_full_prompt_strings(self, _mock_excerpts):
        system_prompt, full_prompt = self.gen.build_refinement_prompts(
            section_name="liability",
            sections_dict=self.sections_dict,
            instruction="Make the causation discussion more forceful.",
        )
        self.assertIsInstance(system_prompt, str)
        self.assertIsInstance(full_prompt, str)
        self.assertIn("defense litigation attorney", system_prompt)
        self.assertIn("Make the causation discussion more forceful.", full_prompt)

    @patch.object(MediationBriefGenerator, "get_style_excerpts", return_value={})
    def test_prompt_includes_current_section_text(self, _mock_excerpts):
        _, full_prompt = self.gen.build_refinement_prompts(
            section_name="liability",
            sections_dict=self.sections_dict,
            instruction="Tighten.",
        )
        # The refinement prompt should reference previously drafted sections
        # (via the PREVIOUSLY DRAFTED SECTIONS context block, which is how the
        # generator already passes prior sections into section prompts).
        # For LIABILITY, STATEMENT_OF_FACTS comes before it in GENERATION_ORDER.
        self.assertIn("March 1, 2025", full_prompt)

    @patch.object(MediationBriefGenerator, "get_style_excerpts", return_value={})
    def test_does_not_mutate_generator_sections(self, _mock_excerpts):
        self.gen.sections = {"introduction": "unchanged"}
        self.gen.build_refinement_prompts(
            section_name="liability",
            sections_dict=self.sections_dict,
            instruction="Tighten.",
        )
        self.assertEqual(self.gen.sections, {"introduction": "unchanged"})

    @patch.object(MediationBriefGenerator, "get_style_excerpts", return_value={})
    def test_does_not_mutate_generator_sections_on_exception(self, _mock_excerpts):
        self.gen.sections = {"introduction": "unchanged"}
        with patch.object(
            MediationBriefGenerator, "_build_section_prompt",
            side_effect=RuntimeError("boom"),
        ):
            with self.assertRaises(RuntimeError):
                self.gen.build_refinement_prompts(
                    section_name="liability",
                    sections_dict=self.sections_dict,
                    instruction="Tighten.",
                )
        self.assertEqual(self.gen.sections, {"introduction": "unchanged"})


if __name__ == "__main__":
    unittest.main()
