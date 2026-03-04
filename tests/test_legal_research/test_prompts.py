"""Tests for legal research LLM prompt templates."""

import unittest

from icharlotte_core.legal_research.prompts import (
    QUERY_PLANNING_PROMPT,
    SYNTHESIS_PROMPT,
    VERIFICATION_PROMPT,
    CITATION_INSTRUCTION,
    build_augmented_system_prompt,
)


class TestQueryPlanningPrompt(unittest.TestCase):
    """Test QUERY_PLANNING_PROMPT constant."""

    def test_query_planning_prompt_exists(self):
        """Prompt contains 'search terms' and 'JSON'."""
        self.assertIn("search terms", QUERY_PLANNING_PROMPT.lower())
        self.assertIn("JSON", QUERY_PLANNING_PROMPT)

    def test_query_planning_prompt_has_output_structure(self):
        """Prompt specifies the expected JSON keys."""
        self.assertIn("case_queries", QUERY_PLANNING_PROMPT)
        self.assertIn("statute_queries", QUERY_PLANNING_PROMPT)
        self.assertIn("legal_doctrines", QUERY_PLANNING_PROMPT)

    def test_query_planning_prompt_mentions_california(self):
        self.assertIn("California", QUERY_PLANNING_PROMPT)


class TestSynthesisPrompt(unittest.TestCase):
    """Test SYNTHESIS_PROMPT constant."""

    def test_synthesis_prompt_exists(self):
        """Prompt contains 'cite'."""
        self.assertIn("cite", SYNTHESIS_PROMPT.lower())

    def test_synthesis_prompt_citation_format(self):
        """Prompt specifies California citation format."""
        self.assertIn("Case Name", SYNTHESIS_PROMPT)
        self.assertIn("Reporter", SYNTHESIS_PROMPT)

    def test_synthesis_prompt_negative_treatment(self):
        self.assertIn("negative treatment", SYNTHESIS_PROMPT.lower())


class TestVerificationPrompt(unittest.TestCase):
    """Test VERIFICATION_PROMPT constant."""

    def test_verification_prompt_exists(self):
        """Prompt contains PASS, FIXED, FLAGGED statuses."""
        self.assertIn("PASS", VERIFICATION_PROMPT)
        self.assertIn("FIXED", VERIFICATION_PROMPT)
        self.assertIn("FLAGGED", VERIFICATION_PROMPT)

    def test_verification_prompt_check_categories(self):
        for category in ("EXISTENCE", "ACCURACY", "SUPPORT", "TREATMENT"):
            self.assertIn(category, VERIFICATION_PROMPT)

    def test_verification_prompt_output_keys(self):
        self.assertIn("verifications", VERIFICATION_PROMPT)
        self.assertIn("corrected_text", VERIFICATION_PROMPT)

    def test_verification_prompt_severity(self):
        self.assertIn("malpractice", VERIFICATION_PROMPT.lower())


class TestCitationInstruction(unittest.TestCase):
    """Test CITATION_INSTRUCTION constant."""

    def test_citation_instruction_anti_hallucination(self):
        """Contains 'MUST ONLY cite' and 'Do NOT fabricate'."""
        self.assertIn("MUST ONLY cite", CITATION_INSTRUCTION)
        self.assertIn("Do NOT fabricate", CITATION_INSTRUCTION)


class TestBuildAugmentedSystemPrompt(unittest.TestCase):
    """Test build_augmented_system_prompt function."""

    def test_build_augmented_system_prompt(self):
        """Combines base prompt, citation instruction, and authority block."""
        base = "You are a legal analyst."
        authority = "Smith v. Jones (2020) 50 Cal.App.5th 100"

        result = build_augmented_system_prompt(base, authority)

        self.assertIn(base, result)
        self.assertIn(CITATION_INSTRUCTION, result)
        self.assertIn(authority, result)

    def test_build_augmented_system_prompt_ordering(self):
        """Authority block appears after citation instruction."""
        base = "Base prompt."
        authority = "AUTHORITY_BLOCK_HERE"

        result = build_augmented_system_prompt(base, authority)

        ci_pos = result.index(CITATION_INSTRUCTION)
        auth_pos = result.index(authority)
        self.assertGreater(auth_pos, ci_pos)


if __name__ == "__main__":
    unittest.main()
