"""
Tests for icharlotte_core.discovery.engine module.

Covers:
- TestDiscoveryEngineStandardMode: generate_standard with real templates
- TestDiscoveryEngineCustomMode: build_llm_prompt variants and generate_custom
- TestDiscoveryEngineParseLLM: parse_llm_response delegation
- TestDiscoveryEngineAdditional: prepare_additional with mocked SetTracker
"""
import os
import unittest
from unittest.mock import patch, MagicMock

from icharlotte_core.discovery.models import (
    CustomStyle,
    DiscoveryMode,
    DiscoveryRequest,
    DiscoverySet,
    DiscoveryType,
    Party,
    PartyRole,
    SetTrackerResult,
)
from icharlotte_core.discovery.engine import DiscoveryEngine

# Path to the discovery template directory
DISCOVERY_DIR = os.path.join(os.path.dirname(__file__), "..", "..", "discovery")
DISCOVERY_DIR = os.path.normpath(DISCOVERY_DIR)

# Check if real template files are available
NEGLIGENCE_SI_PATH = os.path.join(
    DISCOVERY_DIR, "Standard Negligence Discovery (5800.070)", "SI(1) tPltf.docx"
)
DEFINED_TERMS_PATH = os.path.join(DISCOVERY_DIR, "DISCOVERY DEFINED TERMS.docx")
TEMPLATES_AVAILABLE = os.path.isfile(NEGLIGENCE_SI_PATH) and os.path.isfile(DEFINED_TERMS_PATH)

# Reusable party fixtures
PLAINTIFF = Party(
    name="John Smith",
    role=PartyRole.PLAINTIFF,
    is_our_client=False,
    abbreviation="Pltf",
)
DEFENDANT = Party(
    name="Servitek Electric, Inc.",
    role=PartyRole.DEFENDANT,
    is_our_client=True,
    abbreviation="Def",
)


# ---------------------------------------------------------------------------
# TestDiscoveryEngineStandardMode
# ---------------------------------------------------------------------------

@unittest.skipUnless(TEMPLATES_AVAILABLE, "Template files not found in discovery/")
class TestDiscoveryEngineStandardMode(unittest.TestCase):
    """Tests for generate_standard using real template files."""

    @classmethod
    def setUpClass(cls):
        cls.engine = DiscoveryEngine(DISCOVERY_DIR)

    def test_generate_standard_si(self):
        """Negligence SI standard generation produces >50 requests, set_number=1."""
        results = self.engine.generate_standard(
            standard_type="negligence",
            discovery_types=[DiscoveryType.SI],
            directed_to=PLAINTIFF,
            propounding_party=DEFENDANT,
        )
        self.assertEqual(len(results), 1)
        ds = results[0]
        self.assertIsInstance(ds, DiscoverySet)
        self.assertEqual(ds.discovery_type, DiscoveryType.SI)
        self.assertEqual(ds.set_number, 1)
        self.assertEqual(ds.previous_count, 0)
        self.assertGreater(len(ds.requests), 50)
        self.assertEqual(ds.requests[0].number, 1)
        self.assertEqual(ds.directed_to, PLAINTIFF)
        self.assertEqual(ds.propounding_party, DEFENDANT)

    def test_generate_standard_multiple_types(self):
        """Generating SI+RPD returns 2 DiscoverySets."""
        results = self.engine.generate_standard(
            standard_type="negligence",
            discovery_types=[DiscoveryType.SI, DiscoveryType.RPD],
            directed_to=PLAINTIFF,
            propounding_party=DEFENDANT,
        )
        self.assertEqual(len(results), 2)
        types_returned = {ds.discovery_type for ds in results}
        self.assertEqual(types_returned, {DiscoveryType.SI, DiscoveryType.RPD})
        for ds in results:
            self.assertEqual(ds.set_number, 1)
            self.assertEqual(ds.previous_count, 0)
            self.assertGreater(len(ds.requests), 0)

    def test_generate_standard_has_definitions(self):
        """Standard generation loads defined terms block."""
        results = self.engine.generate_standard(
            standard_type="negligence",
            discovery_types=[DiscoveryType.SI],
            directed_to=PLAINTIFF,
            propounding_party=DEFENDANT,
        )
        ds = results[0]
        self.assertTrue(len(ds.definitions_block) > 0, "Expected non-empty definitions")

    def test_generate_standard_has_instructions(self):
        """Standard generation loads instructions block."""
        results = self.engine.generate_standard(
            standard_type="negligence",
            discovery_types=[DiscoveryType.SI],
            directed_to=PLAINTIFF,
            propounding_party=DEFENDANT,
        )
        ds = results[0]
        self.assertTrue(len(ds.instructions_block) > 0, "Expected non-empty instructions")

    def test_generate_standard_with_variables(self):
        """Variable substitution is applied to definitions block."""
        results = self.engine.generate_standard(
            standard_type="negligence",
            discovery_types=[DiscoveryType.SI],
            directed_to=PLAINTIFF,
            propounding_party=DEFENDANT,
            variables={
                "responding_party_name": "JOHN SMITH",
                "propounding_party_name": "SERVITEK ELECTRIC, INC.",
            },
        )
        ds = results[0]
        # If the defined terms originally had ____ blanks near party names,
        # they should now be filled
        # We just verify no crash and definitions are still present
        self.assertTrue(len(ds.definitions_block) > 0)

    def test_generate_standard_invalid_type_raises(self):
        """Unknown standard_type raises ValueError."""
        with self.assertRaises(ValueError):
            self.engine.generate_standard(
                standard_type="unknown_type",
                discovery_types=[DiscoveryType.SI],
                directed_to=PLAINTIFF,
                propounding_party=DEFENDANT,
            )


# ---------------------------------------------------------------------------
# TestDiscoveryEngineCustomMode
# ---------------------------------------------------------------------------

class TestDiscoveryEngineCustomMode(unittest.TestCase):
    """Tests for custom mode prompt building and generate_custom."""

    @classmethod
    def setUpClass(cls):
        cls.engine = DiscoveryEngine(DISCOVERY_DIR)

    def test_build_custom_only_prompt(self):
        """CUSTOM_ONLY prompt contains type name, user instructions, context."""
        prompt = self.engine.build_llm_prompt(
            discovery_type=DiscoveryType.SI,
            user_instructions="Focus on medical treatment history.",
            custom_style=CustomStyle.CUSTOM_ONLY,
            context_text="Plaintiff was in a car accident on 1/1/2025.",
            start_number=1,
        )
        self.assertIn("SPECIAL INTERROGATORY", prompt)
        self.assertIn("Focus on medical treatment history", prompt)
        self.assertIn("car accident", prompt)
        self.assertIn("NO. 1", prompt)  # start number reference

    def test_build_standard_plus_custom_prompt(self):
        """STANDARD_PLUS_CUSTOM prompt includes start number and notes standard set."""
        prompt = self.engine.build_llm_prompt(
            discovery_type=DiscoveryType.RPD,
            user_instructions="Add requests about surveillance footage.",
            custom_style=CustomStyle.STANDARD_PLUS_CUSTOM,
            start_number=36,
            standard_requests_text="REQUEST FOR PRODUCTION NO. 1\nAll documents...",
        )
        self.assertIn("REQUEST FOR PRODUCTION", prompt)
        self.assertIn("NO. 36", prompt)
        self.assertIn("surveillance footage", prompt)
        # Should reference that standard set ends before this number
        self.assertIn("35", prompt)

    def test_build_modified_standard_prompt(self):
        """MODIFIED_STANDARD prompt includes standard requests text for LLM to modify."""
        standard_text = (
            "SPECIAL INTERROGATORY NO. 1\nState your full name.\n\n"
            "SPECIAL INTERROGATORY NO. 2\nIdentify all witnesses.\n"
        )
        prompt = self.engine.build_llm_prompt(
            discovery_type=DiscoveryType.SI,
            user_instructions="Make them more aggressive in tone.",
            custom_style=CustomStyle.MODIFIED_STANDARD,
            start_number=1,
            standard_requests_text=standard_text,
        )
        self.assertIn("SPECIAL INTERROGATORY", prompt)
        self.assertIn("State your full name", prompt)
        self.assertIn("more aggressive", prompt)
        self.assertIn("modify", prompt.lower())

    def test_build_prompt_empty_context(self):
        """Prompt works with empty context_text."""
        prompt = self.engine.build_llm_prompt(
            discovery_type=DiscoveryType.RFA,
            user_instructions="Focus on liability admissions.",
            custom_style=CustomStyle.CUSTOM_ONLY,
        )
        self.assertIn("REQUEST FOR ADMISSION", prompt)
        self.assertIn("liability admissions", prompt)

    def test_build_prompt_rpd_type(self):
        """RPD prompt uses correct header template."""
        prompt = self.engine.build_llm_prompt(
            discovery_type=DiscoveryType.RPD,
            user_instructions="Request all emails.",
            custom_style=CustomStyle.CUSTOM_ONLY,
            start_number=5,
        )
        self.assertIn("REQUEST FOR PRODUCTION NO.", prompt)
        self.assertIn("NO. 5", prompt)

    @unittest.skipUnless(TEMPLATES_AVAILABLE, "Template files not found in discovery/")
    def test_generate_custom_standard_plus_custom(self):
        """STANDARD_PLUS_CUSTOM loads standard requests; custom requests are empty (LLM fills)."""
        results = self.engine.generate_custom(
            custom_style=CustomStyle.STANDARD_PLUS_CUSTOM,
            standard_type="negligence",
            discovery_types=[DiscoveryType.SI],
            directed_to=PLAINTIFF,
            propounding_party=DEFENDANT,
            user_instructions="Add questions about surveillance.",
        )
        self.assertEqual(len(results), 1)
        ds = results[0]
        self.assertEqual(ds.discovery_type, DiscoveryType.SI)
        # Should have standard requests loaded
        self.assertGreater(len(ds.requests), 0)
        self.assertEqual(ds.set_number, 1)

    def test_generate_custom_custom_only(self):
        """CUSTOM_ONLY returns DiscoverySets with empty requests."""
        engine = DiscoveryEngine(DISCOVERY_DIR)
        results = engine.generate_custom(
            custom_style=CustomStyle.CUSTOM_ONLY,
            standard_type="negligence",
            discovery_types=[DiscoveryType.SI],
            directed_to=PLAINTIFF,
            propounding_party=DEFENDANT,
            user_instructions="Write interrogatories about damages.",
        )
        self.assertEqual(len(results), 1)
        ds = results[0]
        self.assertEqual(ds.discovery_type, DiscoveryType.SI)
        # Custom-only: requests are empty (LLM fills them)
        self.assertEqual(len(ds.requests), 0)
        self.assertEqual(ds.set_number, 1)

    def test_generate_custom_modified_standard(self):
        """MODIFIED_STANDARD returns DiscoverySets with empty requests."""
        engine = DiscoveryEngine(DISCOVERY_DIR)
        results = engine.generate_custom(
            custom_style=CustomStyle.MODIFIED_STANDARD,
            standard_type="negligence",
            discovery_types=[DiscoveryType.RPD],
            directed_to=PLAINTIFF,
            propounding_party=DEFENDANT,
            user_instructions="Simplify the language.",
        )
        self.assertEqual(len(results), 1)
        ds = results[0]
        self.assertEqual(ds.discovery_type, DiscoveryType.RPD)
        self.assertEqual(len(ds.requests), 0)


# ---------------------------------------------------------------------------
# TestDiscoveryEngineParseLLM
# ---------------------------------------------------------------------------

class TestDiscoveryEngineParseLLM(unittest.TestCase):
    """Tests for parse_llm_response."""

    @classmethod
    def setUpClass(cls):
        cls.engine = DiscoveryEngine(DISCOVERY_DIR)

    def test_parse_si_response(self):
        """Parses LLM response with SI headers into DiscoveryRequest list."""
        response = (
            "SPECIAL INTERROGATORY NO. 1:\n"
            "State your full name and all aliases used in the past 10 years.\n\n"
            "SPECIAL INTERROGATORY NO. 2:\n"
            "Identify each person who witnessed the INCIDENT.\n"
        )
        requests = self.engine.parse_llm_response(response, DiscoveryType.SI)
        self.assertEqual(len(requests), 2)
        self.assertEqual(requests[0].number, 1)
        self.assertEqual(requests[1].number, 2)
        self.assertIn("full name", requests[0].text)

    def test_parse_rpd_response(self):
        """Parses LLM response with RPD headers."""
        response = (
            "REQUEST FOR PRODUCTION NO. 5:\n"
            "All documents relating to the INCIDENT.\n\n"
            "REQUEST FOR PRODUCTION NO. 6:\n"
            "All photographs of the accident scene.\n"
        )
        requests = self.engine.parse_llm_response(response, DiscoveryType.RPD)
        self.assertEqual(len(requests), 2)
        self.assertEqual(requests[0].number, 5)
        self.assertEqual(requests[1].number, 6)

    def test_parse_empty_response(self):
        """Empty response returns empty list."""
        requests = self.engine.parse_llm_response("", DiscoveryType.SI)
        self.assertEqual(requests, [])


# ---------------------------------------------------------------------------
# TestDiscoveryEngineAdditional
# ---------------------------------------------------------------------------

class TestDiscoveryEngineAdditional(unittest.TestCase):
    """Tests for prepare_additional with mocked SetTracker."""

    def test_prepare_additional_uses_set_tracker(self):
        """prepare_additional creates DiscoverySets using SetTracker results."""
        engine = DiscoveryEngine(DISCOVERY_DIR)

        mock_result = SetTrackerResult(
            next_set_number=3,
            last_request_number=70,
            previous_definitions="PRIOR DEFINITIONS",
            previous_instructions="PRIOR INSTRUCTIONS",
            resolution_method="docx_extraction",
            source_file="/path/to/SI (2) tPltf.docx",
        )

        with patch('icharlotte_core.discovery.engine.SetTracker') as MockTracker:
            mock_tracker = MockTracker.return_value
            mock_tracker.resolve_previous_set.return_value = mock_result

            results = engine.prepare_additional(
                case_path="/some/case/path",
                discovery_types=[DiscoveryType.SI],
                directed_to=PLAINTIFF,
                propounding_party=DEFENDANT,
            )

        self.assertEqual(len(results), 1)
        ds = results[0]
        self.assertEqual(ds.discovery_type, DiscoveryType.SI)
        self.assertEqual(ds.set_number, 3)
        self.assertEqual(ds.previous_count, 70)
        self.assertEqual(ds.definitions_block, "PRIOR DEFINITIONS")
        self.assertEqual(ds.instructions_block, "PRIOR INSTRUCTIONS")
        # Requests empty -- LLM fills them
        self.assertEqual(len(ds.requests), 0)

    def test_prepare_additional_multiple_types(self):
        """prepare_additional handles multiple discovery types."""
        engine = DiscoveryEngine(DISCOVERY_DIR)

        si_result = SetTrackerResult(
            next_set_number=2, last_request_number=50,
            previous_definitions="SI DEFS", previous_instructions="SI INSTR",
        )
        rpd_result = SetTrackerResult(
            next_set_number=3, last_request_number=30,
            previous_definitions="RPD DEFS", previous_instructions="RPD INSTR",
        )

        with patch('icharlotte_core.discovery.engine.SetTracker') as MockTracker:
            mock_tracker = MockTracker.return_value
            mock_tracker.resolve_previous_set.side_effect = [si_result, rpd_result]

            results = engine.prepare_additional(
                case_path="/some/case/path",
                discovery_types=[DiscoveryType.SI, DiscoveryType.RPD],
                directed_to=PLAINTIFF,
                propounding_party=DEFENDANT,
            )

        self.assertEqual(len(results), 2)
        self.assertEqual(results[0].discovery_type, DiscoveryType.SI)
        self.assertEqual(results[0].set_number, 2)
        self.assertEqual(results[0].previous_count, 50)
        self.assertEqual(results[1].discovery_type, DiscoveryType.RPD)
        self.assertEqual(results[1].set_number, 3)
        self.assertEqual(results[1].previous_count, 30)


if __name__ == "__main__":
    unittest.main()
