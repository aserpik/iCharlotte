"""Tests for the SetTracker module."""
import os
import re
import tempfile
import unittest
from unittest.mock import patch, MagicMock

from icharlotte_core.discovery.models import (
    DiscoveryType,
    Party,
    PartyRole,
    SetTrackerResult,
)
from icharlotte_core.discovery.set_tracker import SetTracker


class TestSetTrackerParseFilename(unittest.TestCase):
    """Test static filename parsing."""

    def test_si_set_one_pltf(self):
        result = SetTracker.parse_discovery_filename("SI (1) tPLF.pdf")
        self.assertIsNotNone(result)
        self.assertEqual(result["type"], "SI")
        self.assertEqual(result["set_number"], 1)
        self.assertEqual(result["party_abbrev"], "PLF")

    def test_rpd_set_two_city(self):
        result = SetTracker.parse_discovery_filename("RPD(2) tCity.pdf")
        self.assertIsNotNone(result)
        self.assertEqual(result["type"], "RPD")
        self.assertEqual(result["set_number"], 2)
        self.assertEqual(result["party_abbrev"], "City")

    def test_rfa_set_three_pltf(self):
        result = SetTracker.parse_discovery_filename("RFA (3) tPltf.pdf")
        self.assertIsNotNone(result)
        self.assertEqual(result["type"], "RFA")
        self.assertEqual(result["set_number"], 3)
        self.assertEqual(result["party_abbrev"], "Pltf")

    def test_random_file_returns_none(self):
        result = SetTracker.parse_discovery_filename("random_file.pdf")
        self.assertIsNone(result)

    def test_non_discovery_doc(self):
        result = SetTracker.parse_discovery_filename("deposition_transcript.pdf")
        self.assertIsNone(result)

    def test_case_insensitive(self):
        result = SetTracker.parse_discovery_filename("si (1) tPLF.pdf")
        self.assertIsNotNone(result)
        self.assertEqual(result["type"], "SI")

    def test_docx_extension(self):
        result = SetTracker.parse_discovery_filename("SI (1) tPLF.docx")
        self.assertIsNotNone(result)
        self.assertEqual(result["type"], "SI")
        self.assertEqual(result["set_number"], 1)


class TestSetTrackerScanFilenames(unittest.TestCase):
    """Test scanning propounded folder for existing discovery sets."""

    def setUp(self):
        self.tmpdir = tempfile.mkdtemp()

    def tearDown(self):
        import shutil
        shutil.rmtree(self.tmpdir, ignore_errors=True)

    def _make_case_structure(self, propounded_folder_name="fOUR Client", files=None,
                              subfolder_files=None):
        """Create a mock case directory structure.

        Parameters
        ----------
        propounded_folder_name : str
            Name of the client folder under DISCOVERY/PROPOUNDED.
        files : list of str or None
            Filenames to create directly in the propounded folder.
        subfolder_files : dict of {subfolder_name: [filenames]} or None
            Filenames to create in party subfolders.
        """
        propounded = os.path.join(self.tmpdir, "DISCOVERY", "PROPOUNDED", propounded_folder_name)
        os.makedirs(propounded, exist_ok=True)

        if files:
            for f in files:
                open(os.path.join(propounded, f), "w").close()

        if subfolder_files:
            for subfolder, fnames in subfolder_files.items():
                sub_path = os.path.join(propounded, subfolder)
                os.makedirs(sub_path, exist_ok=True)
                for f in fnames:
                    open(os.path.join(sub_path, f), "w").close()

        return propounded

    def test_single_set_si(self):
        """One SI (1) file -> next set is 2."""
        self._make_case_structure(files=["SI (1) tPLF.pdf"])
        tracker = SetTracker(self.tmpdir)
        party = Party(name="John Smith", role=PartyRole.PLAINTIFF,
                       is_our_client=False, abbreviation="PLF")
        next_set = tracker.get_next_set_number(DiscoveryType.SI, party)
        self.assertEqual(next_set, 2)

    def test_multiple_sets(self):
        """SI(1), SI(2), RPD(1) -> SI next=3, RPD next=2."""
        self._make_case_structure(files=[
            "SI (1) tPLF.pdf",
            "SI (2) tPLF.pdf",
            "RPD (1) tPLF.pdf",
        ])
        tracker = SetTracker(self.tmpdir)
        party = Party(name="John Smith", role=PartyRole.PLAINTIFF,
                       is_our_client=False, abbreviation="PLF")
        self.assertEqual(tracker.get_next_set_number(DiscoveryType.SI, party), 3)
        self.assertEqual(tracker.get_next_set_number(DiscoveryType.RPD, party), 2)

    def test_party_subfolder(self):
        """Files in tPltf/ subfolder are found correctly."""
        self._make_case_structure(
            subfolder_files={"tPltf": ["SI (1) tPltf.pdf", "SI (2) tPltf.pdf"]}
        )
        tracker = SetTracker(self.tmpdir)
        party = Party(name="John Smith", role=PartyRole.PLAINTIFF,
                       is_our_client=False, abbreviation="Pltf")
        self.assertEqual(tracker.get_next_set_number(DiscoveryType.SI, party), 3)

    def test_no_previous_sets(self):
        """Empty folder -> next set is 1."""
        self._make_case_structure(files=[])
        tracker = SetTracker(self.tmpdir)
        party = Party(name="John Smith", role=PartyRole.PLAINTIFF,
                       is_our_client=False, abbreviation="PLF")
        self.assertEqual(tracker.get_next_set_number(DiscoveryType.SI, party), 1)

    def test_case_insensitive_folder(self):
        """'discovery/Propounded/four client' found via case-insensitive matching."""
        # Create with non-standard casing
        self._make_case_structure(
            propounded_folder_name="four client",
            files=["SI (1) tPLF.pdf"]
        )
        tracker = SetTracker(self.tmpdir)
        party = Party(name="John Smith", role=PartyRole.PLAINTIFF,
                       is_our_client=False, abbreviation="PLF")
        next_set = tracker.get_next_set_number(DiscoveryType.SI, party)
        self.assertEqual(next_set, 2)

    def test_mixed_case_discovery_folder(self):
        """'Discovery/propounded/fOUR Client' found with mixed casing."""
        # Manually create with mixed casing
        propounded = os.path.join(self.tmpdir, "Discovery", "propounded", "fOUR Client")
        os.makedirs(propounded, exist_ok=True)
        open(os.path.join(propounded, "RPD (1) tDef.pdf"), "w").close()

        tracker = SetTracker(self.tmpdir)
        party = Party(name="Acme Corp", role=PartyRole.DEFENDANT,
                       is_our_client=False, abbreviation="Def")
        self.assertEqual(tracker.get_next_set_number(DiscoveryType.RPD, party), 2)

    def test_no_propounded_folder(self):
        """No DISCOVERY folder at all -> next set is 1."""
        tracker = SetTracker(self.tmpdir)
        party = Party(name="John Smith", role=PartyRole.PLAINTIFF,
                       is_our_client=False, abbreviation="PLF")
        self.assertEqual(tracker.get_next_set_number(DiscoveryType.SI, party), 1)

    def test_only_matching_party(self):
        """Files for different parties are not counted for the target party."""
        self._make_case_structure(files=[
            "SI (1) tDef.pdf",
            "SI (2) tDef.pdf",
            "SI (1) tPLF.pdf",
        ])
        tracker = SetTracker(self.tmpdir)
        pltf = Party(name="John Smith", role=PartyRole.PLAINTIFF,
                      is_our_client=False, abbreviation="PLF")
        deft = Party(name="Acme Corp", role=PartyRole.DEFENDANT,
                      is_our_client=False, abbreviation="Def")
        self.assertEqual(tracker.get_next_set_number(DiscoveryType.SI, pltf), 2)
        self.assertEqual(tracker.get_next_set_number(DiscoveryType.SI, deft), 3)


class TestSetTrackerGetLastRequestNumber(unittest.TestCase):
    """Test reading last request number from document content."""

    def setUp(self):
        self.tmpdir = tempfile.mkdtemp()

    def tearDown(self):
        import shutil
        shutil.rmtree(self.tmpdir, ignore_errors=True)

    def test_returns_zero_when_no_files(self):
        """No propounded folder -> 0."""
        tracker = SetTracker(self.tmpdir)
        party = Party(name="John Smith", role=PartyRole.PLAINTIFF,
                       is_our_client=False, abbreviation="PLF")
        self.assertEqual(tracker.get_last_request_number(DiscoveryType.SI, party), 0)

    @patch.object(SetTracker, '_read_file_text')
    def test_parses_si_request_numbers(self, mock_read):
        """Should find highest SPECIAL INTERROGATORY NO. in the file."""
        propounded = os.path.join(self.tmpdir, "DISCOVERY", "PROPOUNDED", "fOUR Client")
        os.makedirs(propounded, exist_ok=True)
        filepath = os.path.join(propounded, "SI (1) tPLF.docx")
        open(filepath, "w").close()

        mock_read.return_value = (
            "SPECIAL INTERROGATORY NO. 1\nState your name.\n\n"
            "SPECIAL INTERROGATORY NO. 2\nState your address.\n\n"
            "SPECIAL INTERROGATORY NO. 35\nDescribe the incident.\n"
        )

        tracker = SetTracker(self.tmpdir)
        party = Party(name="John Smith", role=PartyRole.PLAINTIFF,
                       is_our_client=False, abbreviation="PLF")
        self.assertEqual(tracker.get_last_request_number(DiscoveryType.SI, party), 35)

    @patch.object(SetTracker, '_read_file_text')
    def test_parses_rpd_request_numbers(self, mock_read):
        """Should find highest REQUEST FOR PRODUCTION NO. in the file."""
        propounded = os.path.join(self.tmpdir, "DISCOVERY", "PROPOUNDED", "fOUR Client")
        os.makedirs(propounded, exist_ok=True)
        filepath = os.path.join(propounded, "RPD (1) tPLF.docx")
        open(filepath, "w").close()

        mock_read.return_value = (
            "REQUEST FOR PRODUCTION NO. 1\nAll documents.\n\n"
            "REQUEST FOR PRODUCTION NO. 15\nAll photos.\n"
        )

        tracker = SetTracker(self.tmpdir)
        party = Party(name="John Smith", role=PartyRole.PLAINTIFF,
                       is_our_client=False, abbreviation="PLF")
        self.assertEqual(tracker.get_last_request_number(DiscoveryType.RPD, party), 15)


class TestSetTrackerResolvePreviousSet(unittest.TestCase):
    """Test full resolution with cascading fallback."""

    def setUp(self):
        self.tmpdir = tempfile.mkdtemp()

    def tearDown(self):
        import shutil
        shutil.rmtree(self.tmpdir, ignore_errors=True)

    def test_no_previous_sets_returns_defaults(self):
        """No files at all -> set 1, request 0, standard_definitions method."""
        # Create a discovery_dir with a defined terms file for the fallback
        tracker = SetTracker(self.tmpdir)
        party = Party(name="John Smith", role=PartyRole.PLAINTIFF,
                       is_our_client=False, abbreviation="PLF")

        # Use a mock for TemplateLoader since we don't have actual template files
        with patch('icharlotte_core.discovery.set_tracker.TemplateLoader') as MockLoader:
            mock_loader = MockLoader.return_value
            mock_loader.load_defined_terms.return_value = "STANDARD DEFINITIONS"
            mock_loader.load_instructions.return_value = "STANDARD INSTRUCTIONS"

            result = tracker.resolve_previous_set(
                DiscoveryType.SI, party, self.tmpdir
            )

        self.assertEqual(result.next_set_number, 1)
        self.assertEqual(result.last_request_number, 0)
        self.assertEqual(result.resolution_method, "standard_definitions")

    @patch.object(SetTracker, '_read_file_text')
    def test_resolves_from_docx(self, mock_read):
        """When a .docx file exists, extract definitions and instructions from it."""
        propounded = os.path.join(self.tmpdir, "DISCOVERY", "PROPOUNDED", "fOUR Client")
        os.makedirs(propounded, exist_ok=True)
        filepath = os.path.join(propounded, "SI (1) tPLF.docx")
        open(filepath, "w").close()

        mock_read.return_value = (
            "DEFINITIONS\n\n"
            '"INCIDENT" means the accident of January 1, 2025.\n\n'
            "INSTRUCTIONS TO ANSWERING PARTY\n\n"
            "Answer each interrogatory separately.\n\n"
            "SPECIAL INTERROGATORIES\n\n"
            "SPECIAL INTERROGATORY NO. 1\nState your name.\n\n"
            "SPECIAL INTERROGATORY NO. 10\nDescribe the incident.\n"
        )

        tracker = SetTracker(self.tmpdir)
        party = Party(name="John Smith", role=PartyRole.PLAINTIFF,
                       is_our_client=False, abbreviation="PLF")

        result = tracker.resolve_previous_set(
            DiscoveryType.SI, party, self.tmpdir
        )

        self.assertEqual(result.next_set_number, 2)
        self.assertEqual(result.last_request_number, 10)
        self.assertIn("docx", result.resolution_method)
        self.assertIsNotNone(result.source_file)


class TestSetTrackerExtractSections(unittest.TestCase):
    """Test extraction of definitions and instructions from document text."""

    def test_extract_definitions(self):
        text = (
            "Some preamble text\n\n"
            "DEFINITIONS\n\n"
            '"INCIDENT" means the accident.\n'
            '"DOCUMENT" means any writing.\n\n'
            "INSTRUCTIONS TO ANSWERING PARTY\n\n"
            "You must answer fully.\n"
        )
        tracker = SetTracker(".")
        defs = tracker._extract_definitions(text)
        self.assertIn("INCIDENT", defs)
        self.assertIn("DOCUMENT", defs)
        # Should not include instructions
        self.assertNotIn("You must answer fully", defs)

    def test_extract_instructions(self):
        text = (
            "INSTRUCTIONS TO ANSWERING PARTY\n\n"
            "Answer each interrogatory separately.\n"
            "Do not object without basis.\n\n"
            "SPECIAL INTERROGATORIES\n\n"
            "SPECIAL INTERROGATORY NO. 1\nState your name.\n"
        )
        tracker = SetTracker(".")
        instr = tracker._extract_instructions(text, DiscoveryType.SI)
        self.assertIn("Answer each interrogatory separately", instr)
        # Should not include the requests themselves
        self.assertNotIn("State your name", instr)

    def test_extract_definitions_no_section(self):
        """If no definitions section, return empty string."""
        text = "SPECIAL INTERROGATORY NO. 1\nState your name.\n"
        tracker = SetTracker(".")
        defs = tracker._extract_definitions(text)
        self.assertEqual(defs, "")

    def test_extract_instructions_no_section(self):
        """If no instructions section, return empty string."""
        text = "SPECIAL INTERROGATORY NO. 1\nState your name.\n"
        tracker = SetTracker(".")
        instr = tracker._extract_instructions(text, DiscoveryType.SI)
        self.assertEqual(instr, "")


if __name__ == '__main__':
    unittest.main()
