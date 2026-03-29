"""Tests for discovery data models."""
import unittest

from icharlotte_core.discovery.models import (
    PartyRole,
    DiscoveryMode,
    CustomStyle,
    DiscoveryType,
    Party,
    DiscoveryRequest,
    DiscoverySet,
    SetTrackerResult,
    number_to_word,
    generate_abbreviation,
)


class TestNumberToWord(unittest.TestCase):
    """Test number_to_word conversion for 1-99."""

    def test_single_digits(self):
        self.assertEqual(number_to_word(1), "One")
        self.assertEqual(number_to_word(2), "Two")
        self.assertEqual(number_to_word(3), "Three")
        self.assertEqual(number_to_word(9), "Nine")

    def test_teens(self):
        self.assertEqual(number_to_word(10), "Ten")
        self.assertEqual(number_to_word(11), "Eleven")
        self.assertEqual(number_to_word(13), "Thirteen")
        self.assertEqual(number_to_word(19), "Nineteen")

    def test_tens(self):
        self.assertEqual(number_to_word(20), "Twenty")
        self.assertEqual(number_to_word(30), "Thirty")
        self.assertEqual(number_to_word(50), "Fifty")
        self.assertEqual(number_to_word(90), "Ninety")

    def test_compound(self):
        self.assertEqual(number_to_word(21), "Twenty-One")
        self.assertEqual(number_to_word(42), "Forty-Two")
        self.assertEqual(number_to_word(99), "Ninety-Nine")

    def test_out_of_range(self):
        with self.assertRaises(ValueError):
            number_to_word(0)
        with self.assertRaises(ValueError):
            number_to_word(100)
        with self.assertRaises(ValueError):
            number_to_word(-1)


class TestGenerateAbbreviation(unittest.TestCase):
    """Test automatic abbreviation generation for parties."""

    def test_single_plaintiff(self):
        pltf = Party(name="John Smith", role=PartyRole.PLAINTIFF, is_our_client=True)
        all_parties = [pltf]
        self.assertEqual(generate_abbreviation(pltf, all_parties), "Pltf")

    def test_single_defendant(self):
        deft = Party(name="Acme Corp", role=PartyRole.DEFENDANT, is_our_client=False)
        all_parties = [deft]
        self.assertEqual(generate_abbreviation(deft, all_parties), "Def")

    def test_single_cross_defendant(self):
        xdef = Party(name="XYZ LLC", role=PartyRole.CROSS_DEFENDANT, is_our_client=False)
        all_parties = [xdef]
        self.assertEqual(generate_abbreviation(xdef, all_parties), "XDef")

    def test_single_cross_complainant(self):
        xcmplnt = Party(name="ABC Inc", role=PartyRole.CROSS_COMPLAINANT, is_our_client=False)
        all_parties = [xcmplnt]
        self.assertEqual(generate_abbreviation(xcmplnt, all_parties), "XCmplnt")

    def test_multiple_defendants_distinctive_words(self):
        d1 = Party(name="City of Los Angeles", role=PartyRole.DEFENDANT, is_our_client=False)
        d2 = Party(name="Servitek Facility Services", role=PartyRole.DEFENDANT, is_our_client=False)
        all_parties = [d1, d2]
        self.assertEqual(generate_abbreviation(d1, all_parties), "City")
        self.assertEqual(generate_abbreviation(d2, all_parties), "Servitek")

    def test_multiple_plaintiffs_distinctive_words(self):
        p1 = Party(name="Jane Doe", role=PartyRole.PLAINTIFF, is_our_client=True)
        p2 = Party(name="John Smith", role=PartyRole.PLAINTIFF, is_our_client=False)
        all_parties = [p1, p2]
        # Should pick distinctive words
        abbr1 = generate_abbreviation(p1, all_parties)
        abbr2 = generate_abbreviation(p2, all_parties)
        self.assertNotEqual(abbr1, abbr2)
        # Both should be non-empty
        self.assertTrue(len(abbr1) > 0)
        self.assertTrue(len(abbr2) > 0)

    def test_city_special_case(self):
        """If first word is 'city', use 'City' as abbreviation."""
        city = Party(name="City of Sacramento", role=PartyRole.DEFENDANT, is_our_client=False)
        other = Party(name="ABC Corp", role=PartyRole.DEFENDANT, is_our_client=False)
        all_parties = [city, other]
        self.assertEqual(generate_abbreviation(city, all_parties), "City")

    def test_skip_common_words(self):
        """Should skip common words like 'of', 'the', 'inc'."""
        d1 = Party(name="The Corporation of Nothing Inc", role=PartyRole.DEFENDANT, is_our_client=False)
        d2 = Party(name="Widget Industries LLC", role=PartyRole.DEFENDANT, is_our_client=False)
        all_parties = [d1, d2]
        abbr1 = generate_abbreviation(d1, all_parties)
        # Should not be 'The', 'of', 'Inc'
        self.assertNotIn(abbr1.lower(), ["the", "of", "inc", "llc"])


class TestPartyRole(unittest.TestCase):
    """Test PartyRole enum values."""

    def test_roles_exist(self):
        self.assertEqual(PartyRole.PLAINTIFF.value, "plaintiff")
        self.assertEqual(PartyRole.DEFENDANT.value, "defendant")
        self.assertEqual(PartyRole.CROSS_DEFENDANT.value, "cross_defendant")
        self.assertEqual(PartyRole.CROSS_COMPLAINANT.value, "cross_complainant")


class TestDiscoveryType(unittest.TestCase):
    """Test DiscoveryType enum properties."""

    def test_si_properties(self):
        si = DiscoveryType.SI
        self.assertEqual(si.abbreviation, "SI")
        self.assertIn("2030", si.ccp_section)
        self.assertTrue(si.request_header_pattern)
        self.assertIn("SPECIAL INTERROGATOR", si.section_heading.upper())

    def test_rpd_properties(self):
        rpd = DiscoveryType.RPD
        self.assertEqual(rpd.abbreviation, "RPD")
        self.assertIn("2031", rpd.ccp_section)
        self.assertIn("REQUEST", rpd.section_heading.upper())

    def test_rfa_properties(self):
        rfa = DiscoveryType.RFA
        self.assertEqual(rfa.abbreviation, "RFA")
        self.assertIn("2033", rfa.ccp_section)
        self.assertIn("ADMISSION", rfa.section_heading.upper())

    def test_document_title_template(self):
        """Document title template should accept party names and set number."""
        si = DiscoveryType.SI
        title = si.document_title_template.format(
            propounding="PLAINTIFF", responding="DEFENDANT", set_word="One"
        )
        self.assertIn("PLAINTIFF", title)
        self.assertIn("DEFENDANT", title)
        self.assertIn("One", title)

    def test_request_header_template(self):
        """Request header template should accept a number."""
        si = DiscoveryType.SI
        header = si.request_header_template.format(number=1)
        self.assertIn("1", header)


class TestParty(unittest.TestCase):
    """Test Party dataclass."""

    def test_basic_creation(self):
        p = Party(name="John Smith", role=PartyRole.PLAINTIFF, is_our_client=True)
        self.assertEqual(p.name, "John Smith")
        self.assertEqual(p.role, PartyRole.PLAINTIFF)
        self.assertTrue(p.is_our_client)
        self.assertEqual(p.abbreviation, "")

    def test_with_abbreviation(self):
        p = Party(name="Acme Corp", role=PartyRole.DEFENDANT, is_our_client=False, abbreviation="Acme")
        self.assertEqual(p.abbreviation, "Acme")

    def test_role_label(self):
        p = Party(name="John Smith", role=PartyRole.PLAINTIFF, is_our_client=True)
        self.assertIn("Plaintiff", p.role_label)
        d = Party(name="Acme Corp", role=PartyRole.DEFENDANT, is_our_client=False)
        self.assertIn("Defendant", d.role_label)
        xd = Party(name="XYZ", role=PartyRole.CROSS_DEFENDANT, is_our_client=False)
        self.assertIn("Cross-Defendant", xd.role_label)
        xc = Party(name="ABC", role=PartyRole.CROSS_COMPLAINANT, is_our_client=False)
        self.assertIn("Cross-Complainant", xc.role_label)

    def test_formal_description(self):
        p = Party(name="John Smith", role=PartyRole.PLAINTIFF, is_our_client=True)
        desc = p.formal_description
        self.assertIn("John Smith", desc)
        self.assertIn("Plaintiff", desc)

    def test_to_dict_from_dict_roundtrip(self):
        p = Party(name="Acme Corp", role=PartyRole.DEFENDANT, is_our_client=False, abbreviation="Acme")
        d = p.to_dict()
        restored = Party.from_dict(d)
        self.assertEqual(restored.name, p.name)
        self.assertEqual(restored.role, p.role)
        self.assertEqual(restored.is_our_client, p.is_our_client)
        self.assertEqual(restored.abbreviation, p.abbreviation)


class TestDiscoveryRequest(unittest.TestCase):
    """Test DiscoveryRequest dataclass."""

    def test_basic_creation(self):
        req = DiscoveryRequest(number=1, text="Identify all documents.")
        self.assertEqual(req.number, 1)
        self.assertEqual(req.text, "Identify all documents.")
        self.assertEqual(req.definitions, [])

    def test_with_definitions(self):
        req = DiscoveryRequest(
            number=5,
            text="State all facts.",
            definitions=["INCIDENT means the accident of January 1, 2025."]
        )
        self.assertEqual(len(req.definitions), 1)


class TestDiscoverySet(unittest.TestCase):
    """Test DiscoverySet dataclass and properties."""

    def _make_set(self, discovery_type, num_requests=10, previous_count=0):
        pltf = Party(name="John Smith", role=PartyRole.PLAINTIFF, is_our_client=True, abbreviation="Pltf")
        deft = Party(name="Acme Corp", role=PartyRole.DEFENDANT, is_our_client=False, abbreviation="Def")
        requests = [
            DiscoveryRequest(number=i + 1, text=f"Request {i + 1}")
            for i in range(num_requests)
        ]
        return DiscoverySet(
            discovery_type=discovery_type,
            set_number=1,
            directed_to=deft,
            propounding_party=pltf,
            requests=requests,
            previous_count=previous_count,
        )

    def test_total_count(self):
        ds = self._make_set(DiscoveryType.SI, num_requests=10, previous_count=25)
        self.assertEqual(ds.total_count, 35)

    def test_total_count_no_previous(self):
        ds = self._make_set(DiscoveryType.SI, num_requests=10, previous_count=0)
        self.assertEqual(ds.total_count, 10)

    # --- needs_declaration tests ---

    def test_si_over_35_needs_declaration(self):
        ds = self._make_set(DiscoveryType.SI, num_requests=10, previous_count=26)
        self.assertEqual(ds.total_count, 36)
        self.assertTrue(ds.needs_declaration)

    def test_si_at_35_no_declaration(self):
        ds = self._make_set(DiscoveryType.SI, num_requests=10, previous_count=25)
        self.assertEqual(ds.total_count, 35)
        self.assertFalse(ds.needs_declaration)

    def test_si_under_35_no_declaration(self):
        ds = self._make_set(DiscoveryType.SI, num_requests=5, previous_count=0)
        self.assertFalse(ds.needs_declaration)

    def test_rfa_over_35_needs_declaration(self):
        ds = self._make_set(DiscoveryType.RFA, num_requests=10, previous_count=26)
        self.assertTrue(ds.needs_declaration)

    def test_rfa_under_35_no_declaration(self):
        ds = self._make_set(DiscoveryType.RFA, num_requests=5, previous_count=0)
        self.assertFalse(ds.needs_declaration)

    def test_rpd_never_needs_declaration(self):
        """RPD never requires a declaration regardless of count."""
        ds = self._make_set(DiscoveryType.RPD, num_requests=50, previous_count=100)
        self.assertEqual(ds.total_count, 150)
        self.assertFalse(ds.needs_declaration)

    def test_rpd_small_count_no_declaration(self):
        ds = self._make_set(DiscoveryType.RPD, num_requests=5, previous_count=0)
        self.assertFalse(ds.needs_declaration)

    # --- set_word ---

    def test_set_word(self):
        ds = self._make_set(DiscoveryType.SI)
        self.assertEqual(ds.set_word, "One")

    def test_set_word_higher_number(self):
        pltf = Party(name="John Smith", role=PartyRole.PLAINTIFF, is_our_client=True, abbreviation="Pltf")
        deft = Party(name="Acme Corp", role=PartyRole.DEFENDANT, is_our_client=False, abbreviation="Def")
        ds = DiscoverySet(
            discovery_type=DiscoveryType.SI,
            set_number=3,
            directed_to=deft,
            propounding_party=pltf,
            requests=[],
        )
        self.assertEqual(ds.set_word, "Three")

    # --- filename ---

    def test_filename(self):
        ds = self._make_set(DiscoveryType.SI)
        fn = ds.filename
        self.assertIn("SI", fn)
        self.assertIn(".docx", fn)

    # --- plain_text ---

    def test_plain_text(self):
        ds = self._make_set(DiscoveryType.SI, num_requests=3)
        text = ds.plain_text()
        self.assertIn("Request 1", text)
        self.assertIn("Request 2", text)
        self.assertIn("Request 3", text)

    # --- defaults ---

    def test_defaults(self):
        pltf = Party(name="John Smith", role=PartyRole.PLAINTIFF, is_our_client=True)
        deft = Party(name="Acme Corp", role=PartyRole.DEFENDANT, is_our_client=False)
        ds = DiscoverySet(
            discovery_type=DiscoveryType.SI,
            set_number=1,
            directed_to=deft,
            propounding_party=pltf,
            requests=[],
        )
        self.assertEqual(ds.definitions_block, "")
        self.assertEqual(ds.instructions_block, "")
        self.assertEqual(ds.previous_count, 0)


class TestSetTrackerResult(unittest.TestCase):
    """Test SetTrackerResult dataclass."""

    def test_basic_creation(self):
        result = SetTrackerResult(
            next_set_number=2,
            last_request_number=35,
        )
        self.assertEqual(result.next_set_number, 2)
        self.assertEqual(result.last_request_number, 35)
        self.assertEqual(result.previous_definitions, "")
        self.assertEqual(result.previous_instructions, "")
        self.assertEqual(result.resolution_method, "")
        self.assertIsNone(result.source_file)

    def test_full_creation(self):
        result = SetTrackerResult(
            next_set_number=3,
            last_request_number=70,
            previous_definitions="INCIDENT means...",
            previous_instructions="You are instructed...",
            resolution_method="filename_pattern",
            source_file="SI_Set_Two_to_Def.docx",
        )
        self.assertEqual(result.source_file, "SI_Set_Two_to_Def.docx")
        self.assertEqual(result.resolution_method, "filename_pattern")


class TestDiscoveryMode(unittest.TestCase):
    """Test DiscoveryMode enum."""

    def test_modes_exist(self):
        self.assertEqual(DiscoveryMode.INITIAL_STANDARD.value, "initial_standard")
        self.assertEqual(DiscoveryMode.INITIAL_CUSTOM.value, "initial_custom")
        self.assertEqual(DiscoveryMode.ADDITIONAL.value, "additional")


class TestCustomStyle(unittest.TestCase):
    """Test CustomStyle enum."""

    def test_styles_exist(self):
        self.assertEqual(CustomStyle.CUSTOM_ONLY.value, "custom_only")
        self.assertEqual(CustomStyle.STANDARD_PLUS_CUSTOM.value, "standard_plus_custom")
        self.assertEqual(CustomStyle.MODIFIED_STANDARD.value, "modified_standard")


if __name__ == '__main__':
    unittest.main()
