"""Tests for declaration generator."""
import unittest

from icharlotte_core.discovery.models import (
    DiscoveryType,
    Party,
    PartyRole,
    DiscoveryRequest,
    DiscoverySet,
)
from icharlotte_core.discovery.declaration import generate_declaration


def _make_set(discovery_type, num_requests, previous_count=0, set_number=1):
    """Helper to build a DiscoverySet for testing."""
    pltf = Party(
        name="John Smith",
        role=PartyRole.PLAINTIFF,
        is_our_client=True,
        abbreviation="Pltf",
    )
    deft = Party(
        name="Acme Corp",
        role=PartyRole.DEFENDANT,
        is_our_client=False,
        abbreviation="Def",
    )
    start = previous_count + 1
    requests = [
        DiscoveryRequest(number=start + i, text=f"Request {start + i}")
        for i in range(num_requests)
    ]
    return DiscoverySet(
        discovery_type=discovery_type,
        set_number=set_number,
        directed_to=deft,
        propounding_party=pltf,
        requests=requests,
        previous_count=previous_count,
    )


class TestSIDeclarationCites2030(unittest.TestCase):
    """test_si_declaration_cites_2030: verify '2030.030' and '2030.070' appear."""

    def test_si_declaration_cites_2030(self):
        ds = _make_set(DiscoveryType.SI, num_requests=40, previous_count=0)
        text = generate_declaration(ds, attorney_name="Jane Attorney")
        self.assertIn("2030.030", text)
        self.assertIn("2030.070", text)


class TestRFADeclarationCites2033(unittest.TestCase):
    """test_rfa_declaration_cites_2033: verify '2033.050' appears."""

    def test_rfa_declaration_cites_2033(self):
        ds = _make_set(DiscoveryType.RFA, num_requests=40, previous_count=0)
        text = generate_declaration(ds, attorney_name="Jane Attorney")
        self.assertIn("2033.030", text)
        self.assertIn("2033.050", text)


class TestDeclarationCountMathInitial(unittest.TestCase):
    """test_declaration_count_math_initial: prev=0, 63 requests."""

    def test_declaration_count_math_initial(self):
        ds = _make_set(DiscoveryType.SI, num_requests=63, previous_count=0)
        text = generate_declaration(ds, attorney_name="Jane Attorney")
        # Previous count is 0 -> "zero (0)"
        self.assertIn("zero (0)", text.lower())
        # Current set has 63 requests
        self.assertIn("63", text)
        # Total is 63
        self.assertIn("63", text)


class TestDeclarationCountMathAdditional(unittest.TestCase):
    """test_declaration_count_math_additional: prev=63 + 15 new -> total 78."""

    def test_declaration_count_math_additional(self):
        ds = _make_set(DiscoveryType.SI, num_requests=15, previous_count=63)
        text = generate_declaration(ds, attorney_name="Jane Attorney")
        self.assertIn("63", text)
        self.assertIn("15", text)
        self.assertIn("78", text)


class TestRPDReturnsEmpty(unittest.TestCase):
    """test_rpd_returns_empty: RPD with any count returns empty string."""

    def test_rpd_returns_empty(self):
        ds = _make_set(DiscoveryType.RPD, num_requests=100, previous_count=50)
        text = generate_declaration(ds, attorney_name="Jane Attorney")
        self.assertEqual(text, "")


class TestUnder35ReturnsEmpty(unittest.TestCase):
    """test_under_35_returns_empty: SI with 20 requests returns empty string."""

    def test_under_35_returns_empty(self):
        ds = _make_set(DiscoveryType.SI, num_requests=20, previous_count=0)
        text = generate_declaration(ds, attorney_name="Jane Attorney")
        self.assertEqual(text, "")


class TestDeclarationContent(unittest.TestCase):
    """Additional tests for declaration content correctness."""

    def test_title_contains_attorney_and_type(self):
        ds = _make_set(DiscoveryType.SI, num_requests=40, previous_count=0)
        text = generate_declaration(ds, attorney_name="Jane Attorney")
        upper = text.upper()
        self.assertIn("DECLARATION OF JANE ATTORNEY", upper)
        self.assertIn("SPECIAL INTERROGATOR", upper)

    def test_rfa_title_contains_type(self):
        ds = _make_set(DiscoveryType.RFA, num_requests=40, previous_count=0)
        text = generate_declaration(ds, attorney_name="Jane Attorney")
        upper = text.upper()
        self.assertIn("ADMISSION", upper)

    def test_includes_firm_name(self):
        ds = _make_set(DiscoveryType.SI, num_requests=40, previous_count=0)
        text = generate_declaration(ds, attorney_name="Jane Attorney", firm_name="Smith & Jones LLP")
        self.assertIn("Smith & Jones LLP", text)

    def test_default_firm_name(self):
        ds = _make_set(DiscoveryType.SI, num_requests=40, previous_count=0)
        text = generate_declaration(ds, attorney_name="Jane Attorney")
        self.assertIn("Bordin Semmer LLP", text)

    def test_perjury_closing(self):
        ds = _make_set(DiscoveryType.SI, num_requests=40, previous_count=0)
        text = generate_declaration(ds, attorney_name="Jane Attorney")
        self.assertIn("penalty of perjury", text.lower())

    def test_date_placeholder(self):
        ds = _make_set(DiscoveryType.SI, num_requests=40, previous_count=0)
        text = generate_declaration(ds, attorney_name="Jane Attorney")
        self.assertIn("{date}", text)

    def test_signature_line(self):
        ds = _make_set(DiscoveryType.SI, num_requests=40, previous_count=0)
        text = generate_declaration(ds, attorney_name="Jane Attorney")
        self.assertIn("Jane Attorney", text)

    def test_propounding_and_directed_parties_mentioned(self):
        ds = _make_set(DiscoveryType.SI, num_requests=40, previous_count=0)
        text = generate_declaration(ds, attorney_name="Jane Attorney")
        self.assertIn("John Smith", text)
        self.assertIn("Acme Corp", text)

    def test_set_number_mentioned(self):
        ds = _make_set(DiscoveryType.SI, num_requests=40, previous_count=0, set_number=3)
        text = generate_declaration(ds, attorney_name="Jane Attorney")
        self.assertIn("Three", text)

    def test_not_improper_purpose_statement(self):
        ds = _make_set(DiscoveryType.SI, num_requests=40, previous_count=0)
        text = generate_declaration(ds, attorney_name="Jane Attorney")
        self.assertIn("improper purpose", text.lower())

    def test_at_35_returns_empty(self):
        """Exactly 35 should NOT need a declaration."""
        ds = _make_set(DiscoveryType.SI, num_requests=35, previous_count=0)
        text = generate_declaration(ds, attorney_name="Jane Attorney")
        self.assertEqual(text, "")

    def test_at_36_returns_declaration(self):
        """36 should need a declaration."""
        ds = _make_set(DiscoveryType.SI, num_requests=36, previous_count=0)
        text = generate_declaration(ds, attorney_name="Jane Attorney")
        self.assertNotEqual(text, "")


if __name__ == "__main__":
    unittest.main()
