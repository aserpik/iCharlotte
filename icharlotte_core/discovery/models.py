"""
Data models for the discovery generation pipeline.

Defines enums for party roles, discovery types, and modes, as well as
dataclasses for parties, requests, sets, and tracker results.
"""
from dataclasses import dataclass, field
from enum import Enum
from typing import List, Optional


# ---------------------------------------------------------------------------
# Enums
# ---------------------------------------------------------------------------

class PartyRole(Enum):
    """Role of a party in litigation."""
    PLAINTIFF = "plaintiff"
    DEFENDANT = "defendant"
    CROSS_DEFENDANT = "cross_defendant"
    CROSS_COMPLAINANT = "cross_complainant"


class DiscoveryMode(Enum):
    """Mode of discovery generation."""
    INITIAL_STANDARD = "initial_standard"
    INITIAL_CUSTOM = "initial_custom"
    ADDITIONAL = "additional"


class CustomStyle(Enum):
    """Style for custom discovery generation."""
    CUSTOM_ONLY = "custom_only"
    STANDARD_PLUS_CUSTOM = "standard_plus_custom"
    MODIFIED_STANDARD = "modified_standard"


class DiscoveryType(Enum):
    """
    Type of written discovery.

    Each member carries metadata used for document generation:
    abbreviation, CCP section, regex pattern for parsing, header/title
    templates, and section heading text.
    """
    SI = "si"
    RPD = "rpd"
    RFA = "rfa"

    @property
    def abbreviation(self) -> str:
        return self.value.upper()

    @property
    def ccp_section(self) -> str:
        _sections = {
            "si": "CCP \u00a7 2030.010 et seq.",
            "rpd": "CCP \u00a7 2031.010 et seq.",
            "rfa": "CCP \u00a7 2033.010 et seq.",
        }
        return _sections[self.value]

    @property
    def request_header_pattern(self) -> str:
        """Regex pattern to match a request header in existing documents."""
        _patterns = {
            "si": r"SPECIAL\s+INTERROGATORY\s+NO\.\s*(\d+)",
            "rpd": r"REQUEST\s+FOR\s+PRODUCTION\s+NO\.\s*(\d+)",
            "rfa": r"REQUEST\s+FOR\s+ADMISSION\s+NO\.\s*(\d+)",
        }
        return _patterns[self.value]

    @property
    def request_header_template(self) -> str:
        """Template for a single request header.  Use .format(number=N)."""
        _templates = {
            "si": "SPECIAL INTERROGATORY NO. {number}",
            "rpd": "REQUEST FOR PRODUCTION NO. {number}",
            "rfa": "REQUEST FOR ADMISSION NO. {number}",
        }
        return _templates[self.value]

    @property
    def document_title_template(self) -> str:
        """
        Template for the document title page.
        Use .format(propounding=..., responding=..., set_word=...).
        """
        _titles = {
            "si": (
                "{propounding}'S SPECIAL INTERROGATORIES, "
                "SET {set_word}, TO {responding}"
            ),
            "rpd": (
                "{propounding}'S REQUESTS FOR PRODUCTION OF DOCUMENTS, "
                "SET {set_word}, TO {responding}"
            ),
            "rfa": (
                "{propounding}'S REQUESTS FOR ADMISSION, "
                "SET {set_word}, TO {responding}"
            ),
        }
        return _titles[self.value]

    @property
    def section_heading(self) -> str:
        _headings = {
            "si": "SPECIAL INTERROGATORIES",
            "rpd": "REQUESTS FOR PRODUCTION OF DOCUMENTS",
            "rfa": "REQUESTS FOR ADMISSION",
        }
        return _headings[self.value]


# ---------------------------------------------------------------------------
# Utility functions
# ---------------------------------------------------------------------------

_ONES = [
    "", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine",
    "Ten", "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen",
    "Seventeen", "Eighteen", "Nineteen",
]

_TENS = [
    "", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety",
]


def number_to_word(n: int) -> str:
    """Convert an integer 1-99 to its English word form (title-cased).

    Examples:
        1 -> "One", 13 -> "Thirteen", 21 -> "Twenty-One"

    Raises:
        ValueError: If *n* is outside the 1-99 range.
    """
    if not (1 <= n <= 99):
        raise ValueError(f"number_to_word supports 1-99, got {n}")
    if n < 20:
        return _ONES[n]
    tens, ones = divmod(n, 10)
    if ones == 0:
        return _TENS[tens]
    return f"{_TENS[tens]}-{_ONES[ones]}"


# Words to skip when picking a distinctive abbreviation from a party name.
_COMMON_WORDS = frozenset({
    "a", "an", "and", "at", "by", "co", "corp", "corporation", "county",
    "department", "dept", "district", "division", "doe", "does", "et", "al",
    "for", "from", "group", "in", "inc", "incorporated", "individual",
    "individually", "llc", "llp", "lp", "ltd", "of", "on", "or", "pa",
    "pc", "the", "to",
})

_ROLE_ABBREVIATIONS = {
    PartyRole.PLAINTIFF: "Pltf",
    PartyRole.DEFENDANT: "Def",
    PartyRole.CROSS_DEFENDANT: "XDef",
    PartyRole.CROSS_COMPLAINANT: "XCmplnt",
}


def generate_abbreviation(party: 'Party', all_parties: List['Party']) -> str:
    """Auto-generate a short abbreviation for *party*.

    Rules:
    - If *party* is the only one with its role among *all_parties*, return
      the standard role abbreviation (Pltf, Def, XDef, XCmplnt).
    - If multiple parties share the same role, pick a distinctive word from
      the party's name (skipping common words).
    - Special case: if the first word is "city", return "City".
    """
    same_role = [p for p in all_parties if p.role == party.role]
    if len(same_role) == 1:
        return _ROLE_ABBREVIATIONS[party.role]

    # Multiple parties with the same role -- pick a distinctive word.
    words = party.name.split()
    if not words:
        return _ROLE_ABBREVIATIONS[party.role]

    # Special case: name starts with "City"
    if words[0].lower() == "city":
        return "City"

    # Filter to meaningful words
    meaningful = [w for w in words if w.lower().strip(".,;") not in _COMMON_WORDS]
    if not meaningful:
        meaningful = words  # fallback to all words

    # Pick the first meaningful word, title-cased
    return meaningful[0].strip(".,;").title()


# ---------------------------------------------------------------------------
# Dataclasses
# ---------------------------------------------------------------------------

@dataclass
class Party:
    """A party in the litigation."""
    name: str
    role: PartyRole
    is_our_client: bool
    abbreviation: str = ""

    @property
    def role_label(self) -> str:
        """Human-readable role label."""
        _labels = {
            PartyRole.PLAINTIFF: "Plaintiff",
            PartyRole.DEFENDANT: "Defendant",
            PartyRole.CROSS_DEFENDANT: "Cross-Defendant",
            PartyRole.CROSS_COMPLAINANT: "Cross-Complainant",
        }
        return _labels[self.role]

    @property
    def formal_description(self) -> str:
        """Formal description for document headings (e.g. 'Defendant, SERVITEK ELECTRIC, INC.')."""
        return f"{self.role_label}, {self.name.upper()}"

    def to_dict(self) -> dict:
        return {
            'name': self.name,
            'role': self.role.value,
            'is_our_client': self.is_our_client,
            'abbreviation': self.abbreviation,
        }

    @classmethod
    def from_dict(cls, data: dict) -> 'Party':
        return cls(
            name=data["name"],
            role=PartyRole(data["role"]),
            is_our_client=data.get('is_our_client', False),
            abbreviation=data.get('abbreviation', ''),
        )


@dataclass
class DiscoveryRequest:
    """A single numbered discovery request."""
    number: int
    text: str
    definitions: List[str] = field(default_factory=list)


@dataclass
class DiscoverySet:
    """A complete set of discovery requests ready for assembly into a document."""
    discovery_type: DiscoveryType
    set_number: int
    directed_to: Party
    propounding_party: Party
    requests: List[DiscoveryRequest] = field(default_factory=list)
    definitions_block: str = ""
    instructions_block: str = ""
    previous_count: int = 0

    @property
    def needs_declaration(self) -> bool:
        """Whether a CCP declaration of necessity is required.

        SI and RFA require a declaration when total_count exceeds 35.
        RPD never requires a declaration.
        """
        if self.discovery_type == DiscoveryType.RPD:
            return False
        return self.total_count > 35

    @property
    def total_count(self) -> int:
        """Total number of requests including those from prior sets."""
        return self.previous_count + len(self.requests)

    @property
    def set_word(self) -> str:
        """Set number as an English word (e.g. 'One', 'Two')."""
        return number_to_word(self.set_number)

    @property
    def filename(self) -> str:
        """Suggested filename for the generated document."""
        abbr = self.discovery_type.abbreviation
        directed_abbr = self.directed_to.abbreviation or self.directed_to.name.split()[0]
        return f"{abbr}_Set_{self.set_word}_to_{directed_abbr}.docx"

    def plain_text(self) -> str:
        """Render requests as plain text (for preview or clipboard)."""
        lines = []
        dt = self.discovery_type
        for req in self.requests:
            lines.append(dt.request_header_template.format(number=req.number))
            lines.append("")
            lines.append(f"\t{req.text}")
            lines.append("")
        return "\n".join(lines)


@dataclass
class SetTrackerResult:
    """Result of scanning existing discovery files to determine next set/request numbers."""
    next_set_number: int
    last_request_number: int
    previous_definitions: str = ""
    previous_instructions: str = ""
    resolution_method: str = ""
    source_file: Optional[str] = None
