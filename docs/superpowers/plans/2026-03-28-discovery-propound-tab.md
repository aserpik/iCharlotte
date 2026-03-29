# Discovery Propound Tab Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a Discovery tab with a Propound sub-tab that generates SI, RPD, and RFA discovery requests using a hybrid template+LLM architecture.

**Architecture:** Deterministic templates for legal boilerplate (caption, instructions, definitions, declarations, signatures), LLM only for substantive request generation in Custom/Additional modes. Six backend modules (`models`, `templates`, `set_tracker`, `declaration`, `assembler`, `engine`) under `icharlotte_core/discovery/`, one UI module (`discovery_tab.py`), registered in `iCharlotte.py`.

**Tech Stack:** Python 3.x, PyQt6/PySide6, python-docx (with lxml OXML for formatting), PyMuPDF (fitz) for PDF text extraction, existing `LLMWorker`/`LLMHandler`/`ModelFetcher` from `icharlotte_core/llm.py`, `CaseDataManager` from `icharlotte_core/case_data_manager.py`.

**Spec:** `docs/superpowers/specs/2026-03-28-discovery-propound-tab-design.md`

---

## File Map

### New Files

| File | Responsibility |
|------|---------------|
| `icharlotte_core/discovery/__init__.py` | Package init — exports public API |
| `icharlotte_core/discovery/models.py` | Data classes: Party, PartyRole, DiscoveryMode, CustomStyle, DiscoveryType, DiscoveryRequest, DiscoverySet, SetTrackerResult |
| `icharlotte_core/discovery/templates.py` | Load template .docx files, extract request text, load DEFINED TERMS, perform variable substitution |
| `icharlotte_core/discovery/set_tracker.py` | Scan propounded folder for filenames, determine next set number + last request number, resolve previous definitions via cascading fallback |
| `icharlotte_core/discovery/declaration.py` | Generate SI (CCP §2030.070) and RFA (CCP §2033.050) declaration text with count math |
| `icharlotte_core/discovery/assembler.py` | Open caption page template, insert title, append all sections, render .docx with proper styles |
| `icharlotte_core/discovery/engine.py` | Orchestrate pipeline: branch by mode, delegate to templates/set_tracker/LLM, return DiscoverySet objects |
| `icharlotte_core/ui/discovery_tab.py` | DiscoveryTab QWidget with Propound/Respond sub-tabs, left pane controls, right pane editor |
| `tests/test_discovery/__init__.py` | Test package |
| `tests/test_discovery/test_models.py` | Tests for data models |
| `tests/test_discovery/test_templates.py` | Tests for template loading and variable substitution |
| `tests/test_discovery/test_set_tracker.py` | Tests for set number detection and previous set resolution |
| `tests/test_discovery/test_declaration.py` | Tests for declaration generation with count math |
| `tests/test_discovery/test_assembler.py` | Tests for .docx output structure |
| `tests/test_discovery/test_engine.py` | Integration tests for the generation pipeline |

### Modified Files

| File | Change |
|------|--------|
| `iCharlotte.py:~70` | Add import for `DiscoveryTab` |
| `iCharlotte.py:~848` | Add `self.discovery_tab` to tab widget |
| `iCharlotte.py:~1350` | Add `load_case()` call on case switch |

### Existing Files Referenced (read-only patterns)

| File | What to Reuse |
|------|--------------|
| `icharlotte_core/ui/tabs.py:122-208` | `ResizableListWidget` class (import directly) |
| `icharlotte_core/ui/tabs.py:846-889` | `add_file()` pattern for document box |
| `icharlotte_core/ui/tabs.py:1117-1176` | `read_files_content()` pattern |
| `icharlotte_core/ui/tabs.py:279-289` | Provider/model combo setup pattern |
| `icharlotte_core/llm.py:402-446` | `LLMWorker` class (import directly) |
| `icharlotte_core/llm.py:448-498` | `ModelFetcher` class (import directly) |
| `icharlotte_core/config.py` | `API_KEYS`, `GEMINI_DATA_DIR` |
| `icharlotte_core/case_data_manager.py` | `CaseDataManager.get_value()`, `save_variable()` |
| `icharlotte_core/utils.py:212-381` | `get_case_path()` for resolving case folders |
| `Scripts/report_generator/assemble.py:427-555` | OXML paragraph/run formatting patterns |
| `discovery/Standard Negligence Discovery (5800.070)/` | Template documents |
| `discovery/DISCOVERY DEFINED TERMS.docx` | Standard definitions |

---

## Task Dependency Graph

```
Task 1 (models)
    │
    ├──→ Task 2 (templates)
    │        │
    ├──→ Task 3 (set_tracker)
    │        │
    ├──→ Task 4 (declaration)
    │        │
    │        ▼
    ├──→ Task 5 (assembler) ← depends on models
    │        │
    │        ▼
    └──→ Task 6 (engine) ← depends on templates, set_tracker, declaration, assembler
             │
             ▼
         Task 7 (UI shell + left pane)
             │
             ▼
         Task 8 (party roster dropdown)
             │
             ▼
         Task 9 (right pane editor + save)
             │
             ▼
         Task 10 (wire generate to engine)
             │
             ▼
         Task 11 (register in iCharlotte.py)
             │
             ▼
         Task 12 (end-to-end test)
```

Tasks 2, 3, 4 can run in parallel (they only depend on Task 1).

---

### Task 1: Data Models

**Files:**
- Create: `icharlotte_core/discovery/__init__.py`
- Create: `icharlotte_core/discovery/models.py`
- Create: `tests/test_discovery/__init__.py`
- Create: `tests/test_discovery/test_models.py`

- [ ] **Step 1: Create package and write model tests**

```python
# tests/test_discovery/__init__.py
# (empty)

# tests/test_discovery/test_models.py
import unittest
from icharlotte_core.discovery.models import (
    PartyRole, Party, DiscoveryMode, CustomStyle, DiscoveryType,
    DiscoveryRequest, DiscoverySet, SetTrackerResult,
    number_to_word, generate_abbreviation,
)


class TestNumberToWord(unittest.TestCase):
    def test_basic_numbers(self):
        self.assertEqual(number_to_word(1), "One")
        self.assertEqual(number_to_word(2), "Two")
        self.assertEqual(number_to_word(10), "Ten")
        self.assertEqual(number_to_word(13), "Thirteen")
        self.assertEqual(number_to_word(21), "Twenty-One")

    def test_set_label(self):
        """Set numbers use 'ONE (1)' format in documents."""
        self.assertEqual(number_to_word(1).upper(), "ONE")
        self.assertEqual(number_to_word(3).upper(), "THREE")


class TestGenerateAbbreviation(unittest.TestCase):
    def test_single_plaintiff(self):
        p = Party(name="Ruxandra Raschkovsky", role=PartyRole.PLAINTIFF, is_our_client=False)
        self.assertEqual(generate_abbreviation(p, all_parties=[p]), "Pltf")

    def test_single_defendant(self):
        d = Party(name="Servitek Electric, Inc.", role=PartyRole.DEFENDANT, is_our_client=True)
        self.assertEqual(generate_abbreviation(d, all_parties=[d]), "Def")

    def test_multiple_defendants_uses_distinctive_word(self):
        d1 = Party(name="City of Los Angeles", role=PartyRole.DEFENDANT, is_our_client=False)
        d2 = Party(name="Servitek Electric, Inc.", role=PartyRole.DEFENDANT, is_our_client=True)
        all_p = [d1, d2]
        abbr1 = generate_abbreviation(d1, all_parties=all_p)
        abbr2 = generate_abbreviation(d2, all_parties=all_p)
        self.assertNotEqual(abbr1, abbr2)
        # Should pick distinctive word — "City" and "Servitek"
        self.assertIn("City", abbr1)
        self.assertIn("Servitek", abbr2)

    def test_cross_defendant(self):
        cd = Party(name="John Doe", role=PartyRole.CROSS_DEFENDANT, is_our_client=False)
        abbr = generate_abbreviation(cd, all_parties=[cd])
        self.assertEqual(abbr, "XDef")


class TestDiscoverySet(unittest.TestCase):
    def test_needs_declaration_si_over_35(self):
        ds = DiscoverySet(
            discovery_type=DiscoveryType.SI,
            set_number=1,
            directed_to=Party(name="Test", role=PartyRole.PLAINTIFF, is_our_client=False),
            propounding_party=Party(name="Our", role=PartyRole.DEFENDANT, is_our_client=True),
            requests=[DiscoveryRequest(number=i, text=f"Request {i}") for i in range(1, 40)],
            definitions_block="",
            instructions_block="",
            previous_count=0,
        )
        self.assertTrue(ds.needs_declaration)
        self.assertEqual(ds.total_count, 39)

    def test_no_declaration_under_35(self):
        ds = DiscoverySet(
            discovery_type=DiscoveryType.SI,
            set_number=1,
            directed_to=Party(name="Test", role=PartyRole.PLAINTIFF, is_our_client=False),
            propounding_party=Party(name="Our", role=PartyRole.DEFENDANT, is_our_client=True),
            requests=[DiscoveryRequest(number=i, text=f"Request {i}") for i in range(1, 20)],
            definitions_block="",
            instructions_block="",
            previous_count=0,
        )
        self.assertFalse(ds.needs_declaration)

    def test_declaration_for_rfa_over_35(self):
        ds = DiscoverySet(
            discovery_type=DiscoveryType.RFA,
            set_number=1,
            directed_to=Party(name="Test", role=PartyRole.PLAINTIFF, is_our_client=False),
            propounding_party=Party(name="Our", role=PartyRole.DEFENDANT, is_our_client=True),
            requests=[DiscoveryRequest(number=i, text=f"Request {i}") for i in range(1, 50)],
            definitions_block="",
            instructions_block="",
            previous_count=0,
        )
        self.assertTrue(ds.needs_declaration)

    def test_rpd_never_needs_declaration(self):
        ds = DiscoverySet(
            discovery_type=DiscoveryType.RPD,
            set_number=1,
            directed_to=Party(name="Test", role=PartyRole.PLAINTIFF, is_our_client=False),
            propounding_party=Party(name="Our", role=PartyRole.DEFENDANT, is_our_client=True),
            requests=[DiscoveryRequest(number=i, text=f"Request {i}") for i in range(1, 100)],
            definitions_block="",
            instructions_block="",
            previous_count=0,
        )
        self.assertFalse(ds.needs_declaration)


if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `python -m pytest tests/test_discovery/test_models.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'icharlotte_core.discovery'`

- [ ] **Step 3: Implement models.py**

```python
# icharlotte_core/discovery/__init__.py
"""Discovery generation package for propounding SI, RPD, and RFA."""

# icharlotte_core/discovery/models.py
"""Data models for the discovery generation pipeline."""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from enum import Enum
from typing import List, Optional


class PartyRole(Enum):
    PLAINTIFF = "Plaintiff"
    DEFENDANT = "Defendant"
    CROSS_DEFENDANT = "Cross-Defendant"
    CROSS_COMPLAINANT = "Cross-Complainant"


class DiscoveryMode(Enum):
    INITIAL_STANDARD = "initial_standard"
    INITIAL_CUSTOM = "initial_custom"
    ADDITIONAL = "additional"


class CustomStyle(Enum):
    CUSTOM_ONLY = "custom_only"
    STANDARD_PLUS_CUSTOM = "standard_plus"
    MODIFIED_STANDARD = "modified"


class DiscoveryType(Enum):
    SI = "Special Interrogatories"
    RPD = "Requests for Production"
    RFA = "Requests for Admission"

    @property
    def abbreviation(self) -> str:
        return self.name  # SI, RPD, RFA

    @property
    def ccp_section(self) -> str:
        return {
            DiscoveryType.SI: "2030.030",
            DiscoveryType.RPD: "2031.010",
            DiscoveryType.RFA: "2033.010",
        }[self]

    @property
    def request_header_pattern(self) -> str:
        """Regex pattern to match request headers in plain text."""
        return {
            DiscoveryType.SI: r"SPECIAL INTERROGATORY NO\.\s*(\d+):",
            DiscoveryType.RPD: r"REQUEST FOR PRODUCTION(?:\s+OF DOCUMENTS)?\s+NO\.\s*(\d+):",
            DiscoveryType.RFA: r"REQUEST FOR ADMISSION NO\.\s*(\d+):",
        }[self]

    @property
    def request_header_template(self) -> str:
        """Template for generating request headers."""
        return {
            DiscoveryType.SI: "SPECIAL INTERROGATORY NO. {num}:",
            DiscoveryType.RPD: "REQUEST FOR PRODUCTION NO. {num}:",
            DiscoveryType.RFA: "REQUEST FOR ADMISSION NO. {num}:",
        }[self]

    @property
    def document_title_template(self) -> str:
        """Template for the document title in the caption."""
        return {
            DiscoveryType.SI: "{propounding_party}'S SPECIAL INTERROGATORIES TO {responding_party}, SET {set_word}",
            DiscoveryType.RPD: "{propounding_party}'S REQUESTS FOR PRODUCTION OF DOCUMENTS, TO {responding_party}, SET {set_word}",
            DiscoveryType.RFA: "{propounding_party}'S REQUESTS FOR ADMISSION TO {responding_party}, SET {set_word}",
        }[self]

    @property
    def section_heading(self) -> str:
        """The heading used above the numbered requests."""
        return {
            DiscoveryType.SI: "SPECIAL INTERROGATORIES, SET {set_word}",
            DiscoveryType.RPD: "REQUEST FOR PRODUCTION OF DOCUMENTS, SET {set_word}",
            DiscoveryType.RFA: "REQUESTS FOR ADMISSION, SET {set_word}",
        }[self]


@dataclass
class Party:
    name: str
    role: PartyRole
    is_our_client: bool = False
    abbreviation: str = ""

    @property
    def role_label(self) -> str:
        """e.g. 'Defendant' or 'Cross-Defendant'."""
        return self.role.value

    @property
    def formal_description(self) -> str:
        """e.g. 'Defendant, SERVITEK ELECTRIC, INC.'"""
        return f"{self.role_label}, {self.name.upper()}"

    def to_dict(self) -> dict:
        return {
            "name": self.name,
            "role": self.role.value,
            "is_our_client": self.is_our_client,
            "abbreviation": self.abbreviation,
        }

    @classmethod
    def from_dict(cls, d: dict) -> Party:
        return cls(
            name=d["name"],
            role=PartyRole(d["role"]),
            is_our_client=d.get("is_our_client", False),
            abbreviation=d.get("abbreviation", ""),
        )


@dataclass
class DiscoveryRequest:
    number: int
    text: str
    definitions: List[str] = field(default_factory=list)


@dataclass
class DiscoverySet:
    discovery_type: DiscoveryType
    set_number: int
    directed_to: Party
    propounding_party: Party
    requests: List[DiscoveryRequest]
    definitions_block: str
    instructions_block: str
    previous_count: int = 0

    @property
    def needs_declaration(self) -> bool:
        if self.discovery_type == DiscoveryType.RPD:
            return False
        return self.total_count > 35

    @property
    def total_count(self) -> int:
        return self.previous_count + len(self.requests)

    @property
    def set_word(self) -> str:
        return number_to_word(self.set_number).upper()

    @property
    def filename(self) -> str:
        abbr = self.directed_to.abbreviation or "Party"
        return f"{self.discovery_type.abbreviation}({self.set_number}) t{abbr}.docx"

    def plain_text(self) -> str:
        """Render the discovery requests as plain text for the editor."""
        set_word = number_to_word(self.set_number).upper()
        heading = self.discovery_type.section_heading.format(set_word=set_word)
        lines = [heading, ""]
        for req in self.requests:
            header = self.discovery_type.request_header_template.format(num=req.number)
            lines.append(header)
            lines.append(req.text)
            for defn in req.definitions:
                lines.append(defn)
            lines.append("")
        return "\n".join(lines)


@dataclass
class SetTrackerResult:
    next_set_number: int
    last_request_number: int
    previous_definitions: str
    previous_instructions: str
    resolution_method: str  # "docx", "pdf", "standard_fallback"
    source_file: str = ""  # path to the file definitions were loaded from


# --- Utility functions ---

_ONES = [
    "", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine",
    "Ten", "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen",
    "Sixteen", "Seventeen", "Eighteen", "Nineteen",
]
_TENS = ["", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety"]


def number_to_word(n: int) -> str:
    """Convert an integer (1-99) to its English word form. E.g. 1 -> 'One', 21 -> 'Twenty-One'."""
    if n < 1 or n > 99:
        return str(n)
    if n < 20:
        return _ONES[n]
    tens, ones = divmod(n, 10)
    if ones == 0:
        return _TENS[tens]
    return f"{_TENS[tens]}-{_ONES[ones]}"


# Common words to skip when generating abbreviations from entity names
_SKIP_WORDS = {
    "of", "the", "a", "an", "and", "inc", "inc.", "llc", "llp", "corp",
    "corporation", "company", "co", "co.", "ltd", "ltd.", "et", "al",
    "al.", "does", "doe", "city", "county", "state",
}

_ENTITY_SUFFIXES = {"inc", "inc.", "llc", "llp", "corp", "corporation", "company", "co", "co.", "ltd", "ltd."}


def generate_abbreviation(party: Party, all_parties: List[Party]) -> str:
    """Auto-generate a short abbreviation for a party.

    Rules:
    - Single plaintiff: "Pltf"
    - Single defendant: "Def"
    - Cross-defendant: "XDef" (single) or distinctive word
    - Cross-complainant: "XCmplnt" (single) or distinctive word
    - Multiple of same role: use the most distinctive word from the name
    """
    same_role = [p for p in all_parties if p.role == party.role]

    if len(same_role) == 1:
        return {
            PartyRole.PLAINTIFF: "Pltf",
            PartyRole.DEFENDANT: "Def",
            PartyRole.CROSS_DEFENDANT: "XDef",
            PartyRole.CROSS_COMPLAINANT: "XCmplnt",
        }[party.role]

    # Multiple parties with same role — pick distinctive word
    words = party.name.replace(",", "").replace(".", "").split()
    # Filter out common/stop words and entity suffixes
    candidates = [w for w in words if w.lower() not in _SKIP_WORDS and len(w) > 1]

    if not candidates:
        candidates = [w for w in words if len(w) > 1]

    if candidates:
        # Prefer capitalized words that aren't generic
        # "City of Los Angeles" → "City" is in skip list, so try "Los" or "Angeles"
        # Special handling: if "city" is the first word, use it anyway (common legal shorthand)
        first_word = words[0] if words else ""
        if first_word.lower() == "city":
            return "City"
        return candidates[0]

    return party.name[:4]
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `python -m pytest tests/test_discovery/test_models.py -v`
Expected: All tests PASS

- [ ] **Step 5: Commit**

```bash
git add icharlotte_core/discovery/__init__.py icharlotte_core/discovery/models.py \
        tests/test_discovery/__init__.py tests/test_discovery/test_models.py
git commit -m "feat(discovery): add data models and utility functions"
```

---

### Task 2: Template Loading and Variable Substitution

**Files:**
- Create: `icharlotte_core/discovery/templates.py`
- Create: `tests/test_discovery/test_templates.py`

**Depends on:** Task 1 (models)

- [ ] **Step 1: Write template loading tests**

```python
# tests/test_discovery/test_templates.py
import os
import unittest
from icharlotte_core.discovery.models import DiscoveryType, Party, PartyRole
from icharlotte_core.discovery.templates import (
    TemplateLoader,
    substitute_variables,
    extract_requests_from_text,
)

DISCOVERY_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", "discovery")


class TestSubstituteVariables(unittest.TestCase):
    def test_replaces_blank_placeholders(self):
        text = "Defendant ____ hereby propounds"
        result = substitute_variables(text, {"defendant_name": "SERVITEK ELECTRIC, INC."})
        # The blanks next to "Defendant" should be replaced
        self.assertIn("SERVITEK ELECTRIC, INC.", result)
        self.assertNotIn("____", result)

    def test_replaces_plaintiff_references(self):
        text = '("PLAINTIFF" refers to Responding Party, ____, and includes agents)'
        result = substitute_variables(text, {"responding_party_name": "RUXANDRA RASCHKOVSKY"})
        self.assertIn("RUXANDRA RASCHKOVSKY", result)

    def test_preserves_non_placeholder_text(self):
        text = "Pursuant to CCP §2030.030, this is important"
        result = substitute_variables(text, {"defendant_name": "TEST"})
        self.assertEqual(text, result)


class TestExtractRequestsFromText(unittest.TestCase):
    def test_extract_si_requests(self):
        text = (
            "SPECIAL INTERROGATORY NO. 1:\n"
            "State all facts about the incident.\n"
            "SPECIAL INTERROGATORY NO. 2:\n"
            "Identify all persons with knowledge.\n"
        )
        requests = extract_requests_from_text(text, DiscoveryType.SI)
        self.assertEqual(len(requests), 2)
        self.assertEqual(requests[0].number, 1)
        self.assertIn("State all facts", requests[0].text)
        self.assertEqual(requests[1].number, 2)

    def test_extract_preserves_inline_definitions(self):
        text = (
            "SPECIAL INTERROGATORY NO. 1:\n"
            "If YOU contend that DEFENDANT's negligence caused the INCIDENT.\n"
            '(The term "YOU" refers to the responding party.)\n'
            '("INCIDENT" refers to facts described in the Complaint.)\n'
            "SPECIAL INTERROGATORY NO. 2:\n"
            "Identify all persons.\n"
        )
        requests = extract_requests_from_text(text, DiscoveryType.SI)
        self.assertEqual(len(requests), 2)
        self.assertEqual(len(requests[0].definitions), 2)
        self.assertIn("responding party", requests[0].definitions[0])

    def test_extract_rpd_requests(self):
        text = (
            "REQUEST FOR PRODUCTION NO. 1:\n"
            "Produce all documents relating to the incident.\n"
            "REQUEST FOR PRODUCTION NO. 2:\n"
            "Produce all medical records.\n"
        )
        requests = extract_requests_from_text(text, DiscoveryType.RPD)
        self.assertEqual(len(requests), 2)


class TestTemplateLoader(unittest.TestCase):
    """Tests that require the actual template files in discovery/ folder."""

    def setUp(self):
        self.loader = TemplateLoader(DISCOVERY_DIR)

    @unittest.skipUnless(
        os.path.exists(os.path.join(
            os.path.dirname(os.path.abspath(__file__)), "..", "..",
            "discovery", "Standard Negligence Discovery (5800.070)", "SI(1) tPltf.docx"
        )),
        "Template files not present",
    )
    def test_load_standard_negligence_si(self):
        requests = self.loader.load_standard_requests("negligence", DiscoveryType.SI)
        self.assertGreater(len(requests), 50)  # Sample has 63
        self.assertEqual(requests[0].number, 1)
        self.assertIn("negligence", requests[0].text.lower())

    @unittest.skipUnless(
        os.path.exists(os.path.join(
            os.path.dirname(os.path.abspath(__file__)), "..", "..",
            "discovery", "Standard Negligence Discovery (5800.070)", "RPD(1) tPltf.docx"
        )),
        "Template files not present",
    )
    def test_load_standard_negligence_rpd(self):
        requests = self.loader.load_standard_requests("negligence", DiscoveryType.RPD)
        self.assertGreater(len(requests), 20)  # Sample has 34

    @unittest.skipUnless(
        os.path.exists(os.path.join(
            os.path.dirname(os.path.abspath(__file__)), "..", "..",
            "discovery", "DISCOVERY DEFINED TERMS.docx"
        )),
        "Defined terms file not present",
    )
    def test_load_defined_terms(self):
        terms = self.loader.load_defined_terms()
        self.assertIn("DOCUMENT", terms)
        self.assertIn("PERSON", terms)
        # Should contain the blank placeholders before substitution
        self.assertIn("____", terms)


if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `python -m pytest tests/test_discovery/test_templates.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'icharlotte_core.discovery.templates'`

- [ ] **Step 3: Implement templates.py**

```python
# icharlotte_core/discovery/templates.py
"""Template loading and variable substitution for discovery generation."""
from __future__ import annotations

import os
import re
from typing import Dict, List, Optional

from docx import Document

from .models import DiscoveryRequest, DiscoveryType


# Map standard type names to folder paths (relative to discovery root)
STANDARD_TEMPLATES = {
    "negligence": "Standard Negligence Discovery (5800.070)",
}

# Map discovery types to template filenames within a standard folder
TEMPLATE_FILENAMES = {
    DiscoveryType.SI: "SI(1) tPltf.docx",
    DiscoveryType.RPD: "RPD(1) tPltf.docx",
}

DEFINED_TERMS_FILE = "DISCOVERY DEFINED TERMS.docx"


class TemplateLoader:
    """Loads discovery templates and definitions from the discovery/ folder."""

    def __init__(self, discovery_dir: str):
        self.discovery_dir = discovery_dir

    def load_standard_requests(
        self, standard_type: str, disc_type: DiscoveryType
    ) -> List[DiscoveryRequest]:
        """Load discovery requests from a standard template document.

        Args:
            standard_type: e.g. "negligence"
            disc_type: SI, RPD, or RFA

        Returns:
            List of DiscoveryRequest with number, text, and inline definitions.
        """
        folder_name = STANDARD_TEMPLATES.get(standard_type)
        if not folder_name:
            raise ValueError(f"Unknown standard type: {standard_type}")

        filename = TEMPLATE_FILENAMES.get(disc_type)
        if not filename:
            raise ValueError(f"No template file for {disc_type.value} in standard {standard_type}")

        path = os.path.join(self.discovery_dir, folder_name, filename)
        if not os.path.exists(path):
            raise FileNotFoundError(f"Template not found: {path}")

        text = self._extract_text(path)
        return extract_requests_from_text(text, disc_type)

    def load_defined_terms(self) -> str:
        """Load the standard discovery defined terms document as text."""
        path = os.path.join(self.discovery_dir, DEFINED_TERMS_FILE)
        if not os.path.exists(path):
            raise FileNotFoundError(f"Defined terms not found: {path}")
        return self._extract_text(path)

    def load_instructions(self, disc_type: DiscoveryType, template_path: str) -> str:
        """Extract the instructions section from a template document."""
        text = self._extract_text(template_path)
        # Instructions are between "INSTRUCTIONS TO ANSWERING PARTY" and the first request
        instr_match = re.search(
            r"INSTRUCTIONS TO ANSWERING PARTY\s*\n(.*?)(?="
            + disc_type.request_header_pattern
            + r"|"
            + re.escape(disc_type.value.upper())
            + r")",
            text,
            re.DOTALL | re.IGNORECASE,
        )
        if instr_match:
            return instr_match.group(1).strip()
        return ""

    def load_preamble(self, disc_type: DiscoveryType, template_path: str) -> str:
        """Extract the preamble paragraph (the 'TO ... AND ATTORNEYS OF RECORD' block)."""
        text = self._extract_text(template_path)
        # Preamble is between "TO " and "INSTRUCTIONS"
        preamble_match = re.search(
            r"(TO .+?ATTORNEYS OF RECORD.*?)(?=INSTRUCTIONS)",
            text,
            re.DOTALL | re.IGNORECASE,
        )
        if preamble_match:
            return preamble_match.group(1).strip()
        return ""

    @staticmethod
    def _extract_text(docx_path: str) -> str:
        """Extract all paragraph text from a .docx file."""
        doc = Document(docx_path)
        paragraphs = []
        for para in doc.paragraphs:
            paragraphs.append(para.text)
        return "\n".join(paragraphs)


def extract_requests_from_text(
    text: str, disc_type: DiscoveryType
) -> List[DiscoveryRequest]:
    """Parse plain text into a list of DiscoveryRequest objects.

    Splits on request headers (e.g., 'SPECIAL INTERROGATORY NO. 1:').
    Lines starting with '(' after a request body are treated as inline definitions.
    """
    pattern = disc_type.request_header_pattern
    # Split the text on request headers, keeping the delimiter
    parts = re.split(f"({pattern})", text)

    requests = []
    i = 1  # Skip text before first request header
    while i < len(parts) - 1:
        header = parts[i]
        body = parts[i + 1] if i + 1 < len(parts) else ""

        # Extract request number from header
        num_match = re.search(r"(\d+)", header)
        if not num_match:
            i += 2
            continue
        number = int(num_match.group(1))

        # Split body into request text and inline definitions
        lines = body.strip().split("\n")
        text_lines = []
        definitions = []
        in_definitions = False

        for line in lines:
            stripped = line.strip()
            if not stripped or stripped == "///":
                continue
            # Lines starting with '(' or '("' are definitions
            if stripped.startswith("(") and not in_definitions:
                in_definitions = True
                definitions.append(stripped)
            elif in_definitions and (stripped.startswith("(") or stripped.startswith('"')):
                definitions.append(stripped)
            elif in_definitions:
                # Continuation of previous definition
                definitions[-1] += " " + stripped
            else:
                text_lines.append(stripped)

        req_text = " ".join(text_lines).strip()
        requests.append(DiscoveryRequest(
            number=number,
            text=req_text,
            definitions=definitions,
        ))
        i += 2

    return requests


def substitute_variables(text: str, variables: Dict[str, str]) -> str:
    """Replace placeholders in template text with case-specific values.

    Handles two placeholder styles:
    1. ____ (four+ underscores) — replaced contextually based on surrounding text
    2. Named placeholders like {defendant_name} — direct replacement
    """
    result = text

    # Direct named placeholders
    for key, value in variables.items():
        result = result.replace(f"{{{key}}}", value)

    # Contextual blank replacement: look for patterns like "Defendant, ____"
    # or "Responding Party, ____"
    if "responding_party_name" in variables:
        result = re.sub(
            r"Responding Party,?\s*____+",
            f"Responding Party, {variables['responding_party_name']}",
            result,
        )
        # Also replace standalone ____ after party-related terms
        result = re.sub(
            r"(?<=Plaintiff\s)____+",
            variables["responding_party_name"],
            result,
        )

    if "propounding_party_name" in variables:
        result = re.sub(
            r"Propounding Party,?\s*____+",
            f"Propounding Party, {variables['propounding_party_name']}",
            result,
        )

    # Generic: replace remaining ____ with empty string (shouldn't normally happen)
    # Leave this as a visible marker so the user notices unresolved placeholders
    return result
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `python -m pytest tests/test_discovery/test_templates.py -v`
Expected: All tests PASS

- [ ] **Step 5: Commit**

```bash
git add icharlotte_core/discovery/templates.py tests/test_discovery/test_templates.py
git commit -m "feat(discovery): add template loading and variable substitution"
```

---

### Task 3: Set Tracker

**Files:**
- Create: `icharlotte_core/discovery/set_tracker.py`
- Create: `tests/test_discovery/test_set_tracker.py`

**Depends on:** Task 1 (models)

- [ ] **Step 1: Write set tracker tests**

```python
# tests/test_discovery/test_set_tracker.py
import os
import tempfile
import shutil
import unittest

from icharlotte_core.discovery.models import DiscoveryType, Party, PartyRole
from icharlotte_core.discovery.set_tracker import SetTracker


class TestSetTrackerScanFilenames(unittest.TestCase):
    """Test filename scanning with mock folder structures."""

    def setUp(self):
        self.tmpdir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.tmpdir)

    def _make_propounded_folder(self, files, party_subfolder=None):
        """Create a mock propounded folder structure with given filenames."""
        # Build: tmpdir/DISCOVERY/PROPOUNDED/fOUR Client/[party_subfolder/]
        path = os.path.join(self.tmpdir, "DISCOVERY", "PROPOUNDED", "fOUR Client")
        if party_subfolder:
            path = os.path.join(path, party_subfolder)
        os.makedirs(path, exist_ok=True)
        for fname in files:
            open(os.path.join(path, fname), "w").close()
        return self.tmpdir

    def test_single_set_si(self):
        case_path = self._make_propounded_folder(["SI (1) tPLF.pdf"])
        party = Party(name="Test Plaintiff", role=PartyRole.PLAINTIFF, abbreviation="PLF")
        tracker = SetTracker(case_path)
        next_num = tracker.get_next_set_number(DiscoveryType.SI, party)
        self.assertEqual(next_num, 2)

    def test_multiple_sets(self):
        case_path = self._make_propounded_folder([
            "SI (1) tPLF.pdf",
            "SI (2) tPLF.pdf",
            "RPD (1) tPLF.pdf",
        ])
        party = Party(name="Test", role=PartyRole.PLAINTIFF, abbreviation="PLF")
        tracker = SetTracker(case_path)
        self.assertEqual(tracker.get_next_set_number(DiscoveryType.SI, party), 3)
        self.assertEqual(tracker.get_next_set_number(DiscoveryType.RPD, party), 2)

    def test_party_subfolder(self):
        case_path = self._make_propounded_folder(
            ["SI (1) tPltf.pdf"], party_subfolder="tPltf"
        )
        party = Party(name="Test", role=PartyRole.PLAINTIFF, abbreviation="Pltf")
        tracker = SetTracker(case_path)
        next_num = tracker.get_next_set_number(DiscoveryType.SI, party)
        self.assertEqual(next_num, 2)

    def test_no_previous_sets(self):
        case_path = self._make_propounded_folder([])
        party = Party(name="Test", role=PartyRole.PLAINTIFF, abbreviation="PLF")
        tracker = SetTracker(case_path)
        next_num = tracker.get_next_set_number(DiscoveryType.SI, party)
        self.assertEqual(next_num, 1)

    def test_case_insensitive_folder(self):
        """The propounded folder search should be case-insensitive."""
        # Create with different casing
        path = os.path.join(self.tmpdir, "discovery", "Propounded", "four client")
        os.makedirs(path, exist_ok=True)
        open(os.path.join(path, "SI (1) tPLF.pdf"), "w").close()
        party = Party(name="Test", role=PartyRole.PLAINTIFF, abbreviation="PLF")
        tracker = SetTracker(self.tmpdir)
        next_num = tracker.get_next_set_number(DiscoveryType.SI, party)
        self.assertEqual(next_num, 2)


class TestSetTrackerParseFilename(unittest.TestCase):
    def test_parse_standard_filename(self):
        result = SetTracker.parse_discovery_filename("SI (1) tPLF.pdf")
        self.assertEqual(result["type"], "SI")
        self.assertEqual(result["set_number"], 1)
        self.assertEqual(result["party_abbrev"], "PLF")

    def test_parse_no_space(self):
        result = SetTracker.parse_discovery_filename("RPD(2) tCity.pdf")
        self.assertEqual(result["type"], "RPD")
        self.assertEqual(result["set_number"], 2)
        self.assertEqual(result["party_abbrev"], "City")

    def test_parse_rfa(self):
        result = SetTracker.parse_discovery_filename("RFA (3) tPltf.pdf")
        self.assertEqual(result["type"], "RFA")
        self.assertEqual(result["set_number"], 3)

    def test_parse_non_discovery_file(self):
        result = SetTracker.parse_discovery_filename("random_file.pdf")
        self.assertIsNone(result)


if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `python -m pytest tests/test_discovery/test_set_tracker.py -v`
Expected: FAIL — `ModuleNotFoundError`

- [ ] **Step 3: Implement set_tracker.py**

```python
# icharlotte_core/discovery/set_tracker.py
"""Scans propounded folder to determine next set numbers and resolve previous sets."""
from __future__ import annotations

import os
import re
from typing import Dict, List, Optional

from .models import DiscoveryType, Party, SetTrackerResult


class SetTracker:
    """Tracks propounded discovery sets for a case folder."""

    # Pattern: TYPE (NUM) tPARTY.ext  or  TYPE(NUM) tPARTY.ext
    _FILENAME_PATTERN = re.compile(
        r"^(SI|RPD|RFA)\s*\((\d+)\)\s*t(\w+)\.\w+$", re.IGNORECASE
    )

    def __init__(self, case_path: str):
        self.case_path = case_path
        self._propounded_path = self._find_propounded_folder()

    def _find_propounded_folder(self) -> Optional[str]:
        """Find the DISCOVERY/PROPOUNDED/fOUR Client folder (case-insensitive)."""
        # Walk the expected path components case-insensitively
        target_parts = ["discovery", "propounded", "four client"]
        current = self.case_path

        for target in target_parts:
            if not os.path.isdir(current):
                return None
            found = None
            for entry in os.listdir(current):
                entry_path = os.path.join(current, entry)
                if os.path.isdir(entry_path) and entry.lower().replace("f", "f") == target:
                    found = entry_path
                    break
                # More flexible matching for "four client" / "fOUR Client"
                if os.path.isdir(entry_path) and target == "four client":
                    if re.match(r"f\s*our\s*client", entry, re.IGNORECASE):
                        found = entry_path
                        break
            if found is None:
                return None
            current = found
        return current

    def get_next_set_number(self, disc_type: DiscoveryType, party: Party) -> int:
        """Determine the next set number for a given discovery type and party."""
        existing = self._scan_existing_sets(disc_type, party)
        if not existing:
            return 1
        return max(existing) + 1

    def get_last_request_number(
        self, disc_type: DiscoveryType, party: Party
    ) -> int:
        """Find the last request number in the most recent set for this type+party.

        Attempts to read the document content. Returns 0 if unable to determine.
        """
        latest_file = self._find_latest_set_file(disc_type, party)
        if not latest_file:
            return 0

        text = self._read_file_text(latest_file)
        if not text:
            return 0

        # Find all request numbers
        matches = re.findall(disc_type.request_header_pattern, text)
        if matches:
            return max(int(m) for m in matches)
        return 0

    def resolve_previous_set(
        self,
        disc_type: DiscoveryType,
        party: Party,
        discovery_dir: str,
    ) -> SetTrackerResult:
        """Full resolution: next set number, last request number, and previous definitions.

        Cascading fallback:
        1. Find .docx version of previous set
        2. Extract from PDF via PyMuPDF
        3. Fall back to standard definitions
        """
        next_set = self.get_next_set_number(disc_type, party)

        if next_set == 1:
            # No previous sets — use standard definitions
            from .templates import TemplateLoader
            loader = TemplateLoader(discovery_dir)
            return SetTrackerResult(
                next_set_number=1,
                last_request_number=0,
                previous_definitions=loader.load_defined_terms(),
                previous_instructions="",
                resolution_method="standard_fallback",
            )

        # Try to find and read the previous set file
        latest_file = self._find_latest_set_file(disc_type, party)
        if latest_file:
            text = self._read_file_text(latest_file)
            if text:
                definitions = self._extract_definitions(text)
                instructions = self._extract_instructions(text, disc_type)
                last_req = 0
                matches = re.findall(disc_type.request_header_pattern, text)
                if matches:
                    last_req = max(int(m) for m in matches)

                ext = os.path.splitext(latest_file)[1].lower()
                method = "docx" if ext == ".docx" else "pdf"

                return SetTrackerResult(
                    next_set_number=next_set,
                    last_request_number=last_req,
                    previous_definitions=definitions,
                    previous_instructions=instructions,
                    resolution_method=method,
                    source_file=latest_file,
                )

        # All fallbacks failed — use standard definitions
        from .templates import TemplateLoader
        loader = TemplateLoader(discovery_dir)
        return SetTrackerResult(
            next_set_number=next_set,
            last_request_number=0,
            previous_definitions=loader.load_defined_terms(),
            previous_instructions="",
            resolution_method="standard_fallback",
        )

    def _scan_existing_sets(
        self, disc_type: DiscoveryType, party: Party
    ) -> List[int]:
        """Return list of existing set numbers for this type+party."""
        if not self._propounded_path:
            return []

        set_numbers = []
        search_dirs = self._get_search_dirs(party)

        for search_dir in search_dirs:
            if not os.path.isdir(search_dir):
                continue
            for fname in os.listdir(search_dir):
                parsed = self.parse_discovery_filename(fname)
                if parsed and parsed["type"].upper() == disc_type.abbreviation:
                    # Match party abbreviation case-insensitively
                    if party.abbreviation and parsed["party_abbrev"].lower() == party.abbreviation.lower():
                        set_numbers.append(parsed["set_number"])
                    elif not party.abbreviation:
                        set_numbers.append(parsed["set_number"])

        return set_numbers

    def _get_search_dirs(self, party: Party) -> List[str]:
        """Get directories to search for a party's discovery files."""
        if not self._propounded_path:
            return []

        dirs = [self._propounded_path]  # Direct files in fOUR Client

        # Also check party-specific subfolder (t + abbreviation)
        if party.abbreviation:
            for entry in os.listdir(self._propounded_path):
                entry_path = os.path.join(self._propounded_path, entry)
                if os.path.isdir(entry_path):
                    # Match tPltf, tCity, etc. case-insensitively
                    if entry.lower() == f"t{party.abbreviation.lower()}":
                        dirs.insert(0, entry_path)  # Prefer subfolder

        return dirs

    def _find_latest_set_file(
        self, disc_type: DiscoveryType, party: Party
    ) -> Optional[str]:
        """Find the file for the latest set, preferring .docx over .pdf.

        Search order:
        1. NOTES/AI OUTPUT/DISCOVERY REQUESTS/ for .docx
        2. Propounded folder for .docx
        3. Propounded folder for .pdf
        """
        existing_sets = self._scan_existing_sets(disc_type, party)
        if not existing_sets:
            return None
        latest_num = max(existing_sets)

        # Search 1: AI OUTPUT drafts folder
        drafts_dir = os.path.join(self.case_path, "NOTES", "AI OUTPUT", "DISCOVERY REQUESTS")
        found = self._find_set_file_in_dir(drafts_dir, disc_type, party, latest_num, ".docx")
        if found:
            return found

        # Search 2: Propounded folder for .docx
        for search_dir in self._get_search_dirs(party):
            found = self._find_set_file_in_dir(search_dir, disc_type, party, latest_num, ".docx")
            if found:
                return found

        # Search 3: Propounded folder for .pdf
        for search_dir in self._get_search_dirs(party):
            found = self._find_set_file_in_dir(search_dir, disc_type, party, latest_num, ".pdf")
            if found:
                return found

        return None

    def _find_set_file_in_dir(
        self, directory: str, disc_type: DiscoveryType, party: Party,
        set_number: int, extension: str,
    ) -> Optional[str]:
        """Find a specific set file in a directory."""
        if not os.path.isdir(directory):
            return None
        for fname in os.listdir(directory):
            if not fname.lower().endswith(extension):
                continue
            parsed = self.parse_discovery_filename(fname)
            if not parsed:
                continue
            if (
                parsed["type"].upper() == disc_type.abbreviation
                and parsed["set_number"] == set_number
                and parsed["party_abbrev"].lower() == party.abbreviation.lower()
            ):
                return os.path.join(directory, fname)
        return None

    @staticmethod
    def _read_file_text(filepath: str) -> str:
        """Read text content from a .docx or .pdf file."""
        ext = os.path.splitext(filepath)[1].lower()
        if ext == ".docx":
            from docx import Document
            doc = Document(filepath)
            return "\n".join(p.text for p in doc.paragraphs)
        elif ext == ".pdf":
            try:
                import fitz
                doc = fitz.open(filepath)
                text = ""
                for page in doc:
                    text += page.get_text() + "\n"
                doc.close()
                return text
            except Exception:
                return ""
        return ""

    @staticmethod
    def _extract_definitions(text: str) -> str:
        """Extract the definitions section from document text."""
        # Definitions are typically parenthetical blocks near the beginning
        # Look for the first definition and collect all subsequent ones
        definitions = []
        in_defs = False
        for line in text.split("\n"):
            stripped = line.strip()
            if re.match(r'^\(.*?".*?".*?\)', stripped) or re.match(r'^"[A-Z]+.*?".*means', stripped):
                in_defs = True
                definitions.append(stripped)
            elif in_defs and stripped.startswith("("):
                definitions.append(stripped)
            elif in_defs and stripped == "":
                continue
            elif in_defs and not stripped.startswith("("):
                # Check if this is a continuation
                if definitions and not stripped[0:1].isupper():
                    definitions[-1] += " " + stripped
                else:
                    break
        return "\n\n".join(definitions)

    @staticmethod
    def _extract_instructions(text: str, disc_type: DiscoveryType) -> str:
        """Extract instructions section from document text."""
        match = re.search(
            r"INSTRUCTIONS TO ANSWERING PARTY\s*\n(.*?)(?="
            + disc_type.request_header_pattern
            + r"|"
            + re.escape(disc_type.section_heading.split(",")[0])
            + r")",
            text,
            re.DOTALL | re.IGNORECASE,
        )
        if match:
            return match.group(1).strip()
        return ""

    @staticmethod
    def parse_discovery_filename(filename: str) -> Optional[Dict]:
        """Parse a discovery filename into components.

        Returns dict with keys: type, set_number, party_abbrev
        or None if the filename doesn't match the pattern.
        """
        match = SetTracker._FILENAME_PATTERN.match(filename)
        if not match:
            return None
        return {
            "type": match.group(1).upper(),
            "set_number": int(match.group(2)),
            "party_abbrev": match.group(3),
        }
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `python -m pytest tests/test_discovery/test_set_tracker.py -v`
Expected: All tests PASS

- [ ] **Step 5: Commit**

```bash
git add icharlotte_core/discovery/set_tracker.py tests/test_discovery/test_set_tracker.py
git commit -m "feat(discovery): add set tracker for propounded folder scanning"
```

---

### Task 4: Declaration Generator

**Files:**
- Create: `icharlotte_core/discovery/declaration.py`
- Create: `tests/test_discovery/test_declaration.py`

**Depends on:** Task 1 (models)

- [ ] **Step 1: Write declaration tests**

```python
# tests/test_discovery/test_declaration.py
import unittest
from icharlotte_core.discovery.models import (
    DiscoveryType, Party, PartyRole, DiscoverySet, DiscoveryRequest,
)
from icharlotte_core.discovery.declaration import generate_declaration


class TestGenerateDeclaration(unittest.TestCase):
    def _make_set(self, disc_type, num_requests, previous_count=0, set_number=1):
        return DiscoverySet(
            discovery_type=disc_type,
            set_number=set_number,
            directed_to=Party(name="Ruxandra Raschkovsky", role=PartyRole.PLAINTIFF),
            propounding_party=Party(name="Servitek Electric, Inc.", role=PartyRole.DEFENDANT, is_our_client=True),
            requests=[DiscoveryRequest(number=i, text=f"Q{i}") for i in range(1, num_requests + 1)],
            definitions_block="",
            instructions_block="",
            previous_count=previous_count,
        )

    def test_si_declaration_cites_2030(self):
        ds = self._make_set(DiscoveryType.SI, 40)
        decl = generate_declaration(ds, attorney_name="Andrei Serpik")
        self.assertIn("2030.030", decl)
        self.assertIn("2030.070", decl)
        self.assertIn("Andrei Serpik", decl)

    def test_rfa_declaration_cites_2033(self):
        ds = self._make_set(DiscoveryType.RFA, 40)
        decl = generate_declaration(ds, attorney_name="Andrei Serpik")
        self.assertIn("2033.050", decl)
        self.assertIn("Andrei Serpik", decl)

    def test_declaration_count_math_initial(self):
        ds = self._make_set(DiscoveryType.SI, 63, previous_count=0)
        decl = generate_declaration(ds, attorney_name="Test")
        self.assertIn("zero (0)", decl.lower())
        self.assertIn("sixty-three (63)", decl.lower())

    def test_declaration_count_math_additional(self):
        ds = self._make_set(DiscoveryType.SI, 15, previous_count=63, set_number=2)
        decl = generate_declaration(ds, attorney_name="Test")
        self.assertIn("63", decl)
        self.assertIn("15", decl)
        self.assertIn("78", decl)  # Total: 63 + 15

    def test_rpd_returns_empty(self):
        ds = self._make_set(DiscoveryType.RPD, 100)
        decl = generate_declaration(ds, attorney_name="Test")
        self.assertEqual(decl, "")

    def test_under_35_returns_empty(self):
        ds = self._make_set(DiscoveryType.SI, 20)
        decl = generate_declaration(ds, attorney_name="Test")
        self.assertEqual(decl, "")


if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `python -m pytest tests/test_discovery/test_declaration.py -v`
Expected: FAIL — `ModuleNotFoundError`

- [ ] **Step 3: Implement declaration.py**

```python
# icharlotte_core/discovery/declaration.py
"""Generate SI and RFA declarations when request count exceeds 35."""
from __future__ import annotations

from .models import DiscoverySet, DiscoveryType, number_to_word


def generate_declaration(
    discovery_set: DiscoverySet,
    attorney_name: str,
    firm_name: str = "Bordin Semmer LLP",
) -> str:
    """Generate declaration text for SI (CCP §2030.070) or RFA (CCP §2033.050).

    Returns empty string if no declaration is needed (RPD, or count <= 35).
    """
    if not discovery_set.needs_declaration:
        return ""

    disc_type = discovery_set.discovery_type
    total = discovery_set.total_count
    this_set_count = len(discovery_set.requests)
    prev_count = discovery_set.previous_count
    responding = discovery_set.directed_to
    propounding = discovery_set.propounding_party

    # CCP section references by type
    if disc_type == DiscoveryType.SI:
        limit_section = "2030.030"
        limit_subdivision = "2030.030, subdivision (a)(1)"
        declaration_section = "2030.070"
        request_type_name = "specially prepared interrogatories"
        request_type_singular = "interrogatory"
        set_label = "Special Interrogatories"
    elif disc_type == DiscoveryType.RFA:
        limit_section = "2033.030"
        limit_subdivision = "2033.030"
        declaration_section = "2033.050"
        request_type_name = "requests for admission"
        request_type_singular = "request for admission"
        set_label = "Requests for Admission"
    else:
        return ""

    prev_word = number_to_word(prev_count).lower() if prev_count > 0 else "zero"
    this_word = number_to_word(this_set_count).lower()
    total_word = number_to_word(total).lower()
    set_word = number_to_word(discovery_set.set_number)

    lines = [
        f"DECLARATION OF {attorney_name.upper()} IN SUPPORT OF ADDITIONAL "
        f"{set_label.upper()} PROPOUNDED ON {responding.role_label.upper()} "
        f"{responding.name.upper()}",
        "",
        f"I, {attorney_name}, declare:",
        "",
        f"I am an attorney with the law firm of {firm_name}, the attorneys of record for "
        f"{propounding.role_label} {propounding.name.upper()} in this action.",
        "",
        f"I am propounding on behalf of {propounding.role_label} {propounding.name.upper()} "
        f"the attached set of {set_label}, Set {set_word}.",
        "",
        f"This set of {set_label} will cause the total number of {request_type_name} "
        f"propounded to the party to whom they are directed to exceed the number of "
        f"{request_type_name} permitted by Code of Civil Procedure section "
        f"{limit_subdivision}.",
        "",
        f"{propounding.role_label} {propounding.name.upper()} previously served "
        f"{prev_word} ({prev_count}) {request_type_name} upon "
        f"{responding.role_label} {responding.name.upper()}.",
        "",
        f"This set of {request_type_name.lower()} contains a total of {this_word} "
        f"({this_set_count}) {request_type_name}, propounded pursuant to Code of Civil "
        f"Procedure section {limit_section}, bringing the total number of "
        f"{request_type_name} to {total_word} ({total}).",
        "",
        "I am familiar with the issues and the previous discovery conducted by all of "
        "the parties in this case.",
        "",
        f"I have personally examined every single {request_type_singular} contained "
        f"in this set of {request_type_name}.",
        "",
        f"This number of questions is warranted under Code of Civil Procedure section "
        f"{declaration_section}, subdivision (a), because the quantity and complexity "
        f"of the issues involved in this case and because of the expedience of using "
        f"this method of discovery.",
        "",
        "///",
        "///",
        "///",
        "///",
        "///",
        "",
        f"The {request_type_name} being propounded in this set of {set_label} are not "
        f"being used for an improper purpose, such as to harass the party or the attorney "
        f"for the party to whom it is directed or to cause unnecessary delay or needless "
        f"increase in the cost of litigation",
        "",
        "I declare under penalty of perjury under the laws of the State of California "
        "that the foregoing is true and correct.",
        "",
        f"Executed this {{date}}, at Los Angeles, California.",
        "",
        "",
        f"\t\t\t\t\t\t_______________________________",
        f"\t\t\t\t\t\t\t\t\t\t{attorney_name}",
    ]

    return "\n".join(lines)
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `python -m pytest tests/test_discovery/test_declaration.py -v`
Expected: All tests PASS

- [ ] **Step 5: Commit**

```bash
git add icharlotte_core/discovery/declaration.py tests/test_discovery/test_declaration.py
git commit -m "feat(discovery): add declaration generator for SI and RFA"
```

---

### Task 5: Document Assembler

**Files:**
- Create: `icharlotte_core/discovery/assembler.py`
- Create: `tests/test_discovery/test_assembler.py`

**Depends on:** Task 1 (models), Task 4 (declaration)

- [ ] **Step 1: Write assembler tests**

```python
# tests/test_discovery/test_assembler.py
import os
import tempfile
import shutil
import unittest

from docx import Document

from icharlotte_core.discovery.models import (
    DiscoveryType, DiscoveryRequest, DiscoverySet, Party, PartyRole,
    number_to_word,
)
from icharlotte_core.discovery.assembler import DiscoveryAssembler


DISCOVERY_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", "discovery")
CAPTION_PAGE = os.path.join(DISCOVERY_DIR, "Caption Page (AS FM).docx")


class TestDiscoveryAssembler(unittest.TestCase):
    def setUp(self):
        self.tmpdir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.tmpdir)

    def _make_discovery_set(self, disc_type=DiscoveryType.SI, num_requests=5):
        return DiscoverySet(
            discovery_type=disc_type,
            set_number=1,
            directed_to=Party(name="Test Plaintiff", role=PartyRole.PLAINTIFF, abbreviation="Pltf"),
            propounding_party=Party(name="Test Defendant, Inc.", role=PartyRole.DEFENDANT, is_our_client=True, abbreviation="Def"),
            requests=[
                DiscoveryRequest(number=i, text=f"State all facts about item {i}.")
                for i in range(1, num_requests + 1)
            ],
            definitions_block='("YOU" refers to the responding party.)',
            instructions_block="Answer fully and completely under oath within 30 days.",
            previous_count=0,
        )

    @unittest.skipUnless(os.path.exists(CAPTION_PAGE), "Caption page not available")
    def test_assemble_creates_docx(self):
        ds = self._make_discovery_set()
        assembler = DiscoveryAssembler(CAPTION_PAGE)
        output_path = os.path.join(self.tmpdir, ds.filename)
        assembler.assemble(ds, output_path)

        self.assertTrue(os.path.exists(output_path))
        doc = Document(output_path)
        full_text = "\n".join(p.text for p in doc.paragraphs)
        self.assertIn("SPECIAL INTERROGATORY NO. 1:", full_text)
        self.assertIn("State all facts about item 1.", full_text)

    @unittest.skipUnless(os.path.exists(CAPTION_PAGE), "Caption page not available")
    def test_assemble_includes_definitions(self):
        ds = self._make_discovery_set()
        assembler = DiscoveryAssembler(CAPTION_PAGE)
        output_path = os.path.join(self.tmpdir, ds.filename)
        assembler.assemble(ds, output_path)

        doc = Document(output_path)
        full_text = "\n".join(p.text for p in doc.paragraphs)
        self.assertIn("YOU", full_text)
        self.assertIn("responding party", full_text)

    @unittest.skipUnless(os.path.exists(CAPTION_PAGE), "Caption page not available")
    def test_assemble_rpd(self):
        ds = self._make_discovery_set(disc_type=DiscoveryType.RPD, num_requests=3)
        assembler = DiscoveryAssembler(CAPTION_PAGE)
        output_path = os.path.join(self.tmpdir, ds.filename)
        assembler.assemble(ds, output_path)

        doc = Document(output_path)
        full_text = "\n".join(p.text for p in doc.paragraphs)
        self.assertIn("REQUEST FOR PRODUCTION NO. 1:", full_text)

    @unittest.skipUnless(os.path.exists(CAPTION_PAGE), "Caption page not available")
    def test_assemble_with_declaration(self):
        ds = self._make_discovery_set(num_requests=40)
        assembler = DiscoveryAssembler(CAPTION_PAGE)
        output_path = os.path.join(self.tmpdir, ds.filename)
        assembler.assemble(ds, output_path, attorney_name="Andrei Serpik")

        doc = Document(output_path)
        full_text = "\n".join(p.text for p in doc.paragraphs)
        self.assertIn("DECLARATION", full_text)
        self.assertIn("2030.070", full_text)


if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `python -m pytest tests/test_discovery/test_assembler.py -v`
Expected: FAIL — `ModuleNotFoundError`

- [ ] **Step 3: Implement assembler.py**

```python
# icharlotte_core/discovery/assembler.py
"""Renders DiscoverySet objects into formatted .docx files using a caption page template."""
from __future__ import annotations

import os
import copy
from typing import Optional

from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from lxml import etree

from .models import DiscoverySet, DiscoveryType, number_to_word
from .declaration import generate_declaration


def _qn(tag: str) -> str:
    """Resolve a qualified name for Word XML namespace."""
    nsmap = {
        "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    }
    prefix, name = tag.split(":")
    return f"{{{nsmap[prefix]}}}{name}"


class DiscoveryAssembler:
    """Assembles a .docx discovery document from a caption page template."""

    def __init__(self, caption_page_path: str):
        if not os.path.exists(caption_page_path):
            raise FileNotFoundError(f"Caption page not found: {caption_page_path}")
        self.caption_page_path = caption_page_path

    def assemble(
        self,
        discovery_set: DiscoverySet,
        output_path: str,
        attorney_name: str = "",
        firm_name: str = "Bordin Semmer LLP",
        date_str: str = "",
    ) -> str:
        """Build a complete .docx file from a DiscoverySet.

        Args:
            discovery_set: The assembled discovery data.
            output_path: Where to save the .docx file.
            attorney_name: For signature block and declaration.
            firm_name: For signature block.
            date_str: Formatted date string for signature block.

        Returns:
            The output_path on success.
        """
        os.makedirs(os.path.dirname(output_path), exist_ok=True)

        doc = Document(self.caption_page_path)
        ds = discovery_set
        set_word = number_to_word(ds.set_number).upper()

        # --- Insert document title into caption ---
        title = ds.discovery_type.document_title_template.format(
            propounding_party=f"DEFENDANT {ds.propounding_party.name.upper()}",
            responding_party=f"{ds.directed_to.role_label.upper()} {ds.directed_to.name.upper()}",
            set_word=set_word,
        )
        self._insert_title(doc, title)

        # --- Add propounding/responding party block ---
        self._add_paragraph(doc, "")
        self._add_paragraph(
            doc,
            f"PROPOUNDING PARTY:\t{ds.propounding_party.formal_description}",
        )
        self._add_paragraph(
            doc,
            f"RESPONDING PARTY:\t{ds.directed_to.formal_description}",
        )
        self._add_paragraph(
            doc,
            f"SET NO.:\t{set_word} ({ds.set_number})",
        )
        self._add_paragraph(doc, "")

        # --- Preamble ---
        preamble = self._build_preamble(ds)
        self._add_paragraph(doc, preamble)
        self._add_paragraph(doc, "")

        # --- Instructions ---
        if ds.instructions_block:
            self._add_paragraph(doc, "INSTRUCTIONS TO ANSWERING PARTY", bold=True, underline=True)
            for para_text in ds.instructions_block.split("\n"):
                if para_text.strip():
                    self._add_paragraph(doc, para_text.strip())
            self._add_paragraph(doc, "")

        # --- Section heading ---
        heading = ds.discovery_type.section_heading.format(set_word=set_word)
        self._add_paragraph(doc, heading, bold=True, underline=True, center=True)
        self._add_paragraph(doc, "")

        # --- Discovery requests ---
        for req in ds.requests:
            header = ds.discovery_type.request_header_template.format(num=req.number)
            self._add_paragraph(doc, header, bold=True, underline=True)
            self._add_paragraph(doc, req.text)
            for defn in req.definitions:
                self._add_paragraph(doc, defn)
            # Add spacing between requests
            self._add_paragraph(doc, "")

        # --- Signature block ---
        if date_str:
            self._add_paragraph(doc, f"Dated:  {date_str}\t      {firm_name.upper()}")
        else:
            self._add_paragraph(doc, f"Dated:  {{date}}\t      {firm_name.upper()}")
        self._add_paragraph(doc, "")
        self._add_paragraph(doc, "")
        self._add_paragraph(doc, f"By:\t\t")
        if attorney_name:
            self._add_paragraph(doc, f"\t\t{attorney_name}")
        self._add_paragraph(
            doc,
            f"Attorneys for {ds.propounding_party.role_label},",
        )
        self._add_paragraph(doc, ds.propounding_party.name.upper())

        # --- Declaration (if needed) ---
        if ds.needs_declaration and attorney_name:
            self._add_page_break(doc)
            decl_text = generate_declaration(ds, attorney_name, firm_name)
            if date_str:
                decl_text = decl_text.replace("{date}", date_str)
            for line in decl_text.split("\n"):
                if line.strip():
                    # Declaration heading gets bold+underline
                    if line.startswith("DECLARATION OF"):
                        self._add_paragraph(doc, line, bold=True, underline=True, center=True)
                    else:
                        self._add_paragraph(doc, line)
                else:
                    self._add_paragraph(doc, "")

        doc.save(output_path)
        return output_path

    def assemble_from_plain_text(
        self,
        plain_text: str,
        discovery_set: DiscoverySet,
        output_path: str,
        attorney_name: str = "",
        firm_name: str = "Bordin Semmer LLP",
        date_str: str = "",
    ) -> str:
        """Build a .docx from user-edited plain text.

        Re-parses the plain text into DiscoveryRequest objects, then delegates to assemble().
        """
        from .templates import extract_requests_from_text
        requests = extract_requests_from_text(plain_text, discovery_set.discovery_type)
        # Update the discovery set with edited requests
        discovery_set.requests = requests
        return self.assemble(discovery_set, output_path, attorney_name, firm_name, date_str)

    def _insert_title(self, doc: Document, title: str):
        """Insert the document title. Looks for the caption table and adds after it."""
        # The caption page has tables — add title after the last table
        if doc.tables:
            # Add a paragraph after the document body content
            self._add_paragraph(doc, "")
            self._add_paragraph(doc, title, bold=True, center=True)
            self._add_paragraph(doc, "")
        else:
            self._add_paragraph(doc, title, bold=True, center=True)

    def _add_paragraph(
        self,
        doc: Document,
        text: str,
        bold: bool = False,
        underline: bool = False,
        center: bool = False,
    ):
        """Add a paragraph with specified formatting."""
        para = doc.add_paragraph()
        if center:
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER

        run = para.add_run(text)
        run.bold = bold
        run.underline = underline

        # Use Times New Roman 12pt to match discovery document convention
        run.font.name = "Times New Roman"
        run.font.size = Pt(12)

    def _add_page_break(self, doc: Document):
        """Add a page break."""
        from docx.oxml.ns import qn as docx_qn
        from lxml import etree as ET
        para = doc.add_paragraph()
        run = para.add_run()
        br = ET.SubElement(run._element, docx_qn("w:br"))
        br.set(docx_qn("w:type"), "page")

    @staticmethod
    def _build_preamble(ds: DiscoverySet) -> str:
        """Build the statutory preamble paragraph."""
        templates = {
            DiscoveryType.SI: (
                f"TO {ds.directed_to.role_label.upper()} {ds.directed_to.name.upper()} AND "
                f"{'HER' if True else 'HIS'} ATTORNEYS OF RECORD:\n"
                f"\tPursuant to California Code of Civil Procedure §2030.030, "
                f"{ds.propounding_party.role_label}, {ds.propounding_party.name.upper()} "
                f'("Propounding Party" or "{ds.propounding_party.role_label}"), hereby '
                f"propounds to {ds.directed_to.role_label}, {ds.directed_to.name.upper()} "
                f'("{ds.directed_to.role_label}" or "Responding Party"), the following '
                f"{number_to_word(ds.set_number)} Set of Special Interrogatories, "
                f"each of which shall be answered fully, separately, in writing, under oath, "
                f"and within thirty (30) days as required by law."
            ),
            DiscoveryType.RPD: (
                f"TO {ds.directed_to.role_label.upper()} {ds.directed_to.name.upper()} AND "
                f"{'HER' if True else 'HIS'} ATTORNEYS OF RECORD:\n"
                f"\tDemand is hereby made by {ds.propounding_party.role_label}, "
                f"{ds.propounding_party.name.upper()} "
                f'("Propounding Party" or "{ds.propounding_party.role_label}"), pursuant to '
                f"Code of Civil Procedure section 2031.010, et seq., that "
                f"{ds.directed_to.role_label}, {ds.directed_to.name.upper()} "
                f'("{ds.directed_to.role_label}" or "Responding Party"), produce and permit '
                f"inspection, photographing, and photocopying of the documents and/or "
                f"inspection, photographing, testing, and sampling of other tangible things "
                f"described herein."
            ),
            DiscoveryType.RFA: (
                f"TO {ds.directed_to.role_label.upper()} {ds.directed_to.name.upper()} AND "
                f"{'HER' if True else 'HIS'} ATTORNEYS OF RECORD:\n"
                f"\tPursuant to California Code of Civil Procedure §2033.010, "
                f"{ds.propounding_party.role_label}, {ds.propounding_party.name.upper()} "
                f'("Propounding Party" or "{ds.propounding_party.role_label}"), hereby '
                f"requests that {ds.directed_to.role_label}, {ds.directed_to.name.upper()} "
                f'("{ds.directed_to.role_label}" or "Responding Party"), admit the truth of '
                f"the following matters within thirty (30) days as required by law."
            ),
        }
        return templates.get(ds.discovery_type, "")

    @staticmethod
    def find_caption_page(case_path: str) -> Optional[str]:
        """Find the Caption Page .docx in a case folder (case-insensitive search)."""
        if not os.path.isdir(case_path):
            return None
        for fname in os.listdir(case_path):
            if fname.lower().endswith(".docx") and "caption page" in fname.lower():
                return os.path.join(case_path, fname)
        return None
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `python -m pytest tests/test_discovery/test_assembler.py -v`
Expected: All tests PASS

- [ ] **Step 5: Commit**

```bash
git add icharlotte_core/discovery/assembler.py tests/test_discovery/test_assembler.py
git commit -m "feat(discovery): add document assembler for .docx generation"
```

---

### Task 6: Generation Engine

**Files:**
- Create: `icharlotte_core/discovery/engine.py`
- Create: `tests/test_discovery/test_engine.py`

**Depends on:** Tasks 1-5

- [ ] **Step 1: Write engine tests**

```python
# tests/test_discovery/test_engine.py
import os
import unittest
from unittest.mock import patch, MagicMock

from icharlotte_core.discovery.models import (
    DiscoveryMode, DiscoveryType, CustomStyle, Party, PartyRole,
)
from icharlotte_core.discovery.engine import DiscoveryEngine

DISCOVERY_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", "discovery")


class TestDiscoveryEngineStandardMode(unittest.TestCase):
    """Test standard mode — no LLM, pure template."""

    def setUp(self):
        self.engine = DiscoveryEngine(discovery_dir=DISCOVERY_DIR)
        self.propounding = Party(
            name="Servitek Electric, Inc.", role=PartyRole.DEFENDANT,
            is_our_client=True, abbreviation="Servitek",
        )
        self.responding = Party(
            name="Ruxandra Raschkovsky", role=PartyRole.PLAINTIFF,
            abbreviation="Pltf",
        )

    @unittest.skipUnless(
        os.path.exists(os.path.join(DISCOVERY_DIR, "Standard Negligence Discovery (5800.070)", "SI(1) tPltf.docx")),
        "Template files not present",
    )
    def test_generate_standard_si(self):
        results = self.engine.generate_standard(
            standard_type="negligence",
            discovery_types=[DiscoveryType.SI],
            directed_to=self.responding,
            propounding_party=self.propounding,
        )
        self.assertEqual(len(results), 1)
        ds = results[0]
        self.assertEqual(ds.discovery_type, DiscoveryType.SI)
        self.assertEqual(ds.set_number, 1)
        self.assertGreater(len(ds.requests), 50)

    @unittest.skipUnless(
        os.path.exists(os.path.join(DISCOVERY_DIR, "Standard Negligence Discovery (5800.070)")),
        "Template files not present",
    )
    def test_generate_standard_multiple_types(self):
        results = self.engine.generate_standard(
            standard_type="negligence",
            discovery_types=[DiscoveryType.SI, DiscoveryType.RPD],
            directed_to=self.responding,
            propounding_party=self.propounding,
        )
        self.assertEqual(len(results), 2)
        types = {r.discovery_type for r in results}
        self.assertEqual(types, {DiscoveryType.SI, DiscoveryType.RPD})


class TestDiscoveryEngineCustomMode(unittest.TestCase):
    """Test custom mode — uses LLM (mocked)."""

    def setUp(self):
        self.engine = DiscoveryEngine(discovery_dir=DISCOVERY_DIR)
        self.propounding = Party(
            name="Test Corp", role=PartyRole.DEFENDANT,
            is_our_client=True, abbreviation="Test",
        )
        self.responding = Party(
            name="Jane Doe", role=PartyRole.PLAINTIFF, abbreviation="Pltf",
        )

    def test_build_custom_only_prompt(self):
        prompt = self.engine.build_llm_prompt(
            discovery_type=DiscoveryType.SI,
            user_instructions="Generate interrogatories about the construction defect",
            custom_style=CustomStyle.CUSTOM_ONLY,
            context_text="Complaint alleges water intrusion from faulty roof.",
            start_number=1,
        )
        self.assertIn("Special Interrogatories", prompt)
        self.assertIn("construction defect", prompt)
        self.assertIn("water intrusion", prompt)
        # Should instruct to start at number 1
        self.assertIn("1", prompt)

    def test_build_standard_plus_custom_prompt(self):
        prompt = self.engine.build_llm_prompt(
            discovery_type=DiscoveryType.SI,
            user_instructions="Add questions about prior roof repairs",
            custom_style=CustomStyle.STANDARD_PLUS_CUSTOM,
            context_text="",
            start_number=64,
        )
        # Should instruct to start numbering at 64
        self.assertIn("64", prompt)

    def test_build_modified_standard_prompt(self):
        prompt = self.engine.build_llm_prompt(
            discovery_type=DiscoveryType.SI,
            user_instructions="Adapt for wrongful termination case",
            custom_style=CustomStyle.MODIFIED_STANDARD,
            context_text="",
            start_number=1,
            standard_requests_text="SPECIAL INTERROGATORY NO. 1:\nState all facts about negligence.",
        )
        self.assertIn("wrongful termination", prompt)
        self.assertIn("State all facts about negligence", prompt)


if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `python -m pytest tests/test_discovery/test_engine.py -v`
Expected: FAIL — `ModuleNotFoundError`

- [ ] **Step 3: Implement engine.py**

```python
# icharlotte_core/discovery/engine.py
"""Main orchestrator for the discovery generation pipeline."""
from __future__ import annotations

import os
from typing import List, Optional

from .models import (
    CustomStyle, DiscoveryMode, DiscoveryRequest, DiscoverySet,
    DiscoveryType, Party, SetTrackerResult, number_to_word,
)
from .templates import TemplateLoader, extract_requests_from_text, substitute_variables
from .set_tracker import SetTracker


class DiscoveryEngine:
    """Coordinates discovery generation across all modes."""

    def __init__(self, discovery_dir: str):
        self.discovery_dir = discovery_dir
        self.template_loader = TemplateLoader(discovery_dir)

    def generate_standard(
        self,
        standard_type: str,
        discovery_types: List[DiscoveryType],
        directed_to: Party,
        propounding_party: Party,
        variables: Optional[dict] = None,
    ) -> List[DiscoverySet]:
        """Generate discovery using standard templates. No LLM involved."""
        results = []
        defined_terms = self.template_loader.load_defined_terms()
        if variables:
            defined_terms = substitute_variables(defined_terms, variables)

        for disc_type in discovery_types:
            requests = self.template_loader.load_standard_requests(standard_type, disc_type)

            # Substitute variables in request text
            if variables:
                for req in requests:
                    req.text = substitute_variables(req.text, variables)
                    req.definitions = [substitute_variables(d, variables) for d in req.definitions]

            # Load instructions from template
            from .templates import STANDARD_TEMPLATES, TEMPLATE_FILENAMES
            folder = STANDARD_TEMPLATES.get(standard_type, "")
            filename = TEMPLATE_FILENAMES.get(disc_type, "")
            template_path = os.path.join(self.discovery_dir, folder, filename)
            instructions = ""
            if os.path.exists(template_path):
                instructions = self.template_loader.load_instructions(disc_type, template_path)
                if variables:
                    instructions = substitute_variables(instructions, variables)

            ds = DiscoverySet(
                discovery_type=disc_type,
                set_number=1,
                directed_to=directed_to,
                propounding_party=propounding_party,
                requests=requests,
                definitions_block=defined_terms,
                instructions_block=instructions,
                previous_count=0,
            )
            results.append(ds)

        return results

    def generate_custom(
        self,
        custom_style: CustomStyle,
        standard_type: str,
        discovery_types: List[DiscoveryType],
        directed_to: Party,
        propounding_party: Party,
        user_instructions: str,
        context_text: str = "",
        variables: Optional[dict] = None,
    ) -> List[DiscoverySet]:
        """Build DiscoverySet shells for custom mode. LLM text must be added separately.

        For STANDARD_PLUS_CUSTOM: includes standard requests, LLM requests to be appended.
        For CUSTOM_ONLY and MODIFIED_STANDARD: returns empty requests (LLM fills them).
        """
        results = []
        defined_terms = self.template_loader.load_defined_terms()
        if variables:
            defined_terms = substitute_variables(defined_terms, variables)

        for disc_type in discovery_types:
            requests = []
            start_number = 1

            if custom_style == CustomStyle.STANDARD_PLUS_CUSTOM:
                # Load standard requests first
                try:
                    requests = self.template_loader.load_standard_requests(standard_type, disc_type)
                    if variables:
                        for req in requests:
                            req.text = substitute_variables(req.text, variables)
                            req.definitions = [substitute_variables(d, variables) for d in req.definitions]
                    start_number = len(requests) + 1
                except (ValueError, FileNotFoundError):
                    start_number = 1

            ds = DiscoverySet(
                discovery_type=disc_type,
                set_number=1,
                directed_to=directed_to,
                propounding_party=propounding_party,
                requests=requests,
                definitions_block=defined_terms,
                instructions_block="",
                previous_count=0,
            )
            results.append(ds)

        return results

    def prepare_additional(
        self,
        case_path: str,
        discovery_types: List[DiscoveryType],
        directed_to: Party,
        propounding_party: Party,
    ) -> List[DiscoverySet]:
        """Prepare DiscoverySet shells for additional discovery.

        Runs SetTracker to determine set numbers and load previous definitions.
        LLM request text must be added separately.
        """
        tracker = SetTracker(case_path)
        results = []

        for disc_type in discovery_types:
            tracker_result = tracker.resolve_previous_set(
                disc_type, directed_to, self.discovery_dir
            )

            ds = DiscoverySet(
                discovery_type=disc_type,
                set_number=tracker_result.next_set_number,
                directed_to=directed_to,
                propounding_party=propounding_party,
                requests=[],  # LLM will fill these
                definitions_block=tracker_result.previous_definitions,
                instructions_block=tracker_result.previous_instructions,
                previous_count=tracker_result.last_request_number,
            )
            results.append(ds)

        return results

    def build_llm_prompt(
        self,
        discovery_type: DiscoveryType,
        user_instructions: str,
        custom_style: CustomStyle = CustomStyle.CUSTOM_ONLY,
        context_text: str = "",
        start_number: int = 1,
        standard_requests_text: str = "",
    ) -> str:
        """Build the LLM prompt for generating discovery requests.

        The prompt instructs the LLM to return ONLY numbered request text,
        no boilerplate, definitions, or instructions.
        """
        type_name = discovery_type.value
        header_template = discovery_type.request_header_template

        prompt_parts = [
            f"You are a California litigation attorney drafting {type_name}.",
            f"Generate numbered {type_name.lower()} using this exact format:",
            f"",
            f"{header_template.format(num='N')}",
            f"[Question text here]",
            f"",
            f"Start numbering at {start_number}.",
            f"Return ONLY the numbered requests. Do not include instructions, ",
            f"definitions, signature blocks, or any other boilerplate.",
            f"",
        ]

        if custom_style == CustomStyle.MODIFIED_STANDARD:
            prompt_parts.extend([
                "Below are the STANDARD requests. Modify and adapt them based on the user's instructions:",
                "",
                standard_requests_text,
                "",
            ])
        elif custom_style == CustomStyle.STANDARD_PLUS_CUSTOM:
            prompt_parts.extend([
                f"The standard set already contains requests numbered 1 through {start_number - 1}.",
                f"Generate ADDITIONAL requests starting at number {start_number}.",
                "",
            ])

        prompt_parts.extend([
            "USER INSTRUCTIONS:",
            user_instructions,
        ])

        if context_text:
            prompt_parts.extend([
                "",
                "CONTEXT DOCUMENTS:",
                context_text,
            ])

        return "\n".join(prompt_parts)

    def parse_llm_response(
        self, response: str, discovery_type: DiscoveryType
    ) -> List[DiscoveryRequest]:
        """Parse LLM response text into DiscoveryRequest objects."""
        return extract_requests_from_text(response, discovery_type)
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `python -m pytest tests/test_discovery/test_engine.py -v`
Expected: All tests PASS

- [ ] **Step 5: Commit**

```bash
git add icharlotte_core/discovery/engine.py tests/test_discovery/test_engine.py
git commit -m "feat(discovery): add generation engine orchestrator"
```

---

### Task 7: UI — Discovery Tab Shell and Left Pane

**Files:**
- Create: `icharlotte_core/ui/discovery_tab.py`

**Depends on:** Task 1 (models)

This is the largest task. It creates the full UI with left pane controls and conditional visibility.

- [ ] **Step 1: Create discovery_tab.py with tab shell and left pane controls**

```python
# icharlotte_core/ui/discovery_tab.py
"""Discovery tab with Propound and Respond sub-tabs."""
from __future__ import annotations

import os
from typing import List, Optional

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QTabWidget, QSplitter, QLabel,
    QComboBox, QRadioButton, QButtonGroup, QCheckBox, QTextEdit,
    QPushButton, QPlainTextEdit, QListWidgetItem, QMenu, QAction,
    QScrollArea, QFrame, QFileDialog, QDialog, QLineEdit, QFormLayout,
    QDialogButtonBox, QGroupBox,
)
from PySide6.QtCore import Qt, QTimer, Signal, QFileInfo
from PySide6.QtGui import QDragEnterEvent, QDropEvent, QIcon

from icharlotte_core.ui.tabs import ResizableListWidget
from icharlotte_core.llm import LLMWorker, ModelFetcher
from icharlotte_core.config import API_KEYS
from icharlotte_core.discovery.models import (
    Party, PartyRole, DiscoveryMode, DiscoveryType, CustomStyle,
    generate_abbreviation,
)


class DiscoveryTab(QWidget):
    """Main Discovery tab containing Propound and Respond sub-tabs."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.file_number = None
        self.case_path = None

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)

        self.sub_tabs = QTabWidget()
        self.propound_tab = PropoundTab()
        self.respond_tab = QWidget()  # Placeholder for future

        self.sub_tabs.addTab(self.propound_tab, "Propound")
        self.sub_tabs.addTab(self.respond_tab, "Respond")

        # Respond tab placeholder
        respond_layout = QVBoxLayout(self.respond_tab)
        respond_layout.addWidget(QLabel("Respond tab — coming soon"))
        respond_layout.addStretch()

        layout.addWidget(self.sub_tabs)

    def load_case(self, file_number: str):
        """Called when the user switches cases."""
        self.file_number = file_number
        self.propound_tab.load_case(file_number)


class PropoundTab(QWidget):
    """The Propound sub-tab with left pane controls and right pane editor."""

    SUPPORTED_EXTENSIONS = (
        ".pdf", ".docx", ".txt", ".msg", ".png", ".jpg", ".jpeg", ".gif", ".webp",
    )

    def __init__(self, parent=None):
        super().__init__(parent)
        self.file_number = None
        self.case_path = None
        self.parties: List[Party] = []
        self.attached_files: List[str] = []
        self.worker: Optional[LLMWorker] = None
        self.fetcher: Optional[ModelFetcher] = None
        self.cached_models = {}
        self.generated_sets = []  # List of DiscoverySet after generation

        self._setup_ui()

    def _setup_ui(self):
        """Build the left-right splitter layout."""
        main_layout = QHBoxLayout(self)
        main_layout.setContentsMargins(0, 0, 0, 0)

        splitter = QSplitter(Qt.Orientation.Horizontal)

        # --- Left Pane (scroll area) ---
        left_scroll = QScrollArea()
        left_scroll.setWidgetResizable(True)
        left_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        left_scroll.setMinimumWidth(300)
        left_scroll.setMaximumWidth(450)

        left_widget = QWidget()
        self.left_layout = QVBoxLayout(left_widget)
        self.left_layout.setContentsMargins(12, 12, 12, 12)
        self.left_layout.setSpacing(10)

        self._build_mode_section()
        self._build_standard_type_section()
        self._build_custom_style_section()
        self._build_discovery_types_section()
        self._build_directed_to_section()
        self._build_context_docs_section()
        self._build_llm_section()
        self._build_prompt_section()
        self._build_generate_button()

        self.left_layout.addStretch()
        left_scroll.setWidget(left_widget)

        # --- Right Pane ---
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        right_layout.setContentsMargins(0, 0, 0, 0)

        # Document sub-tabs
        self.doc_tabs = QTabWidget()
        self.doc_tabs.setTabsClosable(False)

        # Toolbar
        toolbar = QHBoxLayout()
        self.save_btn = QPushButton("Save as .docx")
        self.save_all_btn = QPushButton("Save All")
        self.status_label = QLabel("")
        self.status_label.setStyleSheet("color: #888; font-size: 11px;")
        toolbar.addWidget(self.save_btn)
        toolbar.addWidget(self.save_all_btn)
        toolbar.addStretch()
        toolbar.addWidget(self.status_label)

        # Empty state
        self.empty_state = QLabel("Configure settings and click Generate to create discovery requests")
        self.empty_state.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.empty_state.setStyleSheet("color: #666; font-size: 14px; padding: 40px;")

        right_layout.addLayout(toolbar)
        right_layout.addWidget(self.doc_tabs)
        right_layout.addWidget(self.empty_state)

        self.doc_tabs.hide()
        self.save_btn.hide()
        self.save_all_btn.hide()

        splitter.addWidget(left_scroll)
        splitter.addWidget(right_widget)
        splitter.setStretchFactor(0, 0)  # Left pane: fixed width
        splitter.setStretchFactor(1, 1)  # Right pane: stretch

        main_layout.addWidget(splitter)

        # Wire up signals
        self._connect_signals()
        # Set initial visibility
        self._on_mode_changed()

    # --- Section Builders ---

    def _build_mode_section(self):
        label = QLabel("DISCOVERY MODE")
        label.setStyleSheet("font-size: 11px; text-transform: uppercase; color: #888;")
        self.left_layout.addWidget(label)

        group_box = QGroupBox()
        group_layout = QVBoxLayout(group_box)
        group_layout.setContentsMargins(8, 8, 8, 8)

        self.mode_group = QButtonGroup(self)
        self.radio_standard = QRadioButton("Initial — Standard")
        self.radio_custom = QRadioButton("Initial — Custom")
        self.radio_additional = QRadioButton("Additional Discovery")
        self.radio_standard.setChecked(True)

        self.mode_group.addButton(self.radio_standard, 0)
        self.mode_group.addButton(self.radio_custom, 1)
        self.mode_group.addButton(self.radio_additional, 2)

        group_layout.addWidget(self.radio_standard)
        group_layout.addWidget(self.radio_custom)
        group_layout.addWidget(self.radio_additional)

        self.left_layout.addWidget(group_box)

    def _build_standard_type_section(self):
        self.standard_type_label = QLabel("STANDARD TYPE")
        self.standard_type_label.setStyleSheet("font-size: 11px; text-transform: uppercase; color: #888;")
        self.left_layout.addWidget(self.standard_type_label)

        self.standard_type_combo = QComboBox()
        self.standard_type_combo.addItem("Standard Negligence")
        self.standard_type_combo.addItem("Standard Wrongful Death (coming soon)")
        # Disable wrongful death for now
        model = self.standard_type_combo.model()
        model.item(1).setEnabled(False)

        self.left_layout.addWidget(self.standard_type_combo)

    def _build_custom_style_section(self):
        self.custom_style_label = QLabel("CUSTOM STYLE")
        self.custom_style_label.setStyleSheet("font-size: 11px; text-transform: uppercase; color: #888;")
        self.left_layout.addWidget(self.custom_style_label)

        self.custom_style_group = QButtonGroup(self)
        custom_box = QGroupBox()
        custom_layout = QVBoxLayout(custom_box)
        custom_layout.setContentsMargins(8, 8, 8, 8)

        self.radio_custom_only = QRadioButton("Custom Only")
        self.radio_standard_plus = QRadioButton("Standard + Custom")
        self.radio_modified = QRadioButton("Modified Standard")
        self.radio_custom_only.setChecked(True)

        self.custom_style_group.addButton(self.radio_custom_only, 0)
        self.custom_style_group.addButton(self.radio_standard_plus, 1)
        self.custom_style_group.addButton(self.radio_modified, 2)

        custom_layout.addWidget(self.radio_custom_only)
        custom_layout.addWidget(self.radio_standard_plus)
        custom_layout.addWidget(self.radio_modified)

        self.custom_style_box = custom_box
        self.left_layout.addWidget(custom_box)

    def _build_discovery_types_section(self):
        label = QLabel("DISCOVERY TYPES")
        label.setStyleSheet("font-size: 11px; text-transform: uppercase; color: #888;")
        self.left_layout.addWidget(label)

        types_layout = QHBoxLayout()
        self.check_si = QCheckBox("SI")
        self.check_rpd = QCheckBox("RPD")
        self.check_rfa = QCheckBox("RFA")
        self.check_si.setChecked(True)
        self.check_rpd.setChecked(True)

        types_layout.addWidget(self.check_si)
        types_layout.addWidget(self.check_rpd)
        types_layout.addWidget(self.check_rfa)
        types_layout.addStretch()

        self.left_layout.addLayout(types_layout)

    def _build_directed_to_section(self):
        label = QLabel("DIRECTED TO")
        label.setStyleSheet("font-size: 11px; text-transform: uppercase; color: #888;")
        self.left_layout.addWidget(label)

        self.directed_to_combo = QComboBox()
        self.directed_to_combo.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.directed_to_combo.contextMenuRequested.connect(self._show_party_context_menu)

        # Add Party button inline
        dt_layout = QHBoxLayout()
        dt_layout.addWidget(self.directed_to_combo, stretch=1)
        self.add_party_btn = QPushButton("+")
        self.add_party_btn.setFixedWidth(30)
        self.add_party_btn.setToolTip("Add Party")
        self.add_party_btn.clicked.connect(self._add_party_dialog)
        dt_layout.addWidget(self.add_party_btn)

        self.left_layout.addLayout(dt_layout)

    def _build_context_docs_section(self):
        label = QLabel("CONTEXT DOCUMENTS")
        label.setStyleSheet("font-size: 11px; text-transform: uppercase; color: #888;")
        self.left_layout.addWidget(label)

        self.file_list = ResizableListWidget()
        self.file_list.setMinimumHeight(60)
        self.file_list.setMaximumHeight(300)
        self.file_list.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.file_list.customContextMenuRequested.connect(self._show_file_context_menu)
        self.left_layout.addWidget(self.file_list)

    def _build_llm_section(self):
        llm_layout = QHBoxLayout()

        provider_vbox = QVBoxLayout()
        self.provider_label = QLabel("PROVIDER")
        self.provider_label.setStyleSheet("font-size: 11px; text-transform: uppercase; color: #888;")
        self.provider_combo = QComboBox()
        self.provider_combo.addItems(["Gemini", "OpenAI", "Claude"])
        provider_vbox.addWidget(self.provider_label)
        provider_vbox.addWidget(self.provider_combo)

        model_vbox = QVBoxLayout()
        self.model_label = QLabel("MODEL")
        self.model_label.setStyleSheet("font-size: 11px; text-transform: uppercase; color: #888;")
        self.model_combo = QComboBox()
        model_vbox.addWidget(self.model_label)
        model_vbox.addWidget(self.model_combo)

        llm_layout.addLayout(provider_vbox, stretch=1)
        llm_layout.addLayout(model_vbox, stretch=2)

        self.llm_widget = QWidget()
        self.llm_widget.setLayout(llm_layout)
        self.left_layout.addWidget(self.llm_widget)

    def _build_prompt_section(self):
        self.prompt_label = QLabel("INSTRUCTIONS / PROMPT")
        self.prompt_label.setStyleSheet("font-size: 11px; text-transform: uppercase; color: #888;")
        self.left_layout.addWidget(self.prompt_label)

        self.prompt_input = QTextEdit()
        self.prompt_input.setPlaceholderText("Describe what discovery requests to generate...")
        self.prompt_input.setMinimumHeight(80)
        self.prompt_input.setMaximumHeight(200)
        self.left_layout.addWidget(self.prompt_input)

    def _build_generate_button(self):
        self.generate_btn = QPushButton("Generate Discovery")
        self.generate_btn.setMinimumHeight(40)
        self.generate_btn.setStyleSheet(
            "QPushButton { background-color: #4a6cf7; color: white; font-size: 14px; "
            "font-weight: bold; border-radius: 6px; }"
            "QPushButton:hover { background-color: #3a5cd7; }"
            "QPushButton:disabled { background-color: #555; color: #888; }"
        )
        self.left_layout.addWidget(self.generate_btn)

    # --- Signal Connections ---

    def _connect_signals(self):
        self.mode_group.buttonClicked.connect(lambda: self._on_mode_changed())
        self.custom_style_group.buttonClicked.connect(lambda: self._on_custom_style_changed())
        self.provider_combo.currentTextChanged.connect(self._update_models)
        self.generate_btn.clicked.connect(self._on_generate)
        self.save_btn.clicked.connect(self._save_current)
        self.save_all_btn.clicked.connect(self._save_all)

    # --- Conditional Visibility ---

    def _on_mode_changed(self):
        mode = self._get_current_mode()
        is_standard = mode == DiscoveryMode.INITIAL_STANDARD
        is_custom = mode == DiscoveryMode.INITIAL_CUSTOM
        is_additional = mode == DiscoveryMode.ADDITIONAL

        # Standard Type: visible only in Standard mode
        self.standard_type_label.setVisible(is_standard)
        self.standard_type_combo.setVisible(is_standard)

        # Custom Style: visible only in Custom mode
        self.custom_style_label.setVisible(is_custom)
        self.custom_style_box.setVisible(is_custom)

        # LLM section: hidden in Standard mode
        self.llm_widget.setVisible(not is_standard)

        # Prompt: hidden in Standard mode
        self.prompt_label.setVisible(not is_standard)
        self.prompt_input.setVisible(not is_standard)

        # Generate button label
        if is_additional:
            self.generate_btn.setText("Generate Additional Discovery")
        else:
            self.generate_btn.setText("Generate Discovery")

        # Fetch models if LLM section just became visible
        if not is_standard and self.model_combo.count() == 0:
            self._update_models(self.provider_combo.currentText())

    def _on_custom_style_changed(self):
        style = self._get_custom_style()
        if style == CustomStyle.STANDARD_PLUS_CUSTOM:
            self.prompt_input.setPlaceholderText(
                "Describe additional requests to generate beyond the standard set..."
            )
        elif style == CustomStyle.MODIFIED_STANDARD:
            self.prompt_input.setPlaceholderText(
                "Describe how to modify the standard requests..."
            )
        else:
            self.prompt_input.setPlaceholderText(
                "Describe what discovery requests to generate..."
            )

    # --- Helper Methods ---

    def _get_current_mode(self) -> DiscoveryMode:
        btn_id = self.mode_group.checkedId()
        return [DiscoveryMode.INITIAL_STANDARD, DiscoveryMode.INITIAL_CUSTOM, DiscoveryMode.ADDITIONAL][btn_id]

    def _get_custom_style(self) -> CustomStyle:
        btn_id = self.custom_style_group.checkedId()
        return [CustomStyle.CUSTOM_ONLY, CustomStyle.STANDARD_PLUS_CUSTOM, CustomStyle.MODIFIED_STANDARD][btn_id]

    def _get_selected_types(self) -> List[DiscoveryType]:
        types = []
        if self.check_si.isChecked():
            types.append(DiscoveryType.SI)
        if self.check_rpd.isChecked():
            types.append(DiscoveryType.RPD)
        if self.check_rfa.isChecked():
            types.append(DiscoveryType.RFA)
        return types

    def _get_directed_to_party(self) -> Optional[Party]:
        idx = self.directed_to_combo.currentIndex()
        if 0 <= idx < len(self.parties):
            non_client = [p for p in self.parties if not p.is_our_client]
            if 0 <= idx < len(non_client):
                return non_client[idx]
        return None

    def _get_our_client(self) -> Optional[Party]:
        for p in self.parties:
            if p.is_our_client:
                return p
        return None

    # --- Case Loading ---

    def load_case(self, file_number: str):
        """Load case data and populate party roster."""
        self.file_number = file_number
        self._load_parties()
        self._refresh_party_combo()

    def _load_parties(self):
        """Load parties from CaseDataManager, seeding from plaintiffs/defendants."""
        self.parties = []
        if not self.file_number:
            return

        try:
            from icharlotte_core.case_data_manager import CaseDataManager
            from icharlotte_core.utils import get_case_path

            cdm = CaseDataManager()
            self.case_path = get_case_path(self.file_number)

            # Check for saved party roster first
            saved_roster = cdm.get_value(self.file_number, "discovery_party_roster")
            if saved_roster and isinstance(saved_roster, list):
                self.parties = [Party.from_dict(d) for d in saved_roster]
                return

            # Seed from existing plaintiffs/defendants
            plaintiffs = cdm.get_value(self.file_number, "plaintiffs") or ""
            defendants = cdm.get_value(self.file_number, "defendants") or ""
            client_name = cdm.get_value(self.file_number, "client_name") or ""

            if isinstance(plaintiffs, str) and plaintiffs:
                for name in self._split_party_names(plaintiffs):
                    self.parties.append(Party(name=name, role=PartyRole.PLAINTIFF))
            if isinstance(defendants, str) and defendants:
                for name in self._split_party_names(defendants):
                    is_client = client_name and client_name.lower() in name.lower()
                    self.parties.append(Party(name=name, role=PartyRole.DEFENDANT, is_our_client=is_client))

            # Generate abbreviations
            for p in self.parties:
                p.abbreviation = generate_abbreviation(p, self.parties)

        except Exception as e:
            print(f"Error loading parties: {e}")

    @staticmethod
    def _split_party_names(names_str: str) -> List[str]:
        """Split a party names string (may be comma or semicolon separated)."""
        # Handle "John Doe; Jane Doe" or "John Doe, Jane Doe"
        for sep in [";", " and ", " AND "]:
            if sep in names_str:
                return [n.strip() for n in names_str.split(sep) if n.strip()]
        return [names_str.strip()] if names_str.strip() else []

    def _refresh_party_combo(self):
        """Update the Directed To dropdown from the parties list."""
        self.directed_to_combo.blockSignals(True)
        self.directed_to_combo.clear()
        for p in self.parties:
            if not p.is_our_client:
                label = f"{p.role_label} — {p.name}"
                self.directed_to_combo.addItem(label)
        self.directed_to_combo.blockSignals(False)

    def _save_parties(self):
        """Persist party roster back to CaseDataManager and sync case variables."""
        if not self.file_number:
            return
        try:
            from icharlotte_core.case_data_manager import CaseDataManager
            cdm = CaseDataManager()

            # Save roster
            roster = [p.to_dict() for p in self.parties]
            cdm.save_variable(self.file_number, "discovery_party_roster", roster, source="user_fix")

            # Sync plaintiffs/defendants back
            plaintiffs = "; ".join(p.name for p in self.parties if p.role == PartyRole.PLAINTIFF)
            defendants = "; ".join(p.name for p in self.parties if p.role == PartyRole.DEFENDANT)
            if plaintiffs:
                cdm.save_variable(self.file_number, "plaintiffs", plaintiffs, source="user_fix")
            if defendants:
                cdm.save_variable(self.file_number, "defendants", defendants, source="user_fix")
        except Exception as e:
            print(f"Error saving parties: {e}")

    # --- Party Management ---

    def _add_party_dialog(self):
        """Show dialog to add a new party."""
        dialog = PartyEditDialog(self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            party = dialog.get_party()
            party.abbreviation = generate_abbreviation(party, self.parties + [party])
            self.parties.append(party)
            self._refresh_party_combo()
            self._save_parties()

    def _show_party_context_menu(self, pos):
        """Right-click menu on the Directed To dropdown for edit/remove."""
        idx = self.directed_to_combo.currentIndex()
        non_client = [p for p in self.parties if not p.is_our_client]
        if idx < 0 or idx >= len(non_client):
            return

        party = non_client[idx]
        menu = QMenu(self)

        edit_action = QAction("Edit Party", self)
        edit_action.triggered.connect(lambda: self._edit_party(party))
        menu.addAction(edit_action)

        remove_action = QAction("Remove Party", self)
        remove_action.triggered.connect(lambda: self._remove_party(party))
        menu.addAction(remove_action)

        menu.exec(self.directed_to_combo.mapToGlobal(pos))

    def _edit_party(self, party: Party):
        dialog = PartyEditDialog(self, party)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            updated = dialog.get_party()
            party.name = updated.name
            party.role = updated.role
            party.abbreviation = generate_abbreviation(party, self.parties)
            self._refresh_party_combo()
            self._save_parties()

    def _remove_party(self, party: Party):
        self.parties.remove(party)
        self._refresh_party_combo()
        self._save_parties()

    # --- File Management (Context Documents) ---

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.accept()
        else:
            event.ignore()

    def dropEvent(self, event: QDropEvent):
        files_to_add = []
        for url in event.mimeData().urls():
            path = url.toLocalFile()
            if os.path.isfile(path) and path.lower().endswith(self.SUPPORTED_EXTENSIONS):
                files_to_add.append(path)
        event.accept()
        if files_to_add:
            QTimer.singleShot(0, lambda: self._process_dropped_files(files_to_add))

    def _process_dropped_files(self, file_paths):
        for path in file_paths:
            self._add_file(path)

    def _add_file(self, path: str):
        if path in self.attached_files:
            return
        self.attached_files.append(path)
        item = QListWidgetItem(os.path.basename(path))
        item.setToolTip(path)
        item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
        item.setCheckState(Qt.CheckState.Checked)
        item.setData(Qt.ItemDataRole.UserRole, path)
        self.file_list.addItem(item)

    def _show_file_context_menu(self, pos):
        item = self.file_list.itemAt(pos)
        if not item:
            return
        menu = QMenu(self)
        remove_action = QAction("Remove", self)
        remove_action.triggered.connect(lambda: self._remove_file(item))
        menu.addAction(remove_action)
        menu.exec(self.file_list.mapToGlobal(pos))

    def _remove_file(self, item):
        path = item.data(Qt.ItemDataRole.UserRole)
        if path in self.attached_files:
            self.attached_files.remove(path)
        row = self.file_list.row(item)
        self.file_list.takeItem(row)

    def read_files_content(self) -> str:
        """Read text content from checked attached files."""
        content = ""
        for i in range(self.file_list.count()):
            item = self.file_list.item(i)
            if item.checkState() != Qt.CheckState.Checked:
                continue
            path = item.data(Qt.ItemDataRole.UserRole)
            if not path:
                continue
            ext = os.path.splitext(path)[1].lower()
            content += f"\n--- FILE: {os.path.basename(path)} ---\n"
            try:
                if ext == ".txt":
                    with open(path, "r", encoding="utf-8", errors="ignore") as f:
                        content += f.read()
                elif ext == ".docx":
                    from docx import Document
                    doc = Document(path)
                    for p in doc.paragraphs:
                        content += p.text + "\n"
                elif ext == ".pdf":
                    import fitz
                    doc = fitz.open(path)
                    for page in doc:
                        content += page.get_text() + "\n"
                    doc.close()
            except Exception as e:
                content += f"[Error reading file: {e}]\n"
        return content

    # --- LLM Provider/Model ---

    def _update_models(self, provider):
        self.model_combo.clear()
        if provider in self.cached_models:
            self.model_combo.addItems(self.cached_models[provider])
            return
        api_key = API_KEYS.get(provider)
        if not api_key and provider != "Claude":
            self.model_combo.addItem(f"No API Key for {provider}")
            return
        self.model_combo.addItem("Fetching models...")
        self.model_combo.setEnabled(False)

        if self.fetcher is not None:
            try:
                self.fetcher.finished.disconnect()
                self.fetcher.error.disconnect()
            except (TypeError, RuntimeError):
                pass
            if self.fetcher.isRunning():
                self.fetcher.wait(1000)

        self.fetcher = ModelFetcher(provider, api_key)
        self.fetcher.finished.connect(self._on_models_fetched)
        self.fetcher.error.connect(lambda err: self._on_models_fetched(provider, [f"Error: {err}"]))
        self.fetcher.start()

    def _on_models_fetched(self, provider, models):
        self.cached_models[provider] = models
        if self.provider_combo.currentText() == provider:
            self.model_combo.clear()
            self.model_combo.addItems(models)
            self.model_combo.setEnabled(True)

    # --- Generate / Save (stubs — wired in Task 10) ---

    def _on_generate(self):
        """Placeholder — wired to engine in Task 10."""
        pass

    def _save_current(self):
        """Placeholder — wired to assembler in Task 9."""
        pass

    def _save_all(self):
        """Placeholder — wired to assembler in Task 9."""
        pass


class PartyEditDialog(QDialog):
    """Dialog for adding or editing a party."""

    def __init__(self, parent=None, party: Optional[Party] = None):
        super().__init__(parent)
        self.setWindowTitle("Edit Party" if party else "Add Party")
        self.setMinimumWidth(350)

        layout = QFormLayout(self)

        self.name_input = QLineEdit()
        self.name_input.setPlaceholderText("Party name")
        if party:
            self.name_input.setText(party.name)

        self.role_combo = QComboBox()
        for role in PartyRole:
            self.role_combo.addItem(role.value)
        if party:
            self.role_combo.setCurrentText(party.role.value)

        self.is_client_check = QCheckBox("Our Client")
        if party:
            self.is_client_check.setChecked(party.is_our_client)

        layout.addRow("Name:", self.name_input)
        layout.addRow("Role:", self.role_combo)
        layout.addRow("", self.is_client_check)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addRow(buttons)

    def get_party(self) -> Party:
        return Party(
            name=self.name_input.text().strip(),
            role=PartyRole(self.role_combo.currentText()),
            is_our_client=self.is_client_check.isChecked(),
        )
```

- [ ] **Step 2: Verify the module imports correctly**

Run: `python -c "from icharlotte_core.ui.discovery_tab import DiscoveryTab; print('OK')"`
Expected: `OK`

- [ ] **Step 3: Commit**

```bash
git add icharlotte_core/ui/discovery_tab.py
git commit -m "feat(discovery): add Discovery tab UI shell with left pane controls"
```

---

### Task 8: Wire Generate Button to Engine

**Files:**
- Modify: `icharlotte_core/ui/discovery_tab.py`

**Depends on:** Tasks 6, 7

- [ ] **Step 1: Implement the _on_generate, _display_results, _save_current, and _save_all methods**

Replace the stub methods in `PropoundTab` with the full implementations. All methods below are instance methods of the `PropoundTab` class.

First, add these imports at the top of `discovery_tab.py`:

```python
from icharlotte_core.discovery.engine import DiscoveryEngine
from icharlotte_core.discovery.assembler import DiscoveryAssembler
```

Then replace each stub method in the `PropoundTab` class with the corresponding implementation below:
def _on_generate(self):
    """Run the discovery generation pipeline."""
    mode = self._get_current_mode()
    types = self._get_selected_types()
    directed_to = self._get_directed_to_party()
    our_client = self._get_our_client()

    if not types:
        return
    if not directed_to or not our_client:
        return

    discovery_dir = os.path.join(os.getcwd(), "discovery")
    engine = DiscoveryEngine(discovery_dir)

    # Build variables dict for template substitution
    variables = {
        "responding_party_name": directed_to.name.upper(),
        "propounding_party_name": our_client.name.upper(),
        "defendant_name": our_client.name.upper(),
    }

    self.generate_btn.setEnabled(False)
    self.generate_btn.setText("Generating...")

    if mode == DiscoveryMode.INITIAL_STANDARD:
        # Synchronous — no LLM
        standard_type = "negligence"  # From combo in future
        results = engine.generate_standard(
            standard_type, types, directed_to, our_client, variables
        )
        self._display_results(results)
        self.generate_btn.setEnabled(True)
        self.generate_btn.setText("Generate Discovery")

    elif mode == DiscoveryMode.INITIAL_CUSTOM:
        custom_style = self._get_custom_style()
        user_prompt = self.prompt_input.toPlainText().strip()
        if not user_prompt:
            self.generate_btn.setEnabled(True)
            self.generate_btn.setText("Generate Discovery")
            return

        # Build shells
        results = engine.generate_custom(
            custom_style, "negligence", types, directed_to, our_client,
            user_prompt, variables=variables,
        )
        self.generated_sets = results

        # For each type, fire LLM
        context_text = self.read_files_content()
        for ds in results:
            start_num = len(ds.requests) + 1
            std_text = ""
            if custom_style == CustomStyle.MODIFIED_STANDARD:
                std_text = ds.plain_text()

            prompt = engine.build_llm_prompt(
                ds.discovery_type, user_prompt, custom_style,
                context_text, start_num, std_text,
            )
            self._run_llm(prompt, ds)

    elif mode == DiscoveryMode.ADDITIONAL:
        user_prompt = self.prompt_input.toPlainText().strip()
        if not user_prompt or not self.case_path:
            self.generate_btn.setEnabled(True)
            self.generate_btn.setText("Generate Additional Discovery")
            return

        results = engine.prepare_additional(
            self.case_path, types, directed_to, our_client,
        )
        self.generated_sets = results

        context_text = self.read_files_content()
        for ds in results:
            start_num = ds.previous_count + 1
            prompt = engine.build_llm_prompt(
                ds.discovery_type, user_prompt,
                CustomStyle.CUSTOM_ONLY, context_text, start_num,
            )
            self._run_llm(prompt, ds)


def _run_llm(self, prompt: str, discovery_set: DiscoverySet):
    """Run LLM generation for a single discovery set."""
    provider = self.provider_combo.currentText()
    model = self.model_combo.currentText()
    settings = {"stream": False}

    self.worker = LLMWorker(provider, model, "", prompt, "", settings)
    self.worker.finished.connect(
        lambda text, ds=discovery_set: self._on_llm_finished(text, ds)
    )
    self.worker.error.connect(self._on_llm_error)
    self.worker.start()


def _on_llm_finished(self, response_text: str, discovery_set: DiscoverySet):
    """Handle LLM response — parse into requests and display."""
    from icharlotte_core.discovery.templates import extract_requests_from_text
    new_requests = extract_requests_from_text(response_text, discovery_set.discovery_type)

    if self._get_custom_style() == CustomStyle.STANDARD_PLUS_CUSTOM:
        discovery_set.requests.extend(new_requests)
    else:
        discovery_set.requests = new_requests

    self._display_results(self.generated_sets)
    self.generate_btn.setEnabled(True)
    mode = self._get_current_mode()
    if mode == DiscoveryMode.ADDITIONAL:
        self.generate_btn.setText("Generate Additional Discovery")
    else:
        self.generate_btn.setText("Generate Discovery")


def _on_llm_error(self, error_msg: str):
    """Handle LLM error."""
    self.generate_btn.setEnabled(True)
    self.status_label.setText(f"Error: {error_msg}")


def _display_results(self, results: List[DiscoverySet]):
    """Display generated discovery sets in right pane sub-tabs."""
    self.generated_sets = results
    self.doc_tabs.clear()
    self.empty_state.hide()
    self.doc_tabs.show()
    self.save_btn.show()
    self.save_all_btn.show()

    for ds in results:
        editor = QPlainTextEdit()
        editor.setPlainText(ds.plain_text())
        editor.setFont(editor.document().defaultFont())  # monospace would be nice
        tab_label = f"{ds.discovery_type.abbreviation}({ds.set_number}) t{ds.directed_to.abbreviation}"
        self.doc_tabs.addTab(editor, tab_label)

    # Update status
    if results:
        ds = results[0]
        self.status_label.setText(
            f"{len(ds.requests)} requests | Set {ds.set_word} | to {ds.directed_to.role_label}"
        )


def _save_current(self):
    """Save the currently visible document tab as .docx."""
    idx = self.doc_tabs.currentIndex()
    if idx < 0 or idx >= len(self.generated_sets):
        return
    ds = self.generated_sets[idx]
    editor = self.doc_tabs.widget(idx)
    plain_text = editor.toPlainText()
    self._save_discovery_set(ds, plain_text)


def _save_all(self):
    """Save all generated document tabs as .docx."""
    for i, ds in enumerate(self.generated_sets):
        editor = self.doc_tabs.widget(i)
        plain_text = editor.toPlainText()
        self._save_discovery_set(ds, plain_text)


def _save_discovery_set(self, ds: DiscoverySet, plain_text: str):
    """Save a single discovery set to .docx."""
    if not self.case_path:
        return

    caption_path = DiscoveryAssembler.find_caption_page(self.case_path)
    if not caption_path:
        self.status_label.setText("Error: Caption Page not found in case folder")
        return

    output_dir = os.path.join(self.case_path, "NOTES", "AI OUTPUT", "DISCOVERY REQUESTS")
    output_path = os.path.join(output_dir, ds.filename)

    assembler = DiscoveryAssembler(caption_path)
    assembler.assemble_from_plain_text(
        plain_text, ds, output_path,
        attorney_name="",  # Could be pulled from config
    )
    self.status_label.setText(f"Saved: {ds.filename}")
```

- [ ] **Step 2: Verify the generate flow works with Standard mode**

Run: `python -c "
from icharlotte_core.ui.discovery_tab import PropoundTab
print('PropoundTab imports OK')
"`
Expected: `PropoundTab imports OK`

- [ ] **Step 3: Commit**

```bash
git add icharlotte_core/ui/discovery_tab.py
git commit -m "feat(discovery): wire generate button to engine and add save functionality"
```

---

### Task 9: Register in iCharlotte.py

**Files:**
- Modify: `iCharlotte.py`

**Depends on:** Task 8

- [ ] **Step 1: Read the current tab registration section**

Read `iCharlotte.py` around lines 68-77 (imports) and lines 824-848 (tab registration) and lines 1341-1350 (case switch handling) to find exact insertion points.

- [ ] **Step 2: Add import**

Add after the existing UI tab imports (around line 77):

```python
from icharlotte_core.ui.discovery_tab import DiscoveryTab
```

- [ ] **Step 3: Add tab registration**

After the existing tab additions (find the pattern `self.tabs.addTab`), add:

```python
# --- Tab: Discovery ---
self.discovery_tab = DiscoveryTab()
self.tabs.addTab(self.discovery_tab, "Discovery")
if self.file_number:
    self.discovery_tab.load_case(self.file_number)
```

- [ ] **Step 4: Add case switch handler**

In the case switch method (around line 1341), add:

```python
if hasattr(self, 'discovery_tab'):
    self.discovery_tab.load_case(self.file_number)
```

- [ ] **Step 5: Test the app launches with the new tab**

Run: `python iCharlotte.py`
Expected: App launches with "Discovery" tab visible. Clicking it shows Propound/Respond sub-tabs. Left pane controls render correctly.

- [ ] **Step 6: Commit**

```bash
git add iCharlotte.py
git commit -m "feat(discovery): register Discovery tab in main window"
```

---

### Task 10: End-to-End Standard Mode Test

**Files:**
- Modify: `tests/test_discovery/test_engine.py`

**Depends on:** Tasks 1-9

- [ ] **Step 1: Add end-to-end integration test**

Add to `test_engine.py`:

```python
class TestEndToEndStandardMode(unittest.TestCase):
    """Full pipeline: generate standard SI → assemble .docx → verify output."""

    def setUp(self):
        self.tmpdir = tempfile.mkdtemp()
        self.engine = DiscoveryEngine(discovery_dir=DISCOVERY_DIR)
        self.propounding = Party(
            name="Servitek Electric, Inc.", role=PartyRole.DEFENDANT,
            is_our_client=True, abbreviation="Servitek",
        )
        self.responding = Party(
            name="Ruxandra Raschkovsky", role=PartyRole.PLAINTIFF,
            abbreviation="Pltf",
        )

    def tearDown(self):
        shutil.rmtree(self.tmpdir)

    @unittest.skipUnless(
        os.path.exists(os.path.join(DISCOVERY_DIR, "Caption Page (AS FM).docx")),
        "Caption page not available",
    )
    def test_standard_si_full_pipeline(self):
        """Generate standard SI, assemble into .docx, verify key content."""
        from icharlotte_core.discovery.assembler import DiscoveryAssembler

        # 1. Generate
        results = self.engine.generate_standard(
            "negligence", [DiscoveryType.SI], self.responding, self.propounding,
        )
        self.assertEqual(len(results), 1)
        ds = results[0]
        self.assertGreater(len(ds.requests), 50)

        # 2. Assemble
        caption_path = os.path.join(DISCOVERY_DIR, "Caption Page (AS FM).docx")
        assembler = DiscoveryAssembler(caption_path)
        output_path = os.path.join(self.tmpdir, ds.filename)
        assembler.assemble(ds, output_path, attorney_name="Andrei Serpik")

        # 3. Verify
        self.assertTrue(os.path.exists(output_path))
        doc = Document(output_path)
        full_text = "\n".join(p.text for p in doc.paragraphs)

        # Verify key sections present
        self.assertIn("SPECIAL INTERROGATORY NO. 1:", full_text)
        self.assertIn("PROPOUNDING PARTY", full_text)
        self.assertIn("RESPONDING PARTY", full_text)
        # Declaration should be present (63 > 35)
        self.assertIn("DECLARATION", full_text)
        self.assertIn("2030.070", full_text)

    @unittest.skipUnless(
        os.path.exists(os.path.join(DISCOVERY_DIR, "Caption Page (AS FM).docx")),
        "Caption page not available",
    )
    def test_standard_rpd_full_pipeline(self):
        """Generate standard RPD, assemble, verify no declaration needed."""
        from icharlotte_core.discovery.assembler import DiscoveryAssembler

        results = self.engine.generate_standard(
            "negligence", [DiscoveryType.RPD], self.responding, self.propounding,
        )
        ds = results[0]

        caption_path = os.path.join(DISCOVERY_DIR, "Caption Page (AS FM).docx")
        assembler = DiscoveryAssembler(caption_path)
        output_path = os.path.join(self.tmpdir, ds.filename)
        assembler.assemble(ds, output_path)

        doc = Document(output_path)
        full_text = "\n".join(p.text for p in doc.paragraphs)
        self.assertIn("REQUEST FOR PRODUCTION NO. 1:", full_text)
        # RPD should NOT have a declaration
        self.assertNotIn("DECLARATION", full_text)
```

- [ ] **Step 2: Run the full test suite**

Run: `python -m pytest tests/test_discovery/ -v`
Expected: All tests PASS

- [ ] **Step 3: Commit**

```bash
git add tests/test_discovery/test_engine.py
git commit -m "test(discovery): add end-to-end integration tests for standard mode"
```

---

### Task 11: Manual Verification and Cleanup

- [ ] **Step 1: Launch the app and verify the Discovery tab**

Run: `python iCharlotte.py`

Verify:
- Discovery tab appears in the main tab bar
- Propound and Respond sub-tabs are visible
- Propound left pane shows all controls
- Switching between Standard/Custom/Additional modes correctly shows/hides controls
- Party roster populates from case data when a case is loaded

- [ ] **Step 2: Test Standard mode generation**

1. Select a case that has template files available
2. Select "Initial — Standard" mode
3. Check SI and RPD
4. Click "Generate Discovery"
5. Verify right pane shows sub-tabs with discovery text
6. Click "Save All" and verify .docx files are created in NOTES/AI OUTPUT/DISCOVERY REQUESTS/

- [ ] **Step 3: Open generated .docx and visually inspect**

Open the saved .docx files and verify:
- Caption page is present
- Document title is correct
- Propounding/Responding party block is correct
- Instructions section is present
- Numbered requests are present and properly formatted
- Declaration is present (for SI with >35 requests)
- Signature block is present

- [ ] **Step 4: Fix any issues found during manual testing**

Iterate until the Standard mode pipeline produces correctly formatted output.

- [ ] **Step 5: Final commit**

```bash
git add -A
git commit -m "feat(discovery): complete Discovery Propound tab initial implementation"
```

---

## Post-Implementation Notes

### What's Not in This Plan (Future Work)
- Custom mode LLM generation needs real-world testing with different prompts
- Additional mode needs testing with actual propounded folder structures
- Respond sub-tab (separate spec needed)
- Form Interrogatories PDF form-filling
- Proof of Service generation
- Rich text editor upgrade
- Word formatting refinements (matching exact styles from sample documents — may need OXML tuning after visual inspection)
