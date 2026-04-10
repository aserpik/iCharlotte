# Discovery Respond Tab Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a "Respond" subtab to the Discovery tab that generates objections and substantive responses to incoming discovery (FI, SI, RFA, RPD) using a three-phase pipeline (parse → select objections → draft responses) with configurable rules.

**Architecture:** Three-phase pipeline housed in `icharlotte_core/discovery/` (response_parser, objection_selector, response_drafter), Word assembly via response_assembler, configurable rules via response_rules + a dialog UI. The RespondTab UI mirrors PropoundTab layout. All LLM calls run in QThread workers.

**Tech Stack:** Python 3.x, PySide6, python-docx, PyMuPDF (fitz), LLMWorker (existing), CaseDataManager (existing)

**Design Spec:** `docs/superpowers/specs/2026-04-10-discovery-respond-tab-design.md`

---

## File Map

| File | Action | Responsibility |
|------|--------|----------------|
| `icharlotte_core/discovery/response_rules.py` | Create | ResponseRules dataclass, serialization, default loading from samples |
| `icharlotte_core/discovery/response_parser.py` | Create | Phase 1: Parse discovery PDFs into ParsedDiscovery/ParsedRequest |
| `icharlotte_core/discovery/objection_selector.py` | Create | Phase 2: Rule-based + LLM objection selection per request |
| `icharlotte_core/discovery/response_drafter.py` | Create | Phase 3: Draft substantive responses, assemble plain-text output |
| `icharlotte_core/discovery/response_assembler.py` | Create | Word document assembly from caption template + response text |
| `icharlotte_core/ui/respond_tab.py` | Create | RespondTab widget: left pane controls + right pane editor |
| `icharlotte_core/ui/response_rules_dialog.py` | Create | Three-tab rules editor dialog |
| `icharlotte_core/ui/discovery_tab.py` | Modify | Replace Respond placeholder with RespondTab; update load_case() |
| `config/response_rules_default.json` | Create | Default rules JSON (generated from sample docs on first run) |
| `tests/test_response_rules.py` | Create | Tests for ResponseRules |
| `tests/test_response_parser.py` | Create | Tests for discovery parser |
| `tests/test_objection_selector.py` | Create | Tests for objection selector |
| `tests/test_response_drafter.py` | Create | Tests for response drafter |
| `tests/test_response_assembler.py` | Create | Tests for Word assembler |
| `tests/test_respond_tab.py` | Create | Tests for RespondTab UI |

---

## Task 1: ResponseRules Data Model & Defaults

**Files:**
- Create: `icharlotte_core/discovery/response_rules.py`
- Create: `tests/test_response_rules.py`

This is the foundation — every other module depends on ResponseRules for boilerplate text and strategy settings.

- [ ] **Step 1: Write test for ResponseRules dataclass and serialization**

```python
# tests/test_response_rules.py
"""Tests for ResponseRules dataclass, serialization, and default loading."""
import json
import os
import unittest
import tempfile

from icharlotte_core.discovery.response_rules import ResponseRules


class TestResponseRules(unittest.TestCase):
    """Verify ResponseRules construction, serialization, and defaults."""

    def test_create_with_defaults(self):
        """ResponseRules() with no args should use sensible defaults."""
        rules = ResponseRules()
        self.assertEqual(rules.objection_aggressiveness, "aggressive")
        self.assertTrue(rules.always_include_privacy_objection)
        self.assertTrue(rules.always_include_privilege_objection)
        self.assertTrue(rules.always_include_burden_objection)
        self.assertTrue(rules.auto_flag_compound)
        self.assertTrue(rules.auto_flag_broad_definitions)
        self.assertEqual(rules.si_response_style, "minimal")
        self.assertEqual(rules.rfa_default_posture, "cautious")
        self.assertEqual(rules.rpd_default_posture, "context_dependent")
        self.assertFalse(rules.fi_17_1_auto_refresh)
        self.assertIn("Subject to and without waiving", rules.waiver_language)
        self.assertIn("Discovery and investigation are ongoing", rules.reservation_clause)
        self.assertEqual(rules.custom_instructions, "")

    def test_to_dict_roundtrip(self):
        """to_dict → from_dict produces identical object."""
        rules = ResponseRules()
        rules.custom_instructions = "Test custom instruction"
        d = rules.to_dict()
        restored = ResponseRules.from_dict(d)
        self.assertEqual(rules.to_dict(), restored.to_dict())

    def test_from_dict_partial(self):
        """from_dict with partial dict uses defaults for missing keys."""
        partial = {"objection_aggressiveness": "conservative", "custom_instructions": "be brief"}
        rules = ResponseRules.from_dict(partial)
        self.assertEqual(rules.objection_aggressiveness, "conservative")
        self.assertEqual(rules.custom_instructions, "be brief")
        # Missing keys get defaults
        self.assertTrue(rules.always_include_privacy_objection)
        self.assertEqual(rules.si_response_style, "minimal")

    def test_save_and_load_json(self):
        """save_to_json and load_from_json roundtrip."""
        rules = ResponseRules()
        rules.objection_aggressiveness = "moderate"
        with tempfile.NamedTemporaryFile(suffix=".json", delete=False, mode="w") as f:
            path = f.name
        try:
            rules.save_to_json(path)
            loaded = ResponseRules.load_from_json(path)
            self.assertEqual(loaded.objection_aggressiveness, "moderate")
            self.assertEqual(loaded.waiver_language, rules.waiver_language)
        finally:
            os.unlink(path)

    def test_preliminary_statements_exist_for_all_types(self):
        """Each discovery type has its own preliminary statement."""
        rules = ResponseRules()
        self.assertTrue(len(rules.preliminary_statement_fi) > 100)
        self.assertTrue(len(rules.preliminary_statement_si) > 100)
        self.assertTrue(len(rules.preliminary_statement_rfa) > 100)
        self.assertTrue(len(rules.preliminary_statement_rpd) > 100)

    def test_intro_templates_have_placeholders(self):
        """Intro templates contain {placeholder} variables."""
        rules = ResponseRules()
        self.assertIn("{responding_party}", rules.intro_template_fi)
        self.assertIn("{propounding_party}", rules.intro_template_fi)
        self.assertIn("{set_word}", rules.intro_template_fi)

    def test_general_objections_only_for_rfa_rpd(self):
        """General objections exist for RFA and RPD but not FI/SI."""
        rules = ResponseRules()
        self.assertTrue(len(rules.general_objections_rfa) > 100)
        self.assertTrue(len(rules.general_objections_rpd) > 100)

    def test_verification_template_exists(self):
        """Verification template has placeholder fields."""
        rules = ResponseRules()
        self.assertIn("{responding_party}", rules.verification_template)


if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 2: Run test to verify it fails**

```bash
python -m pytest tests/test_response_rules.py -v
```
Expected: FAIL — `ModuleNotFoundError: No module named 'icharlotte_core.discovery.response_rules'`

- [ ] **Step 3: Implement ResponseRules**

Create `icharlotte_core/discovery/response_rules.py` with:

```python
"""
Response rules configuration for the discovery response pipeline.

Defines the ResponseRules dataclass that holds all configurable strategy
settings, boilerplate text, and custom instructions used when generating
discovery responses. Supports JSON serialization for per-case persistence
and default loading from sample documents.
"""
import json
import os
from dataclasses import dataclass, field, fields, asdict
from typing import Optional


# ---------------------------------------------------------------------------
# Default boilerplate text (extracted from sample response documents)
# ---------------------------------------------------------------------------

_DEFAULT_WAIVER = (
    "Subject to and without waiving the foregoing objections, "
    "Responding Party responds as follows:"
)

_DEFAULT_RESERVATION = (
    "Discovery and investigation are ongoing and Responding Party reserves "
    "the right to amend, modify and/or supplement this response in the "
    "future in the event that additional documents, facts and/or "
    "information are discovered, or their relevance becomes apparent."
)

_DEFAULT_PRELIMINARY_FI = (
    "These responses are made solely for the purpose of this action. "
    "Each response is subject to all appropriate objections, including "
    "competency, relevancy, materiality, propriety and admissibility, "
    "which would require the exclusion of any response set forth herein "
    "if the question were asked of, or any response were made by, a "
    "witness present and testifying in court. Additionally, each response "
    "is subject to all objections listed in the responses to the "
    "Interrogatories, which shall be incorporated herein by reference. "
    "All such objections are reserved and may be interposed at the time "
    "of trial.\n\n"
    "This Responding Party has not completed its investigation of the "
    "facts relating to this action, has not yet completed its discovery "
    "in this action, and has not yet completed preparation for trial. "
    "Consequently, the following responses are given without prejudice "
    "to this Party's right to allege and/or produce evidence of any "
    "subsequently-discovered facts or circumstances.\n\n"
    "Except for facts explicitly admitted herein, no admission of any "
    "nature is to be implied or inferred. The fact that any interrogatory "
    "herein has been answered should not be taken as an admission, or a "
    "confusion of the existence of any facts set forth or assumed by such "
    "interrogatory or that such response constitutes any fact thus set "
    "forth or assumed. All responses are given on the basis of a good "
    "faith effort to locate the requested information.\n\n"
    "This Party relies on well-established California authority to the "
    "effect that interrogatories cannot be unilaterally designated as "
    "continuing in nature, and serves notice that we will not voluntarily "
    "provide further responses to these interrogatories if additional "
    "information is acquired by us after these responses are served. "
    "Notwithstanding the above, this Responding Party reserves the right "
    "to change any and all responses herein as additional facts and "
    "further information is obtained, new analyses are made, and legal "
    "research is completed. The information contained herein is given in "
    "a good faith effort to supply as much factual material as is "
    "presently known by Responding Party, but should in no way prejudice "
    "this Responding Party's right to make new contentions or provide "
    "additional facts or additional information derived from further "
    "discovery, investigation, research and/or legal analysis. This "
    "preliminary statement shall apply to each and every response given "
    "herein, and shall be incorporated by reference as though fully set "
    "forth in all of the interrogatory responses appearing on the "
    "following pages."
)

_DEFAULT_PRELIMINARY_SI = (
    "These responses are made solely for the purpose of this action. "
    "Each response is subject to all appropriate objections, including, "
    "but not limited to, objections concerning competency, relevancy, "
    "materiality, propriety, and admissibility, which would require the "
    "exclusion of any statement contained herein if the interrogatories "
    "were asked of, or any statement contained herein were made by, a "
    "witness present and testifying in a court. All such objections and "
    "grounds therefore are reserved and may be interposed at the time of "
    "trial.\n\n"
    "This Responding Party has not completed its investigation of the "
    "facts relating to this action, has not yet completed discovery in "
    "this action, and has not yet completed preparation for trial. The "
    "following answers are therefore given without prejudice to this "
    "party's right to allege and/or produce evidence of any subsequently "
    "discovered facts or circumstances.\n\n"
    "Except for facts explicitly admitted herein, no admission of any "
    "nature is to be implied or inferred. The fact that any interrogatory "
    "herein has been answered should not be taken as an admission, or "
    "confession of the existence of any facts set forth or assumed by such "
    "interrogatory or that such response constitutes evidence of any fact "
    "thus set forth or assumed. All responses are given on the basis of a "
    "good faith effort to locate the requested information.\n\n"
    "This party relies on well-established California authority to the "
    "effect that interrogatories cannot unilaterally be designated "
    "continuing in nature and serve notice that we will not voluntarily "
    "provide further responses to these interrogatories if additional "
    "information is acquired by us after these responses are served.\n\n"
    "Notwithstanding the above, this Responding Party reserves the right "
    "to change any and all responses herein as additional facts and "
    "further information is obtained, new analyses are made, and legal "
    "research is completed. The information contained herein is given in "
    "a good faith effort to supply as much factual material as is "
    "presently known by Responding Party, but should in no way prejudice "
    "this Responding Party's right to make new contentions or provide "
    "additional facts or additional information derived from further "
    "discovery, investigation, research and/or legal analysis.\n\n"
    "This preliminary statement shall apply to each and every response "
    "given herein, and shall be incorporated by reference as though fully "
    "set forth in all of the demand responses appearing on the following "
    "pages."
)

_DEFAULT_PRELIMINARY_RFA = (
    "These responses are made solely for the purpose of this action. "
    "Each response is subject to all appropriate objections, including "
    "competency, relevancy, materiality, propriety and admissibility, "
    "which would require the exclusion of any response set forth herein "
    "if the question were asked of, or any response were made by, a "
    "witness present and testifying in court. All such objections are "
    "reserved and may be interposed at the time of trial.\n\n"
    "This Responding Party has not completed its investigation of the "
    "facts relating to this action, has not yet completed its discovery "
    "in this action, and has not yet completed preparation for trial. "
    "Consequently, the following responses are given without prejudice to "
    "this party's right to allege and/or produce evidence of any "
    "subsequently-discovered facts or circumstances.\n\n"
    "Except for facts explicitly admitted herein, no admission of any "
    "nature is to be implied or inferred. The fact that any demand or "
    "request herein has been answered should not be taken as an admission, "
    "or a confusion of the existence of any facts set forth or assumed by "
    "such demand or that such response constitutes any fact thus set forth "
    "or assumed. All responses are given on the basis of a good faith "
    "effort to locate the requested information.\n\n"
    "This party relies on well-established California authority to the "
    "effect that demands and requests cannot be unilaterally designated as "
    "continuing in nature, and serves notice that we will not voluntarily "
    "provide further responses if additional information is acquired by us "
    "after these responses are served.\n\n"
    "Notwithstanding the above, this Responding Party reserves the right "
    "to change any and all responses herein as additional facts and "
    "further information is obtained, new analyses are made, and legal "
    "research is completed.\n\n"
    "The information contained herein is given in a good faith effort to "
    "supply as much factual material as is presently known by Responding "
    "Party, but should in no way prejudice this responding party's right "
    "to make new contentions or provide additional facts or additional "
    "information derived from further discovery, investigation, research "
    "and/or legal analysis.\n\n"
    "This preliminary statement shall apply to each and every response "
    "given herein, and shall be incorporated by reference as though fully "
    "set forth in all of the demand responses appearing on the following "
    "pages."
)

_DEFAULT_PRELIMINARY_RPD = (
    'The responses set forth below represent the present knowledge of '
    '{responding_party} (hereinafter, "Responding Party") based on '
    "discovery, investigation, and case preparation to date. Responding "
    "Party has made reasonable efforts to respond to the Requests, as it "
    "understands and interprets each Request, and the contents hereof are "
    "based on the information obtained from these efforts. Responding "
    "Party will make a reasonable effort to gather information responsive "
    "to each Request as it understands and interprets each Request. If "
    "Propounding Party subsequently asserts a different interpretation, "
    "Responding Party reserves the right to supplement its objections "
    "and/or responses. Responding Party's investigation into this matter "
    "is, and will continue to be, ongoing. Responding Party may locate "
    "additional responsive information or documents at a later date, and "
    "it may assert appropriate objections to the use of the information or "
    "documents identified herein. Responding Party, therefore, expressly "
    "reserves the right to modify or supplement these responses and to "
    "rely on additional responsive information or documents, whether "
    "located in the course of its continuing investigation or in the "
    "course of discovery, at all future hearings and at trial, and the "
    "right to object on appropriate grounds to the use of any information "
    "or documents produced in response to the Requests."
)

_DEFAULT_INTRO_FI = (
    '{responding_party} ("{responding_short}" or "Responding Party") '
    "hereby responds and objects to the {set_word_title} Set of Form "
    "Interrogatories served by {propounding_party} "
    '("{propounding_short}" or "Propounding Party"), as follows:'
)

_DEFAULT_INTRO_SI = (
    '{responding_party} ("{responding_short}" or "Responding Party") '
    "hereby responds and objects to the {set_word_title} Set of Special "
    "Interrogatories served by {propounding_party} "
    '("{propounding_short}" or "Propounding Party"), as follows:'
)

_DEFAULT_INTRO_RFA = (
    '{responding_party} ("{responding_short}" or "Responding Party") '
    "hereby responds and objects to the {set_word_title} Set of Requests "
    "for Admission served by {propounding_party} "
    '("{propounding_short}" or "Propounding Party"), as follows:'
)

_DEFAULT_INTRO_RPD = (
    '{responding_party} ("{responding_short}" or "Responding Party") '
    "hereby responds and objects to the {set_word_title} Set of Request "
    "for Production of Documents served by {propounding_party} "
    '("{propounding_short}" or "Propounding Party"), as follows:'
)

_DEFAULT_GENERAL_OBJECTIONS_RFA = (
    "GENERAL OBJECTIONS INCORPORATED INTO EACH RESPONSE\n\n"
    "To the extent that any request may be construed as calling for "
    "information subject to a claim of privilege, including without "
    "limitation, the attorney-client and work product privilege, "
    "Responding Party claims such privilege and objects to such request "
    "on that basis.\n\n"
    "Responding Party will make reasonable efforts to respond, to the "
    "extent that a request has not been objected to, as the Responding "
    "Party understands and interprets each request. If the Propounding "
    "Party subsequently supplies an interpretation of the request which "
    "differs from that of the Responding Party, Responding Party reserves "
    "the right to supplement its objections and/or responses.\n\n"
    "Responding Party has not completed its investigation of the facts "
    "related to this case, has not completed discovery in this action, and "
    "is not completely prepared for trial. Further, all responses are on "
    "information and belief. Responding Party objects generally to any "
    "request with which the Propounding Party seeks to prejudice the "
    "Responding Party's right to produce evidence of any facts discovered "
    "subsequently to the preparation of responses."
)

_DEFAULT_GENERAL_OBJECTIONS_RPD = (
    "GENERAL OBJECTIONS AND OBJECTIONS TO DEFINITIONS\n\n"
    "1. Responding Party objects to each Request to the extent that it "
    "attempts or purports to call for the production of any information or "
    "document which would disclose Responding Party's trade secrets or "
    "other confidential research, development, or confidential information, "
    "which may be protected by the Uniform Trade Secrets Act enacted as "
    "Civil Code section 3426, et seq., a right of privacy under the United "
    "States Constitution or Article One of the Constitution of the State "
    "of California, or any other applicable law.\n\n"
    "2. Responding Party objects to each Request to the extent that it "
    "attempts or purports to call for the production of any information or "
    "document which is privileged, which was prepared in anticipation of "
    "litigation or for trial, which reveals communications between "
    "Responding Party and its legal counsel, which otherwise constitutes "
    "attorney work product, or which is otherwise privileged or immune "
    "from discovery.\n\n"
    "3. Responding Party objects to the Requests to the extent that they "
    "call for a legal conclusion.\n\n"
    "4. Responding Party objects to the Requests to the extent that they "
    "seek confidential customer information protected by industry "
    "regulation, statute, or law.\n\n"
    "5. Responding Party objects to these Requests to the extent that they "
    "are overly broad and unduly burdensome.\n\n"
    "6. Any and all information produced by Responding Party in response "
    "to the Requests is subject to all objections as to relevance, "
    "materiality, propriety, and admissibility, as well as to any and all "
    "other objections on any grounds that would require the exclusion of "
    "the information or any portion thereof if such information was offered "
    "in evidence, all of which objections and grounds are hereby expressly "
    "reserved and may be interposed at the time of any deposition or at or "
    "before any hearing or trial in this matter.\n\n"
    "7. Responding Party objects generally to the Requests on the ground "
    "that they are overly broad and unduly burdensome, to the extent that "
    "they attempt or purport to impose duties and obligations on Responding "
    "Party beyond the scope of permissible discovery by attempting or "
    "purporting to call for the production of documents that are beyond "
    "those that are in Responding Party's possession, custody or control.\n\n"
    "8. Responding Party objects generally to the Requests on the ground "
    "that they are vague and ambiguous.\n\n"
    "9. No incidental or implied admissions are intended by these responses "
    "including to the extent that Responding Party agrees to produce any "
    "documents described in the Requests.\n\n"
    "10. Responding Party objects to the Requests to the extent that they "
    "seek information with regard to third persons, which is personal and "
    "confidential and protected by the individual right to privacy.\n\n"
    "11. Responding Party objects generally to the \"Definitions\" in the "
    "Requests on the ground that they are vague and ambiguous, overbroad, "
    "oppressive, and purport to impose duties and obligations on Responding "
    "Party beyond the scope of permissible discovery."
)

_DEFAULT_VERIFICATION = (
    "VERIFICATION\n\n"
    "I have read the foregoing {document_title}, and know its contents.\n\n"
    "I am a representative to a party to this action. The matters stated "
    "in the foregoing document are true of my own knowledge except as to "
    "those matters which are stated on information and belief, and as to "
    "those matters I am informed and believe that they are true.\n\n"
    "I declare under penalty of perjury under the laws of the State of "
    "California that the foregoing is true and correct.\n\n"
    "Executed on this ___ day of ______________ 20__, "
    "at ______________________, California.\n\n\n"
    "_________________________________\n"
    "{verifier_name}"
)

# Fixed FI objections (same for every form interrogatory)
DEFAULT_FI_OBJECTIONS = (
    "Responding Party objects to this Interrogatory on the grounds that it "
    "is compound, calls for speculation, and is vague, ambiguous, uncertain "
    "and overbroad. Responding Party objects to this request to the extent "
    "this interrogatory violates the attorney-client or work-product privilege."
)

# Fixed FI 1.1 response
DEFAULT_FI_1_1_RESPONSE = (
    "Responding Party and its attorneys of record, {firm_name}, "
    "{firm_address}; {firm_phone}."
)

# Fixed FI 15.1 response
DEFAULT_FI_15_1_RESPONSE = (
    "Responding Party objects to this Interrogatory on the grounds that it "
    "is vague and ambiguous as to the term \"material.\" Responding Party "
    "further objects to the extent this Interrogatory invades the "
    "attorney-client privilege and work product doctrine. Responding Party "
    "further objects to this Interrogatory on the grounds that it calls for "
    "an expert opinion and a legal conclusion, and seeks the legal reasoning "
    "and theories of Responding Party's contentions. Responding Party is not "
    "required to prepare the Propounding Party's case. Discovery is "
    "continuing and Responding Party reserves the right to amend this "
    "response upon discovery of additional facts and information. Subject to "
    "and without waiving the foregoing objections, Responding Party responds "
    "as follows:\n"
    "A general denial is interposed as a matter of right based in part on "
    "California Code of Civil Procedure § 431.30. As to affirmative "
    "defenses, this interrogatory is premature at this time. Discovery and "
    "investigation are continuing so Responding Party reserves the right to "
    "change, modify or supplement these responses should additional "
    "information be ascertained."
)

# Fixed FI 16.x response
DEFAULT_FI_16_RESPONSE = (
    "Responding Party objects to this Interrogatory on the grounds that it "
    "is vague, ambiguous, uncertain, and overly broad. Responding Party "
    "further objects to this Interrogatory on the grounds that it calls for "
    "an expert opinion and a legal conclusion and seeks the legal reasoning "
    "and theories of Responding Party's contentions; Responding Party is not "
    "required to prepare the Propounding Party's case. Responding Party "
    "further objects on the grounds that the Interrogatory calls for "
    "speculation, seeks premature disclosure of expert opinion and violates "
    "the attorney work-product privilege. Subject to and without waiving the "
    "foregoing objections, Responding Party responds as follows:\n"
    "Pursuant to instruction 2(d) to the official form interrogatories, the "
    "interrogatories in section 16.0 should not be used until the defendant "
    "has had a reasonable opportunity to conduct an investigation or "
    "discovery into plaintiff's injuries and damages. At this time, "
    "responding party has yet to have an opportunity to depose the "
    "Plaintiffs, obtain an IME and obtain all of the Plaintiffs' medical "
    "records. As such, Responding Party is not in a position to fully "
    "answer this interrogatory at this time."
)


@dataclass
class ResponseRules:
    """Configuration for discovery response generation.
    
    Holds strategy toggles, editable boilerplate text, and free-form
    custom instructions. Serializes to/from JSON for persistence.
    """
    # --- Objection Strategy ---
    objection_aggressiveness: str = "aggressive"
    always_include_privacy_objection: bool = True
    always_include_privilege_objection: bool = True
    always_include_burden_objection: bool = True
    auto_flag_compound: bool = True
    auto_flag_broad_definitions: bool = True

    # --- Substantive Response Strategy ---
    si_response_style: str = "minimal"
    rfa_default_posture: str = "cautious"
    rpd_default_posture: str = "context_dependent"
    fi_17_1_auto_refresh: bool = False

    # --- Editable Boilerplate Text ---
    waiver_language: str = field(default_factory=lambda: _DEFAULT_WAIVER)
    reservation_clause: str = field(default_factory=lambda: _DEFAULT_RESERVATION)
    preliminary_statement_fi: str = field(default_factory=lambda: _DEFAULT_PRELIMINARY_FI)
    preliminary_statement_si: str = field(default_factory=lambda: _DEFAULT_PRELIMINARY_SI)
    preliminary_statement_rfa: str = field(default_factory=lambda: _DEFAULT_PRELIMINARY_RFA)
    preliminary_statement_rpd: str = field(default_factory=lambda: _DEFAULT_PRELIMINARY_RPD)
    intro_template_fi: str = field(default_factory=lambda: _DEFAULT_INTRO_FI)
    intro_template_si: str = field(default_factory=lambda: _DEFAULT_INTRO_SI)
    intro_template_rfa: str = field(default_factory=lambda: _DEFAULT_INTRO_RFA)
    intro_template_rpd: str = field(default_factory=lambda: _DEFAULT_INTRO_RPD)
    general_objections_rfa: str = field(default_factory=lambda: _DEFAULT_GENERAL_OBJECTIONS_RFA)
    general_objections_rpd: str = field(default_factory=lambda: _DEFAULT_GENERAL_OBJECTIONS_RPD)
    verification_template: str = field(default_factory=lambda: _DEFAULT_VERIFICATION)

    # --- Fixed FI Responses (not user-editable via dialog, but stored here) ---
    fi_objections: str = field(default_factory=lambda: DEFAULT_FI_OBJECTIONS)
    fi_1_1_response: str = field(default_factory=lambda: DEFAULT_FI_1_1_RESPONSE)
    fi_15_1_response: str = field(default_factory=lambda: DEFAULT_FI_15_1_RESPONSE)
    fi_16_response: str = field(default_factory=lambda: DEFAULT_FI_16_RESPONSE)

    # --- Custom Instructions ---
    custom_instructions: str = ""

    def to_dict(self) -> dict:
        """Serialize to a JSON-safe dictionary."""
        return asdict(self)

    @classmethod
    def from_dict(cls, data: dict) -> 'ResponseRules':
        """Deserialize from dict, using defaults for missing keys."""
        defaults = cls()
        merged = asdict(defaults)
        merged.update(data)
        # Filter to only valid field names
        valid_keys = {f.name for f in fields(cls)}
        filtered = {k: v for k, v in merged.items() if k in valid_keys}
        return cls(**filtered)

    def save_to_json(self, path: str) -> None:
        """Write rules to a JSON file."""
        os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
        with open(path, "w", encoding="utf-8") as f:
            json.dump(self.to_dict(), f, indent=2, ensure_ascii=False)

    @classmethod
    def load_from_json(cls, path: str) -> 'ResponseRules':
        """Load rules from a JSON file. Returns defaults if file missing."""
        if not os.path.isfile(path):
            return cls()
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        return cls.from_dict(data)
```

- [ ] **Step 4: Run tests to verify they pass**

```bash
python -m pytest tests/test_response_rules.py -v
```
Expected: All 8 tests PASS.

- [ ] **Step 5: Commit**

```bash
git add icharlotte_core/discovery/response_rules.py tests/test_response_rules.py
git commit -m "feat(discovery): add ResponseRules dataclass with boilerplate defaults"
```

---

## Task 2: Discovery Response Parser (Phase 1)

**Files:**
- Create: `icharlotte_core/discovery/response_parser.py`
- Create: `tests/test_response_parser.py`

Parses incoming discovery PDFs into structured `ParsedDiscovery` / `ParsedRequest` objects. Uses LLM for extraction and rule-based post-processing for compound detection and definition term extraction.

- [ ] **Step 1: Write tests for ParsedDiscovery/ParsedRequest models and parsing helpers**

```python
# tests/test_response_parser.py
"""Tests for the discovery response parser (Phase 1)."""
import unittest

from icharlotte_core.discovery.response_parser import (
    ParsedDiscovery,
    ParsedRequest,
    detect_discovery_type,
    detect_compound,
    extract_defined_terms,
    build_parse_prompt,
    parse_llm_response,
)
from icharlotte_core.discovery.models import DiscoveryType


class TestDetectDiscoveryType(unittest.TestCase):
    """Verify discovery type auto-detection from document text."""

    def test_detect_fi(self):
        text = "FORM INTERROGATORY NO. 1.1:\nState the name..."
        self.assertEqual(detect_discovery_type(text), "FI")

    def test_detect_si(self):
        text = "SPECIAL INTERROGATORY NO. 1:\nDescribe in detail..."
        self.assertEqual(detect_discovery_type(text), "SI")

    def test_detect_rfa(self):
        text = "REQUEST FOR ADMISSION NO. 1:\nAdmit that..."
        self.assertEqual(detect_discovery_type(text), "RFA")

    def test_detect_rpd(self):
        text = "REQUEST FOR PRODUCTION NO. 1:\nAll documents..."
        self.assertEqual(detect_discovery_type(text), "RPD")

    def test_detect_unknown_returns_none(self):
        text = "This is some random legal document text."
        self.assertIsNone(detect_discovery_type(text))


class TestDetectCompound(unittest.TestCase):
    """Verify compound question detection."""

    def test_simple_question_not_compound(self):
        text = "State the name of each witness."
        self.assertFalse(detect_compound(text))

    def test_and_conjunction_compound(self):
        text = "State all facts AND identify all documents."
        self.assertTrue(detect_compound(text))

    def test_multiple_subparts_compound(self):
        text = "State: (a) the name; (b) the address; (c) the phone number."
        # Subparts with (a)(b)(c) are standard form — not compound
        self.assertFalse(detect_compound(text))

    def test_multiple_action_verbs_compound(self):
        text = "Identify each person and describe the basis for your contention and state all facts supporting your claim."
        self.assertTrue(detect_compound(text))


class TestExtractDefinedTerms(unittest.TestCase):
    """Verify extraction of defined terms used in requests."""

    def test_all_caps_terms(self):
        text = "Describe the INCIDENT involving the VEHICLE."
        terms = extract_defined_terms(text)
        self.assertIn("INCIDENT", terms)
        self.assertIn("VEHICLE", terms)

    def test_ignores_short_caps(self):
        text = "State if YOU are A corporation."
        terms = extract_defined_terms(text)
        # "A" should be excluded (too short), "YOU" should be included
        self.assertIn("YOU", terms)
        self.assertNotIn("A", terms)

    def test_no_defined_terms(self):
        text = "State the name of the witness."
        terms = extract_defined_terms(text)
        self.assertEqual(terms, [])


class TestBuildParsePrompt(unittest.TestCase):
    """Verify LLM prompt construction."""

    def test_prompt_contains_instructions(self):
        prompt = build_parse_prompt("Some discovery text here")
        self.assertIn("discovery type", prompt.lower())
        self.assertIn("propounding party", prompt.lower())
        self.assertIn("JSON", prompt)

    def test_prompt_includes_document_text(self):
        prompt = build_parse_prompt("SPECIAL INTERROGATORY NO. 1: Describe...")
        self.assertIn("SPECIAL INTERROGATORY NO. 1", prompt)


class TestParseLlmResponse(unittest.TestCase):
    """Verify parsing of structured LLM JSON response."""

    def test_valid_json_response(self):
        llm_json = '''{
            "discovery_type": "SI",
            "propounding_party": "Plaintiff JOHN DOE",
            "set_number": 1,
            "case_number": "23STCV12345",
            "requests": [
                {"number": "1", "text": "Describe in detail how the INCIDENT occurred."},
                {"number": "2", "text": "State all facts supporting your contention AND identify all witnesses."}
            ]
        }'''
        parsed = parse_llm_response(llm_json, our_client_name="Defendant ACME Corp")
        self.assertEqual(parsed.discovery_type, "SI")
        self.assertEqual(parsed.propounding_party, "Plaintiff JOHN DOE")
        self.assertEqual(parsed.responding_party, "Defendant ACME Corp")
        self.assertEqual(parsed.set_number, 1)
        self.assertEqual(parsed.case_number, "23STCV12345")
        self.assertEqual(len(parsed.requests), 2)
        # Compound detection should flag request 2
        self.assertFalse(parsed.requests[0].is_compound)
        self.assertTrue(parsed.requests[1].is_compound)
        # Defined terms should be extracted
        self.assertIn("INCIDENT", parsed.requests[0].defined_terms_used)

    def test_malformed_json_raises(self):
        with self.assertRaises(ValueError):
            parse_llm_response("not valid json", our_client_name="Test")


if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 2: Run tests to verify they fail**

```bash
python -m pytest tests/test_response_parser.py -v
```
Expected: FAIL — `ModuleNotFoundError`

- [ ] **Step 3: Implement response_parser.py**

Create `icharlotte_core/discovery/response_parser.py`:

```python
"""
Phase 1: Discovery Response Parser.

Parses incoming discovery PDFs into structured ParsedDiscovery/ParsedRequest
objects. Uses LLM for extraction and rule-based post-processing for compound
detection and definition term extraction.
"""
import json
import re
from dataclasses import dataclass, field
from typing import List, Optional


# ---------------------------------------------------------------------------
# Data models
# ---------------------------------------------------------------------------

@dataclass
class ParsedRequest:
    """A single parsed discovery request."""
    number: str                          # "1.1" for FI, "1" for SI/RFA/RPD
    text: str                            # Full request text
    definitions: List[str] = field(default_factory=list)
    is_compound: bool = False
    defined_terms_used: List[str] = field(default_factory=list)


@dataclass
class ParsedDiscovery:
    """Structured result of parsing an incoming discovery document."""
    discovery_type: str                  # "FI", "SI", "RFA", "RPD"
    propounding_party: str
    responding_party: str
    set_number: int
    set_word: str                        # "ONE", "TWO", etc.
    case_number: str
    requests: List[ParsedRequest] = field(default_factory=list)


# ---------------------------------------------------------------------------
# Discovery type detection
# ---------------------------------------------------------------------------

_TYPE_PATTERNS = {
    "FI": re.compile(r"FORM\s+INTERROGATOR", re.IGNORECASE),
    "SI": re.compile(r"SPECIAL\s+INTERROGATOR", re.IGNORECASE),
    "RFA": re.compile(r"REQUEST\s+FOR\s+ADMISSION", re.IGNORECASE),
    "RPD": re.compile(r"REQUEST\s+FOR\s+PRODUCTION", re.IGNORECASE),
}


def detect_discovery_type(text: str) -> Optional[str]:
    """Auto-detect discovery type from document text. Returns 'FI'/'SI'/'RFA'/'RPD' or None."""
    for dtype, pattern in _TYPE_PATTERNS.items():
        if pattern.search(text):
            return dtype
    return None


# ---------------------------------------------------------------------------
# Compound question detection
# ---------------------------------------------------------------------------

# Patterns indicating a compound question (multiple independent action verbs
# joined by AND, or multiple "state...and...identify...and...describe" chains)
_COMPOUND_PATTERN = re.compile(
    r'\b(state|identify|describe|list|explain|set forth)\b.*?\bAND\b.*?'
    r'\b(state|identify|describe|list|explain|set forth)\b',
    re.IGNORECASE | re.DOTALL,
)


def detect_compound(text: str) -> bool:
    """Return True if the request text appears to be a compound question.
    
    Standard subparts like (a), (b), (c) are NOT considered compound —
    those are a single question with structured response format.
    Compound means multiple independent questions joined by AND.
    """
    return bool(_COMPOUND_PATTERN.search(text))


# ---------------------------------------------------------------------------
# Defined term extraction
# ---------------------------------------------------------------------------

# Common all-caps words that are NOT defined terms
_CAPS_STOPWORDS = frozenset({
    "A", "AN", "AND", "ARE", "AS", "AT", "BE", "BY", "DO", "FOR", "FROM",
    "HAS", "HIS", "HER", "IF", "IN", "IS", "IT", "NO", "NOT", "OF", "ON",
    "OR", "SO", "THE", "TO", "WAS", "SET", "ONE", "TWO", "THREE",
    "INTERROGATORY", "INTERROGATORIES", "REQUEST", "ADMISSION", "PRODUCTION",
    "RESPONSE", "SPECIAL", "FORM", "ADMIT", "DENY", "STATE", "DESCRIBE",
    "IDENTIFY", "ALL", "EACH", "ANY", "THAT", "THIS", "WHICH", "WHAT",
    "WHEN", "WHERE", "WHO", "HOW", "DOES", "DID", "WERE",
})


def extract_defined_terms(text: str) -> List[str]:
    """Extract ALL-CAPS defined terms (3+ chars) from request text.
    
    Defined terms in discovery are typically written in ALL CAPS
    (e.g., INCIDENT, VEHICLE, PLAINTIFF, DOCUMENTS, PERSON).
    """
    words = re.findall(r'\b([A-Z]{3,})\b', text)
    seen = set()
    result = []
    for w in words:
        if w not in _CAPS_STOPWORDS and w not in seen:
            seen.add(w)
            result.append(w)
    return result


# ---------------------------------------------------------------------------
# LLM prompt construction
# ---------------------------------------------------------------------------

def build_parse_prompt(document_text: str) -> str:
    """Build the LLM prompt for Phase 1 parsing.
    
    The prompt instructs the LLM to extract structured data from the
    discovery document text and return it as JSON.
    """
    return f"""You are a legal document parser. Extract structured data from this discovery document.

Return a JSON object with these fields:
- "discovery_type": one of "FI" (Form Interrogatories), "SI" (Special Interrogatories), "RFA" (Requests for Admission), "RPD" (Requests for Production)
- "propounding_party": the full name of the propounding party (e.g., "Plaintiff JOHN DOE")
- "set_number": integer (1, 2, etc.)
- "case_number": the case number (e.g., "23STCV12345")
- "requests": array of objects, each with:
  - "number": string (e.g., "1.1" for form interrogatories, "1" for others)
  - "text": the full text of the request/interrogatory
  - "definitions": array of any inline definition footnotes associated with this request

Return ONLY the JSON object, no other text. If a field cannot be determined, use null.

DOCUMENT TEXT:
{document_text}"""


# ---------------------------------------------------------------------------
# LLM response parsing
# ---------------------------------------------------------------------------

_SET_WORDS = {
    1: "ONE", 2: "TWO", 3: "THREE", 4: "FOUR", 5: "FIVE",
    6: "SIX", 7: "SEVEN", 8: "EIGHT", 9: "NINE", 10: "TEN",
}


def parse_llm_response(
    llm_json: str,
    our_client_name: str,
) -> ParsedDiscovery:
    """Parse the LLM's JSON response into a ParsedDiscovery object.
    
    Also runs post-processing: compound detection and defined term extraction.
    
    Raises ValueError if the JSON is malformed.
    """
    # Strip markdown fences if present
    text = llm_json.strip()
    if text.startswith("```"):
        text = re.sub(r'^```(?:json)?\s*', '', text)
        text = re.sub(r'\s*```$', '', text)

    try:
        data = json.loads(text)
    except json.JSONDecodeError as e:
        raise ValueError(f"Failed to parse LLM response as JSON: {e}") from e

    set_number = int(data.get("set_number", 1) or 1)
    set_word = _SET_WORDS.get(set_number, str(set_number))

    requests = []
    for req_data in data.get("requests", []):
        req_text = req_data.get("text", "")
        req = ParsedRequest(
            number=str(req_data.get("number", "")),
            text=req_text,
            definitions=req_data.get("definitions", []) or [],
            is_compound=detect_compound(req_text),
            defined_terms_used=extract_defined_terms(req_text),
        )
        requests.append(req)

    return ParsedDiscovery(
        discovery_type=data.get("discovery_type", ""),
        propounding_party=data.get("propounding_party", ""),
        responding_party=our_client_name,
        set_number=set_number,
        set_word=set_word,
        case_number=data.get("case_number", ""),
        requests=requests,
    )
```

- [ ] **Step 4: Run tests to verify they pass**

```bash
python -m pytest tests/test_response_parser.py -v
```
Expected: All tests PASS.

- [ ] **Step 5: Commit**

```bash
git add icharlotte_core/discovery/response_parser.py tests/test_response_parser.py
git commit -m "feat(discovery): add Phase 1 response parser with compound/term detection"
```

---

## Task 3: Objection Selector (Phase 2)

**Files:**
- Create: `icharlotte_core/discovery/objection_selector.py`
- Create: `tests/test_objection_selector.py`

Rule-based + LLM hybrid objection selection. Form interrogatories use fixed objections. SI/RFA/RPD use rule-based pre-selection merged with LLM selection.

- [ ] **Step 1: Write tests for objection selector**

```python
# tests/test_objection_selector.py
"""Tests for the objection selector (Phase 2)."""
import unittest

from icharlotte_core.discovery.objection_selector import (
    ObjectionMenu,
    select_fi_objections,
    rule_based_preselect,
    build_objection_prompt,
    parse_objection_response,
    merge_objections,
)
from icharlotte_core.discovery.response_parser import ParsedRequest
from icharlotte_core.discovery.response_rules import ResponseRules


class TestObjectionMenu(unittest.TestCase):
    """Verify ObjectionMenu loading and lookup."""

    def test_default_menu_has_objections(self):
        menu = ObjectionMenu.load_defaults()
        self.assertGreaterEqual(len(menu.objections), 10)

    def test_get_by_id(self):
        menu = ObjectionMenu.load_defaults()
        obj = menu.get(1)
        self.assertIn("speculation", obj.lower())
        self.assertIn("vague", obj.lower())


class TestSelectFiObjections(unittest.TestCase):
    """FI objections are always the same fixed text."""

    def test_returns_fixed_objections(self):
        rules = ResponseRules()
        objections = select_fi_objections(rules)
        self.assertIn("compound", objections.lower())
        self.assertIn("vague", objections.lower())


class TestRuleBasedPreselect(unittest.TestCase):
    """Verify rule-based pre-selection flags."""

    def test_compound_flag(self):
        req = ParsedRequest(
            number="1", text="State all facts AND identify all documents.",
            is_compound=True, defined_terms_used=[],
        )
        rules = ResponseRules()
        menu = ObjectionMenu.load_defaults()
        ids = rule_based_preselect(req, rules, menu)
        # Should include compound objection (ID 9)
        self.assertIn(9, ids)

    def test_privacy_always_included(self):
        req = ParsedRequest(
            number="1", text="State your income.",
            is_compound=False, defined_terms_used=[],
        )
        rules = ResponseRules(always_include_privacy_objection=True)
        menu = ObjectionMenu.load_defaults()
        ids = rule_based_preselect(req, rules, menu)
        self.assertIn(2, ids)  # Privacy objection is ID 2

    def test_broad_definition_flag(self):
        req = ParsedRequest(
            number="1", text="Produce all DOCUMENTS.",
            is_compound=False, defined_terms_used=["DOCUMENTS"],
        )
        rules = ResponseRules(auto_flag_broad_definitions=True)
        menu = ObjectionMenu.load_defaults()
        ids = rule_based_preselect(req, rules, menu)
        # Should include definitional objection (ID 10 or 11)
        self.assertTrue(ids & {10, 11})


class TestMergeObjections(unittest.TestCase):
    """Verify union merge of rule-based and LLM-selected objections."""

    def test_union(self):
        rule_ids = {1, 2, 3}
        llm_ids = {2, 4, 5}
        merged = merge_objections(rule_ids, llm_ids)
        self.assertEqual(merged, {1, 2, 3, 4, 5})


class TestBuildObjectionPrompt(unittest.TestCase):
    """Verify LLM prompt for objection selection."""

    def test_prompt_contains_request_and_menu(self):
        menu = ObjectionMenu.load_defaults()
        prompt = build_objection_prompt(
            "State all facts about the INCIDENT.",
            menu,
            "aggressive",
        )
        self.assertIn("INCIDENT", prompt)
        self.assertIn("speculation", prompt.lower())


class TestParseObjectionResponse(unittest.TestCase):
    """Verify parsing LLM objection selection response."""

    def test_parse_id_list(self):
        response = "1, 2, 4, 6"
        ids = parse_objection_response(response)
        self.assertEqual(ids, {1, 2, 4, 6})

    def test_parse_json_array(self):
        response = "[1, 3, 5]"
        ids = parse_objection_response(response)
        self.assertEqual(ids, {1, 3, 5})


if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 2: Run tests to verify they fail**

```bash
python -m pytest tests/test_objection_selector.py -v
```
Expected: FAIL — `ModuleNotFoundError`

- [ ] **Step 3: Implement objection_selector.py**

Create `icharlotte_core/discovery/objection_selector.py` with:

- `ObjectionMenu` class with `load_defaults()` (hardcoded from DISCOVERY OBJECTIONS.docx), `load_from_docx(path)` for reload, `get(id)`, `all_text()`, `objections` dict.
- `select_fi_objections(rules)` — returns fixed FI objection text from `rules.fi_objections`.
- `rule_based_preselect(request, rules, menu)` — returns `Set[int]` of objection IDs based on flags: `is_compound` → 9, `always_include_privacy` → 2, `always_include_privilege` → 4, `always_include_burden` → 6, `auto_flag_broad_definitions` + `defined_terms_used` → 10/11, expert keywords → 3.
- `build_objection_prompt(request_text, menu, aggressiveness)` — builds LLM prompt listing all objections by ID and asking LLM to select applicable ones. Aggressiveness level modulates instructions.
- `parse_objection_response(llm_text)` — parses comma-separated IDs or JSON array from LLM response into `Set[int]`.
- `merge_objections(rule_ids, llm_ids)` — union of both sets.
- `format_objections(objection_ids, menu)` — joins selected objection texts into a single formatted string.

The 12 default objections are hardcoded as a dict mapping ID to full text (from DISCOVERY OBJECTIONS.docx):
```python
_DEFAULT_OBJECTIONS = {
    1: "Responding Party objects to this Request on the grounds that it calls for speculation and is vague, ambiguous, uncertain and overbroad.",
    2: "Responding Party objects to this Request on the grounds that it is not relevant and not reasonably calculated to lead to the discovery of admissible evidence and seeks to invade Responding Party's privacy.",
    3: "Responding Party further objects to this Request on the grounds that it seeks premature disclosure of expert opinion and/or a legal conclusion.",
    4: "Responding Party further objects to this Request on the grounds that it seeks to invade attorney client privilege and/or violates the attorney work-product privilege.",
    5: "Responding Party further objects to this Request on the grounds that it, as phrased, is argumentative and requires the adoption of an assumption, which is improper; the question assumes facts which may or may not be true, but the form of the question requires that the answer adopt the assumption.",
    6: "Responding Party objects to this Request on the grounds that it is unduly burdensome and so overly broad and unlimited as to time and scope as to be an unwarranted annoyance, embarrassment, and is oppressive; to comply with the Request would be an undue burden and expense on Responding Party and is calculated to annoy and harass Responding Party.",
    7: "Responding Party objects to this Interrogatory on the grounds that it is burdensome and harassing in that it calls for production of a list or summary where no such list or summary presently exists.",
    8: "Responding Party objects to this Interrogatory on the grounds that it has, in substance, been previously propounded. Continuous discovery into the same matter is oppressive, harassing, burdensome, and contrary to the legislative intent of the Discovery Act.",
    9: "Responding Party objects to this Interrogatory on the grounds that it is compound in form.",
    10: 'Responding Party specifically objects to this Request on the grounds that the term "{term}" is undefined and therefore vague, ambiguous, uncertain, confusing, unintelligible and overbroad.',
    11: 'Responding Party specifically objects to Propounding Party\'s definition of "{term}" on the grounds that it is vague, ambiguous, uncertain, confusing, unintelligible and overbroad, and is argumentative and requires the adoption of an assumption, which is improper.',
    12: "Responding Party objects to this Interrogatory on the grounds that the information and/or documents sought are equally and/or more readily available to Propounding Party.",
}
```

- [ ] **Step 4: Run tests to verify they pass**

```bash
python -m pytest tests/test_objection_selector.py -v
```
Expected: All tests PASS.

- [ ] **Step 5: Commit**

```bash
git add icharlotte_core/discovery/objection_selector.py tests/test_objection_selector.py
git commit -m "feat(discovery): add Phase 2 objection selector with rule-based + LLM hybrid"
```

---

## Task 4: Response Drafter (Phase 3)

**Files:**
- Create: `icharlotte_core/discovery/response_drafter.py`
- Create: `tests/test_response_drafter.py`

Assembles the complete plain-text response document by combining parsed requests, selected objections, and substantive responses (fixed templates for FI, LLM-drafted for SI/RFA/RPD).

- [ ] **Step 1: Write tests for response drafter**

```python
# tests/test_response_drafter.py
"""Tests for the response drafter (Phase 3)."""
import unittest

from icharlotte_core.discovery.response_drafter import (
    format_single_response,
    is_fi_series,
    get_fi_fixed_response,
    build_si_prompt,
    build_rfa_prompt,
    build_rpd_prompt,
    build_fi_17_1_prompt,
    assemble_plain_text,
)
from icharlotte_core.discovery.response_parser import ParsedDiscovery, ParsedRequest
from icharlotte_core.discovery.response_rules import ResponseRules


class TestFormatSingleResponse(unittest.TestCase):
    """Verify single response formatting."""

    def test_fi_format(self):
        result = format_single_response(
            disc_type="FI",
            request_number="1.1",
            request_text="State the name...",
            objections="Objection text here.",
            substantive="Attorney info here.",
            waiver="Subject to and without waiving...",
            reservation="Discovery and investigation...",
        )
        self.assertIn("FORM INTERROGATORY NO. 1.1:", result)
        self.assertIn("RESPONSE TO FORM INTERROGATORY NO. 1.1:", result)
        self.assertIn("Objection text here.", result)
        self.assertIn("Subject to and without waiving...", result)
        self.assertIn("Attorney info here.", result)
        self.assertIn("Discovery and investigation...", result)

    def test_si_format(self):
        result = format_single_response(
            disc_type="SI",
            request_number="5",
            request_text="Describe the incident.",
            objections="Objection text.",
            substantive="Substantive response.",
            waiver="Subject to...",
            reservation="Discovery...",
        )
        self.assertIn("SPECIAL INTERROGATORY NO. 5:", result)
        self.assertIn("RESPONSE TO SPECIAL INTERROGATORY NO. 5:", result)


class TestIsFiSeries(unittest.TestCase):
    """Verify form interrogatory series detection."""

    def test_series_3(self):
        self.assertTrue(is_fi_series("3.1", [3]))
        self.assertTrue(is_fi_series("3.2", [3]))
        self.assertFalse(is_fi_series("3.1", [4]))

    def test_series_16(self):
        self.assertTrue(is_fi_series("16.1", [16]))
        self.assertTrue(is_fi_series("16.9", [16]))

    def test_non_series(self):
        self.assertFalse(is_fi_series("6.1", [3, 4, 7, 12, 13, 14, 20]))


class TestGetFiFixedResponse(unittest.TestCase):
    """Verify fixed FI response lookup."""

    def test_1_1_is_fixed(self):
        rules = ResponseRules()
        resp = get_fi_fixed_response("1.1", rules)
        self.assertIsNotNone(resp)
        self.assertIn("attorneys of record", resp.lower())

    def test_15_1_is_fixed(self):
        rules = ResponseRules()
        resp = get_fi_fixed_response("15.1", rules)
        self.assertIsNotNone(resp)
        self.assertIn("general denial", resp.lower())

    def test_16_x_is_fixed(self):
        rules = ResponseRules()
        resp = get_fi_fixed_response("16.1", rules)
        self.assertIsNotNone(resp)
        self.assertIn("section 16.0", resp.lower())

    def test_17_1_is_placeholder(self):
        rules = ResponseRules()
        resp = get_fi_fixed_response("17.1", rules)
        self.assertIsNotNone(resp)
        self.assertIn("PENDING", resp)

    def test_non_fixed_returns_none(self):
        rules = ResponseRules()
        resp = get_fi_fixed_response("6.1", rules)
        self.assertIsNone(resp)

    def test_not_applicable_for_irrelevant_fi(self):
        """FI numbers not applicable to the case type return 'Not applicable'."""
        rules = ResponseRules()
        # is_applicable is determined at draft time based on case type,
        # so get_fi_fixed_response returns None for these — the drafter
        # handles N/A detection separately via detect_inapplicable_fi()
        from icharlotte_core.discovery.response_drafter import detect_inapplicable_fi
        # Wrongful death FIs (e.g., 10.x series) in a non-wrongful-death case
        self.assertTrue(detect_inapplicable_fi("10.1", case_type="negligence"))


class TestBuildPrompts(unittest.TestCase):
    """Verify LLM prompt construction for each discovery type."""

    def test_si_prompt_minimal(self):
        rules = ResponseRules(si_response_style="minimal")
        prompt = build_si_prompt(
            "Describe the incident.",
            "Context: accident at intersection.",
            rules,
        )
        self.assertIn("narrowly", prompt.lower())
        self.assertIn("few words", prompt.lower())

    def test_rfa_prompt_cautious(self):
        rules = ResponseRules(rfa_default_posture="cautious")
        prompt = build_rfa_prompt(
            "Admit that defendant was negligent.",
            "Context: disputed liability.",
            rules,
        )
        self.assertIn("Admit", prompt)
        self.assertIn("Deny", prompt)
        self.assertIn("insufficient", prompt.lower())

    def test_rpd_prompt(self):
        rules = ResponseRules(rpd_default_posture="context_dependent")
        prompt = build_rpd_prompt(
            "All documents relating to training.",
            "Context: employee records available.",
            rules,
        )
        self.assertIn("unable to comply", prompt.lower())
        self.assertIn("will comply", prompt.lower())


class TestAssemblePlainText(unittest.TestCase):
    """Verify full plain-text document assembly."""

    def test_includes_header_and_preliminary(self):
        parsed = ParsedDiscovery(
            discovery_type="SI",
            propounding_party="Plaintiff JOHN DOE",
            responding_party="Defendant ACME CORP",
            set_number=1, set_word="ONE",
            case_number="23STCV12345",
            requests=[
                ParsedRequest(number="1", text="Describe the incident."),
            ],
        )
        rules = ResponseRules()
        objections_map = {"1": "Objection text."}
        responses_map = {"1": "Substantive response."}

        text = assemble_plain_text(parsed, rules, objections_map, responses_map)
        self.assertIn("PROPOUNDING PARTY:", text)
        self.assertIn("RESPONDING PARTY:", text)
        self.assertIn("SET NO.:", text)
        self.assertIn("PRELIMINARY STATEMENT", text)
        self.assertIn("SPECIAL INTERROGATORY NO. 1:", text)
        self.assertIn("RESPONSE TO SPECIAL INTERROGATORY NO. 1:", text)


if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 2: Run tests to verify they fail**

```bash
python -m pytest tests/test_response_drafter.py -v
```
Expected: FAIL — `ModuleNotFoundError`

- [ ] **Step 3: Implement response_drafter.py**

Create `icharlotte_core/discovery/response_drafter.py` with:

- Header/label constants mapping discovery type to header templates:
  ```python
  _HEADERS = {
      "FI": ("FORM INTERROGATORY NO. {n}:", "RESPONSE TO FORM INTERROGATORY NO. {n}:"),
      "SI": ("SPECIAL INTERROGATORY NO. {n}:", "RESPONSE TO SPECIAL INTERROGATORY NO. {n}:"),
      "RFA": ("REQUEST FOR ADMISSION NO. {n}:", "RESPONSE TO REQUEST FOR ADMISSION NO. {n}:"),
      "RPD": ("REQUEST FOR PRODUCTION NO. {n}:", "RESPONSE TO REQUEST NO. {n}:"),
  }
  ```
- `format_single_response(disc_type, request_number, request_text, objections, substantive, waiver, reservation)` — formats one request+response block using the template structure.
- `is_fi_series(number_str, series_list)` — checks if FI number belongs to a given series (e.g., `is_fi_series("3.1", [3])` → True).
- `get_fi_fixed_response(number_str, rules)` — returns fixed text for FI 1.1, 15.1, 16.x, 17.1 (placeholder), or None for LLM-drafted FI.
- `build_si_prompt(request_text, context_text, rules)` — builds SI response prompt with style-specific instructions from `rules.si_response_style`.
- `build_rfa_prompt(request_text, context_text, rules)` — builds RFA response prompt with posture instructions from `rules.rfa_default_posture`. Lists the three allowed responses verbatim.
- `build_rpd_prompt(request_text, context_text, rules)` — builds RPD response prompt with posture instructions from `rules.rpd_default_posture`. Lists the two allowed responses verbatim.
- `detect_inapplicable_fi(fi_number, case_type)` — returns True if the FI number is not applicable to the given case type (e.g., wrongful death FIs 10.x in a negligence case). Used by the drafter to auto-respond "Not applicable."
- `build_fi_17_1_prompt(rfa_responses_text, rules)` — builds prompt to generate FI 17.1 based on completed RFA responses.
- `assemble_plain_text(parsed, rules, objections_map, responses_map)` — assembles the full plain-text response document:
  1. Party identification block
  2. "TO" line
  3. Introduction paragraph (from rules intro template with placeholders filled)
  4. Preliminary statement
  5. General objections (RFA/RPD only)
  6. Individual responses (using `format_single_response` for each)

- [ ] **Step 4: Run tests to verify they pass**

```bash
python -m pytest tests/test_response_drafter.py -v
```
Expected: All tests PASS.

- [ ] **Step 5: Commit**

```bash
git add icharlotte_core/discovery/response_drafter.py tests/test_response_drafter.py
git commit -m "feat(discovery): add Phase 3 response drafter with FI templates and LLM prompts"
```

---

## Task 5: Response Assembler (Word Document)

**Files:**
- Create: `icharlotte_core/discovery/response_assembler.py`
- Create: `tests/test_response_assembler.py`

Assembles final .docx from caption template + response text. Reuses existing style constants and `find_caption_page()` from `assembler.py`.

- [ ] **Step 1: Write tests for response assembler**

```python
# tests/test_response_assembler.py
"""Tests for the response assembler (Word document generation)."""
import os
import tempfile
import unittest

from icharlotte_core.discovery.response_assembler import (
    ResponseAssembler,
    build_response_title,
    build_response_filename,
    extract_signature_block,
    parse_response_text,
)
from icharlotte_core.discovery.response_parser import ParsedDiscovery
from icharlotte_core.discovery.response_rules import ResponseRules


class TestBuildResponseTitle(unittest.TestCase):
    """Verify document title construction."""

    def test_fi_title(self):
        title = build_response_title(
            responding_name="DEFENDANT USA WASTE OF CALIFORNIA, INC.",
            propounding_role="Plaintiff",
            disc_type="FI",
            set_word="ONE",
        )
        self.assertIn("DEFENDANT USA WASTE", title)
        self.assertIn("FORM INTERROGATORIES", title)
        self.assertIn("SET ONE", title)
        self.assertIn("PLAINTIFF'S", title.upper())

    def test_rfa_title(self):
        title = build_response_title(
            responding_name="DEFENDANT ACME CORP",
            propounding_role="Plaintiff",
            disc_type="RFA",
            set_word="TWO",
        )
        self.assertIn("REQUESTS FOR ADMISSION", title)
        self.assertIn("SET TWO", title)


class TestBuildResponseFilename(unittest.TestCase):
    """Verify output filename convention."""

    def test_fi_filename(self):
        fn = build_response_filename("USA Waste", "FI", 1)
        self.assertEqual(fn, "Def USA Waste's Resp to FI(1).docx")

    def test_rpd_filename(self):
        fn = build_response_filename("Acme", "RPD", 2)
        self.assertEqual(fn, "Def Acme's Resp to RPD(2).docx")


class TestParseResponseText(unittest.TestCase):
    """Verify re-parsing of editor plain text into request/response pairs."""

    def test_parse_si_responses(self):
        text = (
            "SPECIAL INTERROGATORY NO. 1:\n"
            "Describe the incident.\n\n"
            "RESPONSE TO SPECIAL INTERROGATORY NO. 1:\n"
            "Objection. Subject to objections, response here.\n\n"
            "SPECIAL INTERROGATORY NO. 2:\n"
            "State all facts.\n\n"
            "RESPONSE TO SPECIAL INTERROGATORY NO. 2:\n"
            "Objection. Response two.\n"
        )
        pairs = parse_response_text(text, "SI")
        self.assertEqual(len(pairs), 2)
        self.assertEqual(pairs[0]["number"], "1")
        self.assertIn("Describe the incident", pairs[0]["request"])
        self.assertIn("response here", pairs[0]["response"])


if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 2: Run tests to verify they fail**

```bash
python -m pytest tests/test_response_assembler.py -v
```
Expected: FAIL — `ModuleNotFoundError`

- [ ] **Step 3: Implement response_assembler.py**

Create `icharlotte_core/discovery/response_assembler.py` with:

- Imports from existing `assembler.py`: style constants (`STYLE_BODY_DOUBLE`, `STYLE_FLUSH_LEFT_DOUBLE`, `STYLE_DISCOVERY_NO`, `STYLE_FLUSH_LEFT`, `SIG_LEFT_INDENT`, `SIG_NAME_LEFT_INDENT`), `_add_para`, `_safe_style`, `find_caption_page` (via `DiscoveryAssembler.find_caption_page`).
- `build_response_title(responding_name, propounding_role, disc_type, set_word)` — constructs title like "DEFENDANT USA WASTE...'S RESPONSES TO PLAINTIFF'S FORM INTERROGATORIES, SET ONE".
- `build_response_filename(abbreviation, disc_type, set_number)` — returns filename like "Def USA Waste's Resp to FI(1).docx".
- `extract_signature_block(doc)` — scans document paragraphs for signature block markers ("Respectfully submitted", "Dated:", firm name patterns). Returns list of extracted paragraphs + their formatting, removes them from doc.
- `parse_response_text(text, disc_type)` — re-parses editor plain text using regex to split on request/response headers. Returns list of dicts `{number, request, response}`.
- `class ResponseAssembler`:
  - `__init__(caption_page_path)` — validates caption exists.
  - `assemble(parsed, response_text, rules, output_path)` — full assembly: load template → replace CAPTION PAGE → extract signature → insert party block → insert intro → insert preliminary → insert general objections → parse and insert individual responses → insert verification → re-insert signature → set footer → validate → save.
  - Uses existing `_set_caption_title` pattern from `DiscoveryAssembler` (searches XML for "CAPTION PAGE" text elements).
  - Reuses style application from existing assembler: `_add_para(doc, text, style_name)`.

- [ ] **Step 4: Run tests to verify they pass**

```bash
python -m pytest tests/test_response_assembler.py -v
```
Expected: All tests PASS.

- [ ] **Step 5: Commit**

```bash
git add icharlotte_core/discovery/response_assembler.py tests/test_response_assembler.py
git commit -m "feat(discovery): add response assembler for Word document generation"
```

---

## Task 6: ResponseRulesDialog UI

**Files:**
- Create: `icharlotte_core/ui/response_rules_dialog.py`
- Create: `tests/test_response_rules_dialog.py`

Three-tab dialog: Strategy toggles, Boilerplate editors, Custom Instructions.

- [ ] **Step 1: Write test for dialog construction and data flow**

```python
# tests/test_response_rules_dialog.py
"""Tests for the ResponseRulesDialog UI."""
import os
import sys
import unittest
from unittest.mock import patch, MagicMock

os.environ["QT_QPA_PLATFORM"] = "offscreen"

from PySide6.QtWidgets import QApplication
app = QApplication.instance() or QApplication(sys.argv)

from icharlotte_core.ui.response_rules_dialog import ResponseRulesDialog
from icharlotte_core.discovery.response_rules import ResponseRules


class TestResponseRulesDialog(unittest.TestCase):

    def test_dialog_creates(self):
        rules = ResponseRules()
        dlg = ResponseRulesDialog(rules)
        self.assertIsNotNone(dlg)

    def test_dialog_loads_rules(self):
        rules = ResponseRules(objection_aggressiveness="conservative")
        dlg = ResponseRulesDialog(rules)
        self.assertTrue(dlg.rb_conservative.isChecked())

    def test_get_rules_returns_modified(self):
        rules = ResponseRules()
        dlg = ResponseRulesDialog(rules)
        dlg.rb_moderate_obj.setChecked(True)
        dlg.custom_instructions_edit.setPlainText("Be extra cautious")
        result = dlg.get_rules()
        self.assertEqual(result.objection_aggressiveness, "moderate")
        self.assertEqual(result.custom_instructions, "Be extra cautious")

    def test_reset_defaults(self):
        rules = ResponseRules(objection_aggressiveness="conservative", custom_instructions="test")
        dlg = ResponseRulesDialog(rules)
        dlg._on_reset_defaults()
        result = dlg.get_rules()
        self.assertEqual(result.objection_aggressiveness, "aggressive")
        self.assertEqual(result.custom_instructions, "")

    def test_three_tabs_exist(self):
        rules = ResponseRules()
        dlg = ResponseRulesDialog(rules)
        self.assertEqual(dlg.tab_widget.count(), 3)
        self.assertEqual(dlg.tab_widget.tabText(0), "Strategy")
        self.assertEqual(dlg.tab_widget.tabText(1), "Boilerplate")
        self.assertEqual(dlg.tab_widget.tabText(2), "Custom Instructions")


if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 2: Run tests to verify they fail**

```bash
python -m pytest tests/test_response_rules_dialog.py -v
```
Expected: FAIL

- [ ] **Step 3: Implement ResponseRulesDialog**

Create `icharlotte_core/ui/response_rules_dialog.py`:

- `class ResponseRulesDialog(QDialog)`:
  - `__init__(rules: ResponseRules, parent=None)` — stores initial rules, builds 3-tab UI.
  - **Tab 1 (Strategy):** QGroupBox per category:
    - "Objection Aggressiveness": 3 radio buttons (`rb_aggressive`, `rb_moderate_obj`, `rb_conservative`)
    - "Always Include" checkboxes: `cb_privacy`, `cb_privilege`, `cb_burden`
    - "Auto-Detection" checkboxes: `cb_compound`, `cb_broad_defs`
    - "SI Response Style": 3 radio buttons (`rb_si_minimal`, `rb_si_moderate`, `rb_si_detailed`)
    - "RFA Default Posture": 3 radio buttons (`rb_rfa_cautious`, `rb_rfa_balanced`, `rb_rfa_cooperative`)
    - "RPD Default Posture": 3 radio buttons (`rb_rpd_unable`, `rb_rpd_comply`, `rb_rpd_context`)
  - **Tab 2 (Boilerplate):** QScrollArea containing labeled QTextEdit fields for each boilerplate block. Each field has a small "Reset" QPushButton that restores the default text for that field.
    - Fields: waiver_language, reservation_clause, preliminary_statement_fi/si/rfa/rpd, intro_template_fi/si/rfa/rpd, general_objections_rfa/rpd, verification_template
  - **Tab 3 (Custom Instructions):** Single large QTextEdit with placeholder text.
  - **Bottom buttons:** OK | Cancel | Reset All to Defaults | Reload from Samples
  - `get_rules() -> ResponseRules` — reads all UI state into a ResponseRules object.
  - `_load_rules(rules)` — populates UI from a ResponseRules object.
  - `_on_reset_defaults()` — calls `_load_rules(ResponseRules())`.
  - `_on_reload_samples()` — placeholder for future sample re-reading.

- [ ] **Step 4: Run tests to verify they pass**

```bash
python -m pytest tests/test_response_rules_dialog.py -v
```
Expected: All tests PASS.

- [ ] **Step 5: Commit**

```bash
git add icharlotte_core/ui/response_rules_dialog.py tests/test_response_rules_dialog.py
git commit -m "feat(ui): add ResponseRulesDialog with 3-tab hybrid editor"
```

---

## Task 7: RespondTab UI Widget

**Files:**
- Create: `icharlotte_core/ui/respond_tab.py`
- Create: `tests/test_respond_tab.py`

The main Respond tab widget with left pane (controls) and right pane (editor). Follows PropoundTab patterns for party management, document lists, LLM integration, and state persistence.

- [ ] **Step 1: Write test for RespondTab construction and basic behavior**

```python
# tests/test_respond_tab.py
"""Tests for the RespondTab UI widget."""
import os
import sys
import unittest
from unittest.mock import patch, MagicMock

os.environ["QT_QPA_PLATFORM"] = "offscreen"

from PySide6.QtWidgets import QApplication
app = QApplication.instance() or QApplication(sys.argv)

from icharlotte_core.ui.respond_tab import RespondTab
from icharlotte_core.discovery.models import Party, PartyRole


class TestRespondTabCreation(unittest.TestCase):

    @patch("icharlotte_core.ui.respond_tab.ModelFetcher")
    def test_creates_without_error(self, mock_fetcher_cls):
        mock_fetcher = MagicMock()
        mock_fetcher.isRunning.return_value = False
        mock_fetcher_cls.return_value = mock_fetcher
        tab = RespondTab()
        self.assertIsNotNone(tab)

    @patch("icharlotte_core.ui.respond_tab.ModelFetcher")
    def test_has_two_document_lists(self, mock_fetcher_cls):
        mock_fetcher = MagicMock()
        mock_fetcher.isRunning.return_value = False
        mock_fetcher_cls.return_value = mock_fetcher
        tab = RespondTab()
        self.assertIsNotNone(tab.discovery_list)
        self.assertIsNotNone(tab.context_list)

    @patch("icharlotte_core.ui.respond_tab.ModelFetcher")
    def test_has_generate_button(self, mock_fetcher_cls):
        mock_fetcher = MagicMock()
        mock_fetcher.isRunning.return_value = False
        mock_fetcher_cls.return_value = mock_fetcher
        tab = RespondTab()
        self.assertEqual(tab.generate_btn.text(), "Generate Responses")

    @patch("icharlotte_core.ui.respond_tab.ModelFetcher")
    def test_has_rules_button(self, mock_fetcher_cls):
        mock_fetcher = MagicMock()
        mock_fetcher.isRunning.return_value = False
        mock_fetcher_cls.return_value = mock_fetcher
        tab = RespondTab()
        self.assertIsNotNone(tab.rules_btn)

    @patch("icharlotte_core.ui.respond_tab.ModelFetcher")
    def test_has_refresh_17_1_button(self, mock_fetcher_cls):
        mock_fetcher = MagicMock()
        mock_fetcher.isRunning.return_value = False
        mock_fetcher_cls.return_value = mock_fetcher
        tab = RespondTab()
        self.assertIsNotNone(tab.refresh_17_1_btn)
        self.assertFalse(tab.refresh_17_1_btn.isEnabled())

    @patch("icharlotte_core.ui.respond_tab.ModelFetcher")
    def test_empty_state_shown_initially(self, mock_fetcher_cls):
        mock_fetcher = MagicMock()
        mock_fetcher.isRunning.return_value = False
        mock_fetcher_cls.return_value = mock_fetcher
        tab = RespondTab()
        self.assertFalse(tab.empty_label.isHidden())


class TestRespondTabLoadCase(unittest.TestCase):

    @patch("icharlotte_core.ui.respond_tab.ModelFetcher")
    @patch("icharlotte_core.ui.respond_tab.CaseDataManager")
    def test_load_case_sets_file_number(self, mock_cdm_cls, mock_fetcher_cls):
        mock_fetcher = MagicMock()
        mock_fetcher.isRunning.return_value = False
        mock_fetcher_cls.return_value = mock_fetcher
        mock_cdm = MagicMock()
        mock_cdm.get_value.return_value = None
        mock_cdm_cls.return_value = mock_cdm

        tab = RespondTab()
        tab.load_case("1234.001")
        self.assertEqual(tab.file_number, "1234.001")


if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 2: Run tests to verify they fail**

```bash
python -m pytest tests/test_respond_tab.py -v
```
Expected: FAIL

- [ ] **Step 3: Implement RespondTab**

Create `icharlotte_core/ui/respond_tab.py`. Follow PropoundTab patterns exactly. Key structure:

```python
"""
Respond sub-tab for the Discovery tab.

Generates objections and substantive responses to incoming discovery
using a three-phase pipeline: parse → select objections → draft responses.
"""
import os
from typing import Dict, List, Optional

from PySide6.QtCore import Qt, QTimer, Signal
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QSplitter, QLabel,
    QPushButton, QComboBox, QTabWidget, QPlainTextEdit,
    QGroupBox, QScrollArea, QFrame,
)

from icharlotte_core.llm import LLMWorker, ModelFetcher
from icharlotte_core.discovery.models import Party, PartyRole, DiscoveryType
from icharlotte_core.discovery.response_rules import ResponseRules
from icharlotte_core.discovery.response_parser import (
    ParsedDiscovery, build_parse_prompt, parse_llm_response,
)
from icharlotte_core.discovery.objection_selector import (
    ObjectionMenu, select_fi_objections, rule_based_preselect,
    build_objection_prompt, parse_objection_response, merge_objections,
    format_objections,
)
from icharlotte_core.discovery.response_drafter import (
    get_fi_fixed_response, build_si_prompt, build_rfa_prompt,
    build_rpd_prompt, assemble_plain_text, format_single_response,
)
from icharlotte_core.ui.response_rules_dialog import ResponseRulesDialog


# Import CaseDataManager from Scripts (same pattern as PropoundTab)
try:
    from Scripts.case_data_manager import CaseDataManager
except ImportError:
    CaseDataManager = None
```

Implement the class with these key sections (following PropoundTab line-for-line patterns):

**Constructor state:**
- `self.file_number: Optional[str] = None`
- `self.parties: List[Party] = []`
- `self.rules: ResponseRules = ResponseRules()`
- `self.cached_models: Dict[str, list] = {}`
- `self._output_save_timer: QTimer` (debounced 800ms)
- `self._llm_workers: list = []`
- `self._parsed_discoveries: Dict[str, ParsedDiscovery] = {}`
- `self._rfa_responses_cache: Optional[str] = None`

**Left pane (`_build_ui`):**
1. `_build_discovery_input()` — ResizableListWidget for discovery PDFs, blue header label, drag-drop (PDF only). Attribute: `self.discovery_list`.
2. `_build_context_documents()` — ResizableListWidget for context docs, gold header label, drag-drop (pdf/docx/txt/msg/images). Attribute: `self.context_list`.
3. `_build_party_section()` — "Our Client" combo + "+" button. Reuses `PartyEditDialog` from discovery_tab.py. Attribute: `self.party_combo`.
4. `_build_llm_section()` — Provider + Model combos. Attribute: `self.provider_combo`, `self.model_combo`. Uses `ModelFetcher` (same as PropoundTab).
5. `_build_buttons()` — "Response Rules..." button (`self.rules_btn`) and "Generate Responses" button (`self.generate_btn`).

**Right pane:**
1. Toolbar: Save as .docx (`self.save_btn`), Save All (`self.save_all_btn`), Clear (`self.clear_btn`), Refresh 17.1 (`self.refresh_17_1_btn`), status label (`self.status_label`).
2. `QTabWidget` (`self.output_tabs`) with `QPlainTextEdit` per generated response.
3. Empty state label (`self.empty_label`).

**Generation flow (`_on_generate`):**
1. Validate inputs (≥1 discovery PDF checked, Our Client selected).
2. Read discovery PDFs via `read_files_content()` (similar to PropoundTab).
3. Read context documents via `read_context_content()`.
4. For each discovery PDF: launch Phase 1 LLM worker to parse.
5. On Phase 1 complete (`_on_parse_finished`): run Phase 2 (objection selection). For FI: rule-based only. For SI/RFA/RPD: launch Phase 2 LLM worker.
6. On Phase 2 complete (`_on_objections_finished`): run Phase 3. For FI: use fixed templates + launch LLM for non-template items. For SI/RFA/RPD: launch Phase 3 LLM worker.
7. On Phase 3 complete (`_on_responses_finished`): call `assemble_plain_text()` and display in editor tab.
8. Cache RFA responses if generated.

**Refresh 17.1 (`_on_refresh_17_1`):**
1. Read current RFA editor tab text.
2. Launch LLM with `build_fi_17_1_prompt()`.
3. On complete: replace placeholder in FI editor tab.

**Save flow (`_save_current`, `_save_all`):**
1. Find caption page via `DiscoveryAssembler.find_caption_page()`.
2. Create `ResponseAssembler` with caption path.
3. Call `assemble()` with current editor text and parsed discovery data.
4. Save to case folder.

**Persistence (`load_case`, `_save_state`, `_load_state`):**
Uses CaseDataManager with variables: `respond_discovery_files`, `respond_context_documents`, `respond_output`, `respond_rules`, `respond_rfa_responses`.
Shares `discovery_party_roster` with PropoundTab.

- [ ] **Step 4: Run tests to verify they pass**

```bash
python -m pytest tests/test_respond_tab.py -v
```
Expected: All tests PASS.

- [ ] **Step 5: Commit**

```bash
git add icharlotte_core/ui/respond_tab.py tests/test_respond_tab.py
git commit -m "feat(ui): add RespondTab with three-phase generation pipeline"
```

---

## Task 8: Integration into DiscoveryTab

**Files:**
- Modify: `icharlotte_core/ui/discovery_tab.py` (lines 1362-1387)
- Modify: `tests/test_discovery_tab.py`

Replace the "Respond — coming soon" placeholder with the real `RespondTab`. Update `load_case()` to delegate to both tabs.

- [ ] **Step 1: Update test for DiscoveryTab to verify RespondTab integration**

Add to `tests/test_discovery_tab.py`:

```python
# Add import at top:
from icharlotte_core.ui.respond_tab import RespondTab

# Add test class:
class TestDiscoveryTabRespondIntegration(unittest.TestCase):
    """Verify RespondTab is integrated into DiscoveryTab."""

    @patch("icharlotte_core.ui.respond_tab.ModelFetcher")
    @patch("icharlotte_core.ui.discovery_tab.ModelFetcher")
    def test_has_respond_tab(self, mock_propound_fetcher, mock_respond_fetcher):
        for m in [mock_propound_fetcher, mock_respond_fetcher]:
            inst = MagicMock()
            inst.isRunning.return_value = False
            m.return_value = inst
        tab = DiscoveryTab()
        self.assertIsInstance(tab.respond_tab, RespondTab)

    @patch("icharlotte_core.ui.respond_tab.ModelFetcher")
    @patch("icharlotte_core.ui.discovery_tab.ModelFetcher")
    @patch("icharlotte_core.ui.respond_tab.CaseDataManager")
    def test_load_case_delegates_to_both(self, mock_cdm, mock_propound_fetcher, mock_respond_fetcher):
        for m in [mock_propound_fetcher, mock_respond_fetcher]:
            inst = MagicMock()
            inst.isRunning.return_value = False
            m.return_value = inst
        mock_cdm_inst = MagicMock()
        mock_cdm_inst.get_value.return_value = None
        mock_cdm.return_value = mock_cdm_inst

        tab = DiscoveryTab()
        tab.load_case("9999.001")
        self.assertEqual(tab.propound_tab.file_number, "9999.001")
        self.assertEqual(tab.respond_tab.file_number, "9999.001")
```

- [ ] **Step 2: Run new tests to verify they fail**

```bash
python -m pytest tests/test_discovery_tab.py::TestDiscoveryTabRespondIntegration -v
```
Expected: FAIL — `RespondTab` not yet wired in, or `respond_tab` attribute doesn't exist.

- [ ] **Step 3: Modify DiscoveryTab to use RespondTab**

In `icharlotte_core/ui/discovery_tab.py`, replace lines 1362-1387:

```python
# Add import at top of file:
from icharlotte_core.ui.respond_tab import RespondTab

# Replace the DiscoveryTab class:
class DiscoveryTab(QWidget):
    """Top-level Discovery tab containing Propound and Respond sub-tabs."""

    def __init__(self, parent=None):
        super().__init__(parent)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)

        self.tabs = QTabWidget()
        layout.addWidget(self.tabs)

        # Propound sub-tab (full implementation)
        self.propound_tab = PropoundTab()
        self.tabs.addTab(self.propound_tab, "Propound")

        # Respond sub-tab
        self.respond_tab = RespondTab()
        self.tabs.addTab(self.respond_tab, "Respond")

    def load_case(self, file_number: str):
        """Delegate to both sub-tabs."""
        self.propound_tab.load_case(file_number)
        self.respond_tab.load_case(file_number)
```

- [ ] **Step 4: Run all discovery tests to verify everything passes**

```bash
python -m pytest tests/test_discovery_tab.py tests/test_respond_tab.py -v
```
Expected: All tests PASS.

- [ ] **Step 5: Commit**

```bash
git add icharlotte_core/ui/discovery_tab.py tests/test_discovery_tab.py
git commit -m "feat(discovery): integrate RespondTab into DiscoveryTab, replacing placeholder"
```

---

## Task 9: Generate Default Rules JSON

**Files:**
- Create: `config/response_rules_default.json`

Generate the default rules JSON file so it's available on first app run.

- [ ] **Step 1: Generate defaults JSON**

```bash
python -c "
from icharlotte_core.discovery.response_rules import ResponseRules
rules = ResponseRules()
rules.save_to_json('config/response_rules_default.json')
print('Default rules saved.')
"
```

- [ ] **Step 2: Verify file was created and is valid JSON**

```bash
python -c "
import json
with open('config/response_rules_default.json') as f:
    d = json.load(f)
print(f'Keys: {len(d)}')
print(f'Waiver starts with: {d[\"waiver_language\"][:40]}...')
"
```
Expected: Key count matches ResponseRules field count, waiver text starts correctly.

- [ ] **Step 3: Commit**

```bash
git add config/response_rules_default.json
git commit -m "feat(config): add default response rules JSON"
```

---

## Task 10: End-to-End Manual Test

**Files:** None (manual verification)

- [ ] **Step 1: Run the app and verify Respond tab loads**

```bash
python iCharlotte.py
```

Navigate to the Discovery tab → Respond sub-tab. Verify:
- Two document drop zones visible (blue "Discovery to Respond To", gold "Context Documents")
- Party dropdown, LLM Provider/Model dropdowns
- "Response Rules..." button opens dialog with 3 tabs
- "Generate Responses" button present
- Empty state label shown in editor pane
- Refresh 17.1 button disabled

- [ ] **Step 2: Test drag-and-drop**

Drag a PDF file onto the "Discovery to Respond To" area. Verify it appears in the list with a checkbox.
Drag a different file onto the "Context Documents" area. Verify it appears separately.

- [ ] **Step 3: Test Response Rules dialog**

Click "Response Rules..." button. Verify:
- Strategy tab: all radio buttons and checkboxes present and functional
- Boilerplate tab: all text fields populated with default text, "Reset" buttons work
- Custom Instructions tab: empty text area with placeholder
- OK saves, Cancel discards changes

- [ ] **Step 4: Test generation (requires LLM API key)**

With a discovery PDF and context documents loaded:
1. Select Our Client from the party dropdown
2. Click "Generate Responses"
3. Verify progress shown in status label
4. Verify editor tabs appear with formatted response text
5. Verify each response has: request text, objections, waiver language, substantive response, reservation clause

- [ ] **Step 5: Test Save as .docx**

Click "Save as .docx" on a generated response. Verify:
- Word document created in the case folder
- Caption page present with correct title
- Party identification block formatted correctly
- Introduction and preliminary statement present
- Individual responses formatted with correct styles
- Verification page at end
- Signature block at end (if caption had one)

- [ ] **Step 6: Run full test suite**

```bash
python -m pytest tests/ -v --timeout=30
```
Expected: All tests pass, no regressions.

- [ ] **Step 7: Final commit**

```bash
git add -A
git commit -m "feat(discovery): complete Respond tab implementation with tests"
```
