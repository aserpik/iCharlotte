# Discovery Response Agent Improvements Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Fix the discovery response agent so that "Conservative + Minimal" settings produce output that is actually conservative, actually grounded in loaded context documents, free of internal contradictions, and free of known grammar/boilerplate bugs.

**Architecture:** Three phases. Phase 1 upgrades the drafter prompt to require grounding and introduces a `[NEEDS HUMAN INPUT:]` refusal protocol; also fixes two mechanical bugs ("further objects" as first sentence, compound detection). Phase 2 introduces a shared `objection_validator` module used by both initial objection selection and a new rule-based post-draft `objection_pruner`, plus few-shot calibration of the LLM objection-selection prompt, plus a UI checkpoint change that moves objection review to after drafting + pruning. Phase 3 adds a batched LLM consistency-check pass that flags contradictions across the response set without auto-reconciling.

**Tech Stack:** Python 3.x, PyQt6, pytest, existing `icharlotte_core.discovery.*` modules. No new external dependencies.

**Design spec:** `docs/superpowers/specs/2026-04-14-discovery-response-agent-improvements-design.md`

**Regression target:** `Z:\Shared\Current Clients\5800 - AMTRUST\017 - Mederos\NOTES\AI OUTPUT\DISCOVERY RESPONSES\Def PREMIER's Resp to SI(1).docx` — rerun after each phase completes and compare against acceptance criteria.

---

## Key file paths referenced throughout

- `icharlotte_core/discovery/response_parser.py` — parses incoming discovery PDFs, owns `ParsedRequest.is_compound` and `detect_compound()`
- `icharlotte_core/discovery/objection_selector.py` — objection library, rule-based pre-selection, LLM prompt, merge, format
- `icharlotte_core/discovery/response_drafter.py` — per-request prompt builders for SI/RFA/RPD (used in tests)
- `icharlotte_core/discovery/response_rules.py` — `ResponseRules` dataclass, default waiver/reservation language
- `icharlotte_core/discovery/engine.py` — pipeline orchestration (still used but the UI path is direct)
- `icharlotte_core/ui/respond_tab.py` — live pipeline for SI/RFA/RPD runs (method `_draft_responses_combined()` at line 885 is the active drafter call; individual `build_si_prompt()` imports at lines 37–38 are currently unused in the live path but kept in sync for test coverage)
- `icharlotte_core/discovery/response_assembler.py` — `_split_response_parts()` at line 528 splits objections/waiver/answer on the waiver marker

---

## Phase 1 — Drafter Grounding + Tactical Cleanup

Phase 1 is fully self-contained. Nothing in Phases 2 or 3 depends on Phase 1's exact implementation details.

### Task 1: Fix "further objects" grammar bug

**Files:**
- Modify: `icharlotte_core/discovery/objection_selector.py:254` (`format_objections()`)
- Create test: `tests/test_discovery/test_format_objections_grammar.py`

**Background:** Several objection texts in `_DEFAULT_OBJECTIONS` begin with "Responding Party further objects" (e.g., IDs 3, 4, 5, 11 — see `objection_selector.py:33-79`). The word "further" is correct when the objection is chained after another but grammatically wrong as the first statement in the chain. The PREMIER output shows this bug on SI 1, 2, 3, 5.

- [ ] **Step 1: Ensure test directory exists**

```bash
ls tests/test_discovery/ 2>&1 || mkdir -p tests/test_discovery
```

If the directory doesn't exist, create it. Check whether there is an `__init__.py` in `tests/` — if so, add one in `tests/test_discovery/` too.

- [ ] **Step 2: Write the failing test**

Create `tests/test_discovery/test_format_objections_grammar.py`:

```python
"""Tests for the grammar fix in format_objections()."""

from icharlotte_core.discovery.objection_selector import (
    ObjectionMenu,
    format_objections,
)


def test_first_objection_further_is_rewritten():
    """A leading 'further objects' objection must drop the 'further'."""
    menu = ObjectionMenu.load_defaults()
    # Objection #3 starts with "Responding Party further objects"
    result = format_objections({3}, menu)
    assert result.startswith("Responding Party objects to this Request")
    assert "Responding Party further objects" not in result


def test_first_objection_non_further_is_unchanged():
    """A leading non-'further' objection stays as written."""
    menu = ObjectionMenu.load_defaults()
    # Objection #1 starts with "Responding Party objects" (no "further")
    result = format_objections({1}, menu)
    assert result.startswith("Responding Party objects to this Request")


def test_subsequent_further_objections_keep_further():
    """Only the FIRST objection loses 'further'; later ones keep it."""
    menu = ObjectionMenu.load_defaults()
    # #1 first (no "further"), then #3 (has "further")
    result = format_objections({1, 3}, menu)
    # The first objection (#1) is unchanged and does not contain "further"
    # The second objection (#3) still reads "Responding Party further objects"
    assert "Responding Party further objects to this Request on the grounds that it seeks premature" in result


def test_multiple_further_objections_only_first_rewritten():
    """When the first objection has 'further', only it gets rewritten."""
    menu = ObjectionMenu.load_defaults()
    # #3 first (has "further"), then #4 (also has "further")
    result = format_objections({3, 4}, menu)
    # First objection loses "further"
    assert result.startswith("Responding Party objects to this Request")
    # Second objection keeps "further"
    assert "Responding Party further objects to this Request on the grounds that it seeks to invade" in result


def test_empty_objection_set_returns_empty():
    """Empty input returns empty string without errors."""
    menu = ObjectionMenu.load_defaults()
    result = format_objections(set(), menu)
    assert result == ""
```

- [ ] **Step 3: Run the test — verify it fails**

```bash
python -m pytest tests/test_discovery/test_format_objections_grammar.py -v
```

Expected: `test_first_objection_further_is_rewritten` and `test_multiple_further_objections_only_first_rewritten` fail with the leading "Responding Party further objects" still present.

- [ ] **Step 4: Apply the grammar fix in `format_objections()`**

Edit `icharlotte_core/discovery/objection_selector.py:254-277`. The current function:

```python
def format_objections(
    objection_ids: Set[int],
    menu: ObjectionMenu,
    term: Optional[str] = None,
) -> str:
    parts = []
    for obj_id in sorted(objection_ids):
        text = menu.get(obj_id)
        if term is not None and "{term}" in text:
            text = text.replace("{term}", term)
        parts.append(text)
    return " ".join(parts)
```

Change the return statement to apply a single-substitution regex to the leading "further objects":

```python
def format_objections(
    objection_ids: Set[int],
    menu: ObjectionMenu,
    term: Optional[str] = None,
) -> str:
    """
    Format selected objections into a single string.

    Args:
        objection_ids: Set of objection IDs to include.
        menu: The ObjectionMenu used for text lookup.
        term: If provided, substituted for ``{term}`` placeholders in
              objections 10 and 11.

    Returns:
        Space-separated objection sentences in ascending ID order.
        The leading "Responding Party further objects" is rewritten to
        "Responding Party objects" because "further" is grammatically
        incorrect as the first objection in the chain.
    """
    parts = []
    for obj_id in sorted(objection_ids):
        text = menu.get(obj_id)
        if term is not None and "{term}" in text:
            text = text.replace("{term}", term)
        parts.append(text)
    joined = " ".join(parts)
    # Rewrite only the leading occurrence — second-and-later objections
    # correctly retain "further".
    joined = re.sub(
        r"^Responding Party further objects",
        "Responding Party objects",
        joined,
        count=1,
    )
    return joined
```

No import changes needed — `re` is already imported at line 12.

- [ ] **Step 5: Run the test — verify it passes**

```bash
python -m pytest tests/test_discovery/test_format_objections_grammar.py -v
```

Expected: all 5 tests pass.

- [ ] **Step 6: Run the full discovery test suite to check for regressions**

```bash
python -m pytest tests/ -k "discovery or objection" -v
```

Expected: no existing tests newly fail.

- [ ] **Step 7: Commit**

```bash
git add icharlotte_core/discovery/objection_selector.py tests/test_discovery/test_format_objections_grammar.py
git commit -m "fix(discovery): strip 'further' from leading objection in format_objections"
```

---

### Task 2: Tighten compound detection in response_parser

**Files:**
- Modify: `icharlotte_core/discovery/response_parser.py:51-59` (`_COMPOUND_PATTERN` and `detect_compound()`)
- Create test: `tests/test_discovery/test_response_parser_compound.py`

**Background:** The Phase 2 validator (built in Phase 2) will use `ParsedRequest.is_compound` as ground truth when checking whether objection #9 (compound) is defensible on a given request. Right now `detect_compound()` uses a narrow regex that requires two imperative verbs with "AND" between them. It returns `False` for almost all PREMIER SIs including ones that ARE genuinely compound (SI 11 asks for name + date + method; SI 24 asks for individuals who supervised + directed). We need to broaden it to catch these AND avoid catching single-question SIs like SI 1 ("employee or independent contractor" — synonymous alternatives, not compound).

**Expected classification against PREMIER SI 1–30:**

| SI | Current `is_compound` | Target `is_compound` | Rationale |
|---|---|---|---|
| 1 | False | **False** | "employee or independent contractor" — synonymous alternatives |
| 2 | False | **True** | "list all dates... for Premier Gunite **and/or** Milo Holte" — two distinct subjects |
| 8 | False | **False** | "collecting or attempting to collect" — single action with attempt qualifier |
| 11 | False | **True** | "identify each customer notified, the date of notification, **and** the method" — three distinct information targets |
| 15 | False | **False** | "recover, disable, or revoke any materials" — three synonyms for the same action |
| 17 | False | **False** | "identify the amount, date received, and project" — borderline; treat as **not** compound because these are all attributes of a single payment (answer the question once per payment) |
| 23 | False | **True** | "identify the date **and** method of revocation" — two distinct attributes; date and method are different information types |
| 24 | False | **True** | "all individuals who supervised **or** directed" — two distinct activities |
| 30 | False | **True** | "by name, address, **and** telephone number" — three distinct attributes per person |

Note the subtle line between SI 17 (not compound) and SI 23/30 (compound). This is a judgment call. For the implementation, the rule is: if the request asks for multiple attributes of one thing, it is NOT compound; if it asks for information about two or more different types of things or activities, it IS compound. The test suite encodes these judgment calls as the expected outputs.

- [ ] **Step 1: Write the regression test fixture**

Create `tests/test_discovery/test_response_parser_compound.py`:

```python
"""Regression tests for tightened compound detection against PREMIER SI 1-30."""

import pytest
from icharlotte_core.discovery.response_parser import detect_compound


# Fixture: PREMIER SI Set One request texts with expected compound classification.
# Tuples are (si_number, request_text, expected_is_compound).
PREMIER_SI_FIXTURES = [
    (
        1,
        "Was Edgar Chavez ever an employee or independent contractor for Premier Gunite?",
        False,
    ),
    (
        2,
        "List all dates during which Edgar Chavez was either an employee or independent contractor for Premier Gunite and/or Milo Holte.",
        True,
    ),
    (
        3,
        "State all facts supporting your contention that Edgar Chavez was not employed by or acting as an agent of Premier Gunite on March 8, 2023.",
        False,
    ),
    (
        4,
        "Identify all dates on which Premier Gunite contends Edgar Chavez ceased working for Premier Gunite.",
        False,
    ),
    (
        5,
        "Identify all documents supporting your contention that Edgar Chavez was not authorized to act on behalf of Premier Gunite on March 8, 2023.",
        False,
    ),
    (
        8,
        "State whether Milo Holte was aware that Edgar Chavez was collecting or attempting to collect payments related to Premier Gunite projects on or about March 8, 2023.",
        False,
    ),
    (
        11,
        "If the response to Special Interrogatory No. 10 is yes, identify each customer notified, the date of notification, and the method of notification.",
        True,
    ),
    (
        15,
        "Identify all steps taken by Premier Gunite to recover, disable, or revoke any materials bearing Premier Gunite's name, logo, or branding that were in Edgar Chavez's possession.",
        False,
    ),
    (
        17,
        "If yes to Special Interrogatory No. 16, identify the amount, date received, and project associated with each payment.",
        False,
    ),
    (
        22,
        "State whether Premier Gunite ever revoked any authority previously granted to Edgar Chavez to interact with customers.",
        False,
    ),
    (
        23,
        "If authority was revoked, identify the date and method of revocation.",
        True,
    ),
    (
        24,
        "Identify all individuals who supervised or directed Edgar Chavez during 2022-2023.",
        True,
    ),
    (
        30,
        "Identify by name, address, and telephone number all Coachella Valley contractors whom you informed, either formally or informally, prior to March 8, 2023, that Edgar Chavez was not employed by Premier Gunite.",
        True,
    ),
]


@pytest.mark.parametrize("si_num,text,expected", PREMIER_SI_FIXTURES)
def test_premier_si_compound_classification(si_num, text, expected):
    """Each PREMIER SI should classify according to the fixture table."""
    actual = detect_compound(text)
    assert actual == expected, (
        f"SI {si_num}: expected is_compound={expected}, got {actual}.\n"
        f"Text: {text}"
    )


def test_synonymous_alternatives_are_not_compound():
    """'or' joining synonyms is not compound."""
    assert detect_compound("Did you recover, disable, or revoke any items?") is False
    assert detect_compound("Was he an employee or contractor?") is False


def test_multiple_information_targets_are_compound():
    """'and' joining distinct information types is compound."""
    assert detect_compound(
        "Identify the date and method of revocation."
    ) is True
    assert detect_compound(
        "State the name, address, and telephone number of each witness."
    ) is True


def test_two_imperative_verbs_is_compound():
    """Two distinct imperative verbs is compound."""
    assert detect_compound(
        "State all facts and identify all witnesses."
    ) is True
```

- [ ] **Step 2: Run the test — verify it fails on broadening cases**

```bash
python -m pytest tests/test_discovery/test_response_parser_compound.py -v
```

Expected: SI 2, 11, 23, 24, 30, and the "multiple information targets" / "two imperative verbs" tests fail because the current regex is too narrow. Tests for SI 1, 3, 4, 5, 8, 15, 17, 22 should pass (current regex already returns False for them).

- [ ] **Step 3: Rewrite `detect_compound()` in response_parser.py**

Edit `icharlotte_core/discovery/response_parser.py:51-59`:

```python
# Compound question detection
#
# A request is compound when it asks for two or more substantively different
# pieces of information. Heuristics (in order):
#   1. Two or more imperative verbs (state, identify, list, describe, explain,
#      set forth, produce) joined by "and" → compound.
#   2. A series of 2+ distinct noun-phrase information targets joined by
#      commas and "and" (e.g., "name, address, and telephone number", "date
#      and method of revocation") → compound.
#   3. Otherwise → not compound.
#
# Synonymous alternatives joined by "or" are NOT compound (e.g., "employee or
# independent contractor", "recover, disable, or revoke").
# Multiple attributes of a single entity are also NOT compound (e.g., "amount,
# date received, and project" when each payment has all three — ambiguous; we
# err on the side of not-compound for attribute lists that describe one thing).

_IMPERATIVE_VERBS = r"(?:state|identify|list|describe|explain|set forth|produce)"

# Two imperative verbs joined by "and"
_TWO_VERBS_PATTERN = re.compile(
    rf"\b{_IMPERATIVE_VERBS}\b[^.]*?\band\b[^.]*?\b{_IMPERATIVE_VERBS}\b",
    re.IGNORECASE,
)

# "X, Y, and Z" or "X and Y" where X/Y/Z are distinct information-target nouns.
# We look for common discovery-request noun targets: date, method, name,
# address, telephone, amount, identity, witness, project, document, subject,
# individual, customer.
_INFO_TARGET_NOUNS = (
    r"(?:date[s]?|method[s]?|name[s]?|address(?:es)?|telephone[s]?|"
    r"amount[s]?|identity|identities|witness(?:es)?|project[s]?|"
    r"document[s]?|individual[s]?|customer[s]?|party|parties|"
    r"time[s]?|location[s]?|reason[s]?)"
)

# Pattern for "X, Y, and Z" or "X and Y" joining info-target nouns
_INFO_TARGETS_LIST_PATTERN = re.compile(
    rf"\b{_INFO_TARGET_NOUNS}\b\s*(?:,\s*\b{_INFO_TARGET_NOUNS}\b\s*)*"
    rf"\band\b\s*\b{_INFO_TARGET_NOUNS}\b",
    re.IGNORECASE,
)


def detect_compound(text: str) -> bool:
    """Return True if the request is compound per the heuristics above."""
    if _TWO_VERBS_PATTERN.search(text):
        return True
    if _INFO_TARGETS_LIST_PATTERN.search(text):
        return True
    return False
```

Delete the old `_COMPOUND_PATTERN` definition at lines 52-56.

- [ ] **Step 4: Run the test again**

```bash
python -m pytest tests/test_discovery/test_response_parser_compound.py -v
```

Expected: all tests pass. If SI 17 fails (currently expected to be False but the new regex may catch "amount, date received, and project" because those are all info-target nouns), note this is a true judgment-call case — the test expects False. If the regex catches it, either (a) add "received" as a qualifier that breaks the pattern, or (b) accept that SI 17 flips to True and update the fixture table AND the Phase 1 acceptance criterion document. Prefer (a) — the rule that "attribute lists of one entity aren't compound" is worth preserving. One way: exclude the pattern if the words immediately before the list contain "each" (as in "each payment") which implies the attributes describe one iteration. Adjust the regex to require that no "each <noun>" phrase precedes the list within 20 characters.

If the tuning takes more than two attempts, split off a follow-up and move on — accept a small number of parser-level compound false positives since the Phase 2 validator will catch them at the objection layer.

- [ ] **Step 5: Run the full discovery test suite**

```bash
python -m pytest tests/ -k "discovery or parser" -v
```

Expected: no existing tests newly fail.

- [ ] **Step 6: Commit**

```bash
git add icharlotte_core/discovery/response_parser.py tests/test_discovery/test_response_parser_compound.py
git commit -m "feat(discovery): broaden compound detection to catch distinct info-target lists"
```

---

### Task 3: Upgrade combined drafter prompt with grounding + refusal protocol

**Files:**
- Modify: `icharlotte_core/ui/respond_tab.py:885-996` (`_draft_responses_combined()`)
- Modify: `icharlotte_core/discovery/response_drafter.py:200-293` (per-request builders `build_si_prompt`, `build_rfa_prompt`, `build_rpd_prompt` — keep in sync for test coverage)
- Modify: `tests/test_response_drafter.py` — add assertions for the new grounding requirement and refusal protocol

**Background:** The live drafter path is `_draft_responses_combined()` in respond_tab.py — it builds a single combined prompt containing all requests and the context, and sends one LLM call per discovery set. The individual `build_si_prompt()` / `build_rfa_prompt()` / `build_rpd_prompt()` functions in `response_drafter.py` are currently only exercised by `tests/test_response_drafter.py`; they are imported in respond_tab.py at lines 37–38 but not called.

Both code paths must receive the grounding requirement and refusal protocol so that: (a) the live run behavior changes, and (b) the test-level behavior stays in sync as a reference.

- [ ] **Step 1: Update `build_si_prompt()` in response_drafter.py**

Edit `icharlotte_core/discovery/response_drafter.py:200-228`. Replace the existing function body:

```python
def build_si_prompt(
    request_text: str,
    context_text: str,
    rules: ResponseRules,
) -> str:
    """Build an LLM prompt for drafting a Special Interrogatory substantive response.

    The prompt requires the LLM to ground its answer in CASE CONTEXT and to
    use the [NEEDS HUMAN INPUT:] refusal token when grounded facts are missing,
    rather than producing hedged non-answers.
    """
    style_instr = _SI_STYLE_INSTRUCTIONS.get(
        rules.si_response_style,
        _SI_STYLE_INSTRUCTIONS["moderate"],
    )

    parts = [
        "You are a California civil litigation defense attorney drafting a substantive",
        "response to a Special Interrogatory. Do not include objections — those are",
        "handled separately.",
        "",
        "GROUNDING REQUIREMENT: Your answer MUST be grounded in the CASE CONTEXT",
        "below. Before writing the answer, scan the CASE CONTEXT for:",
        "  - specific dates, names, and facts relevant to the interrogatory",
        "  - document titles, bates numbers, or filenames that would support the answer",
        "  - direct quotes or specific passages that address the question",
        "",
        "REFUSAL PROTOCOL: If the CASE CONTEXT does not contain the specific facts",
        "needed to answer this interrogatory, DO NOT fabricate or hedge with 'discovery",
        "is ongoing.' Instead, respond with exactly this token:",
        "",
        "  [NEEDS HUMAN INPUT: <one-line description of what facts are missing>]",
        "",
        "Examples of correct refusal:",
        "  [NEEDS HUMAN INPUT: specific project names where X worked in 2022-2023]",
        "  [NEEDS HUMAN INPUT: exact termination date and termination letter filename]",
        "",
        f"DRAFTING STYLE: {style_instr}",
        "",
        "Write only the substantive answer. Use specific facts from the CASE CONTEXT",
        "wherever possible. If you find a grounded answer, state it directly. Do not",
        "include the 'Subject to and without waiving the foregoing objections,",
        "Responding Party responds as follows:' transition — it is inserted by the",
        "pipeline when the response is assembled.",
        "",
        f"INTERROGATORY:\n{request_text}",
        "",
        f"CASE CONTEXT:\n{context_text}",
    ]

    if rules.custom_instructions:
        parts += ["", f"ADDITIONAL INSTRUCTIONS:\n{rules.custom_instructions}"]

    return "\n".join(parts)
```

- [ ] **Step 2: Update `build_rfa_prompt()` in response_drafter.py**

Edit the function at `response_drafter.py:231-264`. The RFA prompt gets the same grounding requirement and refusal protocol, but the refusal format is slightly different — RFAs must pick exactly one of Admit / Deny / Insufficient Information, so the refusal token applies when even picking one of those three is not supportable from context.

```python
def build_rfa_prompt(
    request_text: str,
    context_text: str,
    rules: ResponseRules,
) -> str:
    """Build an LLM prompt for drafting a Request for Admission response."""
    posture_instr = _RFA_POSTURE_INSTRUCTIONS.get(
        rules.rfa_default_posture,
        _RFA_POSTURE_INSTRUCTIONS["cautious"],
    )

    parts = [
        "You are a California civil litigation defense attorney drafting a response",
        "to a Request for Admission. Do not include objections — those are handled",
        "separately.",
        "",
        "You must choose exactly one of the three California-approved response forms:",
        "  1. Admit",
        "  2. Deny",
        "  3. After a reasonable inquiry concerning the matter in this request, the",
        "     information known or readily obtainable is insufficient to enable this",
        "     party to admit the matter.",
        "",
        "GROUNDING REQUIREMENT: Your choice MUST be supported by the CASE CONTEXT",
        "below. Scan the context for facts that would confirm or refute the matter",
        "requested. Only use 'Admit' when the context affirmatively establishes the",
        "fact. Only use 'Deny' when the context refutes the fact.",
        "",
        "REFUSAL PROTOCOL: If the CASE CONTEXT does not contain facts relevant to",
        "the matter requested, use response form #3 (insufficient information). If",
        "the context is so thin that you cannot even determine which form to pick,",
        "respond with exactly this token instead:",
        "",
        "  [NEEDS HUMAN INPUT: <one-line description of what facts are missing>]",
        "",
        f"POSTURE: {posture_instr}",
        "",
        f"REQUEST FOR ADMISSION:\n{request_text}",
        "",
        f"CASE CONTEXT:\n{context_text}",
    ]

    if rules.custom_instructions:
        parts += ["", f"ADDITIONAL INSTRUCTIONS:\n{rules.custom_instructions}"]

    return "\n".join(parts)
```

- [ ] **Step 3: Update `build_rpd_prompt()` in response_drafter.py**

Edit the function at `response_drafter.py:266-293`:

```python
def build_rpd_prompt(
    request_text: str,
    context_text: str,
    rules: ResponseRules,
) -> str:
    """Build an LLM prompt for drafting a Request for Production response."""
    posture_instr = _RPD_POSTURE_INSTRUCTIONS.get(
        rules.rpd_default_posture,
        _RPD_POSTURE_INSTRUCTIONS["context_dependent"],
    )

    parts = [
        "You are a California civil litigation defense attorney drafting a response",
        "to a Request for Production. Do not include objections — those are handled",
        "separately.",
        "",
        "You must choose exactly one of the two California-approved response forms:",
        "  1. Will comply — Responding Party will comply and produce non-privileged",
        "     documents in Responding Party's possession, custody, and control.",
        "  2. Unable to comply — after diligent search and reasonable inquiry, no",
        "     responsive documents are in Responding Party's possession, custody,",
        "     or control.",
        "",
        "GROUNDING REQUIREMENT: Your choice MUST be supported by the CASE CONTEXT",
        "below. Scan the context for indications of whether responsive documents",
        "exist in Responding Party's possession.",
        "",
        "REFUSAL PROTOCOL: If the CASE CONTEXT does not indicate whether responsive",
        "documents exist, respond with exactly this token instead:",
        "",
        "  [NEEDS HUMAN INPUT: <one-line description of what facts are missing>]",
        "",
        f"POSTURE: {posture_instr}",
        "",
        f"REQUEST FOR PRODUCTION:\n{request_text}",
        "",
        f"CASE CONTEXT:\n{context_text}",
    ]

    if rules.custom_instructions:
        parts += ["", f"ADDITIONAL INSTRUCTIONS:\n{rules.custom_instructions}"]

    return "\n".join(parts)
```

- [ ] **Step 4: Update `_draft_responses_combined()` in respond_tab.py**

Edit `icharlotte_core/ui/respond_tab.py:885-996`. The type-specific `type_instruction` block at lines 904-948 needs the grounding requirement and refusal protocol added to each branch. The simplest approach is to define a shared pre-amble constant above the function and include it in each branch.

Find the `_draft_responses_combined` method at line 885. Before the `disc_type = parsed.discovery_type.upper()` line, add this local constant inside the function:

```python
        _GROUNDING_PREAMBLE = (
            "GROUNDING REQUIREMENT: Every answer you produce MUST be grounded in "
            "the CASE CONTEXT below. For each request, scan the CASE CONTEXT for "
            "specific dates, names, facts, document titles, and passages that "
            "address the question.\n\n"
            "REFUSAL PROTOCOL: If the CASE CONTEXT does not contain the specific "
            "facts needed to answer a particular request, DO NOT fabricate or "
            "hedge with 'discovery is ongoing' or vague generalities. Instead, "
            "for that specific response, output exactly:\n\n"
            "  [NEEDS HUMAN INPUT: <one-line description of what facts are missing>]\n\n"
            "Examples:\n"
            "  RESPONSE 5: [NEEDS HUMAN INPUT: specific document titles supporting the contention]\n"
            "  RESPONSE 19: [NEEDS HUMAN INPUT: specific project names during 2022-2023]\n\n"
            "It is better to refuse with the token than to produce a vapid answer.\n"
        )
```

Then modify each `type_instruction` branch to prepend the preamble. For SI (around line 904):

```python
        if disc_type == "SI":
            from icharlotte_core.discovery.response_drafter import _SI_STYLE_INSTRUCTIONS
            style = _SI_STYLE_INSTRUCTIONS.get(
                self.rules.si_response_style, _SI_STYLE_INSTRUCTIONS["moderate"]
            )
            type_instruction = (
                _GROUNDING_PREAMBLE
                + f"DRAFTING STYLE: {style}\n"
                "Draft substantive responses only — objections are handled separately.\n"
                "Write only the factual, responsive answer for each interrogatory.\n"
                "Do not include the 'Subject to and without waiving the foregoing "
                "objections, Responding Party responds as follows:' transition — "
                "it is inserted by the pipeline when the response is assembled."
            )
```

For RFA:

```python
        elif disc_type == "RFA":
            from icharlotte_core.discovery.response_drafter import _RFA_POSTURE_INSTRUCTIONS
            posture = _RFA_POSTURE_INSTRUCTIONS.get(
                self.rules.rfa_default_posture, _RFA_POSTURE_INSTRUCTIONS["cautious"]
            )
            type_instruction = (
                _GROUNDING_PREAMBLE
                + "For each Request for Admission, respond with EXACTLY one of:\n"
                '- "Admit"\n'
                '- "Deny"\n'
                '- "After a reasonable inquiry concerning the matter in this request, '
                "the information known or readily obtainable to Responding Party is "
                'insufficient to enable Responding Party to admit the matter."\n\n'
                "Only use 'Admit' when the CASE CONTEXT affirmatively establishes "
                "the fact; only use 'Deny' when the CASE CONTEXT refutes it; "
                "use response form 3 when the context is silent.\n\n"
                f"POSTURE: {posture}"
            )
```

For RPD:

```python
        elif disc_type == "RPD":
            from icharlotte_core.discovery.response_drafter import _RPD_POSTURE_INSTRUCTIONS
            posture = _RPD_POSTURE_INSTRUCTIONS.get(
                self.rules.rpd_default_posture, _RPD_POSTURE_INSTRUCTIONS["context_dependent"]
            )
            type_instruction = (
                _GROUNDING_PREAMBLE
                + "For each Request for Production, respond with EXACTLY one of:\n"
                '- "Responding Party will comply with this request and produce all '
                "non-privileged documents in Responding Party's possession, custody "
                "and control that Responding Party understands to be responsive to this "
                "Request. Responding Party identifies and refers to the documents "
                'produced concurrently herewith."\n'
                '- "Upon a diligent search and reasonable inquiry made in an effort to '
                "locate the item(s) requested, Responding Party is unable to comply "
                "with this request at this time because the documents responsive to "
                "this request, if they exist, are not in the possession, custody or "
                'control of Responding Party."\n\n'
                f"POSTURE: {posture}"
            )
```

Leave the `else` branch (`type_instruction = "Draft substantive responses."`) alone — it's a fallback for unknown discovery types.

- [ ] **Step 5: Add test assertions for the new drafter prompt content**

Edit `tests/test_response_drafter.py` — add tests (or update existing tests) asserting that `build_si_prompt()`, `build_rfa_prompt()`, and `build_rpd_prompt()` include the grounding and refusal markers. If the file doesn't yet exist in that form, add a new test file `tests/test_discovery/test_drafter_prompts.py`:

```python
"""Tests that drafter prompts include grounding requirement and refusal protocol."""

import pytest
from icharlotte_core.discovery.response_drafter import (
    build_si_prompt,
    build_rfa_prompt,
    build_rpd_prompt,
)
from icharlotte_core.discovery.response_rules import ResponseRules


@pytest.fixture
def rules():
    return ResponseRules()


def test_si_prompt_has_grounding_requirement(rules):
    prompt = build_si_prompt(
        request_text="Was X an employee?",
        context_text="X was hired on 2022-01-01.",
        rules=rules,
    )
    assert "GROUNDING REQUIREMENT" in prompt
    assert "[NEEDS HUMAN INPUT:" in prompt


def test_si_prompt_has_refusal_protocol(rules):
    prompt = build_si_prompt(
        request_text="Was X an employee?",
        context_text="X was hired on 2022-01-01.",
        rules=rules,
    )
    assert "REFUSAL PROTOCOL" in prompt
    assert "do not fabricate" in prompt.lower() or "DO NOT fabricate" in prompt


def test_si_prompt_suppresses_waiver_transition(rules):
    prompt = build_si_prompt(
        request_text="Was X an employee?",
        context_text="X was hired.",
        rules=rules,
    )
    assert "Subject to and without waiving" in prompt  # mentioned as text-to-avoid
    assert "inserted by the pipeline" in prompt


def test_rfa_prompt_has_grounding_requirement(rules):
    prompt = build_rfa_prompt(
        request_text="Admit X was an employee.",
        context_text="X was hired.",
        rules=rules,
    )
    assert "GROUNDING REQUIREMENT" in prompt
    assert "[NEEDS HUMAN INPUT:" in prompt


def test_rpd_prompt_has_grounding_requirement(rules):
    prompt = build_rpd_prompt(
        request_text="Produce all documents relating to X.",
        context_text="X was hired.",
        rules=rules,
    )
    assert "GROUNDING REQUIREMENT" in prompt
    assert "[NEEDS HUMAN INPUT:" in prompt


def test_si_prompt_includes_context_text(rules):
    context = "UNIQUE_CONTEXT_MARKER_12345"
    prompt = build_si_prompt(
        request_text="Was X an employee?",
        context_text=context,
        rules=rules,
    )
    assert "UNIQUE_CONTEXT_MARKER_12345" in prompt
    assert "CASE CONTEXT:" in prompt


def test_si_prompt_includes_request_text(rules):
    request = "UNIQUE_REQUEST_MARKER_67890"
    prompt = build_si_prompt(
        request_text=request,
        context_text="context",
        rules=rules,
    )
    assert "UNIQUE_REQUEST_MARKER_67890" in prompt
```

- [ ] **Step 6: Run the drafter prompt tests**

```bash
python -m pytest tests/test_discovery/test_drafter_prompts.py -v
python -m pytest tests/test_response_drafter.py -v
```

Expected: all tests pass. If the existing `test_response_drafter.py` had assertions about the old prompt wording that no longer hold, update those test assertions to match the new structure — the old tests were pinning the old behavior; the new tests pin the new behavior.

- [ ] **Step 7: Smoke-test that respond_tab.py still parses**

```bash
python -c "from icharlotte_core.ui import respond_tab; print('OK')"
```

Expected: prints "OK". If there's a syntax error, fix it.

- [ ] **Step 8: Commit**

```bash
git add icharlotte_core/discovery/response_drafter.py icharlotte_core/ui/respond_tab.py tests/test_discovery/test_drafter_prompts.py tests/test_response_drafter.py
git commit -m "feat(discovery): add grounding requirement + NEEDS HUMAN INPUT refusal protocol to drafter prompts"
```

---

### Task 4: UI — display and resolve NEEDS HUMAN INPUT flags

**Files:**
- Modify: `icharlotte_core/ui/respond_tab.py` — the output display area around `_display_result()` (line 1143) and `_save_response_doc()` (line 1235), plus the on-combined-draft handler `_on_combined_draft_finished()` (line 998)
- Create or extend: `icharlotte_core/ui/respond_tab.py` — a helper method `_has_unresolved_human_input_flags()` and a warning dialog path in `_save_response_doc()` or the assembly trigger

**Background:** When the drafter returns `[NEEDS HUMAN INPUT: <description>]` as part of its response text, the UI must make this visually distinct, let the user resolve it, and block final document assembly until resolved.

The current `_on_combined_draft_finished()` at line 998 parses the combined LLM output into `responses_map[num] = text`. The `[NEEDS HUMAN INPUT:` marker lives inside that text. All we need is: a detection helper, visual treatment when displaying, and a gate before assembly.

- [ ] **Step 1: Locate the response display area**

Read `icharlotte_core/ui/respond_tab.py:1143-1200` (the `_display_result` method and surrounding code) to see how responses are currently shown to the user. Understand whether responses go into a QTextEdit, a QListWidget, or something else.

```bash
python -c "import ast; tree=ast.parse(open('icharlotte_core/ui/respond_tab.py').read()); [print(n.name, n.lineno) for n in ast.walk(tree) if isinstance(n, ast.FunctionDef) and '_display' in n.name]"
```

Read the identified function(s) to understand the widget being used.

- [ ] **Step 2: Add a helper module-level function in respond_tab.py**

After the imports block (around line 80, before `class RespondTab`):

```python
# ---------------------------------------------------------------------------
# NEEDS HUMAN INPUT flag helpers
# ---------------------------------------------------------------------------

_NEEDS_HUMAN_INPUT_PATTERN = re.compile(
    r"\[NEEDS HUMAN INPUT:\s*([^\]]*)\]", re.IGNORECASE
)


def has_human_input_flag(response_text: str) -> bool:
    """Return True if the response contains a [NEEDS HUMAN INPUT:] token."""
    return bool(_NEEDS_HUMAN_INPUT_PATTERN.search(response_text))


def extract_human_input_descriptions(response_text: str) -> list[str]:
    """Extract all '<description>' values from NEEDS HUMAN INPUT tokens."""
    return [m.strip() for m in _NEEDS_HUMAN_INPUT_PATTERN.findall(response_text)]
```

Make sure `import re` is already at the top of the file — if not, add it.

- [ ] **Step 3: Visual treatment when displaying a response**

In `_display_result()` at line 1143 (or the equivalent response-rendering path — read the function first), check each response for the flag. If the flag is present, wrap the rendered text with a visual marker suitable for the existing widget. If the display uses HTML (QTextEdit with `setHtml`), wrap the affected response in a span with a yellow background:

```python
# Inside the response-rendering loop, after you have `text` for one response:
if has_human_input_flag(text):
    # Yellow background for flagged responses
    display_text = (
        f'<div style="background-color: #fff3b0; padding: 4px; '
        f'border-left: 3px solid #e0a800;">'
        f'<b>[HUMAN INPUT NEEDED]</b><br>{text}'
        f'</div>'
    )
else:
    display_text = text
```

The exact widget API call depends on the existing code (`setHtml`, `insertHtml`, list-item styling, etc.). Preserve the existing code path and add the flag-aware branch — do not rewrite the display function wholesale.

- [ ] **Step 4: Block assembly when unresolved flags exist**

Locate the code path that triggers document assembly — likely `_save_response_doc()` at line 1235 or a sibling method triggered by a "Save" button handler. Before invoking the assembler, scan `responses_map` for human-input flags:

```python
# Inside _save_response_doc or the handler that triggers assembly, BEFORE
# calling the assembler:
flagged = [
    (num, extract_human_input_descriptions(text))
    for num, text in responses_map.items()
    if has_human_input_flag(text)
]
if flagged:
    from PyQt6.QtWidgets import QMessageBox
    flag_summary = "\n".join(
        f"  • {num}: {'; '.join(descs) if descs else '(no description)'}"
        for num, descs in flagged
    )
    msg = QMessageBox(self)
    msg.setWindowTitle("Unresolved NEEDS HUMAN INPUT flags")
    msg.setIcon(QMessageBox.Icon.Warning)
    msg.setText(
        "The following responses contain unresolved [NEEDS HUMAN INPUT:] flags:"
    )
    msg.setInformativeText(flag_summary + "\n\nEdit the affected responses in the editor before saving, or click 'Save Anyway' to proceed with the flags in the document.")
    save_btn = msg.addButton("Save Anyway", QMessageBox.ButtonRole.DestructiveRole)
    cancel_btn = msg.addButton("Cancel", QMessageBox.ButtonRole.RejectRole)
    msg.setDefaultButton(cancel_btn)
    msg.exec()
    if msg.clickedButton() == cancel_btn:
        return  # abort save
```

The exact PyQt import may already be at the top of the file — check before adding a duplicate import.

- [ ] **Step 5: Manual UI smoke test**

Start the app:

```bash
python iCharlotte.py
```

Navigate to the Respond tab. Load a discovery file and some context documents. Click Generate. When the drafter returns responses, manually edit one of them in the editor to contain the literal string `[NEEDS HUMAN INPUT: test flag]`. Attempt to save — verify the warning dialog appears listing the affected SI number and the description "test flag". Click Cancel — verify the save aborts. Remove the flag and attempt save again — verify the save proceeds normally.

- [ ] **Step 6: Commit**

```bash
git add icharlotte_core/ui/respond_tab.py
git commit -m "feat(respond): detect NEEDS HUMAN INPUT flags and block assembly until resolved"
```

---

### Task 5: Phase 1 manual regression test against PREMIER SI Set One

**Files:** No code changes — this is an end-to-end verification run.

- [ ] **Step 1: Locate the PREMIER SI Set One source file**

The PREMIER SI Set One (the plaintiff's interrogatories that the agent responds to) is what the Respond tab accepts as input. Find it on the Z drive:

```bash
ls '/z/Shared/Current Clients/5800 - AMTRUST/017 - Mederos/' 2>&1 | head -20
```

Look for a file containing "SI" or "Special Interrogatories" **propounded** by the plaintiff. The file used to generate the current buggy output is at `/z/Shared/Current Clients/5800 - AMTRUST/017 - Mederos/NOTES/AI OUTPUT/DISCOVERY RESPONSES/Def PREMIER's Resp to SI(1).docx` — that's the agent's output. The input is whatever PDF/DOCX the plaintiff served; ask the user for the exact path if you can't locate it.

- [ ] **Step 2: Start the app and load the PREMIER SI**

```bash
python iCharlotte.py
```

Open the Respond tab. Load the PREMIER SI Set One PDF as the discovery input. Load whatever case context documents the user had loaded in the original run (ask the user if unclear — the original run lives at the path above; the context docs should be listed in the respond tab's context panel state for that case).

- [ ] **Step 3: Configure Rules — Conservative + Minimal**

Click the Rules button. Set `objection_aggressiveness = conservative`. Set `si_response_style = minimal`. No "always include" options checked. No "auto detection" options checked. Same settings as the original PREMIER run.

- [ ] **Step 4: Generate responses**

Click Generate. Wait for the drafter to complete.

- [ ] **Step 5: Verify Phase 1 acceptance criteria**

Check each criterion from the design spec section 4.5 against the generated output:

**Criterion 1 (grounding):** At least 5 responses that previously said "Payroll records" / "Various projects" / similar generic non-answers now either (a) contain specific facts from the loaded context docs OR (b) show a `[NEEDS HUMAN INPUT:]` flag. Zero responses may still be generic non-answers without a flag.

  - Check SI 5 — original said "Payroll records". New output should either list specific documents OR show a flag.
  - Check SI 14 — original said "Premier Gunite terminated his employment". New output should name specific steps OR show a flag.
  - Check SI 19 — original said "Various projects during his employment". New output should name specific projects OR show a flag.
  - Check SI 24 — original said "Milo Holte". This was an accurate answer; should remain accurate.

**Criterion 2 (grammar fix):** No response's first objection begins with "further." Scan every response's objections section.

**Criterion 3 (parser compound tightening):** Open a Python shell and verify:

```python
from icharlotte_core.discovery.response_parser import detect_compound
# Paste SI 1 text:
assert detect_compound("Was Edgar Chavez ever an employee or independent contractor for Premier Gunite?") is False
# SI 2 is compound:
assert detect_compound("List all dates during which Edgar Chavez was either an employee or independent contractor for Premier Gunite and/or Milo Holte.") is True
print("parser OK")
```

Note: the live LLM may still add objection #9 to SI 1/SI 8 because Phase 1 doesn't yet calibrate the LLM. That's expected and will be fixed in Phase 2.

**Criterion 4 (no regression on other discovery types):** Run a small FI set, a small RFA set, and a small RPD set against an existing case. Verify the runs complete without crashes and the output is structurally intact (has caption, preliminary statement, numbered responses, verification block).

- [ ] **Step 6: Document the results**

Write a short markdown note at `docs/superpowers/plans/phase-1-regression-notes.md` capturing: which SIs now have grounded answers, which have flags, which criteria passed, which criteria failed, any surprises.

- [ ] **Step 7: Commit the regression notes**

```bash
git add docs/superpowers/plans/phase-1-regression-notes.md
git commit -m "docs(discovery): phase 1 PREMIER regression notes"
```

**If any Phase 1 criterion failed, stop and fix before starting Phase 2.** The subsequent phases assume Phase 1's grounding/refusal is working.

---

## Phase 2 — Objection Discipline

Phase 2 introduces the shared validator layer, calibrates the LLM objection-selection prompt with few-shot examples, adds a rule-based post-draft prune pass, and moves the UI objection review checkpoint to after pruning.

### Task 6: Scaffold the `objection_validator` module

**Files:**
- Create: `icharlotte_core/discovery/objection_validator.py`
- Create test: `tests/test_discovery/test_objection_validator_scaffold.py`

- [ ] **Step 1: Write the scaffold test**

Create `tests/test_discovery/test_objection_validator_scaffold.py`:

```python
"""Scaffold tests for objection_validator module."""

from icharlotte_core.discovery.objection_validator import (
    ObjectionValidationResult,
    filter_objection_ids,
    VALIDATORS,
)
from icharlotte_core.discovery.response_parser import ParsedRequest


def test_validation_result_construction():
    ok = ObjectionValidationResult(valid=True)
    assert ok.valid is True
    assert ok.reason == ""

    bad = ObjectionValidationResult(valid=False, reason="no expert keyword")
    assert bad.valid is False
    assert bad.reason == "no expert keyword"


def test_validators_dispatcher_contains_expected_gates():
    # Objections 3, 5, 6, 7, 9 are the validated ones per the design spec.
    # Objection 4 (privilege) is intentionally NOT in the dispatcher.
    assert set(VALIDATORS.keys()) == {3, 5, 6, 7, 9}


def test_filter_objection_ids_returns_tuple():
    req = ParsedRequest(number="1", text="Was X an employee?")
    kept, dropped = filter_objection_ids({1}, req)
    assert isinstance(kept, set)
    assert isinstance(dropped, list)


def test_filter_passes_through_unknown_ids():
    req = ParsedRequest(number="1", text="Was X an employee?")
    # Objection 1 has no gate, so it should be passed through unchanged.
    kept, dropped = filter_objection_ids({1}, req)
    assert 1 in kept
    assert dropped == []


def test_filter_preserves_privilege_objection():
    """Objection 4 (privilege) is NEVER filtered, even with no request context."""
    req = ParsedRequest(number="1", text="Was X an employee?")
    kept, dropped = filter_objection_ids({4}, req)
    assert 4 in kept
    assert dropped == []
```

- [ ] **Step 2: Run the test — verify it fails**

```bash
python -m pytest tests/test_discovery/test_objection_validator_scaffold.py -v
```

Expected: ImportError because `objection_validator.py` does not exist.

- [ ] **Step 3: Create the scaffold module**

Create `icharlotte_core/discovery/objection_validator.py`:

```python
"""
Phase 2 shared validator for discovery objections.

This module defines per-objection validity gates. It is used in two places:

1. After initial objection selection (rule-based pre-select + LLM select),
   to drop indefensible picks.
2. Inside the post-draft prune pass (see `objection_pruner.py`), to drop
   objections contradicted or rendered moot by the substantive answer.

The same gate functions are used in both modes. When called with
`answer_text=None`, the gate evaluates only the request-side constraints.
When called with `answer_text=<string>`, the gate additionally applies
answer-aware rules.

Objection IDs NOT in this validator:

- #1 (speculation / vague) — general-purpose, no hard gate
- #2 (privacy) — user-controlled via ResponseRules.always_include_privacy_objection
- #4 (attorney-client / work product) — NEVER validated or pruned, per design
- #8 (previously propounded) — rare, no automated gate
- #10, #11 (undefined terms / vague definitions) — require term-level context
- #12 (equally available) — no automated gate
"""
from dataclasses import dataclass
from typing import Callable, Dict, List, Optional, Set, Tuple

from icharlotte_core.discovery.response_parser import ParsedRequest


@dataclass
class ObjectionValidationResult:
    valid: bool
    reason: str = ""


# Gate signature: (request, answer_text=None) -> ObjectionValidationResult
GateFn = Callable[[ParsedRequest, Optional[str]], ObjectionValidationResult]


# Gate functions are defined in subsequent tasks. The dispatcher is filled
# in as each gate is implemented.
VALIDATORS: Dict[int, GateFn] = {}


def filter_objection_ids(
    objection_ids: Set[int],
    request: ParsedRequest,
    answer_text: Optional[str] = None,
) -> Tuple[Set[int], List[Tuple[int, str]]]:
    """
    Filter a set of objection IDs through the shared validator gates.

    Returns:
        (kept_ids, dropped_with_reasons)
        - kept_ids is the subset that passed validation (or had no gate)
        - dropped_with_reasons is a list of (id, reason) for dropped ones

    Objection IDs not in VALIDATORS are passed through unchanged — the
    validator only enforces gates for objections that have one.

    Objection #4 (privilege) is ALWAYS preserved regardless of VALIDATORS
    state, per design decision.
    """
    kept: Set[int] = set()
    dropped: List[Tuple[int, str]] = []

    for obj_id in objection_ids:
        if obj_id == 4:
            kept.add(obj_id)
            continue
        gate = VALIDATORS.get(obj_id)
        if gate is None:
            kept.add(obj_id)
            continue
        result = gate(request, answer_text)
        if result.valid:
            kept.add(obj_id)
        else:
            dropped.append((obj_id, result.reason))

    return kept, dropped
```

- [ ] **Step 4: Run the scaffold test — verify it passes**

```bash
python -m pytest tests/test_discovery/test_objection_validator_scaffold.py -v
```

Expected: `test_validators_dispatcher_contains_expected_gates` FAILS because VALIDATORS is still empty — this is expected for the scaffold commit; we'll fill VALIDATORS in Tasks 7–11. The other four tests should pass.

Update `test_validators_dispatcher_contains_expected_gates` to skip for now:

```python
import pytest

@pytest.mark.skip(reason="VALIDATORS populated in Tasks 7-11")
def test_validators_dispatcher_contains_expected_gates():
    ...
```

Rerun:

```bash
python -m pytest tests/test_discovery/test_objection_validator_scaffold.py -v
```

Expected: 4 passed, 1 skipped.

- [ ] **Step 5: Commit**

```bash
git add icharlotte_core/discovery/objection_validator.py tests/test_discovery/test_objection_validator_scaffold.py
git commit -m "feat(discovery): scaffold objection_validator module with filter_objection_ids"
```

---

### Task 7: Implement `validate_obj_3_expert_opinion` gate

**Files:**
- Modify: `icharlotte_core/discovery/objection_validator.py`
- Create test: `tests/test_discovery/test_validator_obj_3.py`

**Rule:** Objection #3 ("premature disclosure of expert opinion and/or legal conclusion") is only defensible when the request explicitly seeks an expert opinion or legal contention. Trigger words in the request: `expert`, `opinion of`, `contention`, `contend`, `legal theory`, `legal basis`. In answer-aware mode, if the answer gives a straightforward factual response, also drop the objection.

- [ ] **Step 1: Write the gate tests**

Create `tests/test_discovery/test_validator_obj_3.py`:

```python
"""Tests for validate_obj_3_expert_opinion."""

from icharlotte_core.discovery.objection_validator import (
    validate_obj_3_expert_opinion,
)
from icharlotte_core.discovery.response_parser import ParsedRequest


def _req(text):
    return ParsedRequest(number="1", text=text)


def test_invalid_on_plain_factual_yes_no():
    result = validate_obj_3_expert_opinion(
        _req("Was Edgar Chavez ever an employee of Premier Gunite?")
    )
    assert result.valid is False
    assert "expert" in result.reason.lower() or "opinion" in result.reason.lower()


def test_invalid_on_identify_documents():
    result = validate_obj_3_expert_opinion(
        _req("Identify all documents supporting the termination date.")
    )
    assert result.valid is False


def test_valid_on_explicit_expert_keyword():
    result = validate_obj_3_expert_opinion(
        _req("State the expert opinion of your accident reconstruction expert.")
    )
    assert result.valid is True


def test_valid_on_contention_request():
    result = validate_obj_3_expert_opinion(
        _req("State all facts supporting your contention that X was negligent.")
    )
    assert result.valid is True


def test_valid_on_legal_theory_request():
    result = validate_obj_3_expert_opinion(
        _req("Describe the legal theory under which you seek recovery.")
    )
    assert result.valid is True


def test_answer_aware_drop_on_factual_answer():
    """Even without a trigger word, if the answer is factual, drop it."""
    result = validate_obj_3_expert_opinion(
        _req("State all facts supporting your contention."),
        answer_text="The employee was terminated on January 13, 2023.",
    )
    # Trigger word 'contention' is present → valid regardless of answer.
    # This test documents that the request-side gate takes precedence.
    assert result.valid is True


def test_answer_aware_drop_when_no_trigger_and_no_answer():
    result = validate_obj_3_expert_opinion(
        _req("Was X an employee?"),
        answer_text=None,
    )
    assert result.valid is False
```

- [ ] **Step 2: Run — verify it fails**

```bash
python -m pytest tests/test_discovery/test_validator_obj_3.py -v
```

Expected: ImportError on `validate_obj_3_expert_opinion`.

- [ ] **Step 3: Implement the gate**

Edit `icharlotte_core/discovery/objection_validator.py`. Add after the `VALIDATORS` declaration:

```python
import re as _re

_EXPERT_TRIGGERS = _re.compile(
    r"\b(expert|opinion\s+of|contention|contend[s]?|legal\s+theory|legal\s+basis)\b",
    _re.IGNORECASE,
)


def validate_obj_3_expert_opinion(
    request: ParsedRequest,
    answer_text: Optional[str] = None,
) -> ObjectionValidationResult:
    """
    Valid only when the request explicitly seeks expert opinion, a legal
    contention, or a legal theory. Pure factual questions do NOT warrant #3.
    """
    if _EXPERT_TRIGGERS.search(request.text):
        return ObjectionValidationResult(valid=True)
    return ObjectionValidationResult(
        valid=False,
        reason="request contains no expert/opinion/contention keyword",
    )


VALIDATORS[3] = validate_obj_3_expert_opinion
```

Move the `import re as _re` to the top of the file if preferred (import `re` and use it directly — adjust all references).

- [ ] **Step 4: Run the test**

```bash
python -m pytest tests/test_discovery/test_validator_obj_3.py -v
```

Expected: all 7 tests pass.

- [ ] **Step 5: Commit**

```bash
git add icharlotte_core/discovery/objection_validator.py tests/test_discovery/test_validator_obj_3.py
git commit -m "feat(discovery): add validator gate for objection #3 (expert opinion)"
```

---

### Task 8: Implement `validate_obj_5_argumentative` gate

**Files:**
- Modify: `icharlotte_core/discovery/objection_validator.py`
- Create test: `tests/test_discovery/test_validator_obj_5.py`

**Rule:** Objection #5 (argumentative / requires adoption of an assumption) is defensible when the request contains an embedded factual assertion the client contests. The simple heuristic: look for "state whether X was aware that Y" or "explain why you did Z" patterns where the phrasing presupposes a disputed fact. In answer-aware mode: if the answer accepts the premise (starts with "Yes" or contains the same factual assertion affirmatively), drop the objection — you cannot claim a premise is argumentative and then accept it.

Because embedded-assertion detection is hard in pure code, the request-side gate is a soft filter: it's valid by default unless the answer-side check explicitly drops it. This is a relaxation from the ideal, but safer than a false-negative on legitimately argumentative requests.

- [ ] **Step 1: Write the tests**

Create `tests/test_discovery/test_validator_obj_5.py`:

```python
"""Tests for validate_obj_5_argumentative."""

from icharlotte_core.discovery.objection_validator import (
    validate_obj_5_argumentative,
)
from icharlotte_core.discovery.response_parser import ParsedRequest


def _req(text):
    return ParsedRequest(number="1", text=text)


def test_valid_when_no_answer_context():
    """Without answer context, default to valid (safe)."""
    result = validate_obj_5_argumentative(
        _req("State whether X was aware Y was collecting payments.")
    )
    assert result.valid is True


def test_invalid_when_answer_accepts_premise():
    """'Yes' answers drop the argumentative objection."""
    result = validate_obj_5_argumentative(
        _req("State whether X was aware Y was collecting payments."),
        answer_text="Yes, during the period of his employment.",
    )
    assert result.valid is False
    assert "accept" in result.reason.lower() or "premise" in result.reason.lower()


def test_valid_when_answer_denies_premise():
    result = validate_obj_5_argumentative(
        _req("State whether X was aware Y was collecting payments."),
        answer_text="No.",
    )
    assert result.valid is True


def test_invalid_when_answer_starts_with_yes_comma():
    result = validate_obj_5_argumentative(
        _req("State whether X was aware Y was collecting payments."),
        answer_text="Yes, X was aware.",
    )
    assert result.valid is False


def test_valid_when_answer_is_not_applicable():
    """'Not applicable' isn't an acceptance of the premise."""
    result = validate_obj_5_argumentative(
        _req("If X occurred, state why."),
        answer_text="Not applicable.",
    )
    assert result.valid is True
```

- [ ] **Step 2: Run — verify it fails**

```bash
python -m pytest tests/test_discovery/test_validator_obj_5.py -v
```

Expected: ImportError.

- [ ] **Step 3: Implement the gate**

Add to `objection_validator.py`:

```python
_YES_PREFIX = _re.compile(r"^\s*yes\b", _re.IGNORECASE)


def validate_obj_5_argumentative(
    request: ParsedRequest,
    answer_text: Optional[str] = None,
) -> ObjectionValidationResult:
    """
    Objection #5 (argumentative / requires assumption) is soft-valid by
    default. It is dropped only when the substantive answer affirmatively
    accepts the premise — you cannot claim a premise is argumentative and
    then accept it in the answer.
    """
    if answer_text is None:
        return ObjectionValidationResult(valid=True)

    stripped = answer_text.strip()
    if _YES_PREFIX.match(stripped):
        return ObjectionValidationResult(
            valid=False,
            reason="answer accepts the premise ('Yes...'), so objection that the premise is argumentative is waived",
        )

    return ObjectionValidationResult(valid=True)


VALIDATORS[5] = validate_obj_5_argumentative
```

- [ ] **Step 4: Run the test**

```bash
python -m pytest tests/test_discovery/test_validator_obj_5.py -v
```

Expected: all 5 tests pass.

- [ ] **Step 5: Commit**

```bash
git add icharlotte_core/discovery/objection_validator.py tests/test_discovery/test_validator_obj_5.py
git commit -m "feat(discovery): add validator gate for objection #5 (argumentative)"
```

---

### Task 9: Implement `validate_obj_6_burden_time` gate

**Files:**
- Modify: `icharlotte_core/discovery/objection_validator.py`
- Create test: `tests/test_discovery/test_validator_obj_6.py`

**Rule:** Objection #6 ("unduly burdensome / overly broad and unlimited as to time and scope") is only defensible when the request has no explicit temporal limit. Temporal limits include: month-name years ("during 2022"), date ranges ("between January 2022 and March 2023"), "prior to" / "after" specific dates, "during the period of employment", explicit year spans ("2021–2023"). Answer-aware: if the answer contains a specific date, also drop.

- [ ] **Step 1: Write the tests**

Create `tests/test_discovery/test_validator_obj_6.py`:

```python
"""Tests for validate_obj_6_burden_time."""

from icharlotte_core.discovery.objection_validator import (
    validate_obj_6_burden_time,
)
from icharlotte_core.discovery.response_parser import ParsedRequest


def _req(text):
    return ParsedRequest(number="1", text=text)


def test_valid_on_unlimited_request():
    result = validate_obj_6_burden_time(
        _req("Identify all documents relating to X.")
    )
    assert result.valid is True


def test_invalid_on_prior_to_date():
    result = validate_obj_6_burden_time(
        _req("Identify all contractors you informed prior to March 8, 2023.")
    )
    assert result.valid is False


def test_invalid_on_after_date():
    result = validate_obj_6_burden_time(
        _req("Identify all communications after January 1, 2023.")
    )
    assert result.valid is False


def test_invalid_on_year_range():
    result = validate_obj_6_burden_time(
        _req("Identify all projects during 2022-2023.")
    )
    assert result.valid is False


def test_invalid_on_during_year():
    result = validate_obj_6_burden_time(
        _req("State whether X was aware Y was making representations during 2022.")
    )
    assert result.valid is False


def test_invalid_when_answer_contains_date():
    """Answer-aware: date in answer implies request was bounded."""
    result = validate_obj_6_burden_time(
        _req("Identify the termination date."),
        answer_text="On or about January 13, 2023.",
    )
    assert result.valid is False


def test_answer_containing_year_only_still_drops():
    result = validate_obj_6_burden_time(
        _req("Identify all projects X worked on."),
        answer_text="Various projects during 2022.",
    )
    # Answer contains "2022" which is a concrete time anchor
    assert result.valid is False
```

- [ ] **Step 2: Run — verify it fails**

```bash
python -m pytest tests/test_discovery/test_validator_obj_6.py -v
```

- [ ] **Step 3: Implement the gate**

Add to `objection_validator.py`:

```python
# Patterns indicating explicit time limits in the request OR a dated answer
_TIME_LIMIT_PATTERNS = [
    _re.compile(r"\bprior\s+to\b", _re.IGNORECASE),
    _re.compile(r"\bafter\b\s+(?:January|February|March|April|May|June|July|August|September|October|November|December|\d)", _re.IGNORECASE),
    _re.compile(r"\bbefore\b\s+(?:January|February|March|April|May|June|July|August|September|October|November|December|\d)", _re.IGNORECASE),
    _re.compile(r"\bbetween\b.+?\band\b", _re.IGNORECASE),
    _re.compile(r"\bduring\b\s+(?:\d{4}|the\s+period|his\s+employment|her\s+employment|their\s+employment)", _re.IGNORECASE),
    _re.compile(r"\d{4}\s*[-–]\s*\d{4}"),  # year range like "2022-2023"
    _re.compile(r"\b(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2},?\s*\d{4}\b", _re.IGNORECASE),
    _re.compile(r"\b\d{1,2}/\d{1,2}/\d{2,4}\b"),
]


def _has_time_limit(text: str) -> bool:
    return any(p.search(text) for p in _TIME_LIMIT_PATTERNS)


def validate_obj_6_burden_time(
    request: ParsedRequest,
    answer_text: Optional[str] = None,
) -> ObjectionValidationResult:
    """
    Valid only when the request has NO explicit temporal limit AND the
    answer (if provided) does NOT contain a concrete date anchor.
    """
    if _has_time_limit(request.text):
        return ObjectionValidationResult(
            valid=False,
            reason="request has an explicit time limit; the 'unlimited as to time' objection is factually wrong",
        )

    if answer_text is not None and _has_time_limit(answer_text):
        return ObjectionValidationResult(
            valid=False,
            reason="answer contains a specific date, so 'unlimited as to time' objection contradicts the answer",
        )

    # Also catch bare year mentions in the answer (e.g., "during 2022")
    if answer_text is not None and _re.search(r"\b\d{4}\b", answer_text):
        return ObjectionValidationResult(
            valid=False,
            reason="answer contains a year anchor",
        )

    return ObjectionValidationResult(valid=True)


VALIDATORS[6] = validate_obj_6_burden_time
```

- [ ] **Step 4: Run — verify tests pass**

```bash
python -m pytest tests/test_discovery/test_validator_obj_6.py -v
```

Expected: all 7 tests pass. If the year-in-answer regex creates false positives on unrelated four-digit numbers (unlikely but possible), tighten by requiring the year to be 19xx or 20xx.

- [ ] **Step 5: Commit**

```bash
git add icharlotte_core/discovery/objection_validator.py tests/test_discovery/test_validator_obj_6.py
git commit -m "feat(discovery): add validator gate for objection #6 (burden/time)"
```

---

### Task 10: Implement `validate_obj_7_list_summary` gate

**Files:**
- Modify: `icharlotte_core/discovery/objection_validator.py`
- Create test: `tests/test_discovery/test_validator_obj_7.py`

**Rule:** Objection #7 ("list/summary not in existence") is only defensible when the request asks for identification of multiple items that would require compilation effort. Answer-aware: if the answer is ≤1 items, "None", or "Not applicable", drop the objection — you cannot claim compiling a list would be burdensome when the list has zero or one items.

- [ ] **Step 1: Write the tests**

Create `tests/test_discovery/test_validator_obj_7.py`:

```python
"""Tests for validate_obj_7_list_summary."""

from icharlotte_core.discovery.objection_validator import (
    validate_obj_7_list_summary,
)
from icharlotte_core.discovery.response_parser import ParsedRequest


def _req(text):
    return ParsedRequest(number="1", text=text)


def test_valid_on_list_request_without_answer():
    result = validate_obj_7_list_summary(
        _req("Identify all communications between X and Y.")
    )
    assert result.valid is True


def test_invalid_when_answer_is_none():
    result = validate_obj_7_list_summary(
        _req("Identify all communications between X and Y."),
        answer_text="None.",
    )
    assert result.valid is False
    assert "none" in result.reason.lower() or "single" in result.reason.lower() or "empty" in result.reason.lower()


def test_invalid_when_answer_is_not_applicable():
    result = validate_obj_7_list_summary(
        _req("Identify all payments."),
        answer_text="Not applicable.",
    )
    assert result.valid is False


def test_invalid_when_answer_is_single_name():
    result = validate_obj_7_list_summary(
        _req("Identify all individuals who supervised X."),
        answer_text="Milo Holte.",
    )
    assert result.valid is False


def test_invalid_when_answer_has_one_item_list_structure():
    """Single-item list also triggers the drop."""
    result = validate_obj_7_list_summary(
        _req("Identify all projects X worked on."),
        answer_text="The Acme pool project.",
    )
    assert result.valid is False


def test_valid_when_answer_has_multiple_items():
    result = validate_obj_7_list_summary(
        _req("Identify all projects X worked on."),
        answer_text="The Acme project, the Smith residence, and the Jones estate.",
    )
    assert result.valid is True


def test_valid_when_no_answer_context():
    result = validate_obj_7_list_summary(
        _req("Identify all documents.")
    )
    assert result.valid is True
```

- [ ] **Step 2: Run — verify it fails**

```bash
python -m pytest tests/test_discovery/test_validator_obj_7.py -v
```

- [ ] **Step 3: Implement the gate**

Add to `objection_validator.py`:

```python
_EMPTY_ANSWERS = {"none", "none.", "not applicable", "not applicable.", "n/a", "n/a."}


def _count_list_items(answer: str) -> int:
    """
    Very rough count of discrete items in an answer. Used to decide whether
    the 'list/summary' objection is defensible.

    Heuristic: count comma-separated segments excluding the trailing
    'Discovery is ongoing' boilerplate. A single item yields 1; 'None'
    yields 0; a full sentence with no commas yields 1.
    """
    # Strip the standard trailer if present
    trailer_markers = [
        "Discovery and investigation are ongoing",
        "Responding Party reserves",
    ]
    working = answer
    for marker in trailer_markers:
        idx = working.find(marker)
        if idx >= 0:
            working = working[:idx]
    working = working.strip().rstrip(".").strip()

    if not working:
        return 0
    if working.lower() in _EMPTY_ANSWERS or working.lower().rstrip(".") in _EMPTY_ANSWERS:
        return 0

    # Count comma-separated segments; subtract 1 if the last segment starts
    # with "and"/"or" (Oxford-comma style: "A, B, and C" = 3 items not 4).
    segments = [s.strip() for s in working.split(",") if s.strip()]
    if len(segments) == 1:
        return 1
    return len(segments)


def validate_obj_7_list_summary(
    request: ParsedRequest,
    answer_text: Optional[str] = None,
) -> ObjectionValidationResult:
    """
    Objection #7 is defensible when the request asks for compilation of
    multiple items. It is NOT defensible when the answer reveals that the
    actual list has zero or one items.
    """
    if answer_text is None:
        return ObjectionValidationResult(valid=True)

    count = _count_list_items(answer_text)
    if count <= 1:
        return ObjectionValidationResult(
            valid=False,
            reason=f"answer contains {count} item(s); 'list/summary not in existence' objection contradicts the answer",
        )

    return ObjectionValidationResult(valid=True)


VALIDATORS[7] = validate_obj_7_list_summary
```

- [ ] **Step 4: Run**

```bash
python -m pytest tests/test_discovery/test_validator_obj_7.py -v
```

Expected: all 7 tests pass.

- [ ] **Step 5: Commit**

```bash
git add icharlotte_core/discovery/objection_validator.py tests/test_discovery/test_validator_obj_7.py
git commit -m "feat(discovery): add validator gate for objection #7 (list/summary)"
```

---

### Task 11: Implement `validate_obj_9_compound` gate

**Files:**
- Modify: `icharlotte_core/discovery/objection_validator.py`
- Create test: `tests/test_discovery/test_validator_obj_9.py`

**Rule:** Objection #9 (compound) is valid only when `ParsedRequest.is_compound` is True (which is now accurate post-Task 2). Answer-aware: if the answer is a single atomic fact, drop the objection regardless of `is_compound`.

- [ ] **Step 1: Write the tests**

Create `tests/test_discovery/test_validator_obj_9.py`:

```python
"""Tests for validate_obj_9_compound."""

from icharlotte_core.discovery.objection_validator import validate_obj_9_compound
from icharlotte_core.discovery.response_parser import ParsedRequest


def test_valid_when_request_is_compound():
    req = ParsedRequest(
        number="11",
        text="Identify each customer notified, the date of notification, and the method.",
        is_compound=True,
    )
    result = validate_obj_9_compound(req)
    assert result.valid is True


def test_invalid_when_request_is_not_compound():
    req = ParsedRequest(
        number="1",
        text="Was X an employee?",
        is_compound=False,
    )
    result = validate_obj_9_compound(req)
    assert result.valid is False


def test_invalid_when_answer_is_atomic_even_if_compound_flagged():
    req = ParsedRequest(
        number="24",
        text="Identify all individuals who supervised or directed X.",
        is_compound=True,
    )
    result = validate_obj_9_compound(req, answer_text="Milo Holte.")
    assert result.valid is False


def test_valid_when_answer_has_multiple_facts_and_compound():
    req = ParsedRequest(
        number="11",
        text="Identify each customer, date, and method.",
        is_compound=True,
    )
    result = validate_obj_9_compound(
        req,
        answer_text="James Lindly, verbally, on or about March 8, 2023.",
    )
    assert result.valid is True
```

- [ ] **Step 2: Run — verify it fails**

```bash
python -m pytest tests/test_discovery/test_validator_obj_9.py -v
```

- [ ] **Step 3: Implement the gate**

Add to `objection_validator.py`:

```python
def validate_obj_9_compound(
    request: ParsedRequest,
    answer_text: Optional[str] = None,
) -> ObjectionValidationResult:
    """
    Valid only when the parser flagged the request as compound. Also
    dropped if the substantive answer is a single atomic fact — you
    cannot object that a request is compound and then answer it once.
    """
    if not request.is_compound:
        return ObjectionValidationResult(
            valid=False,
            reason="parser did not flag this request as compound",
        )

    if answer_text is not None:
        count = _count_list_items(answer_text)
        if count <= 1:
            return ObjectionValidationResult(
                valid=False,
                reason="answer is a single fact; cannot claim compound when answered atomically",
            )

    return ObjectionValidationResult(valid=True)


VALIDATORS[9] = validate_obj_9_compound
```

- [ ] **Step 4: Run**

```bash
python -m pytest tests/test_discovery/test_validator_obj_9.py -v
```

- [ ] **Step 5: Re-enable the scaffold dispatcher test**

Edit `tests/test_discovery/test_objection_validator_scaffold.py`. Remove the `@pytest.mark.skip` decorator from `test_validators_dispatcher_contains_expected_gates` now that all five gates are populated.

```bash
python -m pytest tests/test_discovery/test_objection_validator_scaffold.py -v
```

Expected: 5 passed.

- [ ] **Step 6: Commit**

```bash
git add icharlotte_core/discovery/objection_validator.py tests/test_discovery/test_validator_obj_9.py tests/test_discovery/test_objection_validator_scaffold.py
git commit -m "feat(discovery): add validator gate for objection #9 (compound) and enable dispatcher test"
```

---

### Task 12: Wire validator into initial objection selection

**Files:**
- Modify: `icharlotte_core/ui/respond_tab.py:753` (`_on_phase2_finished()`) — after `merge_objections`, apply `filter_objection_ids`
- Create test: `tests/test_discovery/test_validator_integration.py`

- [ ] **Step 1: Read the current `_on_phase2_finished` handler**

```bash
```

Open `icharlotte_core/ui/respond_tab.py` around line 753 and read the method. Understand how merged objections are stored per request.

- [ ] **Step 2: Write an integration test**

Create `tests/test_discovery/test_validator_integration.py`:

```python
"""Integration test for filter_objection_ids wiring into the selection flow."""

from icharlotte_core.discovery.objection_validator import filter_objection_ids
from icharlotte_core.discovery.response_parser import ParsedRequest


def test_merged_set_dropped_to_empty_on_plain_yes_no():
    """SI 1: 'Was X an employee?' — LLM picks {3, 9}, validator drops both."""
    req = ParsedRequest(
        number="1",
        text="Was Edgar Chavez ever an employee for Premier Gunite?",
        is_compound=False,
    )
    merged = {3, 9}
    kept, dropped = filter_objection_ids(merged, req)
    assert kept == set()
    assert len(dropped) == 2


def test_privilege_preserved_on_time_bounded_request():
    """Privilege (#4) is never dropped even if request has time limit."""
    req = ParsedRequest(
        number="30",
        text="Identify contractors informed prior to March 8, 2023.",
        is_compound=False,
    )
    merged = {4, 6}
    kept, dropped = filter_objection_ids(merged, req)
    assert 4 in kept
    assert 6 not in kept  # dropped because request has "prior to <date>"
```

- [ ] **Step 3: Run — should pass from the validator work already done**

```bash
python -m pytest tests/test_discovery/test_validator_integration.py -v
```

Expected: both pass.

- [ ] **Step 4: Wire validator into `_on_phase2_finished`**

Open `icharlotte_core/ui/respond_tab.py` and find the handler (starts around line 753). After the point where merged objections are assigned to a request (look for `merge_objections` or the use of `format_objections`), add a filter step. Concretely:

At the top of respond_tab.py, in the existing import from `icharlotte_core.discovery.objection_selector`, also import the validator:

```python
from icharlotte_core.discovery.objection_validator import filter_objection_ids
```

Inside `_on_phase2_finished`, wherever the merged objection ID set is computed per request, wrap it:

```python
# Before: merged_ids was the union of rule-based and LLM-selected objections.
# Now: filter through the validator to drop indefensible picks before storing.
merged_ids = merge_objections(rule_ids, llm_ids)
kept_ids, dropped = filter_objection_ids(merged_ids, req)
if dropped:
    logger.debug(
        "Phase 2 filter dropped objections for SI %s: %s",
        req.number,
        dropped,
    )
# Use kept_ids from here on instead of merged_ids
```

If the handler stores the formatted string directly (via `format_objections(merged_ids, menu)`), pass `kept_ids` to `format_objections` instead.

- [ ] **Step 5: Smoke test the app**

```bash
python iCharlotte.py
```

Run the respond tab on a small SI set (FI test case or any real case) and confirm no crashes. The live regression run happens in Task 18.

- [ ] **Step 6: Commit**

```bash
git add icharlotte_core/ui/respond_tab.py tests/test_discovery/test_validator_integration.py
git commit -m "feat(discovery): filter initial objection selection through validator"
```

---

### Task 13: Create `objection_pruner` module with answer-aware prune pass

**Files:**
- Create: `icharlotte_core/discovery/objection_pruner.py`
- Create test: `tests/test_discovery/test_objection_pruner.py`

- [ ] **Step 1: Write tests**

Create `tests/test_discovery/test_objection_pruner.py`:

```python
"""Tests for the post-draft answer-aware objection prune pass."""

from icharlotte_core.discovery.objection_pruner import prune_objections_against_answer
from icharlotte_core.discovery.response_parser import ParsedRequest


def _req(text, is_compound=False):
    return ParsedRequest(number="1", text=text, is_compound=is_compound)


def test_prunes_list_summary_on_single_item_answer():
    kept, dropped = prune_objections_against_answer(
        {7},
        _req("Identify all individuals who supervised X."),
        answer_text="Milo Holte.",
    )
    assert kept == set()
    assert dropped[0][0] == 7


def test_prunes_compound_on_atomic_answer():
    kept, dropped = prune_objections_against_answer(
        {9},
        _req("Identify all individuals who supervised or directed X.", is_compound=True),
        answer_text="Milo Holte.",
    )
    assert kept == set()


def test_prunes_burden_time_on_dated_answer():
    kept, dropped = prune_objections_against_answer(
        {6},
        _req("Identify termination steps."),
        answer_text="On or about January 13, 2023.",
    )
    assert kept == set()


def test_preserves_privilege_on_no_answer():
    """Privilege objection is NEVER pruned, even on 'No' answers."""
    kept, dropped = prune_objections_against_answer(
        {4},
        _req("State whether X conducted internal investigation."),
        answer_text="No.",
    )
    assert 4 in kept
    assert dropped == []


def test_preserves_privacy_objection():
    """Objection #2 (privacy) is user-controlled and never pruned."""
    kept, dropped = prune_objections_against_answer(
        {2},
        _req("State whether X was disciplined."),
        answer_text="No.",
    )
    assert 2 in kept
    assert dropped == []


def test_passes_through_when_answer_is_human_input_flag():
    """Don't prune anything when the answer is a refusal token."""
    kept, dropped = prune_objections_against_answer(
        {3, 6, 7, 9},
        _req("Identify all projects X worked on."),
        answer_text="[NEEDS HUMAN INPUT: specific project names]",
    )
    assert kept == {3, 6, 7, 9}
    assert dropped == []


def test_drops_expert_opinion_on_factual_request():
    """Validator-level drop still applies in prune mode."""
    kept, dropped = prune_objections_against_answer(
        {3},
        _req("Was X an employee?"),
        answer_text="Yes.",
    )
    assert kept == set()
```

- [ ] **Step 2: Run — verify it fails**

```bash
python -m pytest tests/test_discovery/test_objection_pruner.py -v
```

Expected: ImportError.

- [ ] **Step 3: Create the pruner module**

Create `icharlotte_core/discovery/objection_pruner.py`:

```python
"""
Phase 2 post-draft objection prune pass.

Runs after the drafter produces a substantive answer. Removes objections
that are contradicted or rendered moot by the answer, using the shared
validator gates in `objection_validator.py`.

Rules NEVER applied (per design):
- Objection #4 (attorney-client / work product) is never pruned.
- Objection #2 (privacy) is user-controlled and never pruned.
- If the answer is a [NEEDS HUMAN INPUT:] refusal token, NO pruning happens
  — we don't yet know what the final answer will look like.
"""
from typing import List, Set, Tuple

from icharlotte_core.discovery.objection_validator import filter_objection_ids
from icharlotte_core.discovery.response_parser import ParsedRequest


def prune_objections_against_answer(
    objection_ids: Set[int],
    request: ParsedRequest,
    answer_text: str,
) -> Tuple[Set[int], List[Tuple[int, str]]]:
    """
    Post-draft objection pruning.

    Returns (kept_ids, dropped_with_reasons). Same semantics as
    filter_objection_ids but always runs in answer-aware mode.
    """
    # Refusal token → preserve everything. We'll re-evaluate after the
    # user resolves the flag.
    if answer_text.strip().startswith("[NEEDS HUMAN INPUT:"):
        return set(objection_ids), []

    return filter_objection_ids(objection_ids, request, answer_text=answer_text)
```

- [ ] **Step 4: Run the tests**

```bash
python -m pytest tests/test_discovery/test_objection_pruner.py -v
```

Expected: all 7 tests pass.

- [ ] **Step 5: Commit**

```bash
git add icharlotte_core/discovery/objection_pruner.py tests/test_discovery/test_objection_pruner.py
git commit -m "feat(discovery): add answer-aware objection pruner"
```

---

### Task 14: Add few-shot calibration to LLM objection selection prompt

**Files:**
- Modify: `icharlotte_core/discovery/objection_selector.py:174-211` (`_AGGRESSIVENESS_INSTRUCTIONS` and `build_objection_prompt()`)
- Create test: `tests/test_discovery/test_objection_prompt_calibration.py`

- [ ] **Step 1: Write tests**

Create `tests/test_discovery/test_objection_prompt_calibration.py`:

```python
"""Tests that the objection prompt uses few-shot calibration blocks."""

from icharlotte_core.discovery.objection_selector import (
    ObjectionMenu,
    build_objection_prompt,
)


def test_conservative_prompt_has_examples():
    menu = ObjectionMenu.load_defaults()
    prompt = build_objection_prompt(
        "Was Edgar Chavez ever an employee?",
        menu,
        aggressiveness="conservative",
    )
    assert "CONSERVATIVE" in prompt.upper()
    assert "Examples" in prompt or "Example" in prompt
    # Must warn about #3 on factual yes/no
    assert "factual yes/no" in prompt.lower() or "expert" in prompt.lower()


def test_conservative_prompt_warns_about_synonymous_alternatives():
    menu = ObjectionMenu.load_defaults()
    prompt = build_objection_prompt(
        "Was X an employee?",
        menu,
        aggressiveness="conservative",
    )
    assert "synonymous alternatives" in prompt.lower() or "synonymous" in prompt.lower()


def test_aggressive_prompt_still_works():
    menu = ObjectionMenu.load_defaults()
    prompt = build_objection_prompt(
        "Was X an employee?",
        menu,
        aggressiveness="aggressive",
    )
    assert "AGGRESSIVE" in prompt.upper() or "aggressive" in prompt.lower()
    # Should encourage over-inclusion
    assert "over-inclusion" in prompt.lower() or "waives" in prompt.lower()


def test_moderate_prompt_exists():
    menu = ObjectionMenu.load_defaults()
    prompt = build_objection_prompt(
        "Was X an employee?",
        menu,
        aggressiveness="moderate",
    )
    assert len(prompt) > 100


def test_unknown_aggressiveness_defaults_to_aggressive():
    menu = ObjectionMenu.load_defaults()
    prompt = build_objection_prompt(
        "Was X an employee?",
        menu,
        aggressiveness="unknown_level",
    )
    # Falls back to aggressive per current behavior (line 200)
    assert "waives" in prompt.lower() or "over-inclusion" in prompt.lower()
```

- [ ] **Step 2: Run — should fail**

```bash
python -m pytest tests/test_discovery/test_objection_prompt_calibration.py -v
```

- [ ] **Step 3: Replace the aggressiveness instructions and prompt builder**

Edit `icharlotte_core/discovery/objection_selector.py:174-211`. Replace the existing `_AGGRESSIVENESS_INSTRUCTIONS` dict and `build_objection_prompt()` function:

```python
_CONSERVATIVE_BLOCK = """AGGRESSIVENESS: CONSERVATIVE

Only select objections that clearly and directly apply to this specific
request. When in doubt, do NOT select the objection. Over-selection under
this setting is a failure.

Specifically DO NOT select:
  - #3 (expert opinion / legal conclusion) unless the request literally
    contains "expert", "opinion of", "contention", or asks the responder
    to state a legal theory. Factual yes/no questions are NOT expert opinion.
  - #6 (burden / overbroad as to time) if the request contains any
    explicit time limit ("during 2022", "prior to March 8, 2023",
    "after January 1", "between X and Y").
  - #7 (list/summary not in existence) if the information requested is a
    small number of discrete facts readily knowable by the client
    (e.g., "who supervised X", "what was the termination date").
  - #9 (compound) unless the request genuinely asks for two or more
    substantively different pieces of information. Synonymous alternatives
    joined by "or" are NOT compound (e.g., "employee or independent
    contractor", "recover, disable, or revoke").

Examples of correct conservative selection:

Request: "Was John Smith ever an employee of Acme Corp?"
Correct objection IDs: (none)
NOT selected: #3 (no expert opinion sought), #9 (not compound — synonymous
alternatives).

Request: "State whether Acme was aware Smith was collecting payments during 2022."
Correct objection IDs: (none, or possibly #5 if client disputes the premise)
NOT selected: #3 (no expert opinion), #6 (time-bounded to 2022), #9
(not compound — single yes/no about awareness).

Request: "Identify the name, address, and telephone number of each witness."
Correct objection IDs: #9 (compound — three distinct information targets)
NOT selected: #3, #6, #7 (a reasonable number of witnesses is not burdensome).

Request: "State all facts supporting your contention of negligence."
Correct objection IDs: #3 (contention request)
"""

_MODERATE_BLOCK = """AGGRESSIVENESS: MODERATE

Select objections that are reasonably applicable. Prefer under-selection
when a judgment call is close — it is better to miss a borderline
objection than to stack the response with objections the opposing counsel
will mock in a meet-and-confer.
"""

_AGGRESSIVE_BLOCK = """AGGRESSIVENESS: AGGRESSIVE

Select ALL objections that might plausibly apply. Err on the side of
over-inclusion — a failure to include an objection waives it permanently.
If an objection might in theory apply, include it.
"""

_AGGRESSIVENESS_INSTRUCTIONS = {
    "aggressive": _AGGRESSIVE_BLOCK,
    "moderate": _MODERATE_BLOCK,
    "conservative": _CONSERVATIVE_BLOCK,
}


def build_objection_prompt(
    request_text: str,
    menu: ObjectionMenu,
    aggressiveness: str = "aggressive",
) -> str:
    """
    Build an LLM prompt for objection selection.

    Args:
        request_text: The full text of the discovery request.
        menu: The ObjectionMenu containing available objections.
        aggressiveness: One of 'aggressive', 'moderate', 'conservative'.

    Returns:
        A prompt string ready to send to an LLM.
    """
    instruction = _AGGRESSIVENESS_INSTRUCTIONS.get(
        aggressiveness, _AGGRESSIVENESS_INSTRUCTIONS["aggressive"]
    )

    return (
        f"You are a California civil litigation defense attorney reviewing "
        f"a discovery request and deciding which objections apply.\n\n"
        f"DISCOVERY REQUEST:\n{request_text}\n\n"
        f"AVAILABLE OBJECTIONS:\n{menu.all_text()}\n\n"
        f"{instruction}\n\n"
        f"Return ONLY the objection IDs that apply as a comma-separated list of "
        f"numbers (e.g., 1, 3, 6). Do not include any other text."
    )
```

- [ ] **Step 4: Run the test**

```bash
python -m pytest tests/test_discovery/test_objection_prompt_calibration.py -v
```

Expected: all 5 pass.

- [ ] **Step 5: Commit**

```bash
git add icharlotte_core/discovery/objection_selector.py tests/test_discovery/test_objection_prompt_calibration.py
git commit -m "feat(discovery): add few-shot calibration to conservative objection prompt"
```

---

### Task 15: Wire the prune pass into the drafting pipeline

**Files:**
- Modify: `icharlotte_core/ui/respond_tab.py:998` (`_on_combined_draft_finished()`) — after parsing responses, run prune pass per request

- [ ] **Step 1: Read the current `_on_combined_draft_finished` handler**

Open `icharlotte_core/ui/respond_tab.py` at line 998. The current flow: parse the combined LLM response, store each answer in `responses_map[num] = text`, then call `_finalize_response()`. The objections are already in `objections_map[num]`. The prune pass needs to run for each (num, request, objection_ids, answer_text) tuple before `_finalize_response`.

**Complication:** `objections_map` stores **formatted strings**, not objection ID sets. The prune pass needs the ID set. We need a parallel map `objection_ids_map: Dict[str, Set[int]]` that carries the IDs forward from `_on_phase2_finished` through to `_on_combined_draft_finished`.

- [ ] **Step 2: Thread the objection ID set through the pipeline**

Edit `_on_phase2_finished` in respond_tab.py. Where it currently builds `objections_map[num] = format_objections(...)`, also store the underlying ID set:

```python
# Existing: objections_map[req.number] = format_objections(kept_ids, self._objection_menu, ...)
# Add a parallel map on self for the IDs
if not hasattr(self, "_objection_ids_map") or self._objection_ids_map is None:
    self._objection_ids_map = {}
self._objection_ids_map[req.number] = kept_ids
```

Initialize `self._objection_ids_map = {}` in `__init__` alongside `self._objection_menu` (around line 101).

- [ ] **Step 3: Run the prune pass in `_on_combined_draft_finished`**

Inside `_on_combined_draft_finished` at line 998, after the response parsing loop that populates `responses_map`, and before `self._finalize_response(...)`:

```python
# Post-draft prune pass: drop objections contradicted or rendered moot
# by the answer. Uses the shared validator via objection_pruner.
from icharlotte_core.discovery.objection_pruner import prune_objections_against_answer
from icharlotte_core.discovery.objection_selector import format_objections

# Build a request lookup for quick access
req_by_number = {req.number: req for req in parsed.requests}

for num, answer_text in responses_map.items():
    req = req_by_number.get(num)
    if req is None:
        continue
    current_ids = self._objection_ids_map.get(num, set())
    if not current_ids:
        continue
    kept_ids, dropped = prune_objections_against_answer(
        current_ids, req, answer_text
    )
    if dropped:
        logger.info(
            "Pruned objections for %s: %s",
            num,
            [(oid, reason) for oid, reason in dropped],
        )
    self._objection_ids_map[num] = kept_ids
    # Re-format the objection string with the pruned set
    objections_map[num] = format_objections(kept_ids, self._objection_menu)

self._finalize_response(parsed, objections_map, responses_map, file_label)
```

- [ ] **Step 4: Smoke test in the app**

```bash
python iCharlotte.py
```

Run a small SI set through the Respond tab. Verify no crashes, and that the log shows "Pruned objections for ..." entries for at least some requests.

- [ ] **Step 5: Commit**

```bash
git add icharlotte_core/ui/respond_tab.py
git commit -m "feat(discovery): run post-draft objection prune pass before finalize"
```

---

### Task 16: UI — post-prune review panel with Restore action

**Files:**
- Modify: `icharlotte_core/ui/respond_tab.py` — the response display (likely `_display_result` around line 1143) and any edit flow

**Background:** With the prune pass in place, users need visibility into what was dropped and the ability to restore any dropped objection they disagree with. The simplest version (matches the design spec's minimal UX): append a "Dropped objections" section to each response's display showing dropped IDs, reasons, and a "Restore" action that re-adds them.

Because the existing respond tab uses a text editor for the final assembled output rather than per-response widgets, the minimum-viable UI implementation for this task is:

1. After the prune pass runs, **log dropped objections to a persistent per-response record**.
2. In the display area, **append a collapsed "Dropped objections" section to each response's text** (visible as a comment-style block the user can manually cut if they want to keep it dropped).
3. For a richer UI (buttons to restore), defer to a follow-up task if the current tab's editor isn't per-response.

This task implements the minimal version. Upgrading to interactive buttons is a follow-up noted in the completion report.

- [ ] **Step 1: Thread dropped objections through to the display**

In `_on_combined_draft_finished`, after the prune pass loop, store dropped details per response:

```python
if not hasattr(self, "_dropped_objections_map"):
    self._dropped_objections_map = {}
self._dropped_objections_map[num] = dropped  # List[Tuple[int, str]]
```

Initialize `self._dropped_objections_map = {}` in `__init__`.

- [ ] **Step 2: Append dropped-objection footer to displayed responses**

In `_display_result()` (or wherever the assembled response text is written to the editor widget), after the main response body, append a small informational block for each response with drops:

```python
# Pseudocode — adapt to the actual display mechanism:
def _build_display_text(self, parsed, objections_map, responses_map):
    lines = []
    for req in parsed.requests:
        num = req.number
        lines.append(f"RESPONSE TO {num}:")
        lines.append(objections_map.get(num, ""))
        lines.append("")
        lines.append(responses_map.get(num, ""))
        # Dropped-objection footer
        dropped = self._dropped_objections_map.get(num, []) if hasattr(self, "_dropped_objections_map") else []
        if dropped:
            lines.append("")
            lines.append("[// Dropped by validator/prune — manually restore if needed:")
            for oid, reason in dropped:
                lines.append(f"//   #{oid}: {reason}")
            lines.append("// ]")
        lines.append("")
        lines.append("")
    return "\n".join(lines)
```

The `[// ... // ]` comment-style block is visually distinct, easy to spot-delete if the user wants to re-add an objection, and does NOT go into the assembled .docx (verify this by checking `_save_response_doc()` — if it copies the editor text verbatim, add a strip pass that removes `[// ... // ]` blocks before saving).

- [ ] **Step 3: Add a strip pass in the save path**

In `_save_response_doc` (line 1235) or wherever the editor text is transferred to the assembler, add:

```python
import re as _re
text = self.editor.toPlainText()
# Strip dropped-objection comment blocks before assembly
text = _re.sub(
    r"\[// Dropped by validator/prune.*?// \]\s*",
    "",
    text,
    flags=_re.DOTALL,
)
# Continue with the existing assembly path using 'text'
```

- [ ] **Step 4: Manual smoke test**

Start the app, run a small SI set, verify that dropped-objection comments appear in the editor view under responses that had pruning, verify that saving to a .docx strips the comments from the output, verify the user can manually remove a `// #<id>` line to "restore" (they'd then need to re-add the objection text — for now, this minimal version just surfaces the drops).

If at this point a follow-up "interactive restore" UI is clearly warranted, document it in Task 17's regression notes.

- [ ] **Step 5: Commit**

```bash
git add icharlotte_core/ui/respond_tab.py
git commit -m "feat(respond): show pruned objections as comment footer with strip-on-save"
```

---

### Task 17: Phase 2 manual regression test against PREMIER SI Set One

**Files:** No code changes — end-to-end verification.

- [ ] **Step 1: Rerun the PREMIER SI Set One through the respond tab**

Same steps as Task 5 — load the PREMIER SI PDF, load the same context documents, set Conservative + Minimal, click Generate.

- [ ] **Step 2: Verify Phase 2 acceptance criteria**

From the spec section 5.6:

**Criterion 1:** Objection #3 appears ONLY on requests literally containing "expert", "opinion of", or "contention". Scan every response. It should NOT appear on SI 1, 2, 3, 5, 20, 22, 26 (which were flagged in the original run).

**Criterion 2:** Objection #7 does not appear on any response whose answer is a single item, "None", or "Not applicable". Check SI 9, 15, 19, 21, 24, 30 in particular.

**Criterion 3:** Objection #6 does not appear on SI 30 (time-bounded) or SI 4 (single date).

**Criterion 4:** Objection #9 appears only on SI 2, 11, 23, 24, 30.

**Criterion 5:** Objection #4 (privilege) is never dropped — verify it appears wherever it appeared in Phase 1 output (SI 12, 25, 26, 28 in the original).

**Criterion 6:** Dropped-objection comment blocks are visible in the editor for responses that had pruning.

**Criterion 7:** Unit test coverage — each validator gate has at least three tests (Tasks 7–11 created these).

- [ ] **Step 3: Append results to regression notes**

Edit `docs/superpowers/plans/phase-1-regression-notes.md` (from Task 5) and add a Phase 2 section. Capture: which SIs had objection drops, which still have over-firing, any criteria failures.

- [ ] **Step 4: Commit**

```bash
git add docs/superpowers/plans/phase-1-regression-notes.md
git commit -m "docs(discovery): phase 2 PREMIER regression notes"
```

**If any Phase 2 criterion failed, stop and fix before starting Phase 3.**

---

## Phase 3 — Cross-Question Consistency

### Task 18: Create `consistency_checker` module

**Files:**
- Create: `icharlotte_core/discovery/consistency_checker.py`
- Create test: `tests/test_discovery/test_consistency_checker.py`

- [ ] **Step 1: Write tests for the flag parser**

Create `tests/test_discovery/test_consistency_checker.py`:

```python
"""Tests for the cross-question consistency check module."""

from icharlotte_core.discovery.consistency_checker import (
    build_consistency_prompt,
    parse_consistency_flags,
    ConsistencyFlag,
)


def test_build_prompt_includes_all_responses():
    responses = {
        "1": ("Was X an employee?", "Yes."),
        "2": ("List dates.", "2022-01-01 through 2023-01-13."),
    }
    prompt = build_consistency_prompt(responses)
    assert "Was X an employee?" in prompt
    assert "Yes." in prompt
    assert "List dates." in prompt
    assert "SI 1:" in prompt or "RESPONSE 1" in prompt
    assert "SI 2:" in prompt or "RESPONSE 2" in prompt


def test_parse_no_contradictions():
    flags = parse_consistency_flags("NO CONTRADICTIONS")
    assert flags == []


def test_parse_single_flag():
    llm_output = """[CONSISTENCY FLAG: SI 10 contradicts SI 30] SI 10 says Yes to customer notification but SI 30 says None for contractors notified prior to March 8, 2023."""
    flags = parse_consistency_flags(llm_output)
    assert len(flags) == 1
    assert flags[0].si_a == "10"
    assert flags[0].si_b == "30"
    assert "customer notification" in flags[0].reason.lower() or "notification" in flags[0].reason.lower()


def test_parse_multiple_flags():
    llm_output = """
[CONSISTENCY FLAG: SI 10 contradicts SI 30] Customer notification vs. contractor notification discrepancy.
[CONSISTENCY FLAG: SI 14 contradicts SI 22] Step to prevent representations conflicts with revocation of authority.
"""
    flags = parse_consistency_flags(llm_output)
    assert len(flags) == 2
    assert flags[0].si_a == "10"
    assert flags[1].si_a == "14"


def test_parse_ignores_non_flag_text():
    llm_output = """
Here is my analysis.

[CONSISTENCY FLAG: SI 1 contradicts SI 2] Simple contradiction.

That's all I found.
"""
    flags = parse_consistency_flags(llm_output)
    assert len(flags) == 1
    assert flags[0].si_a == "1"
    assert flags[0].si_b == "2"
```

- [ ] **Step 2: Run — verify it fails**

```bash
python -m pytest tests/test_discovery/test_consistency_checker.py -v
```

- [ ] **Step 3: Create the module**

Create `icharlotte_core/discovery/consistency_checker.py`:

```python
"""
Phase 3 cross-question consistency checker.

After all requests in a set have been drafted and pruned, a single LLM call
reviews the entire set looking for factual contradictions between responses.
The LLM does NOT rewrite or reconcile — it only emits [CONSISTENCY FLAG:]
markers. The UI surfaces these to the user who resolves them manually before
the document is assembled.
"""
import re
from dataclasses import dataclass
from typing import Dict, List, Tuple


@dataclass
class ConsistencyFlag:
    si_a: str
    si_b: str
    reason: str


_FLAG_PATTERN = re.compile(
    r"\[CONSISTENCY FLAG:\s*SI\s+(\S+)\s+contradicts\s+SI\s+(\S+)\s*\]\s*([^\n\[]*)",
    re.IGNORECASE,
)


def build_consistency_prompt(
    responses: Dict[str, Tuple[str, str]],
) -> str:
    """
    Build an LLM prompt that reviews a full response set for contradictions.

    Args:
        responses: dict mapping SI number (as string) to (request_text, answer_text).

    Returns:
        Prompt string ready to send to an LLM.
    """
    header = (
        "You are reviewing a set of California discovery responses for "
        "internal consistency. Your job is to identify factual contradictions "
        "between responses — for example, one response says 'yes' to a premise "
        "that another response denies, or two responses give incompatible "
        "dates, or the steps described in one response contradict a fact "
        "asserted in another.\n\n"
        "DO NOT rewrite any response. DO NOT auto-reconcile. Your only output "
        "is a list of consistency flags.\n\n"
        "For each contradiction you find, output one line in this exact format:\n\n"
        "  [CONSISTENCY FLAG: SI <X> contradicts SI <Y>] <one-line reason>\n\n"
        "Example:\n"
        "  [CONSISTENCY FLAG: SI 10 contradicts SI 30] SI 10 says Yes to "
        "customer notification but SI 30 lists None when asked which contractors "
        "were informed prior to March 8, 2023.\n\n"
        "If you find no contradictions, output exactly: NO CONTRADICTIONS\n\n"
        "RESPONSES TO REVIEW:\n"
    )

    body_lines = []
    for num in sorted(responses.keys(), key=lambda k: (len(k), k)):
        req, ans = responses[num]
        body_lines.append(f"SI {num}: {req}")
        body_lines.append(f"ANSWER: {ans}")
        body_lines.append("")

    return header + "\n".join(body_lines)


def parse_consistency_flags(llm_text: str) -> List[ConsistencyFlag]:
    """Parse LLM output into ConsistencyFlag objects."""
    text = llm_text.strip()
    if text == "NO CONTRADICTIONS":
        return []

    flags: List[ConsistencyFlag] = []
    for match in _FLAG_PATTERN.finditer(text):
        si_a = match.group(1).strip()
        si_b = match.group(2).strip()
        reason = match.group(3).strip()
        flags.append(ConsistencyFlag(si_a=si_a, si_b=si_b, reason=reason))

    return flags
```

- [ ] **Step 4: Run tests**

```bash
python -m pytest tests/test_discovery/test_consistency_checker.py -v
```

Expected: all 5 pass.

- [ ] **Step 5: Commit**

```bash
git add icharlotte_core/discovery/consistency_checker.py tests/test_discovery/test_consistency_checker.py
git commit -m "feat(discovery): add cross-question consistency checker with flag parser"
```

---

### Task 19: Wire consistency check into the drafting pipeline

**Files:**
- Modify: `icharlotte_core/ui/respond_tab.py:998` (`_on_combined_draft_finished()`) — after prune pass, run consistency check asynchronously

- [ ] **Step 1: Add async consistency check worker**

In `_on_combined_draft_finished`, after the prune loop and before `_finalize_response`, launch a consistency check as a background LLM call. Because this is an extra LLM roundtrip, it should run asynchronously and not block the user's first view of the pruned responses.

Approach: call `_finalize_response` immediately to let the user start reviewing, then kick off a separate `LLMWorker` for the consistency check that updates the display when done.

```python
# Inside _on_combined_draft_finished, after the prune loop:
self._finalize_response(parsed, objections_map, responses_map, file_label)

# Launch consistency check in the background
self._launch_consistency_check(parsed, responses_map)
```

Add a new method `_launch_consistency_check`:

```python
def _launch_consistency_check(self, parsed, responses_map):
    """Kick off an async LLM call to review the response set for contradictions."""
    from icharlotte_core.discovery.consistency_checker import build_consistency_prompt

    # Build {num: (request_text, answer_text)}
    responses = {}
    for req in parsed.requests:
        num = req.number
        ans = responses_map.get(num, "")
        if ans.strip().startswith("[NEEDS HUMAN INPUT:"):
            continue  # skip flagged responses in the consistency check
        responses[num] = (req.text, ans)

    if len(responses) < 2:
        return  # nothing to cross-check

    prompt = build_consistency_prompt(responses)

    provider = self.provider_combo.currentText()
    model = self.model_combo.currentText()

    worker = LLMWorker(
        provider=provider,
        model=model,
        system="You are a California litigation defense attorney reviewing discovery responses for consistency.",
        user=prompt,
        files=[],
        settings={"stream": False},
    )
    worker.finished.connect(self._on_consistency_check_finished)
    worker.error.connect(lambda err: logger.warning("Consistency check failed: %s", err))
    self._llm_workers.append(worker)
    worker.start()


def _on_consistency_check_finished(self, response_text: str):
    """Parse consistency flags and surface them in the display."""
    from icharlotte_core.discovery.consistency_checker import parse_consistency_flags

    flags = parse_consistency_flags(response_text)
    if not flags:
        logger.info("Consistency check: no contradictions found")
        return

    logger.info("Consistency check: %d flag(s) found", len(flags))
    # Append a summary block to the editor
    from PyQt6.QtWidgets import QMessageBox
    summary_lines = ["// CONSISTENCY FLAGS FROM CROSS-QUESTION REVIEW:"]
    for flag in flags:
        summary_lines.append(f"//   [SI {flag.si_a} ↔ SI {flag.si_b}] {flag.reason}")
    summary_lines.append("// Review and resolve manually before saving.")
    summary = "\n".join(summary_lines)

    # Append to the editor as an informational comment block
    current = self.editor.toPlainText()
    self.editor.setPlainText(current + "\n\n" + summary + "\n")

    # Also show a non-modal info dialog
    msg = QMessageBox(self)
    msg.setWindowTitle("Consistency check")
    msg.setIcon(QMessageBox.Icon.Information)
    msg.setText(f"Cross-question consistency check found {len(flags)} flag(s).")
    msg.setDetailedText("\n".join(f"SI {f.si_a} ↔ SI {f.si_b}: {f.reason}" for f in flags))
    msg.setStandardButtons(QMessageBox.StandardButton.Ok)
    msg.exec()
```

- [ ] **Step 2: Strip consistency flag comments on save**

Extend the save-path strip regex from Task 16 to also remove consistency flag blocks:

```python
text = _re.sub(
    r"// CONSISTENCY FLAGS FROM CROSS-QUESTION REVIEW:.*?// Review and resolve manually before saving\.\s*",
    "",
    text,
    flags=_re.DOTALL,
)
```

- [ ] **Step 3: Manual smoke test**

Start the app, run the PREMIER SI Set One again, wait for the consistency check to return. Expected: at least one flag on the SI 10 / SI 30 pair (customer notification vs. contractor notification).

- [ ] **Step 4: Commit**

```bash
git add icharlotte_core/ui/respond_tab.py
git commit -m "feat(respond): async cross-question consistency check with editor-level flag display"
```

---

### Task 20: Phase 3 manual regression test against PREMIER SI Set One

**Files:** No code changes.

- [ ] **Step 1: Rerun PREMIER SI Set One end-to-end**

Full run with Conservative + Minimal. Wait for draft + prune + consistency check.

- [ ] **Step 2: Verify Phase 3 acceptance criteria**

From spec section 6.3:

1. At least one consistency flag on the SI 10 / SI 30 pair.
2. Any additional flags reviewed manually.
3. If zero flags produced on a set known to contain contradictions, investigate prompt quality.
4. Flags display correctly in editor.
5. Flag comments stripped from saved .docx.

- [ ] **Step 3: Final regression note**

Append to `docs/superpowers/plans/phase-1-regression-notes.md` — Phase 3 section.

- [ ] **Step 4: Commit**

```bash
git add docs/superpowers/plans/phase-1-regression-notes.md
git commit -m "docs(discovery): phase 3 PREMIER regression notes"
```

---

## Self-review checklist (for the implementer after all tasks)

After completing all tasks:

- [ ] Run the full test suite: `python -m pytest tests/test_discovery/ -v`
- [ ] Full PREMIER rerun with Conservative + Minimal — confirm all three phases' acceptance criteria hold simultaneously.
- [ ] Grep for any lingering `[NEEDS HUMAN INPUT:` tokens in recent generated output that weren't resolved.
- [ ] Verify objection #4 (privilege) was never incorrectly dropped across the entire PREMIER rerun.
- [ ] Sanity-check that FI flows still work (Task 5 step 6 smoke test plus one full FI run).
- [ ] Confirm no linting errors introduced in touched files.

## Known follow-ups (deferred, not blockers)

- Interactive "Restore" buttons for dropped objections (Task 16 ships a comment-footer minimum-viable version).
- Fact-extraction pre-pass for richer case context (deferred from Phase 1 brainstorming — revisit if grounding-prompt-only approach proves insufficient across multiple real cases).
- Auto-reconciliation of consistency flags (deferred — flag-and-resolve is the current approach).
- Grouping of related questions for consistency-aware drafting (deferred).
