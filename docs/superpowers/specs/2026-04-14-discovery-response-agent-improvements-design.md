# Discovery Response Agent Improvements — Design

**Date:** 2026-04-14
**Status:** Approved for implementation planning
**Regression target:** `Z:\Shared\Current Clients\5800 - AMTRUST\017 - Mederos\NOTES\AI OUTPUT\DISCOVERY RESPONSES\Def PREMIER's Resp to SI(1).docx`

---

## 1. Problem

Running the discovery response agent with "Conservative" objection aggressiveness and "Minimal" SI response style produced output that was neither conservative nor well-grounded. Concrete failures observed in the PREMIER SI Set One run:

1. **Objection over-firing despite "conservative" setting.**
   - Objection #3 (expert opinion / legal conclusion) added to pure factual yes/no questions (SI 1, 2, 3, 5, 20, 22, 26) where the request contains no "expert"/"opinion" keyword. The LLM is hallucinating this objection.
   - Objection #9 (compound) fires on non-compound requests like SI 1 ("ever an employee or independent contractor") and SI 8 ("collecting or attempting to collect").
   - Objection #7 (list/summary) fires on requests whose actual answers are single items — "Milo Holte" (SI 24), "None" (SI 9, 15, 21, 30).
   - Objection #6 (burden / overbroad time) fires on time-bounded requests (SI 30 is limited to "prior to March 8, 2023"; SI 4 asks for a single date).

2. **Grammar bug.** SI 1, 2, 3, 5 begin with "Responding Party **further** objects" — grammatically wrong as the first objection in a chain.

3. **Internal inconsistencies across the response set.**
   - SI 10 (customer notified) = Yes; SI 13 (publicly corrected) = No; SI 11 names "James Lindly on or about March 8, 2023"; SI 30 (contractors notified prior to 3/8) = None. Each answered independently with no cross-check.
   - SI 14 ("steps taken to prevent Chavez from holding himself out") answered with "terminated his employment" — conflating termination with active prevention, handing the plaintiff free ammunition.

4. **Substantive answers too thin / risk motion to compel under CCP 2030.220.**
   - SI 5: "Payroll records." — no specific document identification.
   - SI 19: "Various projects during his employment." — textbook evasive.
   - The drafter has access to context documents (loaded via the respond tab's context panel, read by `respond_tab.py:438 read_context_content()`) but is not instructed to mine them for specifics.

5. **Contradictions between objections and answers.** The "compiling a list would be burdensome" objection on SI 24 is followed by a one-item answer ("Milo Holte"). The objector didn't know what the answer would be.

### Root causes

| Symptom | Root cause |
|---|---|
| Objections ignore aggressiveness setting | `rule_based_preselect()` in `objection_selector.py:130` doesn't take aggressiveness; LLM prompt has only a one-sentence instruction with no calibration examples |
| Objections contradict answers | Phase 2 (objections) runs before Phase 3 (answers); no post-draft pruning |
| Cross-question inconsistency | Each request drafted in isolation; no awareness of prior answers in the set |
| Generic substantive answers | Drafter prompt in `response_drafter.py:200 build_si_prompt()` passes context as a bare `CASE CONTEXT:` block with no grounding instructions and no refusal protocol for missing facts |
| "Further objects" as first sentence | `format_objections()` in `objection_selector.py:254` concatenates by ID without touching leading words |
| Compound false positives | Parser flags compound on any "or"/"and" conjunction, catching synonymous alternatives |

---

## 2. Goals and non-goals

### Goals

- "Conservative + Minimal" settings produce output that is actually conservative and actually minimal.
- Drafter grounds substantive answers in loaded context documents wherever possible.
- When context is insufficient, the drafter flags the request for human input instead of producing vapid prose or hedged non-answers.
- Objections don't contradict the answers they precede.
- Cross-question inconsistencies are surfaced to the user for resolution before the document is assembled.
- Grammar and compound-detection bugs are fixed.

### Non-goals (explicitly out of scope)

- **Document structure / caption placement.** Initial suspicion that the caption was at the end of the document was wrong — the assembler at `response_assembler.py:234` loads the caption template first and appends. The weird ordering in extracted text was an artifact of `get_document_text()` reading Word table structure.
- **Trailer deduplication** ("Discovery and investigation are ongoing…"). Leave on every response.
- **Preliminary statement trimming.** Full 6-paragraph version stays regardless of style.
- **Privilege objection (#4) pruning.** The prune pass does not touch objection #4 at all, even when the answer is "No" or "Not applicable."
- **"Not applicable" de-padding.** Conditional follow-up answers keep their full objection stack.
- **Reversing Phase 2 / Phase 3 order.** Objections are still selected before the answer is drafted. The prune pass cleans up afterward.
- **Sequential or grouped drafting for consistency.** Drafter stays independent per request. A batched LLM pass at the end flags inconsistencies.
- **Auto-reconciliation of inconsistencies.** We flag, the user resolves manually.
- **Retrieval / RAG / fact-extraction pre-pass over context docs.** Deferred to future work. If Phase 1 prompt engineering proves insufficient, this becomes a follow-up project.
- **Grouping infrastructure for related questions.** Deferred.

---

## 3. Architecture overview

The pipeline stays logically the same:

```
parse request → select objections → draft answer → assemble
```

Three new insertions:

```
parse request
   ↓
select objections  ──→  validator layer (new, shared)  ──→  pruned initial set
   ↓
draft answer (with upgraded prompt + grounding + refusal protocol)
   ↓
rule-based prune pass (new)  ──→  uses validator layer + answer-aware rules
   ↓
(after all requests drafted)
   ↓
batched LLM consistency pass (new)  ──→  flags contradictions, does not rewrite
   ↓
user review checkpoint (shows pruned set + NEEDS HUMAN INPUT flags + CONSISTENCY flags)
   ↓
assemble
```

The validator layer is a shared module used in two places: (a) filtering the initial objection selection for sanity, and (b) informing the rule-based prune pass. It lives in one file and has one authoritative set of per-objection gates.

---

## 4. Phase 1 — Drafter grounding + tactical cleanup

### 4.1 Drafter prompt upgrade

**Location:** `response_drafter.py:200 build_si_prompt()`, and symmetric changes to `build_rfa_prompt()` (`:231`) and `build_rpd_prompt()` (`:268`).

**New prompt structure for SI (schematic):**

```
You are a California civil litigation defense attorney drafting a substantive
response to a Special Interrogatory. Do not include objections — those are
handled separately.

GROUNDING REQUIREMENT: Your answer MUST be grounded in the CASE CONTEXT
below. Before writing the answer, scan the CASE CONTEXT for:
  - specific dates, names, and facts relevant to the interrogatory
  - document titles, bates numbers, or filenames that would support the answer
  - direct quotes or specific passages that address the question

REFUSAL PROTOCOL: If the CASE CONTEXT does not contain the specific facts
needed to answer this interrogatory, DO NOT fabricate or hedge with "discovery
is ongoing." Instead, respond with exactly this token:

  [NEEDS HUMAN INPUT: <one-line description of what facts are missing>]

Examples:
  [NEEDS HUMAN INPUT: specific project names where Chavez worked in 2022–2023]
  [NEEDS HUMAN INPUT: exact termination date and termination letter filename]

DRAFTING STYLE: <minimal/moderate/detailed instruction, unchanged from current>

INTERROGATORY:
<text>

CASE CONTEXT:
<text>

Write only the substantive answer. Use specific facts from the CASE CONTEXT
wherever possible. If you find a grounded answer, state it directly. Do not
include the "Subject to and without waiving the foregoing objections,
Responding Party responds as follows:" transition — it is inserted by the
pipeline from `ResponseRules.waiver_language` when the response is assembled
(see `response_rules.py:14` and `response_assembler.py:528` `_split_response_parts`).
```

RFA and RPD prompts get the same grounding requirement, the same refusal protocol, and the same anti-hedging instruction, adapted to their respective response forms.

### 4.2 Confidence gate — UI flow

Flag detection is a simple string check: does the drafter's response start with `[NEEDS HUMAN INPUT:`?

**UI behavior when a flag is present:**

1. The response row in the review list shows a yellow background.
2. The review panel body shows the missing-fact description prominently.
3. A "Resolve" button opens an inline editor where the user types the correct answer.
4. Until resolved, the response cannot be included in the assembled document.
5. Attempting to assemble with unresolved flags shows a warning listing the affected SI numbers.
6. Resolved answers drop the `[NEEDS HUMAN INPUT:]` prefix and become normal responses.

Implementation touches the respond tab review widgets only. The drafter and assembler don't need flag-awareness — they just pass the string through.

### 4.3 "Further objects" grammar fix (Bucket D item #1)

**Location:** `objection_selector.py:254 format_objections()`.

After joining the objection texts, apply a single regex substitution to the leading occurrence only:

```python
joined = " ".join(parts)
joined = re.sub(
    r'^Responding Party further objects',
    'Responding Party objects',
    joined,
    count=1,
)
return joined
```

Does not affect second-or-later objections — they correctly retain "further."

**Unit tests:**
- Single objection whose text starts with "further" — gets rewritten.
- Multiple objections, first starting with "further" — only first rewritten.
- Multiple objections, first not starting with "further" — no change.
- Empty objection set — returns empty string cleanly.

### 4.4 Compound detection tightening (Bucket D item #8)

**Location:** `response_parser.py`, the `is_compound` field of `ParsedRequest`.

**New heuristic:**

1. Count interrogative verbs in the request: `state`, `identify`, `list`, `describe`, `explain`, `set forth`, `produce`. If ≥2, flag as compound.
2. Otherwise, look for "and" joining two distinct direct objects where each is a separate piece of information requested — e.g., "identify the date and method of revocation." Use a pattern of "<noun1>, <noun2>, and <noun3>" or "the <noun> and <noun>" where both nouns are information targets.
3. Otherwise, not compound. "Employee or independent contractor" — single categorical question, not compound. "Collecting or attempting to collect" — single action described with an attempt qualifier, not compound.

**Regression test against PREMIER SI 1–30:**
- Correctly flag as compound: SI 2, 11, 23, 24, 30.
- Correctly not flag: SI 1, 8, 15, 17, 22 (these are currently mis-flagged).
- Borderline SI to verify after implementation: SI 6 ("during 2021–2023" period specified, factual yes/no — not compound).

The Phase 2 validator layer will catch any compound false positives that slip through as a safety net.

### 4.5 Phase 1 acceptance criteria

After Phase 1 ships, rerun the PREMIER SI Set One with "Conservative + Minimal" settings. The rerun must show:

1. At least 5 responses that previously contained generic non-answers ("Payroll records" on SI 5, "Various projects" on SI 19, etc.) now either contain specific facts drawn from loaded context docs, OR show a `[NEEDS HUMAN INPUT:]` flag. Zero responses may still be generic non-answers without a flag.
2. No response's first objection begins with "further."
3. The **rule-based** compound pre-selection no longer fires on SI 1 or SI 8 (verified via unit test against the parser). Note: the LLM may still independently add objection #9 to these SIs because Phase 1 does not yet include LLM calibration. Full elimination of compound false positives on SI 1 and SI 8 is a Phase 2 acceptance criterion.
4. The parser still correctly flags compound on SI 2, 11, 23, 24, 30.
5. Smoke test FI/RFA/RPD flows against an existing case — no crashes, output structurally intact.

Phase 1 is self-contained. Nothing in Phases 2 or 3 depends on Phase 1's exact implementation details.

---

## 5. Phase 2 — Objection discipline

### 5.1 Validator layer — new shared module `objection_validator.py`

A single module that knows the hard requirements for each objection to be legitimately applicable. Used in two places: (a) to filter the initial selection, and (b) to inform the prune pass.

**Module structure:**

```python
# icharlotte_core/discovery/objection_validator.py

from typing import Optional
from icharlotte_core.discovery.response_parser import ParsedRequest

class ObjectionValidationResult:
    def __init__(self, valid: bool, reason: str = ""):
        self.valid = valid
        self.reason = reason

# Per-objection gate functions. Each returns ObjectionValidationResult.
# Each takes (request: ParsedRequest, answer_text: Optional[str] = None).
# When answer_text is None, this is initial-selection validation.
# When answer_text is provided, this is post-draft pruning.

def validate_obj_3_expert_opinion(request, answer_text=None):
    # Valid if request contains "expert", "opinion of", "contention",
    # or asks for a legal conclusion. Otherwise invalid.
    ...

def validate_obj_5_argumentative(request, answer_text=None):
    # Valid if the request contains an embedded factual assertion.
    # If answer_text is provided and accepts the assertion
    # (e.g., "Yes, during his employment"), return invalid — you can't
    # object that the premise is argumentative and then accept it.
    ...

def validate_obj_6_burden_time(request, answer_text=None):
    # Valid only if the request lacks an explicit temporal limit.
    # If answer_text is provided and contains a specific date
    # (regex: \b(January|February|...|\d{1,2}/\d{1,2}/\d{2,4})\b),
    # return invalid.
    ...

def validate_obj_7_list_summary(request, answer_text=None):
    # Valid only if the request asks for identification of multiple items.
    # If answer_text is provided and contains a single fact or "None" /
    # "Not applicable", return invalid.
    ...

def validate_obj_9_compound(request, answer_text=None):
    # Delegates to ParsedRequest.is_compound (tightened in Phase 1).
    # If answer_text is provided and is a single atomic fact, return invalid.
    ...

# Dispatcher used by both initial selection and prune pass.
VALIDATORS = {
    3: validate_obj_3_expert_opinion,
    5: validate_obj_5_argumentative,
    6: validate_obj_6_burden_time,
    7: validate_obj_7_list_summary,
    9: validate_obj_9_compound,
}

def filter_objection_ids(
    objection_ids: set[int],
    request: ParsedRequest,
    answer_text: Optional[str] = None,
) -> tuple[set[int], list[tuple[int, str]]]:
    """
    Returns (kept_ids, dropped_with_reasons).
    Objections #1, #2, #4, #8, #10, #11, #12 have no gate — they're either
    always valid or validated by different means. Unknown IDs are passed through.
    """
    ...
```

**Important design notes:**
- Objection #4 (attorney-client privilege / work product) is NOT in the validator dispatcher. It is never filtered, never pruned, regardless of answer content. Decision per user in brainstorming.
- Objection #2 (privacy) is not validated — it's controlled by the `always_include_privacy_objection` rule flag in `ResponseRules`. User controls this directly.
- Objections #1, #8, #10, #11, #12 have no answer-aware pruning rules — either they're general-purpose or they depend on parser flags that don't change based on answer content.

### 5.2 Initial-selection integration

**Location:** `objection_selector.py`, after `merge_objections()` at `:245`.

After `rule_based_preselect()` and LLM selection are merged, run `filter_objection_ids(merged_ids, request, answer_text=None)` to drop indefensible initial picks. Log dropped objections and reasons at DEBUG level.

### 5.3 Rule-based prune pass — new module `objection_pruner.py`

**Location:** new file `icharlotte_core/discovery/objection_pruner.py`.

**Pipeline position:** after the drafter returns a substantive answer, before the response is stored for UI review.

**Logic:**

```python
# icharlotte_core/discovery/objection_pruner.py

from icharlotte_core.discovery.objection_validator import filter_objection_ids
from icharlotte_core.discovery.response_parser import ParsedRequest

def prune_objections_against_answer(
    objection_ids: set[int],
    request: ParsedRequest,
    answer_text: str,
) -> tuple[set[int], list[tuple[int, str]]]:
    """
    Post-draft objection pruning. Drops objections rendered moot or
    contradicted by the substantive answer.

    NEVER drops objection #4 (privilege) — even if the answer is "No".
    NEVER drops objection #2 (privacy) — user-controlled.
    """
    # If the answer is a NEEDS HUMAN INPUT flag, don't prune anything —
    # we don't know what the final answer will look like.
    if answer_text.strip().startswith("[NEEDS HUMAN INPUT:"):
        return objection_ids, []

    # Delegate to the shared validator with answer context.
    return filter_objection_ids(objection_ids, request, answer_text=answer_text)
```

**Additional answer-aware rules evaluated inside the validator gates (not as separate rules) — summarized:**
- If the answer contains a month+year or MM/DD/YYYY date → #6 dropped.
- If the answer is ≤1 items or "None"/"Not applicable"/"Yes"/"No" → #7 dropped.
- If the answer is a single atomic fact → #9 dropped (even if parser flagged compound).
- If the answer accepts the premise of the request → #5 dropped.
- If the request had no "expert"/"opinion"/"contention" keyword → #3 dropped regardless of answer.

### 5.4 Few-shot calibration for LLM objection selection

**Location:** `objection_selector.py:174 _AGGRESSIVENESS_INSTRUCTIONS` and `:184 build_objection_prompt()`.

Replace the one-sentence aggressiveness instruction with a multi-part block for each aggressiveness level. The "conservative" version is the most important.

**Conservative block (new):**

```
AGGRESSIVENESS: CONSERVATIVE

Only select objections that clearly and directly apply to this specific
request. When in doubt, do NOT select the objection. Over-selection is
worse than under-selection under this setting.

Specifically DO NOT select:
  - #3 (expert opinion / legal conclusion) unless the request literally
    contains "expert", "opinion of", or asks the responder to state
    a legal contention. Factual yes/no questions are not expert opinion.
  - #6 (burden / overbroad as to time) if the request contains any
    explicit time limit ("during 2022", "prior to March 8, 2023",
    "after January 1").
  - #7 (list/summary not in existence) if the information requested
    is a small number of discrete facts readily knowable by the client
    (e.g., "who supervised X", "what was the termination date").
  - #9 (compound) unless the request genuinely asks for two or more
    substantively different pieces of information. Synonymous alternatives
    joined by "or" (e.g., "employee or independent contractor") are NOT
    compound.

Examples of correct conservative selection:

Request: "Was John Smith ever an employee of Acme Corp?"
Correct objections: (none)
NOT selected: #3 (no expert opinion sought), #9 (not compound — synonymous
alternatives would be a different question but this is a plain factual yes/no).

Request: "State whether Acme Corp was aware that Smith was collecting
payments during 2022."
Correct objections: (none or possibly #5 if client disputes the premise)
NOT selected: #3 (no expert opinion), #6 (time-bounded to 2022), #9 (not
compound).

Request: "List all dates and methods by which Acme terminated Smith's
authority, and identify all witnesses to such termination."
Correct objections: #9 (compound — three separate information requests)
NOT selected: #3 (no expert opinion), #6 (undefined time scope, could be
argued but time isn't the dominant issue here).
```

**Moderate block:** similar structure but with softer guidance.

**Aggressive block:** leave roughly as-is (the current one-sentence version is adequate for aggressive mode since the failure mode there is under-selection, which the current prompt doesn't cause).

### 5.5 UI review checkpoint — post-prune visibility

**Location:** `respond_tab.py`, objection review widget.

**Current flow:** LLM picks objections → user reviews the LLM-selected set → user can add/remove objections manually → drafter runs.

**New flow:** LLM picks objections → validator filters (initial selection) → drafter runs → prune pass filters (post-draft) → user reviews the **already-pruned** set with the drafted answer visible → user can restore pruned objections via a "Restore" action that surfaces the dropped objections with reasons → assemble.

**UI changes:**
1. The objection review panel shows the post-prune set as the primary display.
2. A "Show dropped objections" collapsible reveals each objection that was dropped along with the validator/pruner reason text.
3. A "Restore" button next to each dropped objection adds it back to the current response's objection set.
4. The review checkpoint happens after drafting, not before, because the prune pass needs the answer text. This is a meaningful UX change from the current "review objections before drafting" flow.

### 5.6 Phase 2 acceptance criteria

After Phase 2 ships, rerun the PREMIER SI Set One with "Conservative + Minimal" settings. The rerun must show:

1. Objection #3 appears ONLY on requests that literally contain "expert", "opinion of", or "contention". It no longer appears on SI 1, 2, 3, 5, 20, 22, 26.
2. Objection #7 does not appear on any response whose answer is a single item, "None", or "Not applicable". Previously on SI 2, 4, 9, 11, 14, 15, 17, 19, 21, 24, 30 — after Phase 2, dropped from at least SI 9, 15, 19, 21, 24, 30 (the ones with one-item or "None" answers).
3. Objection #6 does not appear on SI 30 (time-bounded) or SI 4 (single date requested).
4. Objection #9 appears only on SI 2, 11, 23, 24, 30 from the PREMIER set.
5. Objection #4 (privilege) is never dropped — it appears wherever it appeared in Phase 1 output.
6. Users can restore any dropped objection via the UI, and restored objections appear in the final assembled document.
7. Unit test coverage: each validator gate function has at least three tests (clearly valid, clearly invalid, borderline).

---

## 6. Phase 3 — Cross-question consistency

### 6.1 Batched LLM consistency pass

**Position in pipeline:** after all requests in a set have been drafted AND pruned, before the UI review checkpoint.

**New module:** `icharlotte_core/discovery/consistency_checker.py`.

**Prompt structure:**

```
You are reviewing a set of California discovery responses for internal
consistency. Your job is to identify factual contradictions between
responses — for example, one response says "yes" to a premise that another
response denies, or two responses give incompatible dates.

DO NOT rewrite any response. DO NOT auto-reconcile anything. Your only
output is a list of consistency flags.

For each contradiction you find, output one line in this exact format:

  [CONSISTENCY FLAG: SI <X> contradicts SI <Y>] <one-line reason>

Example:
  [CONSISTENCY FLAG: SI 10 contradicts SI 30] SI 10 says "Yes, Premier
  notified customers that Chavez was not authorized" but SI 30 says "None"
  when asked which Coachella Valley contractors were informed prior to
  March 8, 2023.

If you find no contradictions, output exactly: NO CONTRADICTIONS

RESPONSES TO REVIEW:
SI 1: <request text>
ANSWER: <answer text>

SI 2: <request text>
ANSWER: <answer text>

... (continues for all requests in the set)
```

**Model choice:** same model family as the objection selection LLM, chosen from `LLMConfig` — cheap but capable. One call per set, not per request.

**Parsing:** a simple regex extracts `[CONSISTENCY FLAG: SI X contradicts SI Y]` markers and their one-line reasons. Flags are attached to **both** affected responses (so the user sees them on SI X and SI Y when reviewing).

### 6.2 UI display of consistency flags

1. A response with an attached consistency flag gets an orange background (distinct from yellow used by `NEEDS HUMAN INPUT`).
2. The review panel body shows the flag reason at the top of the response.
3. A "Resolve" button opens a dialog showing both affected responses side-by-side, with the reason, and lets the user edit either or both.
4. Resolving a flag clears it from both affected responses.
5. Unresolved consistency flags do NOT block assembly — they are warnings, not errors. The assembly confirmation dialog lists unresolved flags so the user knows what they're shipping.

### 6.3 Phase 3 acceptance criteria

After Phase 3 ships, rerun the PREMIER SI Set One. The rerun must show:

1. At least one consistency flag on the known-problematic pair (SI 10 "Yes" / SI 30 "None" around customer notification).
2. Additional flags on any other contradictions the model identifies — these get reviewed manually, no preset expectation.
3. If the batched pass produces zero flags on a set known to contain contradictions, investigate prompt quality before shipping.
4. The UI correctly displays, resolves, and clears flags.
5. A clean SI set with no contradictions produces "NO CONTRADICTIONS" and zero orange-highlighted responses.

---

## 7. Sequencing and rollout

**Sequence 3** (decided in brainstorming): highest leverage first.

1. **Phase 1 first** — drafter grounding + confidence gate + grammar fix + compound detection. Ship, rerun PREMIER, verify acceptance criteria.
2. **Phase 2 second** — validator layer + few-shot calibration + rule-based prune pass + UI checkpoint changes. Ship, rerun PREMIER, verify acceptance criteria.
3. **Phase 3 third** — batched LLM consistency pass + UI for consistency flags. Ship, rerun PREMIER, verify acceptance criteria.

**Regression gate between phases:** if the PREMIER rerun for a phase fails any acceptance criterion, fix before starting the next phase. Each phase produces a checkpoint commit on a dedicated branch so we can bisect if something regresses.

---

## 8. Testing strategy

### 8.1 Unit tests

- `tests/test_discovery/test_objection_validator.py` — one test module for the validator layer. Each gate function gets at least three tests (valid, invalid, borderline). Tests exercise both `answer_text=None` (initial selection) and `answer_text="..."` (post-draft pruning) modes.
- `tests/test_discovery/test_objection_pruner.py` — tests the answer-aware pruning rules against synthetic before/after pairs. Includes a test that verifies objection #4 is NEVER dropped even on "No" / "Not applicable" answers.
- `tests/test_discovery/test_response_parser_compound.py` — regression tests for the tightened compound detection using the PREMIER SI 1–30 texts as fixtures.
- `tests/test_discovery/test_format_objections_grammar.py` — tests the "further objects" fix in isolation.
- `tests/test_discovery/test_consistency_checker.py` — tests the consistency flag parser against fixture LLM outputs (not the LLM call itself).

### 8.2 Integration tests

- `tests/test_discovery/test_premier_regression.py` — golden-file regression against the PREMIER SI Set One. Keeps the current (buggy) output as `baseline.txt` and the target (post-fix) output as `expected.txt`. Runs the full pipeline against a mocked LLM that returns canned responses, asserts structural properties (no leading "further", no #3 without expert keyword, etc.).

### 8.3 Manual verification

- After each phase ships, rerun the real PREMIER SI Set One against the real LLM with "Conservative + Minimal" settings. Visually compare against the current `Def PREMIER's Resp to SI(1).docx`. Verify acceptance criteria for that phase.
- Smoke test a small FI, a small RFA, and a small RPD set on an existing case to confirm no cross-form regressions.

---

## 9. Risks and open questions

### Risks

- **Few-shot examples may bias the LLM toward the example topics.** Mitigation: use examples from diverse request shapes (yes/no, identification, list, contention). Avoid over-fitting to the PREMIER-style request phrasing.
- **The post-draft prune pass changes the UX from "review objections, then draft" to "draft, then review pruned objections."** Users who relied on the pre-draft review step will see a behavior change. Mitigation: the "Restore" action preserves the ability to add back any objection the prune pass removed, so users retain full control of the final output.
- **Validator false negatives.** If a gate function is too strict, legitimate objections will be dropped. Mitigation: every drop is logged with a reason, and the UI surfaces dropped objections under "Show dropped objections" so the user can always see what was removed and why.
- **Consistency pass may produce false positives** (flagging non-contradictions as contradictions). Mitigation: flags are warnings, not errors — the user can dismiss them. The prompt explicitly tells the model not to auto-reconcile.
- **Drafter refusals may cluster if loaded context is thin.** A case with minimal context docs will see many `[NEEDS HUMAN INPUT:]` flags. This is the correct behavior (preferable to fabrication), but the first user encounter may feel like the agent "stopped working." Mitigation: documentation and the UI flag count surfaced prominently so the user understands the agent is asking for help, not failing.

### Open questions (to resolve during implementation, not now)

- Exact date-detection regex for the #6 pruning rule. Needs to catch "March 8, 2023", "3/8/2023", "on or about January 13, 2023", and "during 2021–2023" without false-positiving on unrelated numbers.
- Whether to show dropped objections collapsed-by-default or expanded-by-default in the UI. Default to collapsed; revisit if users complain they can't find restored objections.
- Whether the consistency pass runs automatically or via a user button. Default to automatic after drafting finishes; users who want to skip it can cancel.

---

## 10. Summary of decisions from brainstorming

| Decision | Chosen approach | Alternative(s) considered |
|---|---|---|
| Prune pass architecture | Hybrid: rule-based per-request + batched LLM pass at end (shares infra with consistency pass) | Pure rules, pure LLM |
| Review checkpoint position | After pruning, with Restore action | Before pruning, or both |
| Drafter context approach | Prompt engineering with grounding + refusal protocol | Retrieval, fact extraction pre-pass, curated fact file |
| Confidence gate | Yes — `[NEEDS HUMAN INPUT:]` flag | No — keep default hedge |
| Objection calibration | Few-shot examples + deterministic validator layer | Stricter per-objection prompt gates, combined LLM call |
| Validator scope | Shared by initial selection and prune pass | Separate modules |
| Cross-question consistency | Independent drafting + batched LLM flag-don't-reconcile | Sequential drafting, grouped drafting |
| Sequencing | Highest leverage first | Cheap wins first, architecture first |
| Regression target | PREMIER SI Set One | Synthetic test case |
| Bucket D cleanup | Items #1 (grammar), #3 (standard clause preserved), #8 (compound), #9 (argumentative via validator) | Plus items #2, #4, #5, #6, #7, #10 (declined) |
| Document structure | No change — not a real bug | N/A |
