# Oppose-a-Motion Redesign — Design

**Date:** 2026-05-26
**Status:** Approved in brainstorming; pending written-spec review
**Supersedes:** `2026-05-26-oppose-motion-wizard-design.md` (original wizard task design)

---

## Goal

Rebuild the Wizard's *Oppose a Motion* task to produce comprehensive, persuasive California civil opposition memoranda with citations the user can trust. The current task hallucinates case holdings (e.g., citing *Sinaiko* for mootness when its actual holding is about waiver of objections) and produces stilted, generic prose. The redesign eliminates pre-draft legal research, lets the LLM draft from its own knowledge of California civil case law, and then back-loads rigorous citation verification against CourtListener (case law) and California Legislative Information (statutes).

Writing voice is improved through workbench-managed *style examples* — the user's own prior oppositions, tagged by motion type, are injected into the drafter prompt as voice exemplars.

---

## Current State

The existing pipeline (`icharlotte_core/opposition/`) performs:

1. **Motion analysis** — LLM extracts metadata. Works.
2. **Outline generation** — LLM generates section outline. Works.
3. **Authority research** — LLM plans CourtListener search queries; CourtListener returns candidate cases; LLM filters. Frequently surfaces irrelevant cases; LLM filter is unreliable.
4. **Drafting** — single LLM call with motion + outline + an "authority block" (case names + reporter cites only). LLM has to confabulate what each case held, leading to hallucinated propositions.
5. **Citation verification** — verifier uses CourtListener search snippets and word-overlap against opinion sentences. Frequently rate-limited; word-overlap is too weak to catch the *Sinaiko*-class error.
6. **Assembly** — Word doc generated from the draft body.

Writing quality is "weird" — generic LLM prose with repetitive transitions and no attorney voice. Citations are often wrong: real cases cited for propositions they don't actually support, or invented entirely.

---

## Selected Approach: Back-loaded checking (draft → verify → flag)

The drafter writes the opposition from its own knowledge of California civil law and case law. No pre-draft CourtListener research. After drafting, every citation (case and statute) is independently verified against authoritative source text, with verdicts surfaced inline in the wizard's output page.

**Why this approach over front-loaded grounding:** the user prefers attorney control over what to do about flagged citations. Verification is faster end-to-end (~2-3 min vs. ~5-7 min for front-loaded grounding). Style transfer is more reliable when the drafter operates on its own knowledge rather than being constrained to a small set of pre-vetted cases.

---

## Pipeline Architecture

Five stages, each driven by a workbench-editable prompt:

```
1. analyze_motion          (existing — minor cleanup)
        ↓
2. generate_outline        (existing — minor cleanup)
        ↓
3. draft_memorandum        (rewritten — uses style examples; no authority list input)
        ↓
   Citation Parser         (new — extracts case + statute cites with propositions)
        ↓
4. verify_citation         (new — per-cite, fans out case path vs statute path)
        ↓
5. find_replacement        (new — optional; fires only on red flags via UI button)
```

The `plan_authority_queries` and `filter_authorities` prompts from today's pipeline are **removed**. No CourtListener search happens before drafting; the drafter's only inputs are the motion text, context documents, outline, and active style examples.

### Citation Parser

A new module `icharlotte_core/opposition/citation_parser.py` scans the drafted body and emits one `Citation` record per cite:

```python
@dataclass
class Citation:
    kind: str               # "case" | "statute" | "rule" | "unknown"
    raw_text: str           # exact substring as it appears in the body
    normalized: str         # canonical form for cache key
    proposition: str        # containing sentence + 1 sentence of prior context
    body_offset: int        # character offset for UI underline placement
    case_name: str = ""
    reporter_citation: str = ""
    year: str = ""
    law_code: str = ""      # e.g., "CCP", "EVID", "BPC"
    section_num: str = ""   # e.g., "2024.020"
```

**Case-cite patterns:** `Name v. Name (YYYY) Vol Reporter Page` and common variants with `*...*` italic markers, parenthetical year, optional pincites and parallel cites. Reporters: `Cal.`, `Cal.App.`, `Cal.2d` through `Cal.6th`, `Cal.App.2d` through `Cal.App.5th`, `Cal.Rptr.`, `Cal.Rptr.2d/3d`, `P.2d`, `P.3d`.

**Statute-cite patterns:** Common forms ("Code Civ. Proc., § 2024.020", "Code of Civil Procedure section 2024.020", "CCP § 2024.020", etc.) map to a normalized `(law_code, section_num)` pair via this abbreviation table:

| Common citation forms | leginfo `lawCode` |
|---|---|
| CCP, Code Civ. Proc., Code of Civil Procedure | `CCP` |
| Evid., Evid. Code, Evidence Code | `EVID` |
| Civ., Civ. Code, Civil Code | `CIV` |
| Pen., Pen. Code, Penal Code | `PEN` |
| Gov., Gov. Code, Government Code | `GOV` |
| Bus. & Prof. Code, B&P | `BPC` |
| Health & Saf. Code, H&S | `HSC` |
| Lab., Lab. Code, Labor Code | `LAB` |
| Veh. Code | `VEH` |
| Fam. Code | `FAM` |
| Prob. Code | `PROB` |

Citation forms not in the table map to `kind="unknown"` and are tagged UNVERIFIED in the UI.

**Rule-cite patterns:** "California Rules of Court, rule N.NNNN" and "CRC rule N.NNNN".

**Proposition extraction:** for each match, the parser walks backward to the start of the containing sentence and forward to the next sentence boundary; concatenates that and the prior sentence. The verifier sees ~2-3 sentences of context, enough to understand the brief's actual claim.

### Verification flow

For each `Citation`:

#### Case path

1. **Existence check.** Call CourtListener `/api/rest/v4/citation-lookup/?text={raw_citation}`. Returns matching cluster IDs or 404.
2. **NOT_FOUND short-circuit.** If no cluster found, emit verdict `NOT_FOUND` with note "this citation does not appear in CourtListener's California reporter index; it may be invented, mis-cited, or unpublished."
3. **Opinion fetch.** GET `/api/rest/v4/opinions/?cluster={cluster_id}`, read `plain_text` of the lead opinion. Cache to `.cache/opinions/{cluster_id}.json`.
4. **Holding comparison.** Single LLM call (see `verify_citation_current.txt` skeleton below).

#### Statute path

1. **URL build.** `https://leginfo.legislature.ca.gov/faces/codes_displaySection.xhtml?lawCode={law_code}&sectionNum={section_num}.`
2. **Fetch and extract.** Parse with BeautifulSoup; extract `.section_content` text. Cache to `.cache/statutes/{law_code}_{section_num}.json`.
3. **NOT_FOUND short-circuit.** If the page 404s or the section-content div is empty/repealed, emit `NOT_FOUND` with note "this statute section was not found at leginfo; it may be invented, repealed, or mis-cited."
4. **Statute comparison.** Same LLM prompt as the case path, with statute text in place of opinion text.

#### Out-of-scope citation types (v1)

Federal cases, federal statutes, local court rules, treatises (Witkin, Rutter), and the California Constitution are tagged `UNVERIFIED` (gray indicator) with a popup note "verifier doesn't yet cover this source; verify manually."

#### Verifier prompt skeleton

The LLM only emits SUPPORTED / PARTIAL / NOT_SUPPORTED verdicts. The `NOT_FOUND` verdict is system-emitted from the existence-check short-circuits above; the LLM is never invoked in that case.

Stored at `Scripts/prompts/oppose_motion/verify_citation_current.txt`:

```
You are auditing a single citation in a California civil opposition memorandum.

Given:
- The brief's PROPOSITION (the sentence(s) around the cite, showing what the
  brief claims the authority stands for).
- The actual AUTHORITY TEXT (opinion text or statute text).

Your job: decide if the authority actually supports what the brief claims.

Return JSON only with these keys:
- verdict: "SUPPORTED" | "PARTIAL" | "NOT_SUPPORTED"
- evidence: 1-2 verbatim sentences from the AUTHORITY TEXT that you relied on
- note: short attorney-facing explanation (≤2 sentences). For PARTIAL, say what's
        accurate and what's overstated. For NOT_SUPPORTED, say what the case
        actually holds.

Be strict. If the authority is on a different issue, says the opposite, or only
glancingly relates: NOT_SUPPORTED. If it supports a broader or narrower version
of the claim: PARTIAL. Reserve SUPPORTED for cases where the authority directly
holds what the brief claims.

PROPOSITION:
{proposition}

CITATION:
{raw_citation_text}

AUTHORITY TEXT:
{opinion_or_statute_text}
```

### Edge cases handled by the parser

- **Pincites** (`226 Cal.App.4th 401, 415`) — pincite stripped for cache key, kept in raw_text for display.
- **Parallel citations** (`226 Cal.App.4th 401, 175 Cal.Rptr.3d 423`) — primary reporter used for lookup, parallel preserved in display.
- **Signal prefixes** ("See", "Cf.", "See also") — preserved in raw_text; ignored for proposition extraction.
- **String cites** (multiple cases per sentence) — each parsed independently; share proposition context.
- **Duplicate cites** — verified once, cached verdict applied to all instances.

### Caching, rate limiting, parallelism

- **Cache keys:** `cluster_id` for opinions; `lawCode_sectionNum` for statutes.
- **CourtListener:** 5,000 requests/day with API key (already in environment). Typical brief: ~10-15 cite-lookup calls + ~5-10 opinion fetches. Cache warms over a few weeks of use.
- **leginfo:** no documented rate limit; self-throttle to 1 request/sec with retry-backoff on 503.
- **Parallelism:** verifier processes citations concurrently with a bounded pool (default 4 simultaneous network ops). Typical verification phase: ~10-15s for ~8 cites.

---

## Workbench Integration

### File layout

```
Scripts/prompts/oppose_motion/
├── analyze_motion_current.txt
├── generate_outline_current.txt
├── draft_memorandum_current.txt          ← main drafter prompt
├── verify_citation_current.txt           ← verifier prompt (case + statute)
├── find_replacement_current.txt          ← optional replacement search
├── style_examples.json                   ← example file paths + tags
└── .cache/                               ← gitignored
    ├── opinions/
    │   └── {cluster_id}.json
    └── statutes/
        └── {law_code}_{section_num}.json
```

This mirrors `Scripts/prompts/mediation_brief/` so the workbench's existing version control, A/B testing, and dashboard machinery pick it up without changes.

### LLMConfig registration

- **Agent ID:** `agent_oppose_motion`
- **Passes:** `analyze_motion`, `generate_outline`, `draft_memorandum`, `verify_citation`, `find_replacement`
- Each pass can have its own model preference sequence (e.g., Gemini Pro for drafting, Gemini Flash for verification).
- Added to `WORKBENCH_TO_AGENT_ID` in `dialogs.py`: `"oppose_motion": "agent_oppose_motion"`.

### Workbench tabs (existing, applied to oppose_motion)

When the user selects `oppose_motion` as the active agent:

- **Editor** — edit each pass's prompt; save as new version or overwrite current.
- **LLM Assistant** — make a prompt more specific / shorter / clearer.
- **A/B Testing** — run two prompt versions side-by-side on the same motion.
- **Version History** — roll back to any prior version.
- **Dashboard** — usage counts, performance scores per version.
- **Model Defaults** — per-pass model preference.

### New: "Style Examples" tab

Visible only when `oppose_motion` is the selected agent. A table:

| Column | Content |
|---|---|
| Label | User-supplied short name (e.g., "MTC Opp - Discovery Sanctions") |
| File Path | Absolute path to a .docx opposition |
| Motion Type Tags | Free-form strings (e.g., `motion to compel`, `discovery`) |
| Active | Checkbox to include/exclude in matching pool |
| Actions | Edit, Remove |

Bottom buttons: **Add Example** (file picker + dialog for label and tags), **Refresh** (re-extract text from .docx files), and the existing workbench Save buttons.

The Add/Edit dialog suggests common motion-type tags from a dropdown (`msj`, `msa`, `demurrer`, `motion to compel`, `motion to compel further`, `anti-slapp`, `motion in limine`, `motion for reconsideration`, `motion to set aside`, `motion to continue`) but accepts free-form input. Multi-tag is supported. Examples with **no tags** are treated as universal.

### Style example auto-matching at draft time

When the wizard runs `draft_memorandum`:

1. Pull `metadata.motion_type` from the analyze stage.
2. Normalize to lowercase.
3. Iterate `style_examples.json`: an example matches if `active=true` AND any of its `motion_types` tags is a substring of the normalized motion type, OR it has no tags (universal).
4. Take up to **3 matching examples** (token-budget cap; more than 3 yields diminishing returns on style transfer).
5. For each match: extract text from the .docx (cached on disk at `.cache/style_examples/{hash_of_path_and_mtime}.txt` so unchanged exemplars don't re-extract).
6. Inject into the drafter prompt as `<style_exemplar_1>...</style_exemplar_1>` blocks with the instruction *"These are exemplar oppositions from this firm; mimic their voice, structure, and rhetorical style — do not copy their facts or citations."*

If zero matches, the drafter runs without exemplars. The wizard status pane reports this so the user knows.

---

## UI Changes

The wizard's three-page flow (Settings → Status → Output) is unchanged structurally. Changes are concentrated in the Status page's verification phase and the Output page.

### Status page during verification

After drafting completes, the worker enters a verification phase that emits per-citation progress:

```
Drafting opposition memorandum...
Drafting...
Verifying citations (12 found)...
  ✓ Cottini v. Enloe Medical Center (2014) 226 Cal.App.4th 401
  ✓ Code Civ. Proc. § 2024.020
  ⚠ Sinaiko Healthcare (2007) 148 Cal.App.4th 390  — partial
  ✗ Smith v. Imaginary (2019) 35 Cal.5th 999       — not found
Verification complete: 8 supported, 2 partial, 2 not_supported.
```

The verifier runs in a worker thread with 4-way parallelism and bounded total time (~30s typical).

### Output page

Adds a **verification summary banner** at the top:

- Counts per verdict (🟢 SUPPORTED, 🟡 PARTIAL, 🔴 NOT_SUPPORTED / NOT_FOUND, ⚪ UNVERIFIED).
- Warning text when any red verdicts exist.
- **Re-verify all** button — re-runs verification on the current body text (useful after manual edits).
- **Export report** button — writes `<brief_name>_verification.txt` alongside the saved opposition with one row per citation: cite, verdict, proposition, evidence quote, note.

The two-pane layout (draft on left, drawer on right) is unchanged, but:

- Citations in the draft now have **verdict-colored underlines**:
  - 🟢 green = SUPPORTED
  - 🟡 yellow = PARTIAL
  - 🔴 red = NOT_SUPPORTED or NOT_FOUND
  - ⚪ gray = UNVERIFIED
- Hovering a citation shows a 1-line tooltip with verdict and key note.
- Clicking opens the extended `CitationDetailDialog`.

### Extended CitationDetailDialog

The popup is verdict-specific:

- **SUPPORTED:** green header, shows brief's proposition, verified holding/statute quote, verifier note, "Open in CourtListener" button.
- **PARTIAL:** yellow header, shows what's accurate vs overstated, evidence quote, opinion link.
- **NOT_SUPPORTED:** red header, shows brief's proposition, then a contrasting quote of what the case actually holds, verifier note explaining the gap, "Find replacement case" button, "Open in CourtListener" button.
- **NOT_FOUND:** red header, lists likely causes (invented / mis-typed / unpublished), "Find replacement case" button.
- **UNVERIFIED:** gray header, "verifier doesn't cover this source" message, no replacement button.

### "Find replacement case" button

The decision whether to **build** this feature in v1 is deferred to the implementation plan; the spec supports both options. When built, the button shows only on red verdicts. Clicking it runs the `find_replacement_current.txt` prompt: given the brief's proposition + failed cite, the LLM proposes ~3 replacement candidates which are each then verified the same way. A small dialog presents the candidates with verdicts; one click swaps the citation into the brief at that location. When deferred, the button is hidden and the user replaces cites manually.

### Save behavior

The Save button is always enabled.

- If any red flags exist, clicking Save first prompts: *"This opposition has N citations flagged as NOT_SUPPORTED. Save anyway?"* with Save / Cancel.
- The .docx is assembled as today.
- The verification report (`<brief_name>_verification.txt`) is written alongside the .docx, regardless of red-flag count.

---

## Module Structure

### New modules

- `icharlotte_core/opposition/citation_parser.py` — citation extraction with proposition windows.
- `icharlotte_core/opposition/statute_verifier.py` — leginfo fetch + cache.
- `icharlotte_core/opposition/style_examples.py` — load/save `style_examples.json`, auto-match by motion type, .docx text extraction.
- `icharlotte_core/ui/dialogs_style_examples.py` — the new Style Examples workbench tab widget.

### Rewritten modules

- `icharlotte_core/opposition/drafter.py` — removes authority_block input; accepts style_examples; loads prompt from `PromptManager`.
- `icharlotte_core/opposition/citation_verifier.py` — replaced with the new case/statute verifier described above.
- `icharlotte_core/opposition/motion_analyzer.py` — minor cleanup; load prompts from `PromptManager`.

### Removed modules

- `icharlotte_core/opposition/authority.py` — no pre-draft research.

### Modified modules

- `icharlotte_core/ui/wizard/pages/oppose_motion_page.py` — Status page emits verification progress; Output page gets summary banner, extended popup, color-coded underlines, save-with-warning behavior.
- `icharlotte_core/ui/dialogs.py` — `PromptsDialog`: add `oppose_motion` to `WORKBENCH_TO_AGENT_ID`; add Style Examples tab when `oppose_motion` is selected.
- `icharlotte_core/llm_config.py` — register `agent_oppose_motion` with its passes.

---

## Testing Strategy

### Unit tests

- `tests/test_opposition/test_citation_parser.py` — case/statute/rule patterns, proposition extraction, pincite stripping, parallel citations, signal prefixes.
- `tests/test_opposition/test_statute_verifier.py` — URL building per law code, HTML extraction, NOT_FOUND handling on repealed sections, cache hit/miss.
- `tests/test_opposition/test_citation_verifier.py` — case path with mocked CourtListener, statute path with mocked leginfo, NOT_FOUND short-circuits, verdict mapping.
- `tests/test_opposition/test_style_examples.py` — JSON load/save, motion-type substring matching, 3-example cap, no-tag universal behavior.
- `tests/test_opposition/test_drafter.py` — prompt loading from `PromptManager`, style-exemplar injection, returns DraftDocument with rejection_reason on invalid LLM output.
- `tests/test_wizard/test_oppose_motion_page.py` — summary banner counts, color-coded anchors per verdict, save-warning on red flags, re-verify path, find-replacement dialog plumbing.
- `tests/test_prompts_dialog_style_examples.py` — Style Examples tab visibility per agent, add/edit/remove, motion-type tag dropdown suggestions.

### Integration test

A scripted end-to-end run against the existing Pinscreen MTC motion (`Plaintiff's MTC Inspection FINAL TBS.pdf`) that:

1. Drives the worker programmatically through analyze → outline → draft → verify.
2. Asserts the drafter receives style-example blocks when matching examples are configured.
3. Asserts each parsed citation is verified and assigned a verdict.
4. Captures the verification report and snapshots it for regression detection.

This test runs only when `COURTLISTENER_API_TOKEN` and `GEMINI_API_KEY` are set; skipped in CI.

---

## Migration / Rollout

- The existing `oppose_motion` registry entry in `WORKBENCH_TO_AGENT_ID` does not exist yet; addition is purely additive.
- No data migration: today's task doesn't persist anything that survives across runs except the wizard preview .docx, which the redesigned task continues to produce.
- The original `authority.py` module is deleted; one in-tree usage in `oppose_motion_page.py` is updated to call the new verifier instead.
- Style examples start as an empty `style_examples.json`; first-run users see no exemplars until they add files via the workbench.

---

## Out of Scope (v1)

- Federal case law / federal statutes.
- Local court rules (county-specific).
- Treatises (Witkin, Rutter, etc.).
- California Constitution citations.
- KeyCite-style "has this case been overruled" check (CourtListener doesn't expose this).
- Pin-cite verification (checking that a specific page reference is correct).
- Google Scholar Case Law as an alternate search source.
- Per-section re-drafting after manual cite swaps.

Each can be added later as additive enhancements without changing the v1 architecture.
