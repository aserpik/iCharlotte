# Discovery Respond Tab — Design Spec

**Date:** 2026-04-10  
**Status:** Approved  
**Author:** Brainstorming session (Opus 4.6)

---

## Overview

Add a "Respond" subtab to the existing Discovery tab in iCharlotte. The Respond tab generates objections and substantive responses to discovery propounded by opposing parties. It mirrors the Propound tab's layout but reverses the workflow — instead of drafting discovery requests, it drafts responses to incoming discovery.

**Supported discovery types:**
- Form Interrogatories (FI)
- Special Interrogatories (SI)
- Requests for Admission (RFA)
- Requests for Production of Documents (RPD)

---

## Architecture

**Approach:** Hybrid — shared core models in `icharlotte_core/discovery/`, response-specific logic in new files within the same package, UI in a new file `icharlotte_core/ui/respond_tab.py`.

### New Files

| File | Purpose |
|------|---------|
| `icharlotte_core/discovery/response_parser.py` | Phase 1: Parse incoming discovery PDFs into structured data |
| `icharlotte_core/discovery/objection_selector.py` | Phase 2: Select objections per request (rule-based + LLM hybrid) |
| `icharlotte_core/discovery/response_drafter.py` | Phase 3: Draft substantive responses using context documents |
| `icharlotte_core/discovery/response_assembler.py` | Assemble final Word document from caption template + response text |
| `icharlotte_core/discovery/response_rules.py` | ResponseRules dataclass, serialization, default loading |
| `icharlotte_core/ui/respond_tab.py` | RespondTab widget (left pane controls + right pane editor) |
| `icharlotte_core/ui/response_rules_dialog.py` | Rules editor dialog (3-tab hybrid form) |
| `config/response_rules_default.json` | Default rules populated from sample documents |

### Modified Files

| File | Change |
|------|--------|
| `icharlotte_core/ui/discovery_tab.py` | Replace Respond placeholder with real `RespondTab`; update `load_case()` to delegate to both tabs |

### Shared Infrastructure (reused from existing code)

- `discovery/models.py` — `Party`, `PartyRole`, `DiscoveryType` enums
- `discovery/assembler.py` — `find_caption_page()` static method, Word style constants
- `ui/discovery_tab.py` — `DiscoveryTab` wrapper hosts both `PropoundTab` and `RespondTab`
- `CaseDataManager` — per-case variable persistence
- `LLMWorker` — background LLM calls in QThread
- `word_validator.py` — post-assembly validation

---

## UI Layout

### Left Pane (Controls)

Top to bottom:

1. **Discovery to Respond To** — drag-and-drop list with checkboxes. Accepts PDF files. Blue-themed header. Each dropped file represents one set of incoming discovery (FI, SI, RFA, or RPD). Discovery type is auto-detected during parsing.

2. **Context Documents** — drag-and-drop list with checkboxes. Accepts PDF, DOCX, TXT, MSG, images. Gold-themed header. Same behavior as the Propound tab's document box — provides factual context for drafting substantive responses.

3. **Settings group:**
   - **Our Client** — dropdown of `Party` objects from shared `discovery_party_roster`. Includes `+` button to add parties and right-click context menu to edit/remove. Same party management as Propound tab.
   - **LLM Provider** — dropdown (Gemini / OpenAI / Claude)
   - **Model** — dropdown, dynamically populated per provider

4. **Response Rules** button — opens the `ResponseRulesDialog`

5. **Generate Responses** button — styled blue, disabled during generation

### Right Pane (Editor)

- **Toolbar:** Save as .docx | Save All | Clear | Refresh 17.1 | status label (right-aligned)
- **Tabs:** One `QPlainTextEdit` tab per generated response set, labeled e.g., "Resp to FI(1)", "Resp to RFA(1)", "Resp to RPD(1)"
- **Empty state:** Centered label "Drop discovery PDFs and click Generate" when no output exists
- **Refresh 17.1 button:** Only enabled when both FI and RFA tabs exist. Triggers LLM to generate FI 17.1 response based on current RFA editor content.

---

## Three-Phase Pipeline

### Phase 1: Discovery Parser (`response_parser.py`)

**Input:** Raw PDF text from each incoming discovery document.

**Output:**
```python
@dataclass
class ParsedDiscovery:
    discovery_type: DiscoveryType  # FI, SI, RFA, RPD
    propounding_party: str         # e.g., "Plaintiff SALAMUDIN JAN"
    responding_party: str          # extracted or inferred from Our Client
    set_number: int                # 1, 2, etc.
    set_word: str                  # "ONE", "TWO", etc.
    case_number: str               # e.g., "21STCV30788"
    requests: List[ParsedRequest]

@dataclass
class ParsedRequest:
    number: str              # "1.1" for FI, "1" for SI/RFA/RPD
    text: str                # full request text
    definitions: List[str]   # inline definitions if any
    is_compound: bool        # flagged by compound detection
    defined_terms_used: List[str]  # e.g., ["INCIDENT", "VEHICLE"]
```

**How it works:**
- LLM call with a focused extraction prompt — identify discovery type, extract party names, set number, case number, and each individual request with its number and full text.
- Discovery type auto-detected from document content — looks for "FORM INTERROGATORY", "SPECIAL INTERROGATORY", "REQUEST FOR ADMISSION", "REQUEST FOR PRODUCTION".
- Compound question detection runs as post-parse step — flags requests containing multiple subparts or conjunctive questions (e.g., "state all facts AND identify all documents").
- Broad definition terms are extracted so Phase 2 can apply definitional objections.
- Uses the model selected in the UI's LLM Provider/Model dropdowns (same as all other phases). A faster model is suitable here since this is extraction, not reasoning, but model selection is the user's choice.

### Phase 2: Objection Selector (`objection_selector.py`)

**Input:** `ParsedDiscovery` + `ResponseRules` + objections menu (loaded from `C:\AI\discovery\DISCOVERY OBJECTIONS.docx`).

**Output:** `Dict[str, List[str]]` — mapping of request number to list of selected objection texts.

**By discovery type:**

- **Form Interrogatories:** 100% rule-based. Loads the fixed objections from the sample FI response document. Every FI gets the same objections. No LLM needed.

- **SI / RFA / RPD:** Hybrid approach:
  1. **Rule-based pre-selection:** Always include objections matching detected flags:
     - `is_compound` → compound objection
     - `defined_terms_used` has broadly-defined terms → definitional objection
     - Request mentions "expert" or "opinion" → expert opinion objection
     - `always_include_privacy_objection` → privacy objection
     - `always_include_privilege_objection` → privilege objection
     - `always_include_burden_objection` → burden objection
  2. **LLM selection:** Send each request + the full objections menu to the LLM. Instructions: "Select ALL objections that might apply. Err on the side of over-inclusion — a failure to include an objection waives it. If an objection might in any way apply, include it." Aggressiveness level from `ResponseRules.objection_aggressiveness` modulates the prompt.
  3. **Merge:** Union of rule-based and LLM-selected objections.

**Objections menu:** Loaded once from `C:\AI\discovery\DISCOVERY OBJECTIONS.docx` on first use and cached. Contains 13 standard objections, each with an ID for reference.

### Phase 3: Response Drafter (`response_drafter.py`)

**Input:** `ParsedDiscovery` + selected objections per request + context document text + `ResponseRules`.

**Output:** Complete plain-text response document per discovery type.

**By discovery type:**

#### Form Interrogatories

| Request | Source | Notes |
|---------|--------|-------|
| 1.1 | Fixed from sample | Always identical |
| 3.x, 4.x, 7.x, 12.x, 13.x, 14.x, 20.x series | Templated from sample format | Placeholders where info unavailable |
| 15.1 | Fixed from sample | Always identical |
| 16.x series | Fixed from sample | Always identical |
| 17.1 | Placeholder | "[PENDING — complete after RFA responses are finalized]" |
| N/A detection | Rule-based | Interrogatories irrelevant to case type → "Not applicable" |
| All others | LLM-drafted | Using context documents |

Every FI response includes:
- Objections (fixed from sample)
- "Subject to and without waiving the foregoing objections, Responding Party responds as follows:"
- Substantive response
- Reservation clause

#### Special Interrogatories

LLM-drafted with explicit instructions controlled by `ResponseRules.si_response_style`:
- **Minimal (default):** "Draft responses as narrowly as possible. Use as few words as possible. Only provide information specifically requested. If the wording allows a vague response, prefer the vague response. Avoid providing information damaging to the client."
- **Moderate:** Standard responsive drafting without extreme minimization.
- **Detailed:** Full responsive answers.

#### Requests for Admission

LLM chooses from three fixed options controlled by `ResponseRules.rfa_default_posture`:
- `"Admit"` — only when the fact is clearly undisputed
- `"Deny"`
- `"After a reasonable inquiry concerning the matter in this request, the information known or readily obtainable to Responding Party is insufficient to enable Responding Party to admit the matter."`

**Cautious (default):** LLM instructed to lean toward Deny or Insufficient — only Admit facts that are definitively not in dispute.
**Balanced:** LLM uses best judgment.
**Cooperative:** LLM leans toward admitting undisputed facts.

#### Requests for Production

LLM chooses from two fixed options controlled by `ResponseRules.rpd_default_posture`:
- `"Upon a diligent search and reasonable inquiry made in an effort to locate the item(s) requested, Responding Party is unable to comply with this request at this time because the documents responsive to this request, if they exist, are not in the possession, custody or control of Responding Party."`
- `"Responding Party will comply with this request and produce all non-privileged documents in Responding Party's possession, custody and control that Responding Party understands to be responsive to this Request. Responding Party identifies and refers to the documents produced concurrently herewith."`

**Unable to Comply (default):** LLM defaults to unable unless context clearly indicates documents exist and are producible.
**Will Comply:** LLM defaults to compliance unless context suggests documents don't exist.
**Context-Dependent:** LLM decides per request based on context.

#### Response Structure (all types)

Every individual response follows this structure:
1. Request header: e.g., `"SPECIAL INTERROGATORY NO. 1:"`
2. Request text (from parsed discovery)
3. Response header: e.g., `"RESPONSE TO SPECIAL INTERROGATORY NO. 1:"`
4. Objections
5. `"Subject to and without waiving the foregoing objections, Responding Party responds as follows:"`
6. Substantive response
7. `"Discovery and investigation are ongoing and Responding Party reserves the right to amend, modify and/or supplement this response in the future in the event that additional documents, facts and/or information are discovered, or their relevance becomes apparent."`

---

## Response Rules System

### Data Model (`response_rules.py`)

```python
@dataclass
class ResponseRules:
    # --- Objection Strategy ---
    objection_aggressiveness: str  # "aggressive" | "moderate" | "conservative"
    always_include_privacy_objection: bool
    always_include_privilege_objection: bool
    always_include_burden_objection: bool
    auto_flag_compound: bool
    auto_flag_broad_definitions: bool

    # --- Substantive Response Strategy ---
    si_response_style: str       # "minimal" | "moderate" | "detailed"
    rfa_default_posture: str     # "cautious" | "balanced" | "cooperative"
    rpd_default_posture: str     # "unable_to_comply" | "will_comply" | "context_dependent"
    fi_17_1_auto_refresh: bool   # default False

    # --- Editable Boilerplate Text ---
    waiver_language: str
    reservation_clause: str
    preliminary_statement_fi: str
    preliminary_statement_si: str
    preliminary_statement_rfa: str
    preliminary_statement_rpd: str
    intro_template_fi: str
    intro_template_si: str
    intro_template_rfa: str
    intro_template_rpd: str
    general_objections_rfa: str
    general_objections_rpd: str
    verification_template: str

    # --- Custom Instructions ---
    custom_instructions: str
```

**Defaults:** Populated from sample documents in `C:\AI\discovery\Discovery Responses\` on first run. Saved to `config/response_rules_default.json`. Per-case overrides stored in `CaseDataManager` variable `respond_rules`.

### Rules Dialog (`response_rules_dialog.py`)

Three-tab dialog opened via "Response Rules" button on Respond tab:

**Tab 1 — Strategy:**
- Objection aggressiveness: radio group (Aggressive / Moderate / Conservative)
- Checkboxes: privacy, privilege, burden always-include; compound auto-flag; broad definition auto-flag
- SI response style: radio group (Minimal / Moderate / Detailed)
- RFA default posture: radio group (Cautious / Balanced / Cooperative)
- RPD default posture: radio group (Unable to Comply / Will Comply / Context-Dependent)

**Tab 2 — Boilerplate:**
- Editable `QTextEdit` fields for each boilerplate block
- Per-field "Reset" button to restore default text from samples
- Fields: waiver language, reservation clause, preliminary statements (4), intro templates (4), general objections (2), verification template

**Tab 3 — Custom Instructions:**
- Large `QTextEdit` for free-form instructions
- Placeholder: "Add any additional instructions for how responses should be drafted. These will be sent to the LLM along with the structured rules above."

**Bottom bar:** OK | Cancel | Reset All to Defaults | Reload from Samples

---

## Word Document Assembly (`response_assembler.py`)

### Class: `ResponseAssembler`

**Input:**
- Caption page path (via `DiscoveryAssembler.find_caption_page()`)
- `ParsedDiscovery` (party names, set number, case info)
- Response text from editor (may have been edited by attorney)
- `ResponseRules` (boilerplate blocks)

**Assembly steps:**

1. **Load caption template** — copy the .docx, preserving all styles and formatting.

2. **Replace "CAPTION PAGE"** — find "CAPTION PAGE" text and replace with document title:
   - Format: `"DEFENDANT [NAME]'S RESPONSES TO [PROPOUNDING PARTY]'S [TYPE], SET [NUMBER]"`
   - e.g., `"DEFENDANT USA WASTE OF CALIFORNIA, INC.'S RESPONSES TO PLAINTIFF'S FORM INTERROGATORIES, SET ONE"`
   - Bold, matching sample formatting conventions.

3. **Extract and relocate signature block** — scan caption for signature block (detected by attorney name patterns, "Respectfully submitted", or firm name). If found, remove from current position and store for end-of-document insertion.

4. **Insert party identification block:**
   ```
   PROPOUNDING PARTY:     [propounding party name]
   RESPONDING PARTY:      [responding party name]
   SET NO.:               [SET_WORD] ([set_number])
   ```

5. **Insert "TO" line:** `"TO [PROPOUNDING PARTY] AND [HIS/HER/THEIR] ATTORNEYS OF RECORD:"`

6. **Insert introduction paragraph** — from `ResponseRules` intro template with placeholders filled with case-specific info.

7. **Insert preliminary statement** — from `ResponseRules` per discovery type.

8. **Insert general objections** — RFA and RPD only, from `ResponseRules`.

9. **Parse and insert individual responses** — re-parse editor plain text to extract request/response pairs. Format with Word styles:
   - Request header: Discovery No. style (bold)
   - Request text: Body Double style (indented, double-spaced)
   - Response header: bold
   - Objections, waiver, substantive response, reservation: Body Double style

10. **Insert verification page** — from `ResponseRules.verification_template` with placeholder fields for verifier name, date, signature line.

11. **Re-insert signature block** — at end of document, after verification.

12. **Set footer** — case number, judge, parties, matching sample format.

13. **Validate** — `word_validator.validate_after_edit()`.

**Output filename:** `"Def [Abbreviation]'s Resp to [Type]([SetNum]).docx"` — e.g., `"Def USA Waste's Resp to FI(1).docx"`

**Styles reused:** Body Double, Flush Left Double, Discovery No., Center Double Bold Underlined, Flush Left — all pre-defined in the caption template.

---

## State Persistence

Per-case variables stored via `CaseDataManager`:

| Variable | Type | Purpose |
|----------|------|---------|
| `respond_discovery_files` | `List[{path, checked}]` | Discovery PDFs to respond to |
| `respond_context_documents` | `List[{path, checked}]` | Context documents |
| `respond_output` | `List[{label, text}]` | Editor tab contents |
| `respond_rules` | `dict` | Serialized `ResponseRules` per-case overrides |
| `respond_rfa_responses` | `dict` | Cached RFA responses for FI 17.1 refresh |

**Party roster:** Shared with Propound tab via `discovery_party_roster` variable.

**Save triggers:** Document list changes, generation complete, editor edits (debounced 800ms via `QTimer`), rules dialog OK.

**Load trigger:** `load_case(file_number)` called when user switches cases.

---

## Generation Flow

```
User clicks "Generate Responses"
│
├── Validate: ≥1 discovery PDF checked, Our Client selected
│
├── For each checked discovery PDF (parallel where independent):
│   │
│   ├── Phase 1: Parse PDF → ParsedDiscovery
│   │   └── LLM call (fast model) — extract structure
│   │
│   ├── Phase 2: Select objections
│   │   ├── FI: rule-based only (fixed from samples)
│   │   └── SI/RFA/RPD: rule-based + LLM hybrid
│   │
│   └── Phase 3: Draft responses
│       ├── FI: fixed templates + LLM for non-template items
│       ├── SI: LLM with narrow/minimal instructions
│       ├── RFA: LLM chooses Admit/Deny/Insufficient
│       └── RPD: LLM chooses Unable/Will Comply
│
├── Assemble plain text per type → display in editor tabs
│
├── If RFA generated, cache responses for FI 17.1
│
└── FI 17.1 shows placeholder until "Refresh 17.1" clicked

User clicks "Refresh 17.1"
│
├── Read current RFA editor text (may have been edited)
├── LLM generates FI 17.1 based on RFA answers
└── Replace placeholder in FI editor tab

User clicks "Save as .docx"
│
├── Find caption page (parent folder scan)
├── ResponseAssembler builds Word doc
├── Validate with word_validator
└── Save to case folder
```

---

## Threading

All LLM calls run in `QThread` via existing `LLMWorker`. Generate button disables during generation. Status label shows progress ("Parsing FI...", "Drafting SI responses...", etc.). Multiple discovery types process in parallel since they are independent. FI 17.1 dependency handled by manual "Refresh 17.1" button.

---

## Sample Data Dependencies

**Read once and cached to `config/response_rules_default.json`:**

| Source | Data extracted |
|--------|---------------|
| `C:\AI\discovery\Discovery Responses\Def USA Waste's Resp to Pltf's FI(1).docx` | Fixed FI objections, FI 1.1/15.1/16.x responses, FI preliminary statement, FI intro template |
| `C:\AI\discovery\Discovery Responses\Def USA Waste's Resp to SI(1).docx` | SI preliminary statement, SI intro template |
| `C:\AI\discovery\Discovery Responses\Def USA Waste's Resp to RFA(1).docx` | RFA preliminary statement, RFA intro template, RFA general objections |
| `C:\AI\discovery\Discovery Responses\Def USA Waste's Resp to RPD(1).docx` | RPD preliminary statement, RPD intro template, RPD general objections |
| `C:\AI\discovery\DISCOVERY OBJECTIONS.docx` | 13 standard objections menu for SI/RFA/RPD |

**"Reload from Samples" button** in rules dialog re-reads these files and regenerates the defaults cache.

---

## Additional Rules (auto-applied)

1. **Verification page** — generated at end of document with placeholder fields for verifier name, date, signature line. Matches sample format.

2. **Definitional objections** — when incoming discovery defines terms broadly (e.g., "DOCUMENT" includes emails, texts, etc.), the "problematic definition" objection is automatically included for requests using those terms.

3. **Compound question detection** — requests containing multiple subparts or conjunctive questions ("state all facts AND identify all documents") are flagged and the "compound" objection is auto-included.

4. **"Not applicable" handling** — form interrogatories irrelevant to the case type (e.g., wrongful death interrogatories in a slip-and-fall) auto-respond with "Not applicable" rather than drafting substantive responses.
