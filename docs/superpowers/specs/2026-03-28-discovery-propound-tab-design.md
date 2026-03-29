# Discovery Tab — Propound Sub-Tab Design Spec

**Date:** 2026-03-28
**Status:** Approved
**Scope:** Discovery Tab UI + generation engine for propounding SI, RPD, and RFA

---

## Overview

Add a Discovery tab to iCharlotte with two sub-tabs: **Propound** (this spec) and **Respond** (future). The Propound tab allows users to generate discovery requests (Special Interrogatories, Requests for Production, Requests for Admission) using a hybrid approach — deterministic templates for legal boilerplate, LLM for substantive request content when needed.

## Architecture: Hybrid Template + LLM

The document is split into deterministic zones and LLM zones:

| Zone | Method | Examples |
|------|--------|---------|
| Caption/header | Template (from case's Caption Page .docx) | Court name, parties, case number |
| Propounding/Responding block | Template generated from Party data | Party names, set number |
| Preamble | Template per discovery type | CCP statutory citation |
| Instructions | Template per type; Additional mode loads from previous set | CCP boilerplate paragraphs |
| Definitions | Template from DEFINED TERMS.docx; Additional loads from previous set | Party names, incident description |
| Discovery requests | Standard: template copy; Custom/Additional: LLM-generated | The actual interrogatories/demands |
| Signature block | Template from case/config data | Firm name, attorneys, date |
| Declaration | Template with computed values (when SI >35 or RFA >35) | Attorney name, count math, CCP cite |

**Rationale:** Legal documents have zero tolerance for errors in boilerplate, citations, and definitions — those should never go through an LLM. The substantive discovery requests are where LLM creativity adds value.

---

## Module Structure

```
icharlotte_core/
├── discovery/                         # New package
│   ├── __init__.py
│   ├── engine.py                      # Main orchestrator — coordinates generation
│   ├── templates.py                   # Template loading, variable substitution
│   ├── set_tracker.py                 # Scans propounded folder, determines next set #
│   ├── declaration.py                 # Generates SI/RFA declarations with count math
│   ├── assembler.py                   # Renders final .docx from caption template + content
│   └── models.py                      # Data models (DiscoveryRequest, DiscoverySet, Party)
│
├── ui/
│   └── discovery_tab.py               # New — DiscoveryTab (Propound + Respond sub-tabs)
│
discovery/                              # Existing folder — templates & samples
├── Caption Page (AS FM).docx           # Sample caption template
├── DISCOVERY DEFINED TERMS.docx        # Standard definitions
├── Standard Negligence Discovery (5800.070)/
│   ├── SI(1) tPltf.docx               # Standard SI template (63 interrogatories)
│   └── RPD(1) tPltf.docx              # Standard RPD template (34 requests)
└── Standard Wrongful Death Discovery/  # Future expansion
```

### Module Responsibilities

- **`engine.py`** — Orchestrates the generation pipeline. Receives user inputs (mode, types, party, prompt, context docs), delegates to templates/set_tracker/LLM as needed, returns `DiscoverySet` objects for display. Runs LLM calls in a QThread.
- **`templates.py`** — Loads template .docx files and DEFINED TERMS.docx. Performs variable substitution (party names, case number, dates, incident description) on template text. Replaces `____` placeholders with case-specific values.
- **`set_tracker.py`** — Scans the propounded folder (`DISCOVERY/PROPOUNDED/fOUR Client/`) to determine next set number and last request number. Resolves previous set's definitions and format via cascading fallback.
- **`declaration.py`** — Generates declaration text for SI (CCP §2030.070) and RFA (CCP §2033.050) when request count exceeds 35. Computes: previously propounded + this set = total.
- **`assembler.py`** — Takes a `DiscoverySet` + edited plain text from the UI and renders a properly formatted .docx. Opens the caption page template, inserts document title, appends all sections with matching styles.
- **`models.py`** — Data classes for the domain model.

---

## Data Models

```python
class PartyRole(Enum):
    PLAINTIFF = "Plaintiff"
    DEFENDANT = "Defendant"
    CROSS_DEFENDANT = "Cross-Defendant"
    CROSS_COMPLAINANT = "Cross-Complainant"

class Party:
    name: str                     # "Ruxandra Raschkovsky"
    role: PartyRole               # PartyRole.PLAINTIFF
    is_our_client: bool           # True for the represented party
    abbreviation: str             # "Pltf", "City", etc. — auto-generated, user-editable

class DiscoveryMode(Enum):
    INITIAL_STANDARD = "initial_standard"
    INITIAL_CUSTOM = "initial_custom"
    ADDITIONAL = "additional"

class CustomStyle(Enum):
    CUSTOM_ONLY = "custom_only"           # Only LLM-generated requests
    STANDARD_PLUS_CUSTOM = "standard_plus" # Standard template + LLM appended
    MODIFIED_STANDARD = "modified"         # LLM rewrites standard requests

class DiscoveryType(Enum):
    SI = "Special Interrogatories"
    RPD = "Requests for Production"
    RFA = "Requests for Admission"

class DiscoveryRequest:
    number: int                   # 1, 2, 3...
    text: str                     # The interrogatory/request text
    definitions: list[str]        # Inline definitions following this request (if any)

class DiscoverySet:
    discovery_type: DiscoveryType
    set_number: int               # 1, 2, 3...
    directed_to: Party
    propounding_party: Party      # Our client
    requests: list[DiscoveryRequest]
    definitions_block: str        # Full definitions section text
    instructions_block: str       # Instructions section text
    needs_declaration: bool       # True if SI >35 or RFA >35
    previous_count: int           # For declaration: previously propounded count
```

---

## Party Management

The "Directed To" dropdown doubles as the party roster manager:

- **Expanded dropdown** shows all parties grouped by "Opposing Parties" (selectable) and "Our Client" (visible but non-selectable)
- Each party has an edit control for inline name/role editing
- Right-click to remove a party
- "+ Add Party" option at the bottom of the dropdown
- **Abbreviation auto-generated** from name/role (single plaintiff → "Pltf", entities → distinctive word like "City", "Servitek"). User can edit via the inline editor.

### Persistence & Sync

- Party roster stored in case data JSON
- On first load: seeds from existing `plaintiffs` and `defendants` case variables
- **All edits sync back to case variables** — adding, editing, or removing a party in the dropdown immediately updates `plaintiffs`/`defendants` in `MasterCaseDatabase` so other features stay current
- This handles Doe amendments: user adds new parties as they enter the case

---

## Generation Pipeline

### Flow

```
User clicks "Generate"
    │
    ▼
1. Gather Inputs — mode, types, target party, prompt, context docs
    │
    ▼
2. Load Case Data — case number, court, judge, attorneys, caption page path
    │
    ▼
3. Branch by mode:
    │
    ├── STANDARD ──────────────────────────────────────────────┐
    │   Copy requests verbatim from template files             │
    │   (e.g., SI(1) tPltf.docx) with variable substitution   │
    │   No LLM involved.                                      │
    │                                                          │
    ├── CUSTOM ────────────────────────────────────────────────┤
    │   Branch by custom style:                                │
    │   ├── Custom Only: LLM generates all requests            │
    │   ├── Standard + Custom: template requests first,        │
    │   │   LLM appends additional (numbering continues)       │
    │   └── Modified Standard: LLM rewrites template           │
    │       requests based on user prompt                      │
    │                                                          │
    ├── ADDITIONAL ────────────────────────────────────────────┤
    │   a) SetTracker scans propounded folder                  │
    │      → next set number                                   │
    │      → last request number (for numbering continuity)    │
    │   b) Resolve previous set's definitions & format         │
    │   c) LLM generates requests starting at next number      │
    │                                                          │
    ▼                                                          │
4. Assemble DiscoverySet(s) ◄──────────────────────────────────┘
    │  For each checked discovery type:
    │  combine template sections + generated/copied requests
    │
    ▼
5. Display plain text in right-pane editor (one sub-tab per type)
    │
    User edits as needed...
    │
    ▼
6. Save as .docx → NOTES/AI OUTPUT/DISCOVERY REQUESTS/
```

### Custom Mode Sub-Options

| Custom Style | Behavior | LLM Usage |
|---|---|---|
| **Custom Only** | LLM generates all requests from user prompt. Standard requests not included. | Full — generates all requests |
| **Standard + Custom** | Standard template requests included verbatim (e.g., SI Nos. 1–63 from negligence template), then LLM generates additional requests appended after. Numbering continues from last standard request (e.g., starting at No. 64). | Partial — generates only the appended requests |
| **Modified Standard** | Standard template requests sent to LLM with user prompt. LLM returns modified/adapted versions. | Full — rewrites all requests |

### LLM Integration

- LLM calls run in a **QThread** (following existing `LLMWorker` pattern) for UI responsiveness
- User selects provider (Gemini/Claude/OpenAI) and model via dropdowns on left pane
- Provider/model dropdowns reuse existing `ModelFetcher` pattern from ChatTab
- Context documents read via same `read_files_content()` pattern (checked files only)
- LLM prompt instructs the model to return **only numbered request text** in the established format — no boilerplate, definitions, or instructions

---

## SetTracker — Previous Set Resolution

For Additional Discovery, the SetTracker resolves the previous set via cascading fallback:

```
SetTracker.resolve_previous_set(case_path, party, discovery_type)
    │
    ├── 1. Scan propounded folder for filenames
    │      → next set number (from filename pattern "SI (1) tPLF")
    │      → identifies previous set file
    │
    ├── 2. Search for .docx version (best text extraction)
    │      a) NOTES/AI OUTPUT/DISCOVERY REQUESTS/ for matching .docx
    │      b) Same propounded folder for .docx alongside PDF
    │      c) Broader case folder search
    │      → extract: definitions, instructions, last request number
    │
    ├── 3. If no .docx → extract from PDF via PyMuPDF
    │      → parse definitions section, instructions, last request #
    │
    ├── 4. If PDF extraction fails → standard definitions fallback
    │      (DISCOVERY DEFINED TERMS.docx with case variables substituted)
    │
    └── 5. If all fail → notify user
           "Could not read previous set. Drag into Context Documents
            or definitions will default to standard template."
```

**Returns:** next set number, last request number, previous definitions, previous instructions, resolution method.

**Resolution method shown in UI** as a subtle indicator: "Definitions loaded from SI(1) tPltf.docx" or "Using standard definitions (previous set not found)".

### Propounded Folder Structure

```
[Case Folder]/DISCOVERY/PROPOUNDED/fOUR Client/
├── SI (1) tPLF.pdf                    # Single opposing party — files directly here
├── RPD (1) tPLF.pdf
│
├── tPlf/                              # Multiple opposing parties — subfolders
│   ├── SI (1) tPLF.pdf
│   └── RPD (1) tPLF.pdf
├── tCity/
│   └── SI (1) tCity.pdf
```

- Folder name matching is case-insensitive
- Party subfolder prefix is "t" + party abbreviation
- Filename pattern: `[TYPE] ([SET_NUM]) t[PARTY_ABBREV].pdf`

---

## Document Assembly (.docx Output)

### Section Mapping

| Final Document Section | Source |
|---|---|
| Caption page (header, court, parties, case #, title, attorneys) | Case's Caption Page .docx template. Document title substituted per type (e.g., "DEFENDANT ___'S SPECIAL INTERROGATORIES TO ___, SET ___") |
| Propounding/Responding party block | Template — generated from Party data and set number |
| Preamble paragraph | Template per discovery type (SI → CCP §2030.030, RPD → CCP §2031.010, RFA → CCP §2033.010) with party names substituted |
| Instructions to Answering Party | Template per type. Additional mode: from previous set |
| Definitions | Standard mode: from DEFINED TERMS.docx. Additional: from previous set. Case variables substituted in both |
| Discovery requests | Parsed from the edited plain text in the right-pane editor |
| Signature block | Template — firm name, attorneys, date from case/config data |
| Declaration (conditional) | Template for SI >35 (CCP §2030.070) or RFA >35 (CCP §2033.050). Computed: previously propounded + this set = total |

### Caption Page Template

- Each case folder contains a .docx with "Caption Page" in the filename (case-insensitive search)
- Assembler opens it with `python-docx`, preserving existing styles, margins, line numbers
- Document title inserted into the caption table
- Subsequent sections appended as paragraphs using styles from the template
- `///` filler lines added where needed (matching sample convention)

### Plain Text → Structured Requests Parsing

The editor displays plain text. On save, the assembler parses it:
- Lines matching `SPECIAL INTERROGATORY NO. X:` / `REQUEST FOR PRODUCTION NO. X:` / `REQUEST FOR ADMISSION NO. X:` mark request boundaries
- Text between headers = request body
- Parenthetical definitions after a request are preserved as part of that request

### Output Location

Files saved to: `[Case Folder]/NOTES/AI OUTPUT/DISCOVERY REQUESTS/`

### File Naming Convention

Follows the sample pattern: `[TYPE]([SET_NUM]) t[PARTY_ABBREV].docx`

Examples:
- `SI(1) tPltf.docx`
- `RPD(1) tPltf.docx`
- `RFA(1) tCity.docx`
- `SI(2) tPltf.docx` (additional set)

---

## UI Design

### Tab Structure

```
Discovery Tab (QTabWidget)
├── Propound Sub-Tab (this spec)
└── Respond Sub-Tab (future — placeholder)
```

### Propound Layout

**QSplitter** divides left pane (controls) and right pane (editor). Left pane scrolls independently.

#### Left Pane (top to bottom)

1. **Discovery Mode** — Radio buttons: Initial Standard / Initial Custom / Additional
2. **Standard Type** — Dropdown (visible only in Standard mode): Standard Negligence, etc.
3. **Custom Style** — Selector (visible only in Custom mode): Custom Only / Standard + Custom / Modified Standard
4. **Discovery Types** — Checkboxes: SI, RPD, RFA (multi-select)
5. **Directed To** — Custom dropdown with integrated party roster (see Party Management)
6. **Context Documents** — Drag-and-drop file list (reuses ChatTab pattern: `ResizableListWidget` with checkboxes, context menu, OCR support)
7. **LLM Provider/Model** — Two dropdowns (hidden in Standard mode)
8. **Instructions / Prompt** — Text area (hidden in Standard mode; required for Custom and Additional)
9. **Generate Button**

#### Right Pane

1. **Sub-tab bar** — One tab per generated document (e.g., "SI(1) tPltf", "RPD(1) tPltf")
2. **Toolbar** — "Save as .docx" button, "Save All" button, status info (request count, set number, target party)
3. **Plain text editor** — `QPlainTextEdit` displaying the discovery requests for editing
4. **Empty state** — "Configure settings and click Generate to create discovery requests"

### Conditional Visibility

| Control | Initial Standard | Initial Custom | Additional |
|---|:---:|:---:|:---:|
| Standard Type dropdown | Visible | Hidden | Hidden |
| Custom Style selector | Hidden | Visible | Hidden |
| Discovery Types checkboxes | Enabled | Enabled | Enabled |
| Context Documents box | Visible | Visible | Visible |
| LLM Provider/Model | Hidden | Visible | Visible |
| Instructions prompt | Hidden | Visible, required | Visible, required |
| Generate button label | "Generate Discovery" | "Generate Discovery" | "Generate Additional Discovery" |

### Additional Behaviors

- **Standard mode** hides LLM and prompt controls to make it clear no AI is involved
- **Custom "Standard + Custom"** prompt placeholder: "Describe additional requests to generate beyond the standard set..."
- **Custom "Modified Standard"** prompt placeholder: "Describe how to modify the standard requests..."
- **Additional mode** shows brief status during SetTracker scan: "Scanning previous discovery... Found SI(1), RPD(1). Next set: Two."
- **Generate button** disables with spinner during generation
- **Resolution indicator** in Additional mode: "Definitions loaded from SI(1) tPltf.docx" or "Using standard definitions"

---

## Declarations

Auto-generated when request count exceeds 35:

- **SI >35**: Declaration per CCP §2030.070 — attorney attests that the number of interrogatories is warranted by case complexity
- **RFA >35**: Declaration per CCP §2033.050 — same structure, different CCP citation

Declaration includes:
- Attorney name and firm
- Party names (propounding and responding)
- Count of previously propounded requests (0 for initial, N for additional)
- Count of requests in this set
- Total count
- Attestation that each request has been personally examined and is warranted

---

## Future Enhancements (Not in Initial Build)

- **Form Interrogatories** — PDF form-filling for Judicial Council DISC-001
- **Proof of Service** — Auto-generate POS listing all served discovery documents
- **Respond sub-tab** — Draft responses to received discovery
- **Standard Wrongful Death** — Additional standard discovery type template
- **PDF export** — Convert final .docx to PDF for electronic filing
- **Rich text editor** — Upgrade from plain text to WYSIWYG editing

---

## Validation

Per project convention (CLAUDE.md), all generated .docx files must be validated:

```python
from icharlotte_core.word_validator import validate_after_edit
result = validate_after_edit(doc, range_start, range_end)
result.print_summary()
```

---

## Testing Strategy

- **Unit tests** for `set_tracker.py` — mock folder structures, verify set number and request number detection
- **Unit tests** for `templates.py` — verify variable substitution in definitions and instructions
- **Unit tests** for `declaration.py` — verify count math and CCP citations
- **Unit tests** for `assembler.py` — verify .docx output structure against expected sections
- **Integration test** — end-to-end: Standard mode generates correct SI from template with case variables
- **Integration test** — Custom mode sends correct prompt to LLM and assembles result
