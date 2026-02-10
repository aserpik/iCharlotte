# Report Review Engine — Design Document

**Date:** 2026-02-10
**Status:** Draft

## Problem

Junior associates submit litigation reports that frequently have issues in three areas:
1. **Formatting** — wrong fonts, missing indentation, headings not styled correctly, metadata table wrong
2. **Content/Completeness** — missing sections, thin analysis, facts omitted
3. **Voice/Tone/Style** — too casual, wrong hedging language, not written from defense counsel perspective

Currently these are reviewed manually. The goal is to automate this review and deliver the result as a redlined (Track Changes) Word document so the associate can see exactly what changed.

## Trigger & UX

### Entry Point
From the existing AI Assistant for Word (hotkey-activated panel in `word_hotkey.py`):
- A new **"Report Review"** button in the assistant panel
- **No text selected** → full document review (Pass 1 + Pass 2)
- **Text selected** → selection-only review (Pass 2 only, full document as context)

### Popup Dialog (`ReportReviewDialog`)
Appears when the button is clicked:

```
┌─────────────────────────────────────┐
│  Report Review Options              │
│                                     │
│  Review scope: ○ Full document      │
│                ○ Selected text only  │
│  (auto-set based on Word selection) │
│                                     │
│  ─── Additional Context ──────────  │
│  ☐ Include case data (AI OUTPUT)    │
│  ☐ Attach additional documents:     │
│     [+ Add files...]                │
│     • Complaint_filed.pdf           │
│     • Depo_Smith.docx         [✕]   │
│                                     │
│  [Start Review]       [Cancel]      │
└─────────────────────────────────────┘
```

- **"Include case data"**: resolves case from document file path, ingests `NOTES/AI OUTPUT` folder using existing `gather.py` logic
- **"Add files"**: file picker filtered to `.docx`, `.doc`, `.pdf`; text extracted via `document_processor.py`
- **Progress**: shown inline in assistant panel — "Pass 1: Scanning structure...", "Pass 2: Reviewing FACTUAL BACKGROUND (2/10)..."

## Architecture: Two-Pass Review

### Pass 1 — Structural Scan (rule-based, fast)

Runs before any LLM calls. No LLM cost. Skipped for selection-only reviews.

**Step 1: Section Detection**
Scan paragraphs for ALL CAPS bold/underlined text. Build a section map with:
- Section name (fuzzy-matched to known 10 sections)
- Start/end character positions in the document
- Heading paragraph reference

Fuzzy matching handles associate variations: "SETTLEMENT" → "SETTLEMENT STATUS", "FACTUAL BACKGROUND AND SUMMARY" → "FACTUAL BACKGROUND".

**Step 2: Formatting Checks**
Reuse existing `word_validator.py` report checks:
- Section headings: ALL CAPS, bold, underline
- Body paragraphs: proper first-line indent (0.5"), correct font/size
- Metadata table: correct fields, proper indentation
- Subheadings: bold, underlined, lettered/numbered
- Empty paragraph bloat
- Salutation, closing block, delivery line

**Step 3: Completeness Check**
Compare detected sections against expected 10-section structure:
- Missing sections entirely
- Sections suspiciously short (< 200 chars when typical is 3000+)
- Sections in wrong order

**Step 4: Summary Display**
Show structured findings in assistant panel before Pass 2:
```
STRUCTURAL SCAN — 5 issues found:
  ERROR: Missing section: EXPERTS
  ERROR: Metadata table missing "Claim No." field
  WARN:  MEDICAL RECORD REVIEW is very thin (180 chars, typical: 4500)
  WARN:  "SETTLEMENT" should be "SETTLEMENT STATUS"
  WARN:  2 subheadings missing bold+underline formatting
```

Pass 2 begins automatically after summary is displayed.

### Pass 2 — Voice/Content Review (LLM, section-by-section)

**Prompt per section:**
```
You are reviewing a litigation report section written by a junior associate.
Your task is to revise it to match the voice, tone, style, and thoroughness
of the senior attorney's writing.

## Style Guide
{style_guide from config/report_style_guide.json}

## Example — how this section should read:
{1-2 gold standard examples from config/report_style_examples/}

## Full document context (for cross-references):
{full document text, truncated if needed}

## The associate's draft for this section:
{section text from the document}

[IF case data enabled:]
## Available case facts (from case file):
{gathered data from AI OUTPUT folder}

[IF additional docs provided:]
## Additional reference documents:
{extracted text from user-selected files}

## Instructions:
- Rewrite this section to match the senior attorney's voice and style
- Fix any formatting issues (headings, structure)
- If case data is provided, add any significant facts/details the associate missed
- Preserve the associate's correct analysis — only change what needs improvement
- Output the revised section text only, no commentary
```

**Key decisions:**
- Each section reviewed independently, but receives full document as context
- LLM returns revised prose (not a list of edits) — RedlineEngine handles the diffing
- Uses `agent_report_review` config in `llm_preferences.json` (same model fallback as `agent_report_refine`)
- Sections flagged as missing in Pass 1 are skipped (can't redline what doesn't exist)

### Selection-Only Mode
- Skip Pass 1 entirely
- Send selected text to LLM with full document as context
- Redline applied only to the selection range
- Same prompt structure, just scoped to the selection

## Redline Application

After each section's LLM call returns revised text:

1. **Identify Word range** — from the section map, get character positions of this section's content (between this heading and next heading)
2. **RedlineEngine** — existing flat-diff approach: flatten original → flatten revised → character-level diff → map to Word paragraphs → apply as Track Changes
3. **Validate** — `validate_after_edit()` on the section range to catch paragraph mark damage or heading corruption
4. **Next section**

**Heading corrections** (e.g., "SETTLEMENT" → "SETTLEMENT STATUS"):
Applied as simple tracked replacements on the heading paragraph, separate from content redline.

**Metadata table fixes:**
Applied as targeted tracked edits to specific table cells.

**Post-review validation:**
After all sections redlined, run full `validate_redline()` suite:
- No revisions outside target ranges
- No paragraph marks destroyed
- No bold headings deleted
- Change ratio reasonable per section

**Completion message:**
```
Review complete — 8 sections reviewed, 47 tracked changes applied.
3 structural issues noted above. Accept/reject changes in Word.
```

## File Structure

### New Files
| File | Purpose |
|------|---------|
| `icharlotte_core/report_reviewer.py` | Main orchestrator class (`ReportReviewer`) |
| `icharlotte_core/ui/report_review_dialog.py` | Popup dialog (`ReportReviewDialog`) |

### Modified Files
| File | Change |
|------|--------|
| `icharlotte_core/word_hotkey.py` | Add "Report Review" button, wire to `ReportReviewer` |
| `icharlotte_core/word_validator.py` | Add section detection + fuzzy matching helpers |
| `config/llm_preferences.json` | Add `agent_report_review` entry |

### No New Dependencies
Everything builds on existing modules: `style_library`, `word_validator`, `RedlineEngine`, `LLMCaller`, `document_processor`, `gather.py`.

## Class Outline

```python
class ReportReviewer:
    def __init__(self, doc_com, case_path=None):
        self.doc_com = doc_com        # win32com Word document
        self.case_path = case_path
        self.style_guide = load_style_guide()
        self.examples = load_style_examples()

    def detect_sections(self) -> list[SectionInfo]:
        """Scan document for ALL CAPS headings, fuzzy-match to known sections."""

    def run_structural_scan(self, sections) -> ValidationResult:
        """Pass 1: rule-based formatting + completeness checks."""

    def review_section(self, section, case_data=None, extra_docs=None) -> str:
        """Pass 2: LLM call for one section, returns revised text."""

    def apply_redline(self, section, revised_text):
        """Apply Track Changes via RedlineEngine for one section."""

    def run(self, selection_range=None, include_case_data=False, extra_doc_paths=None,
            progress_callback=None):
        """Main entry point — full pipeline or selection-only."""
```

## Data Flow

```
ReportReviewDialog
    │
    ▼
ReportReviewer.run()
    │
    ├─► detect_sections()          → list[SectionInfo]
    ├─► run_structural_scan()      → ValidationResult (displayed to user)
    │
    ├─► [if case data] gather.py   → case context text
    ├─► [if extra docs] doc_proc   → extracted text
    │
    ├─► for each section:
    │     ├─► review_section()     → revised text (LLM)
    │     ├─► apply_redline()      → Track Changes in Word
    │     └─► validate_after_edit()→ catch formatting damage
    │
    └─► validate_redline() (full)  → final integrity check
```
