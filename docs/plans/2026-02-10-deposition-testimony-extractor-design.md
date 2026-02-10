# Deposition Testimony Extractor — Design Document

**Date**: 2026-02-10
**Status**: Approved

## Overview

A new "Depositions" tab in iCharlotte for extracting verbatim deposition testimony by topic. The user loads a PDF transcript, enters a prompt describing what testimony to extract (topics, categories, questions), and receives formatted Q/A pairs with accurate page:line citations — ready for use in legal briefs.

## Requirements

1. **Verbatim accuracy** — extracted testimony must not be changed, truncated, altered, or summarized in any way
2. **Accurate citations** — page and line numbers must reflect the actual transcript locations
3. **Specific format** — Q/A pairs indented 0.5", citations in `(Exh. __, ([LastName] Depo. Trns.) at p. [pg]:[lines].)` format
4. **Both transcript formats** — handles full-size (1 per page) and condensed (4 per page) PDF layouts
5. **Line number/timestamp stripping** — removed from extracted text but tracked as metadata
6. **Highlight feature** — creates `[H.AI] filename.pdf` copy with yellow highlights on extracted text; cumulative across multiple prompts

## Architecture: Three-Stage Pipeline

```
[PDF Transcript] → PARSE → [Q/A Index] → SELECT (LLM) → [Relevant IDs] → OUTPUT → [Word Doc + Highlighted PDF]
```

### Stage 1: Parse (Deterministic)

Python parser reads the PDF and builds a structured Q/A index. No AI involved.

**Format detection** (first few pages):
- **Full-size**: 1 transcript page per PDF page. Page number at top, line numbers 1-25 along left margin.
- **Condensed**: 4 transcript pages per 1 PDF page. Detect by finding multiple page-number headers per PDF page. Split text on page-number boundaries before processing.

**Line-by-line processing**:
1. Strip line numbers: regex `^\s*\d{1,2}\s+`
2. Strip timestamps: regex `\d{1,2}:\d{2}(:\d{2})?`
3. Detect Q/A markers: lines starting with `Q.` or `A.`
4. Continuation lines (no marker) append to current Q or A
5. Track page transitions via page-number headers
6. Skip colloquy, objections (`MR./MS. [NAME]:`, `THE COURT:`, `THE WITNESS:`), and instructions — but track their line positions so subsequent Q/A citations stay accurate

**Output — Q/A Index** (list of exchange objects):
```python
@dataclass
class QAExchange:
    id: int
    question: str        # verbatim, cleaned of line numbers/timestamps
    answer: str          # verbatim, cleaned
    page_start: int      # transcript page number
    line_start: int      # first line of Q
    page_end: int        # transcript page number
    line_end: int        # last line of A
```

**Deponent extraction**: Parse transcript header for "DEPOSITION OF [NAME]", date, case info — reuse patterns from existing `DeponentExtractor` in `summarize_deposition.py`.

**Caching**: Save parsed index as `{transcript_name}.depo_index.json` alongside the PDF. Invalidate if PDF modification time changes.

### Stage 2: Select (LLM)

Send the full Q/A index (with verbatim text and IDs) to the LLM along with the user's prompt. The LLM reads all testimony to judge relevance but returns **only the IDs** of relevant exchanges.

**LLM prompt structure**:
```
You are a legal assistant. Given the following deposition transcript exchanges, identify which ones are relevant to the user's request.

User's request: {prompt}

Return ONLY the IDs of relevant exchanges as a JSON array. Do not modify or reproduce the testimony text.

Exchanges:
[ID: 1] Page 51, Lines 23-25
Q. Did she have any problems with her back...
A. No.

[ID: 2] Page 52, Lines 23-25
Q. Do you know whether your mom had any pain...
A. Prior to the accident, no.
...
```

**LLM response**: `[1, 2, 3, 5, 8, 12]`

**Model**: New agent `agent_depo_extract` in `llm_preferences.json`. Default sequence: Gemini 2.0 Flash (1M context) → Claude Sonnet → GPT-4o.

**Chunking**: For transcripts exceeding context limits (~800+ pages), split the Q/A index into chunks and run selection on each chunk separately. Merge results.

### Stage 3: Output

**Word document** (python-docx):
- Each Q/A exchange formatted with 0.5" left indent
- `Q.` + tab (0.5" tab stop) + question text (hanging indent for wrapped lines)
- `A.` + tab (0.5" tab stop) + answer text (hanging indent)
- **Consecutive grouping**: Adjacent selected exchanges share one citation block
- **Citation format**: `(Exh. __, ([LastName] Depo. Trns.) at p. [page]:[start_line]-[end_line].)`
  - Multi-page: `p. 51:23-52:12`
  - Multiple ranges in one citation: `p. 51:23-52:12; 52:23-25; 53:12-18.`
- **Exhibit number**: Left as `__` (user fills in manually) unless specified in UI
- **Section headers**: Each prompt's results preceded by the prompt text as a header
- **Filename**: `[Extracted] [Deponent Name] Depo Trns.docx`

**PDF highlighting** (PyMuPDF/fitz):
- On first extraction: copy PDF → `[H.AI] original_filename.pdf`
- Locate extracted text on correct PDF pages using `page.search_for(text)`
- Apply yellow highlight annotations: `page.add_highlight_annot(rect)`
- On subsequent extractions: open existing `[H.AI]` copy, add new highlights, save
- Condensed format: search within correct quadrant of PDF page

## UI Design: Depositions Tab

**Split-panel layout** (QSplitter):

```
┌──────────────────────────────────────────────────────────────────────┐
│ [Load Transcript] [Export to Word]  ☐ Highlight text  Deponent: ... │
├─────────────────────────────┬────────────────────────────────────────┤
│   TRANSCRIPT VIEWER         │   EXTRACTION RESULTS                   │
│   (left panel)              │   (right panel)                        │
│                             │                                        │
│   Parsed text with          │   Formatted Q/A extractions            │
│   page/line numbers         │   grouped by prompt                    │
│   Highlights on extracted   │   with citations                       │
│   passages                  │   Clickable → scrolls left panel       │
├─────────────────────────────┴────────────────────────────────────────┤
│ Prompt: [                                                 ] [Extract]│
└──────────────────────────────────────────────────────────────────────┘
```

**Left panel — Transcript Viewer** (QTextEdit, read-only):
- Shows parsed transcript with page headers and line numbers
- Yellow background on text that has been extracted
- Clicking an extraction result in the right panel scrolls here

**Right panel — Extraction Results** (QTextEdit, read-only):
- Formatted Q/A pairs with citations
- Grouped by prompt (prompt text as header)
- New prompt results append below previous ones

**Bottom bar**:
- QLineEdit for prompt input
- QPushButton "Extract"
- Progress bar during LLM call

**Top toolbar**:
- "Load Transcript" — file dialog (PDF filter)
- "Export to Word" — saves all results to .docx
- "Highlight text" checkbox — toggles PDF highlight feature
- Deponent name label (auto-detected)

**Cross-linking**: Clicking a result in the right panel scrolls the left panel to the source location. This uses the page:line metadata from the Q/A index to find the corresponding position in the transcript viewer.

## File Structure

```
icharlotte_core/
├── ui/
│   └── deposition_tab.py          # DepositionTab QWidget
├── deposition/
│   ├── __init__.py
│   ├── transcript_parser.py       # PDF → Q/A index
│   ├── testimony_selector.py      # LLM selection stage
│   ├── testimony_formatter.py     # Word output + PDF highlighting
│   └── models.py                  # QAExchange, TranscriptIndex, DeponentInfo
```

**Integration points**:
- `iCharlotte.py`: Add `DepositionTab` to tab widget, wire `load_case()` callback
- `config/llm_preferences.json`: Add `agent_depo_extract` agent config
- `config/prompts/DEPO_EXTRACT_PROMPT.txt`: LLM system prompt for selection

## Key Technical Decisions

| Decision | Choice | Rationale |
|----------|--------|-----------|
| Verbatim source | Parser index (never LLM) | LLMs alter text even when told "verbatim" |
| Citation source | Parser metadata | Deterministic, mathematically correct |
| LLM role | Relevance selection only | Returns IDs, not text |
| Parse caching | JSON file alongside PDF | Fast re-use across multiple prompts |
| Highlight accumulation | Single `[H.AI]` copy | Multiple prompts build on same highlighted file |
| Q/A scope | Exchanges only | Skip colloquy, objections, instructions |
| Multi-prompt results | Append in right panel | Build up complete extraction document |

## Dependencies

- **PyMuPDF (fitz)** — PDF highlighting (may need `pip install pymupdf`)
- **pypdf** — PDF text extraction (already installed)
- **python-docx** — Word output (already installed)
- **LLMCaller** — existing multi-provider LLM infrastructure

## Testing Strategy

1. **Parser tests**: Parse both full-size and condensed transcripts, verify Q/A count, page/line accuracy
2. **Citation tests**: Verify page:line format matches expected output for known exchanges
3. **Format tests**: Verify Word output matches sample formatting (indentation, tab stops, citation blocks)
4. **Highlight tests**: Verify PDF copy creation, highlight positioning, cumulative behavior
5. **Test transcripts**: Use the Saltarelli transcripts (full + condensed) as test fixtures
