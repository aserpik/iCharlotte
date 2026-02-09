# Litigation Report Generator — Design Document

**Date:** 2026-02-07
**Status:** Approved

## Problem

iCharlotte has strong content-generating agents (liability, exposure, medical records, discovery, complaint, docket) but no automated way to assemble their outputs into a properly formatted litigation report. The current workflow is manual copy-paste from agent outputs into Word templates, with moderate reshaping of content to match the report's voice and formatting conventions.

## Goal

A 5-stage pipeline that takes raw agent outputs and case metadata, refines them into report-ready prose matching the user's writing style, and assembles a final .docx report — indistinguishable from a manually written one.

---

## Pipeline Overview

```
Stage 1: GATHER
  → Pull case metadata from database (header fields, parties, case numbers)
  → Collect agent outputs from case folder (AI OUTPUT files)
  → Determine report type (FSR vs Status Report)
  → Identify which sections have content available

Stage 2: REFINE (per-section, parallelizable)
  → For each section with content, LLM reshapes raw output into report-ready prose
  → Uses style guide + example sections for voice matching
  → For sections without data, generates placeholder paragraph
  → For status reports, receives prior report for delta-focused writing

Stage 3: ASSEMBLE
  → Load docxtpl template (extracted from existing report)
  → Slot refined sections into template
  → Populate header metadata from case database
  → Handle section inclusion/exclusion based on report type

Stage 4: POLISH
  → LLM reads full assembled draft
  → Adds transitional phrases and cross-references between sections
  → Ensures tonal consistency
  → Returns targeted edits (not full rewrite) to preserve formatting

Stage 5: OUTPUT
  → Render final .docx file
  → Save to case folder
```

---

## Style Learning (One-Time Setup)

Before the pipeline can refine sections, it learns the user's writing style from ~35 example reports stored in `C:\geminiterminal2\autodownloadreports`.

### Process

1. **Parse all reports** — extract text section by section (Factual Background, Liability, Exposure, etc.)
2. **Build style reference library** — for each section type, store 3-5 best examples
3. **Distill a style guide** — LLM analyzes all examples and produces a concise guide covering:
   - Typical sentence structures and transitions
   - Hedging/qualifying patterns ("We believe...", "It is anticipated that...")
   - Detail level per section type
   - How uncertainty or missing information is handled
   - Formatting conventions (lettered sub-sections, bullet vs. prose, etc.)
4. **Store as config files** — `config/report_style_guide.json` + example sections in `config/report_style_examples/`

### Updatable

The style library can be updated over time by adding new reports to the example corpus and re-running the style extraction. New reports are added to the reference pool, and the style guide is regenerated.

---

## Template Extraction

The docxtpl template is extracted from one of the existing reports:

1. Pick a representative FSR/ILP report (has all sections)
2. Strip case-specific content, replace with Jinja2 template variables:
   - Header: `{{ date }}`, `{{ recipient_name }}`, `{{ bs_file_no }}`, `{{ case_no }}`, etc.
   - Caption: `{{ plaintiff_name }}` v. `{{ defendant_name }}`
   - Section bodies: `{{ factual_background }}`, `{{ procedural_history }}`, etc.
3. Preserve all Word formatting (fonts, margins, styles, spacing, signature block)
4. Add conditional blocks for optional sections:
   ```
   {% if discovery %}
   DISCOVERY
   {{ discovery }}
   {% endif %}
   ```
5. Save as `templates/litigation_report.docx`

**Note:** The existing "report agent" button in the Case View already auto-populates metadata. The template extraction will build on that existing infrastructure.

---

## Section Refinement (Stage 2 Detail)

Each section gets an independent LLM refinement call.

### Inputs

- Raw agent output (e.g., liability.py output for the case)
- Distilled style guide
- 2-3 example sections of the same type from the style library
- Report type context (FSR vs. Status Report)
- For status reports: the prior report's version of this section

### LLM Instructions

- Match tone, sentence structure, and detail level of examples
- FSR: comprehensive coverage of all available information
- Status Report: focus on new developments, reference prior report
- **No hallucination guardrail:** preserve all factual content and legal citations from raw output — do not add information
- Format per section conventions (Liability: lettered sub-headings per COA; Medical: chronological with billing tables; etc.)

### Output

Report-ready prose for that section.

---

## Polish Pass (Stage 4 Detail)

### What It Does

- Adds transitional phrases between sections
- Ensures Exposure references specific findings from Medical Record Review and Liability
- Catches tonal inconsistencies between independently refined sections
- Adjusts Further Case Handling to reference outstanding items from earlier sections

### What It Does NOT Do

- Add new facts or analysis
- Remove or substantially rewrite sections
- Change structure or section ordering

### Technical Approach

- Assembled draft converted to text with section markers
- LLM returns targeted edits (insertions/modifications), not a full rewrite
- Edits applied back to the docx, preserving formatting

---

## Report Type Logic

### FSR / Initial Litigation Plan

- Include ALL sections where agent output exists
- For sections without data, include placeholder: *"Discovery responses have not yet been received. This section will be updated in a subsequent status report once responses are obtained."*
- System checks which agent outputs exist in case folder / database

### Status Report

- **Always include:** Factual Background (brief recap), Procedural History (updated), Evaluation of Liability, Evaluation of Exposure, Settlement Status, Further Case Handling
- **Conditionally include:** Discovery, Medical Record Review — only if new outputs exist since last report
- **Skip:** sections unchanged since last report (not in "always include" list)
- LLM refinement receives prior report to enable delta-focused writing

### New Content Detection

- Compare agent output file timestamps against date of last generated report
- Or compare against a log tracking which outputs were included in previous reports

---

## Agent Output → Report Section Mapping

| Report Section | Agent Source | Output Location |
|---|---|---|
| Factual Background | `complaint.py` | AI OUTPUT folder |
| Procedural History | `docket.py` | AI OUTPUT folder |
| Investigation | (not yet automated) | Manual / future agent |
| Discovery | `summarize_discovery.py` | AI OUTPUT folder |
| Medical Record Review | `med_record.py` + `med_chron.py` | AI OUTPUT folder |
| Evaluation of Liability | `liability.py` | AI OUTPUT folder |
| Evaluation of Exposure | `exposure.py` | AI OUTPUT folder |
| Settlement Status | (not yet automated) | Manual input |
| Further Case Handling | (partial) | Manual input + timeline |

---

## File Structure

```
Scripts/
  report_generator/
    __init__.py
    pipeline.py           # Main orchestrator (5-stage pipeline)
    gather.py             # Stage 1: collect metadata + agent outputs
    refine.py             # Stage 2: per-section LLM refinement
    assemble.py           # Stage 3: docxtpl template rendering
    polish.py             # Stage 4: final LLM polish pass
    style_library.py      # Style guide management (build, update, query)
    template_extractor.py # One-time: extract template from existing report

config/
  report_style_guide.json    # Distilled style rules
  report_style_examples/     # Best example sections by type
  report_section_mapping.json # Maps agent output files → report sections

templates/
  litigation_report.docx     # The docxtpl template
```

---

## Implementation Order

1. **Template extractor** — build the .docx template from an existing report
2. **Style library builder** — parse the 35 reports, distill the style guide
3. **Gather stage** — pull metadata + agent outputs for a case
4. **Refine stage** — per-section refinement with style matching
5. **Assemble stage** — slot everything into the template
6. **Polish stage** — final smoothing pass
7. **Wire together in `pipeline.py`** with CLI interface
8. **Later: UI integration** into the Case View "report agent" button

## Dependencies

- `docxtpl` — Word template rendering with Jinja2
- `python-docx` — Word document manipulation (likely already installed)
- Existing iCharlotte infrastructure: `LLMCaller`, `MasterCaseDatabase`, agent output files

---

## Approach: Standalone Script First, Then UI

The pipeline will first be built as a standalone script (`pipeline.py`) with CLI interface for rapid iteration on template quality, refinement prompts, and style matching. Once the output quality is validated, it will be integrated into the Case View as a button alongside the existing "report agent" functionality.
