# Report Validation & Inspection Tools

**Date:** 2026-02-08
**Scope:** Add validation and inspection capabilities to the report generator pipeline
**Status:** Design

## Problem

Formatting issues in generated reports take 29+ iterations and 3-4 hours to fix because:
1. Claude cannot see the generated documents
2. The user must manually compare output to reference reports and describe issues
3. No automated way to verify formatting correctness after code changes

## Solution: Three Components

### 1. Reference Profiler (Word MCP - interactive, no code)

Claude uses the Word MCP tools (`get_document_xml`, `get_document_outline`, `get_paragraph_text_from_document`) to analyze gold standard reports selected by the user from `autodownloadreports/`.

**What gets profiled:**
- Paragraph styles used for each document region (headings, body, closing)
- Spacing values: `space_after`, `space_before` on key paragraph types
- Indentation: `first_line_indent`, `left_indent`, `hanging_indent`
- Font properties: name, size, bold, underline, color
- Table structure: row/column counts, column widths, table indent
- Subheading patterns: lettered (A., B.) vs numbered (1., 2.) formatting
- Metadata section layout: field order, alignment, salutation format

**Output:** `config/report_reference_profile.json`

```json
{
  "gold_standard_files": ["Carrier001 (Lit Plan).docx", "Carrier001 (FSR).docx"],
  "profiled_date": "2026-02-08",
  "section_heading": {
    "all_caps": true,
    "bold": true,
    "style": "Body",
    "space_before_twips": 0,
    "space_after_twips": 0
  },
  "body_paragraph": {
    "first_line_indent_twips": 720,
    "space_after_twips": 0,
    "font_name": "Times New Roman",
    "font_size_pt": 12
  },
  "subheading_l1": {
    "bold": true,
    "underline": true,
    "hanging_indent_twips": 720,
    "prefix_pattern": "A., B., C."
  },
  "subheading_l2": {
    "bold": true,
    "underline": true,
    "hanging_indent_twips": 720,
    "prefix_pattern": "1., 2., 3."
  },
  "metadata_table": {
    "columns": 1,
    "indent_twips": 1440,
    "col_width_inches": 4.0,
    "borders": "none"
  },
  "salutation": {
    "space_after_twips": 240
  },
  "intro_paragraph": {
    "first_line_indent_twips": 720,
    "space_after_twips": 240
  },
  "closing": {
    "style": "zClosing"
  }
}
```

**When to run:** Whenever the user identifies new gold standard reports or wants to update the profile. Claude does this interactively using MCP tools during a session.

### 2. Validator (`Scripts/report_generator/validate.py`)

Python module using python-docx that checks a generated report against the reference profile and hand-crafted rules.

**Architecture:**

```python
# Each rule is a simple function
def check_metadata_table(doc, profile):
    findings = []
    # ... inspect table properties ...
    return findings

# Rules are just a list - add new ones by appending
RULES = [
    check_metadata_table,
    check_section_headings,
    check_body_paragraphs,
    check_subheadings,
    check_closing,
    check_no_empty_paragraphs,
]

def validate_report(doc_path, profile_path=None):
    """Main entry point. Returns ValidationResult."""
    doc = Document(doc_path)
    profile = load_profile(profile_path)
    findings = []
    for rule in RULES:
        findings.extend(rule(doc, profile))
    return ValidationResult(findings)
```

**Finding object:**

```python
@dataclass
class Finding:
    level: str          # "PASS", "FAIL", "WARN"
    category: str       # "metadata_table", "body_paragraph", etc.
    message: str        # Human-readable description
    paragraph_index: int | None  # Which paragraph, if applicable
    expected: Any       # What the profile says
    actual: Any         # What the document has
```

**Output example:**

```
=== Report Validation: ILP_2026-02-08.docx ===
[PASS] Metadata table indent: 1440 twips
[PASS] Metadata table columns: 1
[FAIL] Paragraph 14 ("Dear Robin,"): space_after=0, expected 240 twips
[PASS] Section heading "FACTUAL BACKGROUND": all-caps, bold, style=Body
[WARN] Body paragraph font: "Calibri" in 3 paragraphs, expected "Times New Roman"
[PASS] Closing style: zClosing
[PASS] No empty paragraphs remaining

Results: 12 PASS, 1 FAIL, 1 WARN
```

**Initial rules (v1):**

| Rule | What it checks |
|------|---------------|
| `check_metadata_table` | Table indent, column width, borders=none, field order, no all-caps names |
| `check_section_headings` | All-caps, bold, correct style, spacing |
| `check_body_paragraphs` | First-line indent (720 twips), font name/size, space-after |
| `check_subheadings` | L1 lettered format, L2 numbered format, hanging indent, bold+underline |
| `check_salutation` | Space-after = 240 twips |
| `check_intro_paragraph` | First-line indent = 720, space-after = 240 |
| `check_closing` | Uses zClosing style, attorney names present |
| `check_no_empty_paragraphs` | No leftover empty paragraphs from template cleanup |

**Adding a new rule:** Write a function, append to `RULES` list. No base classes or registries.

**Accessing formatting values in python-docx:**

```python
# High-level API
para.paragraph_format.first_line_indent   # EMUs
para.paragraph_format.space_after         # EMUs
para.paragraph_format.left_indent         # EMUs
run.font.name, run.font.size, run.font.bold

# Direct XML for properties not in API
pPr = para._element.find(qn('w:pPr'))
ind = pPr.find(qn('w:ind'))
hanging_twips = ind.get(qn('w:hanging'))  # string, e.g. "720"
```

**Unit conversions:**
- 1 inch = 914400 EMUs = 1440 twips
- 1 pt = 12700 EMUs = 20 twips
- python-docx uses EMUs internally; OXML uses twips
- Validator normalizes everything to twips for comparison against profile

### 3. Interactive Inspector (Word MCP workflow)

Not code - a workflow Claude follows during development sessions.

**When a formatting issue is reported:**
1. Open the generated report with `get_document_outline` to see paragraph structure
2. Use `get_paragraph_text_from_document` to inspect specific paragraphs
3. Use `get_document_xml` to see raw OXML when needed
4. Compare against reference report using the same tools
5. Identify the exact difference
6. Fix the code in `assemble.py`
7. Regenerate, re-inspect, confirm fix

**When profiling a new gold standard report:**
1. `get_document_outline` for full paragraph + style map
2. `get_document_xml` for spacing, indentation, font details
3. `get_document_info` for metadata
4. Extract values and update `config/report_reference_profile.json`

## Integration

### Pipeline integration

```
gather → refine → assemble → [VALIDATE] → polish → output
```

Validation runs after assemble (before polish) so we catch formatting issues before the LLM polish pass modifies text. Failures print warnings but do not block generation.

```python
# In pipeline.py run_pipeline()
doc_path = assemble_report(...)

# Optional validation
if not skip_validate:
    result = validate_report(doc_path)
    result.print_summary()
    if result.fail_count > 0:
        logger.warning(f"Validation found {result.fail_count} issues")

polished_path = polish_report(doc_path, ...)
```

### Standalone CLI

```bash
# Validate a specific report
python -m Scripts.report_generator.validate path/to/report.docx

# Validate with custom profile
python -m Scripts.report_generator.validate report.docx --profile config/custom_profile.json

# Verbose output (show PASS results too)
python -m Scripts.report_generator.validate report.docx --verbose
```

## File Changes

| File | Change |
|------|--------|
| `Scripts/report_generator/validate.py` | **NEW** - Validator module |
| `config/report_reference_profile.json` | **NEW** - Reference profile (generated by MCP profiling) |
| `Scripts/report_generator/pipeline.py` | **MODIFY** - Add optional validate step after assemble |

## Existing Pipeline (Unchanged)

The following remain exactly as-is:
- `templates/litigation_report.docx` (docxtpl template)
- `Scripts/report_generator/assemble.py` (document assembly with OXML formatting)
- `Scripts/report_generator/gather.py` (data collection)
- `Scripts/report_generator/refine.py` (LLM refinement)
- `Scripts/report_generator/polish.py` (final polish)
- `Scripts/report_generator/style_library.py` (style guide management)
- `Scripts/report_generator/template_extractor.py` (template creation)

## Implementation Order

1. **Profile gold standard reports** - Use Word MCP to analyze user-selected reference reports, create `report_reference_profile.json`
2. **Build validator v1** - Implement `validate.py` with initial rule set
3. **Test validator** - Run against both reference reports (should pass) and known-bad reports (should catch issues)
4. **Integrate into pipeline** - Add optional validate step to `pipeline.py`
5. **Iterate** - Add rules as new formatting issues are discovered
