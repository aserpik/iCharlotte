# Metadata Section Formatting Fixes - Report Generator

**Date:** 2026-02-08
**File:** `Scripts/report_generator/assemble.py`
**Objective:** Match Report Agent metadata section formatting exactly

---

## Problems Encountered

### 1. Blank Line Insertion Failure
**Problem:** Empty paragraphs inserted via OXML were being removed by `_clean_empty_paragraphs()`

**Attempted Solutions (all failed):**
- Creating empty `<w:p>` elements
- Adding empty runs with space characters
- Using `xml:space='preserve'` with spaces
- Using non-breaking space characters (`\u00A0`)

**Root Cause:** The `_clean_empty_paragraphs()` function runs at the end of `assemble_report()` and removes all empty paragraphs to clean up template bloat. Our manually inserted blank paragraphs were being removed in this cleanup.

**Solution:** Use paragraph spacing properties instead of empty paragraphs
```python
# Add space_after to create visual blank line (12pt = one line)
spacing = OxmlElement('w:spacing')
spacing.set(qn('w:after'), '240')  # 12pt = 240 twips
pPr.append(spacing)
```

### 2. Table Width and Indentation
**Problem:** Metadata table was too narrow and not indented properly

**Solution:**
- Changed from 2-column to 1-column table
- Set table indent to 1.5" (1440 twips)
- Set table width to 4.0" (so 1.5" + 4.0" = 5.5" from left margin)

```python
# Set table left indent to 1.5 inches
tblInd = OxmlElement('w:tblInd')
tblInd.set(qn('w:w'), '1440')  # 1440 twips = 1.5 inches
tblInd.set(qn('w:type'), 'dxa')
tblPr.append(tblInd)

# Set column width
table.columns[0].width = Inches(4.0)
```

### 3. All-Caps Names
**Problem:** Client names from database were in all caps ("MICHELLE MATTHEWS")

**Solution:** Convert to proper case
```python
if client_name and client_name.isupper():
    client_name = client_name.title()
```

### 4. Introduction Paragraph Appearing Twice
**Problem:** Factual background was being used as intro paragraph AND appearing in FACTUAL BACKGROUND section

**Solution:** Use standard intro text instead of factual_background variable
```python
intro = (
    f'As you recall, Bordin Semmer LLP has been retained to represent the interests of {client_name} '
    f'("Defendant") in the above-referenced matter. Please allow the following to serve as our '
    f'Litigation Plan in this matter.'
)
```

### 5. "Re:" Line Formatting
**Problem:** Had leading spaces and incorrect spacing before case name

**Solution:** No leading spaces, align values at position ~20 characters
```python
for label, value in fields:
    if value:
        spacing = max(1, 20 - len(label))
        text = label + " " * spacing + value
    else:
        text = label + " " * 14
```

---

## Final Working Implementation

### Metadata Section Structure
```
VIA EMAIL

[2 blank lines via empty <w:p> elements]

[Table with 1.5" indent, 4.0" width]
    Re:              [Case Name]
    Client\Insured:  [Client Name]
    Claim No.:       [Claim Number]
    BS File No.:     [File Number]
    Case No.:        [Case Number]
    Date of Loss:    [Incident Date]
    Subject:         [Subject]

Dear [Adjuster],                    [space_after = 12pt]

As you recall, Bordin Semmer LLP... [first-line indent = 0.5", space_after = 12pt]

FACTUAL BACKGROUND
```

### Key Code Sections (assemble.py lines 786-850)

**Salutation with spacing:**
```python
sal_para = OxmlElement('w:p')
sal_pPr = OxmlElement('w:pPr')

# Add space_after to create blank line
sal_spacing = OxmlElement('w:spacing')
sal_spacing.set(qn('w:after'), '240')  # 12pt = 240 twips
sal_pPr.append(sal_spacing)
sal_para.append(sal_pPr)

sal_run = OxmlElement('w:r')
sal_text = OxmlElement('w:t')
sal_text.text = salutation
sal_run.append(sal_text)
sal_para.append(sal_run)
```

**Intro paragraph with indent and spacing:**
```python
intro_para = OxmlElement('w:p')
intro_pPr = OxmlElement('w:pPr')

# First-line indent
intro_ind = OxmlElement('w:ind')
intro_ind.set(qn('w:firstLine'), '720')  # 720 twips = 0.5 inch
intro_pPr.append(intro_ind)

# Space after
intro_spacing = OxmlElement('w:spacing')
intro_spacing.set(qn('w:after'), '240')  # 12pt = 240 twips
intro_pPr.append(intro_spacing)

intro_para.append(intro_pPr)
```

---

## Critical Learnings

### Don't Fight the Cleanup Function
- `_clean_empty_paragraphs()` runs at the end and removes empty paragraphs
- Use paragraph properties (space_after, space_before) instead of empty paragraphs
- This is more reliable and Word-native

### OXML Twips Conversions
- 1 inch = 1440 twips
- 1 pt = 20 twips
- 0.5 inch (first-line indent) = 720 twips
- 12pt (blank line) = 240 twips

### Table in Cell Pattern
- Metadata fields are paragraphs INSIDE a table cell, not a traditional table
- This allows proper indentation of the entire metadata block
- Use single-column table with no borders

### Testing Without Full Infrastructure
The issue was we couldn't see the XML output to debug. Created test scripts:
- `test_report_gen.py` - Minimal test with fake data
- `test_full_pipeline.py` - Full pipeline with real case data

But still couldn't inspect XML structure directly - this is why we're adding Office-Word-MCP-Server.

---

## Files Modified

1. **`Scripts/report_generator/assemble.py`**
   - `_build_metadata_table()` - Complete rewrite (lines 627-853)
   - Changed table structure from 2-column to 1-column
   - Added proper spacing using paragraph properties
   - Fixed indentation and width
   - Added proper case conversion

2. **`Scripts/report_generator/gather.py`**
   - Line 172: Changed `variables.pop()` to `variables.get()` to keep factual_background in metadata

3. **`Scripts/report_generator/refine.py`**
   - Updated Discovery section instructions to exclude "Related Litigation Documents"

---

## Next Steps (With MCP Tools)

Once Office-Word-MCP-Server is installed, we can:

1. **Create inspection helpers:**
   - `inspect_metadata_section(docx_path)` - Validate structure
   - `compare_with_reference(generated, reference)` - Show differences
   - `dump_paragraph_properties(docx_path, start, end)` - Debug spacing/indentation

2. **Add validation tests:**
   - Check table indent = 1.5"
   - Check table width = 4.0"
   - Check paragraph spacing after salutation = 12pt
   - Check first-line indent on intro = 0.5"

3. **Document patterns:**
   - When to use space_after vs empty paragraphs
   - How to debug OXML issues systematically
   - Testing workflow for Word document generation

---

## Reference Documents

- Target format: `Z:\Shared\Current Clients\3800- NATIONWIDE\3850\084 - Dudash\STATUS\[DRAFT] Carrier 3850.084.docx`
- Report Agent (legacy): `report_agent.py`
- Template: `templates/litigation_report.docx`
