# Quote Insertion for Mediation Brief — Design Spec

**Date:** 2026-04-10
**Status:** Draft

---

## Overview

After a mediation brief has been generated, add a way for users to search deposition transcripts for relevant testimony and insert selected quotes into the brief. Users upload one or more transcripts, describe what they're looking for, and the LLM finds matching Q&A passages. The user selects which quotes to insert, chooses the target section/subsection, and picks between Quick Insert (standalone block) or Weave In (section regeneration with natural integration).

---

## UI: Quote Insertion Dialog

**Trigger:** An "Add Quotes" button appears in the chat area after a mediation brief is generated. Only visible/enabled when `med_brief_generator.is_active` is True and no worker is running. Clicking opens the `QuoteInsertionDialog`.

### Dialog Layout (top to bottom)

1. **Transcript upload area** — File list widget with "Add Transcript(s)" button and drag-and-drop support. Shows filenames with remove buttons. Accepts PDF and DOCX. Multiple files allowed.

2. **Search description** — Multi-line text field with placeholder: "Describe what testimony you're looking for (e.g., 'where plaintiff admits he saw the plastic sheeting before the fall')".

3. **Placement controls** — Horizontal row:
   - Section dropdown — populated from the brief's current sections (e.g., "LIABILITY", "DAMAGES")
   - Subsection dropdown — populated from the selected section's actual subsections (e.g., "A. No Duty of Care"). Defaults to "Auto" (LLM decides placement).
   - Insertion mode toggle — "Quick Insert" / "Weave In" radio buttons

4. **"Search" button** — Triggers LLM search. Shows progress spinner while running.

5. **Results panel** — Scrollable area with quote cards:
   - Checkbox (checked by default)
   - Deponent name and transcript filename
   - Q&A text (Q. / A. formatted)
   - Page:line citation
   - "Edit" button to tweak quote text before insertion

6. **Bottom buttons** — "Insert Selected" (enabled when at least one checked) and "Cancel"

---

## LLM Quote Search

### Process

1. Read all uploaded transcripts (PDF via PyMuPDF, DOCX via python-docx)
2. Send transcript text + user's search description to LLM
3. LLM returns structured results for each match:
   - Deponent last name
   - Verbatim Q&A exchange (Q. and A. lines exactly as they appear in the transcript — no paraphrasing, rewording, or cleanup)
   - Page number and line number range
   - Brief note on why this passage is relevant

### Prompt Structure

- System prompt: transcript analyst role, strict instruction to return testimony exactly as it appears — no changes to wording whatsoever
- Main prompt: user's search description + formatting instructions (return as structured blocks with markers, deponent name, page:line)
- Text: concatenated transcript content with each transcript labeled by filename for attribution

### Multiple Transcripts

Each transcript is labeled with its filename in the prompt. Results include the source filename so quotes can be attributed to the correct deponent.

### Configuration

- Reuses `agent_mediation_brief` from `llm_preferences.json`
- New prompt file: `Scripts/prompts/mediation_brief/quote_search_current.txt` — editable via Workbench
- Registered in `Scripts/prompts/registry.json`

---

## Quote Insertion Logic

### Quick Insert Mode

1. Parse each selected quote into the standard depo quote format (Q./A. lines + citation)
2. Determine insertion point:
   - If user selected a specific subsection → insert after the last paragraph of that subsection
   - If "Auto" → LLM makes a short call to determine which subsection the quote best supports, then insert after the last paragraph of that subsection
3. Append the formatted quote block to the relevant section's text in `self.sections`
4. Reassemble the Word document via `assemble_document()`

### Weave In Mode

1. Take the affected section's current text and the new quote(s)
2. Regenerate that section via `generate_section()` with a refinement instruction: "Incorporate the following deposition testimony into this section at the most appropriate location. Weave it into the argument naturally with proper context and transitions." Include the verbatim Q&A text and citation.
3. If the affected section is in `_INTRO_TRIGGERS` (Liability, Damages, Statement of Facts, Conclusion), also regenerate the Introduction
4. Reassemble the Word document

### Save Behavior

- If a previously saved file path exists on the generator (`self.saved_path`), overwrite it automatically
- Display a "Save As..." link in the chat for saving a copy elsewhere
- If no prior save path exists, open a Save As dialog

---

## Architecture

### New Files

- **`icharlotte_core/ui/quote_dialog.py`** — `QuoteInsertionDialog` class (the dialog UI) and `QuoteSearchWorker(QThread)` for background LLM search
- **`Scripts/prompts/mediation_brief/quote_search_current.txt`** — Editable search prompt

### Modified Files

- **`icharlotte_core/mediation_brief.py`** — Add:
  - `search_quotes(transcripts, description)` method — reads transcripts, calls LLM, returns structured quote results
  - `insert_quotes_quick(quotes, section_name, subsection_title)` method — appends quote blocks to section text
  - `self.saved_path` attribute — tracks last saved file path for overwrite behavior
- **`icharlotte_core/ui/tabs.py`** — Add:
  - "Add Quotes" button (visible after brief generation)
  - Handler to open `QuoteInsertionDialog`
  - Post-insertion chat display and save logic
- **`Scripts/prompts/registry.json`** — Register `mediation_brief:quote_search` prompt

### Background Execution

- **LLM search:** Runs in `QuoteSearchWorker(QThread)` to keep dialog responsive. Progress spinner in dialog.
- **Quick Insert:** Synchronous — just text manipulation + document assembly (fast).
- **Weave In:** Runs section regeneration in a `QThread` since it involves LLM calls. Progress shown in chat.

### State Management

- `MediationBriefGenerator.sections` dict is updated with modified section text after insertion
- Subsequent refinement commands or additional quote insertions work against the latest version
- `self.saved_path` is set after first Save As, used for subsequent overwrites

---

## Error Handling

- **No transcripts uploaded:** "Search" button disabled until at least one transcript is added
- **No matches found:** Display "No matching testimony found. Try broadening your search description." in the results panel
- **LLM failure:** Display error in results panel, allow retry
- **Transcript read failure:** Skip unreadable files with warning, proceed with remaining files
- **Assembly failure:** Display error in chat, sections remain updated in memory for retry
