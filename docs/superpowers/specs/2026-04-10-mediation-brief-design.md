# Mediation Brief Generator — Design Spec

**Date:** 2026-04-10
**Status:** Draft

---

## Overview

Add a "Mediation Brief" option to the Templates dropdown in the Chat tab. When selected, the app drafts a comprehensive, persuasive defense-side mediation brief using the documents selected in the document box. The brief is generated section-by-section, streamed to the chat window, and assembled into a formatted Word document using the case's existing caption template. After generation, the user can refine individual sections through conversational follow-up.

---

## Output Format

- **Chat window:** Markdown-rendered text streamed section-by-section as each is generated
- **Word document (.docx):** Formatted document built on the case's caption template with line numbers, proper heading styles, indented deposition quotes, and signature block

---

## Architecture

### New Files

- **`icharlotte_core/mediation_brief.py`** — Main module containing:
  - `MediationBriefGenerator` class — pipeline orchestration, section generation, Word assembly, conversational refinement
  - `MediationBriefWorker(QThread)` — background thread for non-blocking generation
  - Caption template handling logic
  - Sample brief style extraction and caching
  - Section text parsing and Word formatting

- **`Scripts/prompts/mediation_brief/`** — 9 versioned prompt files:
  - `planning_current.txt` — metadata/argument extraction from documents
  - `introduction_current.txt` — Introduction section
  - `statement_of_facts_current.txt` — Statement of Facts section
  - `procedural_status_current.txt` — Procedural Status section
  - `liability_current.txt` — Liability section
  - `damages_current.txt` — Damages section
  - `settlement_position_current.txt` — Settlement Position section
  - `conclusion_current.txt` — Conclusion section
  - `routing_current.txt` — Refinement routing (identify which sections to regenerate)

### Modified Files

- **`icharlotte_core/ui/tabs.py`** — Add "Mediation Brief" to Templates menu, confirmation dialog, generation hooks, refinement routing
- **`icharlotte_core/ui/dialogs.py`** — Register "mediation_brief" agent in Workbench mappings
- **`icharlotte_core/chat/models.py`** — Add `builtin_mediation_brief` to `BUILTIN_PROMPTS`
- **`config/llm_preferences.json`** — Register `agent_mediation_brief` for model sequence config
- **`Scripts/prompts/registry.json`** — Register all 9 prompt passes

---

## Pipeline Flow

```
User selects "Mediation Brief" from Templates menu
  -> Confirmation dialog (shows doc count, case name)
  -> User confirms
  -> Find caption template (parent folder, or file picker fallback)
  -> Read & cache sample brief excerpts (first run only)
  -> Read all selected documents
  -> Section generation (sequential, Introduction last):
      0. Planning pass (extract facts, arguments, quotes, dates)
      1. Statement of Facts
      2. Procedural Status
      3. Liability (with level-two subheadings)
      4. Damages (with level-two subheadings)
      5. Settlement Position
      6. Conclusion
      7. Introduction (generated last, sees all other sections)
  -> Each section streams to chat as generated
  -> Assemble .docx from caption template + all sections
  -> Save As dialog (defaults to case parent folder)
  -> Enter refinement mode (user chats to revise sections)
```

---

## UI Changes

### Templates Menu

- Add "Mediation Brief" as a new entry in the Templates dropdown (under built-in prompts)
- Unlike other templates that insert text into the chat input, this entry triggers generation directly

### Confirmation Dialog

When "Mediation Brief" is selected, show a confirmation dialog:
- Display: case name, number of checked documents, list of document filenames
- Buttons: "Generate" and "Cancel"
- On confirm: begin the pipeline

### Chat Display During Generation

- Status updates between sections: "Generating Liability section (4 of 7)..."
- Each section streams tokens to chat as generated
- Visual separators (horizontal rule or bold section label) between sections
- On completion: "Mediation brief generated. Save As dialog opening..."

---

## Caption Template Handling

### Finding the Template

1. Determine the case's parent folder from the selected case path
2. Search for any `.docx` file with "caption" in the filename (case-insensitive)
3. If multiple matches, use the first found
4. If no match, show a file picker dialog: "No caption template found in [folder]. Please select one."

### Processing the Template

1. **Copy** the caption document to a temp file (never modify the original)
2. **Body replacement:** Search all `w:t` elements in the document XML (including nested tables — the caption template may use nested table layouts that `doc.paragraphs` won't reach). Replace "CAPTION PAGE" text with "DEFENDANT'S CONFIDENTIAL MEDIATION BRIEF" — bold, all caps, with "CONFIDENTIAL" additionally underlined. Reuse the XML iteration approach from `icharlotte_core/discovery/assembler.py:_set_caption_title()`
3. **Footer replacement:** Iterate through all section footers, find "CAPTION PAGE" in footer paragraphs, replace with the same styled text
4. **Signature block detection:** Scan from the bottom of the document for signature block indicators (lines with "By:", "DATED:", attorney name patterns, firm name, bar numbers). Extract these paragraphs — they will be appended after the Conclusion section
5. **Content insertion:** After the caption page content and a page break, insert the brief sections

---

## Section Structure & Formatting

### Level 1 Headings

Format: `I.     INTRODUCTION`
- Roman numeral (capitalized) + period + 0.5" tab + heading title in ALL CAPS
- Both roman numeral and title are **bold**
- Title is additionally **underlined**
- Implemented as a paragraph with hanging indent (numeral at left margin, title at 0.5")

### Level 2 Headings (Liability & Damages sections only)

Format: `A.     This Is The Title Of The Subheading`
- Letter (capitalized) + period + 0.5" tab + Title Case heading
- Both letter and title are **bold**
- Title is additionally **underlined**
- Letter numbering restarts at "A." for each Level 1 section
- Same hanging indent pattern as Level 1

### Body Paragraphs

- Normal style matching the caption template's base font
- Standard paragraph spacing (use `space_after`, not empty paragraphs)

### Deposition Quotes

- Left indent 0.5" from margin
- Verbatim from transcript, cleaned up (remove extra numbers, characters, dashes)
- Followed by citation: `(LastName Depo Trns., at p. PgNum:LineNum.)`

---

## Section Content Requirements

### I. INTRODUCTION (generated last)

1. **Paragraph 1:** One-paragraph summary of the basic facts of the case
2. **Paragraph 2:** State whether we dispute or concede liability for mediation purposes. Summarize all key arguments to challenge liability, or if conceding, arguments for comparative fault
3. **Paragraph 3:** State that we challenge the nature and scope of Plaintiff's claimed injuries and damages. Summarize all key arguments to challenge damages
4. **Paragraph 4 (optional):** Any other favorable defense arguments not covered above
5. **Final sentence:** One sentence — we come to mediation in good faith but plaintiff must recognize the problems in their case

The Introduction must be persuasive and comprehensive, summarizing all arguments made in the rest of the brief. Generated last so it can reference actual arguments from other sections.

### II. STATEMENT OF FACTS

Key facts necessary for the reader to understand the case and the arguments in the rest of the brief. Chronological narrative with supporting evidence.

### III. PROCEDURAL STATUS

- Trial date
- Dates of party depositions (if available in documents)

### IV. LIABILITY

- 1-2 sentence introduction paragraph
- Level 2 subheadings — one per key liability argument
- Under each subheading: brief statement of the law, then detailed persuasive application of law to facts explaining why plaintiff cannot establish liability at trial

### V. DAMAGES

- 1-2 sentence introduction paragraph
- Level 2 subheadings — one per key damages argument (letter numbering restarts at A.)
- Under each subheading: detailed persuasive argument for why plaintiff cannot recover claimed damages, referencing specific case facts

### VI. SETTLEMENT POSITION

- Policy limits summary
- Prior offers and demands (to the extent available in documents)

### VII. CONCLUSION

- 1-2 paragraphs summarizing our position and why we believe we will prevail at trial

---

## Sample Brief Style Extraction

### First-Run Extraction

1. Read the 4 sample PDFs from `C:\AI\Mediation Briefs` using PyMuPDF
2. Extract full text from each
3. For each of the 7 sections, identify and extract that section's content by matching the roman numeral + heading pattern
4. Keep the best 1-2 examples per section (selected by length/completeness)
5. Cache to `Scripts/prompts/mediation_brief/style_cache.json` with source file hashes for cache invalidation

### Cache Structure

```json
{
  "source_hashes": {"file1.pdf": "abc123", ...},
  "sections": {
    "introduction": ["excerpt from sample 1", "excerpt from sample 2"],
    "statement_of_facts": ["..."],
    "liability": ["..."],
    "damages": ["..."],
    "procedural_status": ["..."],
    "settlement_position": ["..."],
    "conclusion": ["..."]
  }
}
```

### Runtime Usage

- Each section's LLM call includes cached excerpt(s) for that section as few-shot examples
- Hard-coded style guide (tone, argument structure, citation patterns) always included
- If the samples folder doesn't exist or is empty: log a warning, proceed with style guide only

---

## Section Generation Details

### Per-Section LLM Call Inputs

Each section generation call receives:

1. **System prompt (hard-coded):** Formatting rules, heading structure, citation format, section-specific structural requirements, defense attorney persona
2. **Main prompt (from Workbench):** Editable tone/style/argument instructions
3. **Planning pass output:** Extracted facts, arguments, quotes, dates
4. **Sample excerpt(s):** Cached example of that section from sample briefs
5. **Style guide:** Hard-coded rules derived from sample analysis
6. **Previously generated sections:** For later sections, the text of all earlier sections
7. **Full document content:** Concatenated source documents

### Generation Order

1. Planning pass (extraction, not displayed as a section)
2. Statement of Facts
3. Procedural Status
4. Liability
5. Damages
6. Settlement Position
7. Conclusion
8. Introduction (sees all 6 other completed sections)

### Missing Information Handling

- Write around missing information naturally — no placeholders
- If deposition transcripts aren't in the documents, the brief simply doesn't include deposition quotes
- If trial date or settlement history isn't available, those sections are written without that information

---

## Word Document Assembly

### Step 1 — Prepare Caption Template

- Copy caption document to temp file
- Apply "CAPTION PAGE" replacements (body + footer)
- Detect and extract signature block paragraphs

### Step 2 — Insert Brief Content

After the caption page content + page break, for each section in document order:

1. Parse the LLM's text output into structured elements:
   - Level 1 headings (roman numeral pattern)
   - Level 2 headings (letter pattern)
   - Body paragraphs
   - Deposition quotes (indentation or citation pattern)
2. Insert each element as a formatted Word paragraph with appropriate runs (bold, underline), tab stops, and indentation

### Step 3 — Append Signature Block

If detected in Step 1, append the original signature block paragraphs (with preserved formatting) after the Conclusion section.

### Step 4 — Save

- Open Save As dialog, defaulting to the case's parent folder
- Suggested filename: "Defendant's Confidential Mediation Brief.docx"
- Run `validate_after_edit` from `word_validator.py`

---

## Conversational Refinement

### How It Works

1. After generation, `MediationBriefGenerator` stores the completed state: all 7 section texts, planning pass output, document content, .docx path, caption template path
2. When the user sends a follow-up message, the ChatTab detects the active refinement mode (via a flag on the generator instance)
3. The user's message is sent to an LLM routing call with the `routing` prompt: given the 7 section names and the user's message, identify which section(s) need regeneration
4. Identified section(s) are regenerated with:
   - Original inputs (documents, planning pass, sample excerpts, style guide)
   - Other sections' current text (unchanged)
   - User's refinement instruction appended to the section's main prompt
5. If Introduction-relevant sections change (Liability, Damages), the Introduction is also regenerated
6. Updated sections stream to chat
7. The .docx is reassembled and a new Save As dialog opens

### Routing

- LLM-driven: the routing prompt identifies which section(s) the user's feedback targets
- If the LLM returns "none" (message isn't about the brief), the message is handled as a normal chat message
- Examples:
  - "Make the Damages section more aggressive" -> regenerates Damages
  - "Add a comparative fault argument" -> regenerates Liability, then Introduction
  - "The trial date is March 15, 2027" -> regenerates Procedural Status

### Exiting Refinement Mode

- User can keep refining indefinitely within the same conversation
- Starting a new conversation or switching cases clears the refinement state
- Non-brief messages pass through to normal chat handling

---

## Prompt Management & Workbench Integration

### Workbench Registration

- Register "mediation_brief" as an agent in the Workbench agent list
- Each prompt file becomes a selectable "pass" under the agent
- All Workbench features available: version history, A/B testing, LLM-assisted improvement
- Add `WORKBENCH_TO_AGENT_ID` mapping: `"mediation_brief": "agent_mediation_brief"`

### LLM Config

- Register `agent_mediation_brief` in `config/llm_preferences.json`
- Model sequence: configurable (default Gemini 2.5 Pro -> Claude Opus fallback)
- High max_tokens, 300s timeout for lengthy section output

### Hard-Coded vs. Workbench

| Content | Location |
|---------|----------|
| Section heading formatting rules | Hard-coded in system prompt |
| Section structure requirements | Hard-coded in system prompt |
| Deposition citation format | Hard-coded in system prompt |
| Defense attorney persona | Hard-coded in system prompt |
| Tone, style, argument instructions | Workbench (per-section prompts) |
| Routing logic | Workbench (`routing_current.txt`) |
| Planning extraction instructions | Workbench (`planning_current.txt`) |

---

## Error Handling

- **LLM failure on a section:** Use `LLMCaller` model fallback. If all models fail, display error in chat and save whatever sections completed successfully
- **Caption template not found:** File picker fallback
- **Sample briefs folder missing:** Warning logged, proceed with style guide only
- **Save As cancelled:** Brief stays in memory, user can trigger save again or continue refining
- **Document reading failure:** Skip unreadable documents with a warning in chat, proceed with remaining documents
