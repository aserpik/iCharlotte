# Mediation Brief Refinement & Quote Insertion from Word AI Assistant

Date: 2026-04-14
Status: Design

## Problem

The chat tab supports two operations on a generated mediation brief:
1. **Conversational refinement** — user types an instruction, `RoutingWorker` picks affected sections, `refine_sections()` regenerates them, `_reassemble_and_save()` rewrites the .docx.
2. **Quote insertion** — "Add Quotes" button opens `QuoteInsertionDialog`, searches deposition transcripts, and inserts formatted Q&A blocks into a chosen section.

Both operations work only from the chat tab and only while `MediationBriefGenerator.is_active == True` in the current session. The user wants the same two operations available from the Word AI assistant popup (Win+V), so they can edit the brief directly in Word without switching back to iCharlotte's chat tab.

## Goals

- Add "Mediation Brief: Refine Section" and "Mediation Brief: Add Quotes" entries to the Word popup's template dropdown when the active document is a brief.
- Edit the live Word document in place (no full .docx regeneration) so manual edits are preserved.
- Work across iCharlotte sessions — no dependency on `generator.is_active` or in-memory state.
- Preserve the existing chat tab flows unchanged.
- Reuse existing components (prompts, search worker, quote result widgets, task manager, redline) rather than duplicating logic.

## Non-goals

- No sidecar JSON for persisting brief state between sessions. The live document is the source of truth.
- No cross-awareness between the chat tab's in-memory generator and Word-initiated edits. The chat tab's `med_brief_generator.sections` may go stale after a Word refine; this is accepted as a known limitation.
- No preservation of `planning_output` across sessions. Refines from Word run without it. If quality degrades noticeably we may revisit; the style excerpts (loaded from disk cache) remain available.
- No edits to the caption page or footer.
- No UI for managing multiple concurrent briefs. If more than one brief is open in Word, the popup operates on whichever document was active when Win+V was pressed (existing popup behavior).

## Architecture overview

The Word popup operates on the live Word document. The `MediationBriefGenerator` is used as a stateless library of prompts and parsing logic — no instance state is carried between invocations.

**Core flow for both operations:**
1. On popup open, detect whether the active document is a brief.
2. Add the two brief entries to the template dropdown if yes.
3. When the user picks one, parse the live document into a `LiveBrief` on demand.
4. Run the operation (LLM refine or quote search) using the parsed state as input.
5. Write the result back into the live document via the existing `TaskManager` insertion path.

## New module: `icharlotte_core/mediation_brief_live.py`

Live-document parser and helpers. No PySide6 UI dependency. Pure logic + Word COM access.

**Dataclasses:**

```python
@dataclass
class LiveSection:
    name: str           # canonical section name (e.g. "liability")
    heading_title: str  # as appears in doc (e.g. "IV. LIABILITY")
    text: str           # body text without heading
    start_para_index: int  # 1-based index into doc.Paragraphs
    end_para_index: int    # 1-based index of last body paragraph

@dataclass
class LiveBrief:
    doc_path: str
    sections: Dict[str, LiveSection]
```

**Functions:**

- `parse_brief_from_word_doc(doc_com) -> LiveBrief` — walks `doc.Paragraphs`, matches each against `_HEADING_PATTERN` from `mediation_brief.py`, maps heading titles to canonical names via `_HEADING_TO_SECTION`, collects body text between headings. Skips unrecognised headings silently.
- `is_mediation_brief(doc_com) -> bool` — returns True if `parse_brief_from_word_doc` finds ≥3 recognised sections. Used to gate the brief template entries.
- `get_word_range_for_section(doc_com, section: LiveSection)` — returns a Word `Range` covering the section body (from end of heading paragraph through the last body paragraph before the next heading).
- `insert_formatted_quotes_at_range(doc_com, range_com, quotes: List[Dict]) -> None` — builds a temp .docx containing the formatted quote blocks by reusing the quote-formatting code from `Scripts/report_generator/assemble.py` (or `mediation_brief.py`'s assembler — whichever holds it), then calls `range_com.InsertFile(tmp_path)`. Deletes the temp file in a `finally` block.

The heading regex and `_HEADING_TO_SECTION` map are imported from `mediation_brief.py` to keep a single source of truth.

## New method: `MediationBriefGenerator.refine_section_standalone`

Stateless sibling of `refine_sections`:

```python
def refine_section_standalone(
    self,
    section_name: str,
    sections_dict: Dict[str, str],
    instruction: str,
) -> str:
    """Run a section refinement without mutating self.sections.

    Uses sections_dict as the context instead of self.sections.
    Reads style excerpts from the on-disk cache via get_style_excerpts().
    Uses self.planning_output if set, empty string otherwise.
    Returns the refined section text.
    """
```

Internally this shares prompt construction with the existing `refine_sections` path. The existing `generate_section` method is refactored to accept an optional `sections_context` parameter, with the default behavior (reading from `self.sections`) unchanged so the chat tab flow is untouched.

## New dialog: `icharlotte_core/ui/quote_dialog_word.py`

`WordQuoteInsertionDialog` — a stripped variant of `QuoteInsertionDialog`.

**Reuses from the existing dialog (imported, not duplicated):** `QuoteSearchWorker`, `QuoteResultWidget`, transcript list widget construction, search description widget, results scroll area.

**Removed from the existing dialog:** section combo, subsection combo, Quick Insert / Weave In radio buttons. Quick Insert is the only mode; the target is the Word cursor/selection range, not a section.

**Signal:** `quotes_to_insert(quotes: List[Dict])` — flat list, no section or mode info.

**Constructor:** `__init__(self, parent=None)` — takes no generator instance. The dialog instantiates a throwaway `MediationBriefGenerator()` internally only for `search_quotes()`, which does not depend on any mutable instance state.

## Word popup changes (`word_hotkey.py`)

**Template dropdown entries (added dynamically):**

On popup open, after `self._original_document` is captured:
1. Call `is_mediation_brief(self._original_document)`.
2. If True, add "Mediation Brief: Refine Section" and "Mediation Brief: Add Quotes" to the existing template dropdown.
3. If False, leave the dropdown unchanged.

These entries are removed if the popup is reused with a different document in which `is_mediation_brief` returns False.

**New section combo:**

A `QComboBox` populated from `SECTION_ORDER` via `SECTION_HEADINGS[...][1]` display names. Hidden by default. Shown only when the user selects "Mediation Brief: Refine Section" from the template dropdown. Hidden again if the user switches away to another template.

**Dispatch in `_do_execute`:**

Branch on the selected template:

- **"Mediation Brief: Refine Section"**:
  1. Parse the live doc via `parse_brief_from_word_doc`.
  2. Look up the picked section in `live.sections`. If missing → status toast "Section not found in this document", abort.
  3. Build `sections_dict = {name: sec.text for name, sec in live.sections.items()}`.
  4. Resolve the target Word range via `get_word_range_for_section`.
  5. Create a task bookmark on the range (`_create_task_bookmark`) with a new task id.
  6. Submit a `TaskData` to `TaskManager`. The task's LLM call is `MediationBriefGenerator().refine_section_standalone(section, sections_dict, instruction)` — executed on `TaskLLMWorkerThread`.
  7. Existing insertion path handles replacement at the bookmark. Redline checkbox stays orthogonal — if active, the existing redline path runs over the bookmarked range.

- **"Mediation Brief: Add Quotes"**:
  1. Capture the current Word selection/cursor as a bookmarked insertion range.
  2. Hide the popup.
  3. Open `WordQuoteInsertionDialog` modally.
  4. On `quotes_to_insert` signal, submit a synthetic task to `TaskManager` whose insertion callback is `insert_formatted_quotes_at_range(doc, resolved_range, quotes)`. The "LLM call" portion is a no-op returning a sentinel — this keeps the status-bar UX and serialization consistent with refine tasks.
  5. Task manager resolves the bookmark, runs the callback, removes the bookmark, deletes the temp file.

## Concurrency

Both operations go through the existing `TaskManager` sequential insertion queue. This gives us:
- Serialized writes into the live Word document, avoiding COM contention.
- Consistent status-bar UX (task rows, spinners, completion signals) across refine and quote insertion.
- Bookmark-based range tracking so a refine on Liability running concurrently with a quote insertion in Damages both land correctly.

Known concurrency limitation: a second refine on the same section started while the first is pending operates on parsed text that predates the first's changes. Same limitation exists in the chat tab today. Accepted.

## Error handling

**Detection failures:**
- `is_mediation_brief` False → brief entries don't appear, user sees the normal popup. No error.

**Refine failures:**
- Section not present in parsed result → inline warning, abort before task submission.
- LLM returns empty → existing `TaskManager` error path fires, bookmark cleanup in `finally`.
- Document closed between Execute and worker completion → existing COM-dead-reference handling in `_InsertionProxy` fires. Log and drop the result.

**Quote insertion failures:**
- No transcripts added → Search button disabled (existing dialog behavior).
- `search_quotes` returns `[]` → "No matches found" status in dialog, user can revise or cancel.
- `Range.InsertFile` fails → catch, status toast, abandon insertion, delete bookmark and temp file in `finally`.

**Word validator gate (per CLAUDE.md mandatory rule):**
- After every refine insertion: `word_validator.validate_after_edit(doc, range_start, range_end)`.
- After every quote insertion: `word_validator.validate_after_edit(doc, range_start, range_end)`.
- If redline mode is active during refine: additionally `word_validator.validate_redline(...)` per the existing redline path.

## Testing

**Unit (offline, no Word COM):**
- `parse_brief_from_word_doc` against a fake doc-like object with mocked `Paragraphs` yielding real heading/body text. Assert sections dict matches expected canonical names and body text.
- `is_mediation_brief` True for ≥3 recognised headings, False otherwise, False for empty doc.
- `refine_section_standalone` with mocked `LLMCaller.call` — assert the prompt includes style excerpts and the current section text, assert the return value matches the mock.
- Temp quote docx helper — generate from known quotes, re-open with python-docx, assert paragraph count, indents, and line spacing match the assembler's existing output.

**Integration (live Word COM, manual checklist):**
- Generate a brief via chat tab, keep Word doc open, Win+V → verify brief entries appear in dropdown.
- Refine Liability with a small instruction → verify target section changes, verify other sections byte-for-byte identical.
- Refine Liability with redline checkbox enabled → verify Track Changes appear, verify unchanged sentences are not redlined.
- Add Quotes at cursor inside Liability, pick 2 quotes → verify Q/A hanging indent, citation line, surrounding paragraphs untouched.
- Close iCharlotte, reopen brief doc in Word, Win+V → verify refine and add-quotes still work (cross-session path).
- Open a non-brief Word doc, Win+V → verify brief entries absent.
- Trigger two concurrent refines on different sections → verify both land without clobbering.

## Files touched

**New:**
- `icharlotte_core/mediation_brief_live.py`
- `icharlotte_core/ui/quote_dialog_word.py`
- `tests/test_mediation_brief_live.py`
- `tests/test_refine_section_standalone.py`

**Modified:**
- `icharlotte_core/mediation_brief.py` — add `refine_section_standalone`, refactor `generate_section` to accept optional `sections_context`.
- `icharlotte_core/word_hotkey.py` — template dropdown entries, section combo, dispatch branches in `_do_execute`, quote insertion path, task bookmark for section ranges.

**Unchanged (important):**
- `icharlotte_core/ui/tabs.py` — chat tab integration stays as is.
- `icharlotte_core/ui/quote_dialog.py` — existing dialog stays as is; new `quote_dialog_word.py` imports shared widgets.

## Risks

- **Temp-docx formatting fidelity.** `Range.InsertFile` inherits the inserted doc's direct formatting. Since the assembler uses direct formatting (not style names), this should match the active doc. Verify during implementation — if style mismatches occur, fall back to COM-level paragraph formatting.
- **Section range boundary off-by-one.** Mapping python-docx section text back to 1-based `doc.Paragraphs` indices is error-prone. Write a focused test that round-trips: parse → get range → read range text → compare against parsed section text.
- **`planning_output` loss** on cross-session refines may degrade quality for Introduction-heavy instructions. Accepted as Approach C's explicit trade-off. Monitor and revisit if needed.
- **Concurrent refine on the same section** uses stale parsed text for the second task. Pre-existing limitation, not worse than the chat tab.
