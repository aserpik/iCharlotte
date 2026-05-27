# Oppose a Motion Wizard Task - Design

**Date:** 2026-05-26
**Status:** Approved in brainstorming; pending written-spec review

---

## Overview

Add a new **Oppose a Motion** task to Wizard Mode. The task drafts a California civil litigation opposition memorandum from one uploaded motion and selected case-context documents. It guides the user through motion selection, context selection, motion-type confirmation, outline review, draft generation, legal citation verification, and user-triggered Save As.

V1 generates the opposition memorandum only. It does not generate declarations, exhibit lists, proposed orders, separate statements, or a full filing package.

The feature is limited to California civil litigation. Legal research should prioritize published California Supreme Court and California Court of Appeal authority and should exclude unpublished or non-California authorities from the draft authority pool by default.

---

## Goals

- Add a Wizard card titled **Oppose a Motion**.
- Prompt the user to select one motion to oppose.
- Prompt the user to select one or more context documents.
- Extract text from the motion and context documents with existing iCharlotte extraction paths.
- Auto-detect the motion type, parties, relief requested, and key issues, then require user confirmation/edit before outline generation.
- Generate an editable three-level opposition outline.
- Start every generated outline item selected by default.
- Let the user select/deselect, edit, add, delete, and reorder outline items.
- Draft a comprehensive opposition memorandum from the selected outline items.
- Use context documents for factual grounding, but do not cite context documents in the memorandum.
- Cite legal authority for key arguments using Free Law Project/CourtListener data.
- Verify each legal citation for existence, normalization, and proposition-specific support where possible.
- Show citation details in a right-side source drawer when the user clicks an inline citation.
- Automatically suggest replacement authorities for failed or unsupported citations, but require user acceptance before changing the draft.
- Reuse the existing case caption/template path when available.
- Save the completed `.docx` only when the user clicks **Save** and selects a destination.
- Validate generated Word output with `icharlotte_core/word_validator.py`.

## Non-Goals

- No multi-jurisdiction support in v1.
- No full filing package in v1.
- No automatic citation replacement in the draft.
- No hard export block for unresolved citation issues.
- No citation verification appendix in the exported Word document.
- No separate citation verification report file.
- No record citations to selected context documents.
- No outline regeneration button in v1.
- No real LLM or real CourtListener calls in automated tests.

---

## Recommended Approach

Use a shared opposition engine plus guided Wizard UI.

Add a new `icharlotte_core/opposition/` service layer for extraction orchestration, motion analysis, outline state, drafting, legal research coordination, citation verification, replacement suggestions, and Word assembly. The Wizard page owns user interaction and display state.

This follows the same direction as the guided **Respond to Discovery** work: keep durable behavior outside the UI, let the Wizard drive focused screens, and keep services independently testable.

---

## Planned Modules

| Module | Purpose |
|---|---|
| `icharlotte_core/opposition/models.py` | Dataclasses for motion metadata, outline nodes, section plans, draft sections, and citation verification records. |
| `icharlotte_core/opposition/extraction.py` | Motion/context text extraction wrappers around `DocumentProcessor`. |
| `icharlotte_core/opposition/motion_analyzer.py` | LLM-backed motion-type and issue detection with structured output. |
| `icharlotte_core/opposition/outline.py` | Outline tree creation, validation, selection, edit, add/delete, reorder, and conversion to drafting section plan. |
| `icharlotte_core/opposition/drafter.py` | LLM-backed memorandum drafting from confirmed metadata, selected outline, context text, and researched authority. |
| `icharlotte_core/opposition/citation_verifier.py` | CourtListener citation lookup mapping, opinion support checks, warning states, and replacement suggestions. |
| `icharlotte_core/opposition/assembler.py` | Word preview/final assembly using caption/template when available and `word_validator.py` validation. |
| `icharlotte_core/ui/wizard/pages/oppose_motion_page.py` | Wizard UI for confirmation, outline, draft review, source drawer, and Save As. |

Existing modules to reuse:

- `icharlotte_core/document_processor.py` for PDF, DOCX, TXT, and MSG extraction.
- `icharlotte_core/legal_research/engine.py` and `sources/courtlistener.py` where practical.
- `icharlotte_core/ui/wizard/in_process_task_tab.py` pattern for in-process Wizard flows.
- `icharlotte_core/ui/wizard/pages/output_page.py` save-as behavior as a reference.
- Existing caption/template lookup used by discovery assembly.
- `icharlotte_core/word_validator.py` for mandatory `.docx` validation.

---

## Wizard Registration

Add a `TASK_REGISTRY` entry:

- `task_id`: `oppose_motion`
- `title`: `Oppose a Motion`
- `description`: `Draft and verify an opposition memorandum for a California civil motion.`
- `icon_glyph`: memo or courthouse-style icon
- `script_name`: empty; run through a custom in-process Wizard builder
- `default_folders`: prefer `MOTIONS`, `PLEADINGS`, `DISCOVERY`, then case root

Task routing should treat `oppose_motion` as a guided in-process task.

---

## User Flow

1. User selects **Oppose a Motion** from Wizard Mode.
2. App opens a motion picker popup.
3. User selects one PDF or DOCX motion.
4. App opens a context picker popup.
5. User selects zero or more PDF, DOCX, TXT, or MSG context files.
6. App extracts the motion and context text.
7. App analyzes the motion and displays a confirmation screen.
8. User confirms or edits motion type, parties, relief requested, key dates, and opposition posture.
9. App generates a proposed opposition outline.
10. App displays the editable three-level outline tree with all generated nodes selected.
11. User edits/selects outline nodes and clicks **Generate Draft**.
12. App researches authority, drafts the memorandum, verifies citations, and builds a Word preview.
13. App displays the editable memorandum preview with a right-side source drawer.
14. User clicks a citation to inspect its support in the source drawer.
15. User may accept suggested replacement citations, edit the draft manually, or leave warnings unresolved.
16. User clicks **Save**, chooses a destination, and the app writes the `.docx`.

---

## File Selection

### Motion To Oppose

Allowed:

- PDF
- DOCX

Exactly one motion file is selected. Canceling the dialog cancels task creation.

### Context Documents

Allowed:

- PDF
- DOCX
- TXT
- MSG

Multiple context documents are allowed. No context is allowed, but the app shows a warning:

```text
No context documents selected. The opposition may lack factual support and may need manual factual development.
```

The memorandum may use facts from context documents, but it must not cite those context documents.

---

## Motion Confirmation Screen

The app auto-detects and displays structured metadata:

- motion type
- moving party
- opposing party / responding party
- relief requested
- hearing date if found
- opposition due date if found or inferable
- procedural posture
- moving party's principal arguments
- proposed opposition posture

Required fields before outline generation:

- motion type
- relief requested
- at least one principal argument or issue

If detection is unclear, the screen still opens, but the missing fields are highlighted and the user must fill them in.

---

## Opposition Outline Screen

The outline is a three-level tree:

1. main headings
2. first-level subheadings
3. second-level subheadings

Every generated node starts selected.

Allowed user actions:

- select/deselect any node
- edit node text
- add a node
- delete a node
- reorder nodes within the same level

V1 does not include outline regeneration. Once the first outline is generated, revisions are manual.

Draft generation uses only selected nodes. If a parent heading is deselected, its selected children are ignored unless the user reselects or moves them.

---

## Drafting Behavior

The draft should be a comprehensive and persuasive California civil opposition memorandum. It should:

- follow the confirmed motion type and relief requested;
- address the moving papers' key arguments;
- use the selected outline structure;
- use factual context documents for substance without record citations;
- cite legal authority for key legal propositions;
- avoid citing any authority not present in researched authority data;
- avoid unpublished/noncitable opinions by default;
- distinguish adverse or weaker authority where surfaced;
- preserve a clean filing-focused document body.

The draft prompt should include strict citation instructions similar to existing legal research prompts:

- do not fabricate case names or reporter details;
- cite only authority supplied by the research step;
- state when authority is insufficient instead of inventing support;
- use California citation format.

---

## Legal Research

Research is per selected argument section or tightly grouped issue. The research step should use the existing legal research engine where practical but may need opposition-specific query planning.

Authority prioritization:

1. California Supreme Court published opinions.
2. California Court of Appeal published opinions.
3. Recent directly relevant authority over older authority, unless older authority is controlling or seminal.

Excluded from v1 draft authority pool by default:

- unpublished California opinions;
- non-California authority;
- federal authority;
- secondary sources;
- law review articles.

The UI should label this as Free Law Project/CourtListener-based legal research and verification.

---

## Citation Verification

V1 verification has three layers.

### 1. Existence And Normalization

Use CourtListener's citation lookup API to parse and look up citations in the draft. Map API statuses to app states:

| API Result | App State |
|---|---|
| found / 200 | `exists` |
| not found / 404 | `not_found` |
| invalid reporter / 400 | `invalid` |
| multiple choices / 300 | `ambiguous` |
| throttled / 429 | `throttled` |

Normalized citations should be stored and displayed when available.

### 2. Proposition-Specific Support

For each legal citation, identify the proposition the citation supports in the draft. Fetch the cited opinion text and locate supporting language.

App-level verification states:

| State | Meaning |
|---|---|
| `verified` | Citation exists and supporting opinion language was found for the cited proposition. |
| `exists_support_unconfirmed` | Citation exists, but the app could not tie it to supporting language. |
| `ambiguous` | Citation maps to multiple possible authorities. |
| `not_found` | Citation could not be found in CourtListener. |
| `invalid` | Citation parser found an invalid reporter or malformed citation. |
| `possible_negative_treatment` | Citation exists but later citing material raises caution signals. |
| `throttled` | Verification was delayed or incomplete due to API limits. |

Existence alone is not enough for `verified`.

### 3. Citation-Network Check

Use CourtListener's citation network where available to inspect later citing opinions. This should surface caution signals and later treatment context. It should not be described as equivalent to proprietary Shepard's or KeyCite.

User-facing wording should be precise, for example:

```text
CourtListener citation-network check found later citing opinions. Review treatment before filing.
```

### Replacement Suggestions

For failed, ambiguous, or support-unconfirmed citations:

- automatically search for replacement candidate authorities;
- show replacement candidates in the source drawer;
- include supporting passage and reason for each candidate;
- require explicit user acceptance before changing the draft.

---

## Source Drawer

The output screen uses a split layout:

- left: editable memorandum preview;
- right: source drawer.

Clicking an inline legal citation opens the drawer for that citation.

The drawer displays:

- citation as drafted;
- normalized citation;
- verification state;
- case name;
- court;
- date;
- published/citable status when available;
- supporting passage;
- link to the CourtListener opinion;
- warning details;
- replacement candidates if any.

The drawer state persists in the Wizard task state. It does not create a separate verification report file.

Unresolved citation issues are shown only inside the app. They are not appended to the exported `.docx`.

---

## Word Output And Save Behavior

The app should build an internal Word preview after drafting and verification. It should reuse the existing caption/template path when available, similar to the **Respond to Discovery** task's template behavior.

If a caption/template is missing:

- show a warning;
- allow a body-only preview if assembly can still proceed;
- keep Save As available with clear warning text.

The generated file is not saved to a user-selected location automatically. The output page **Save** button opens a Save As dialog. If the user cancels, no final user-selected file is written.

The app must validate generated `.docx` output with `word_validator.py` before presenting it as ready. If validation fails, show the validation summary.

---

## Persistence

Persist in existing per-case Wizard state:

- selected motion path;
- selected context paths;
- extracted motion metadata;
- confirmed metadata edits;
- selected and edited outline tree;
- draft preview path if any;
- draft text state if needed for restore;
- citation verification records;
- source drawer state.

Do not persist full context document text unless an existing Wizard persistence pattern already permits it. Prefer file paths plus regenerated extraction when possible.

Do not create:

- a separate citation report file;
- a verification appendix;
- a final `.docx` outside the user's Save As choice.

---

## Error Handling

- **No case loaded:** show the existing Wizard-style open-case-first message.
- **Motion selection canceled:** no task tab is created.
- **Unsupported motion file type:** require PDF or DOCX.
- **Unsupported context file type:** ignore or reject with a clear message.
- **Unreadable motion:** keep the task on the confirmation step and show extraction details.
- **Unreadable context file:** continue if other context files succeeded; surface a warning.
- **No context selected:** allow continue with warning.
- **Motion type unclear:** require manual confirmation/edit.
- **CourtListener token missing:** warn before legal drafting; allow user to continue only with unverified/no legal citations.
- **CourtListener unavailable:** show retry/wait option; mark unresolved citations as not fully verified.
- **Citation lookup throttled:** show wait/retry status and preserve partial verification.
- **No supporting passage found:** mark as `exists_support_unconfirmed`.
- **Replacement search fails:** keep original citation warning visible.
- **Caption/template missing:** warn and allow body-only preview when possible.
- **Word validation fails:** show validation summary and do not claim the document is ready.

---

## Testing

Automated tests must avoid real LLM and real API calls.

### Unit Tests

- motion/context file-type validation;
- motion metadata serialization;
- outline tree select/edit/add/delete/reorder behavior;
- selected outline to section-plan conversion;
- citation lookup response mapping;
- support-passage verification states;
- replacement suggestion state;
- Wizard-state round trip for draft and source drawer metadata.

### UI Tests

- Wizard registry shows **Oppose a Motion**;
- selecting the card opens motion picker, then context picker;
- confirmation screen blocks continuation until required fields are present;
- outline screen starts with all generated items selected;
- three-level outline editing works;
- output screen opens the right-side source drawer when a citation is clicked;
- Save opens a Save As dialog and does not save automatically.

### Integration-Style Tests With Mocks

- motion PDF/DOCX extraction produces confirmation metadata;
- outline generation produces a selected section plan;
- draft generation uses factual context but does not cite context documents;
- legal citations are limited to mocked research authorities;
- unsupported citation gets replacement candidates but is not auto-replaced;
- Word preview assembly calls `word_validator.py`.

### Manual Verification

- Run iCharlotte in Wizard Mode against a real California civil motion.
- Confirm caption/template reuse.
- Confirm source drawer opens from inline citation clicks.
- Confirm source drawer links to supporting CourtListener opinion passages.
- Confirm unresolved citation warnings appear only in app.
- Confirm Save asks for a destination.
- Confirm no separate verification report or appendix is created.
- Confirm no Microsoft Word windows are closed during verification.

---

## Open Questions Resolved During Brainstorming

- Jurisdiction scope: California civil litigation only.
- V1 output: opposition memorandum only.
- Motion type: auto-detect, then require confirmation/edit.
- Outline depth: main headings plus two subheading levels.
- Outline defaults: all generated items selected.
- Outline regeneration: omitted in v1.
- Citation review layout: right-side source drawer.
- Citation export blocking: no hard block; unresolved issues are warnings.
- Verification appendix: not included in exported `.docx`.
- Separate verification report: not created.
- Failed citations: automatic replacement suggestions, no automatic draft replacement.
- Save behavior: user-triggered Save As only.
- Factual context: use for grounding but do not cite context documents.
- Motion file types: PDF and DOCX.
- Context file types: PDF, DOCX, TXT, MSG.
- Legal authority pool: published California appellate authority by default.
