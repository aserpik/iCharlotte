# Generate Motion — Wizard Task

**Date:** 2026-06-02
**Status:** Approved design, pending implementation plan
**Scope:** New Wizard task that drafts a California civil motion from scratch using user-supplied target documents and context. Sibling of the existing "Oppose a Motion" task, reusing the `icharlotte_core/opposition/` engine.

## Problem

The Wizard can *oppose* a motion but cannot *bring* one. Attorneys frequently draft affirmative motions (compel, demurrer, strike) from a target document plus their grounds. There is no tool that drafts the Notice of Motion + Memorandum of Points & Authorities with grounded, verified citations.

## Goals

- Add a "Generate Motion" Wizard task that drafts a CA civil motion from scratch.
- Be document-driven: the user uploads the target document(s); the engine analyzes them and **proposes** the grounds/relief and a section outline, both editable before drafting.
- Maximize reuse of the existing opposition engine (research, drafting, citation verification, assembly).
- Ship 3 motion types fully configured, with a generic fallback for any other type.

## Non-Goals (v1)

- No auto-drafting of attachments (meet-and-confer declaration, separate statement) — these are emitted as **labeled placeholders**.
- No Motion for Summary Judgment/Adjudication config in v1 (handled by the generic fallback for now).
- No e-filing, no proposed-order generation beyond placeholders.
- No changes to the Oppose a Motion task.

## Task Wiring

- In-process task (`script_name=""`), like Oppose a Motion.
- `registry.py` entry: `task_id="generate_motion"`, title "Generate a Motion", category **"Motions & Drafting"**, icon glyph ⚖️, `default_folders=["MOTIONS", "PLEADINGS", "DISCOVERY"]`.
- Custom settings page + output page factories (mirrors the oppose_motion registration pattern).

## Settings Page (intake)

A `QStackedWidget` modeled on `OpposeMotionSettingsPage`, with three sub-steps:

1. **Intake**
   - Motion-type dropdown: *Motion to Compel Further Responses* / *Demurrer* / *Motion to Strike* / *Other (generic)*.
   - Target documents via `ContextFilesDialog` (the canonical multi-folder context picker), with type-specific guidance text:
     - Compel → the discovery requests + the served responses at issue
     - Demurrer / Strike → the complaint (or cross-complaint) being challenged
     - Generic → any relevant documents
   - Caption & parties auto-filled from case metadata (`master_db`); editable. Optional hearing date / department / reservation number.
2. **Analyze** (runs on Continue)
   - An LLM pass over the target documents produces (a) **proposed grounds/relief** and (b) a **section outline**.
   - Runs as a Phase-1-style analyze step (the settings page drives it, like the deposition/med-chron analyze flows). Status surfaced to the user.
3. **Review**
   - Editable grounds/relief text — pre-filled with the LLM proposal; the user can edit or append **custom grounds/relief**.
   - Editable outline tree (reuses the Oppose outline editor widgets / `OutlineNode`).
   - **Generate** button → full draft pipeline.

## Pipeline

`intake → analyze target docs → grounds + outline (editable) → research authority → draft sections → verify citations → assemble Word → output`

### New code
- **Per-type target analyzer** — extracts the motion-specific grounds from the target documents:
  - Compel → deficient responses at issue
  - Demurrer → causes of action that fail to state facts sufficient to constitute a cause of action
  - Strike → improper / irrelevant / false matter (e.g. punitive-damages allegations)
  - Generic → general grounds from user-supplied documents
- **`MotionTypeConfig`** dataclass (one per configured type + generic):
  - `type_id`, `display_name`
  - `target_doc_guidance` — intake helper text
  - `legal_standard_hint` — statute + standard to ground the Legal Standard section
  - `section_plan` — default outline nodes
  - `placeholder_attachments` — labels emitted as placeholders
  - `analyzer_prompt` / `grounds_prompt` — per-type LLM guidance
- A new module, e.g. `icharlotte_core/motion_generation/` (config registry, analyzer, glue), keeping it separate from `opposition/` while importing the shared engine pieces.

### Reused from `icharlotte_core/opposition/`
- `argument_research.research_arguments` + local case-law corpus grounding (`local_case_verifier`, `legal_research/local_corpus/`)
- `drafter.draft_memorandum`
- `citation_parser.extract_citations`
- `verifier` / `case_verifier` / `statute_verifier`
- `assembler` (extended for the moving-motion document shape)
- `outline` / `models` (`OutlineNode`, metadata types)
- The anchor-wrapped-citation output page pattern

## Per-Type Configs (v1)

| Type | Legal standard hint | Analyzer extracts | Placeholder attachments |
|------|---------------------|-------------------|-------------------------|
| Compel Further Responses | CCP 2030.300 / 2031.310 | deficient responses at issue | meet-and-confer declaration, separate statement |
| Demurrer | CCP 430.10(e) | causes of action failing to state facts | meet-and-confer declaration |
| Motion to Strike | CCP 435–436 | improper / irrelevant matter | meet-and-confer declaration |
| Generic | (none; user-described) | general grounds | — |

## Output Document

- Notice of Motion + Memorandum of Points & Authorities:
  - Introduction · Statement of Facts · Legal Standard · Argument · Conclusion
- Per-type attachments emitted as **labeled placeholders** (e.g. a "[MEET AND CONFER DECLARATION — to be completed]" section, "[SEPARATE STATEMENT — to be completed]").
- Caption assembled from case metadata.
- MANDATORY: validate the produced `.docx` via `icharlotte_core/word_validator.py` (`validate_report` / `validate_after_edit`), per project rules.

## Citation Grounding & Verification

Reuse the opposition flow exactly: research supporting authority against the local CA case-law corpus when available (falling back to the live CourtListener path), draft with grounded citations, then run the citation/statute verifier over the body. Verified citations are anchor-wrapped in the output page.

## Components / Boundaries

- `registry.py` — task registration only.
- `motion_generation/` (new) — owns motion-type configs, the target analyzer, and the orchestration glue; depends on `opposition/` for the shared engine and on `legal_research/local_corpus/` for grounding.
- `ui/wizard/pages/generate_motion_page.py` (new) — intake/analyze/review settings page + output page; depends on `motion_generation/` and reuses Oppose outline widgets.

## Testing

- **Pure logic (no Qt):**
  - Each `MotionTypeConfig` carries the correct statute string and a non-empty section plan; the configured set is exactly {compel, demurrer, strike, generic}.
  - Target analyzer (mocked LLM) returns expected grounds structures for representative compel/demurrer/strike sample inputs.
  - Generic fallback path: an unconfigured type routes to the generic config and still produces an outline + grounds.
  - Placeholder attachments are present in the assembled document model for each type.
- **pytest-qt** (mirrors `tests/test_wizard/` patterns):
  - Intake → analyze → review → generate page flow; grounds/outline are editable and custom grounds are carried into the draft request.
- **Reuse existing** opposition citation-verification tests for the shared verifier path.

## Risks / Notes

- Legal-standard hints are starting points; tune per type over time.
- Caption/party auto-fill depends on case metadata being populated; intake fields remain editable as the fallback.
- The generic fallback will produce a structurally valid but less type-aware motion — acceptable for v1, flagged in the output.
- Attachment placeholders must be visually obvious so they are never mistaken for completed sections.
