# Respond to Discovery Speed-First Generation - Design

**Date:** 2026-05-25
**Status:** Approved in brainstorming; pending implementation plan

---

## Overview

Improve the Wizard Mode **Respond to Discovery** task so it generates more applicable proposed objections and more accurate substantive responses without adding meaningful friction to the user flow.

The current wizard should stay visually close to the existing workflow: select incoming discovery, choose rules, select context files, review proposed responses, and finalize. The improvement belongs inside proposal generation. Instead of combining simple objection heuristics with one broad substantive-response pass, the wizard should build a request-scoped analysis bundle for each parsed discovery request.

The goal is better first drafts with the same fast review experience.

## Goals

- Keep the wizard speed-first and avoid a new evidence-review workflow.
- Improve conditional objection selection by evaluating each request against the selected rule library and relevant context.
- Improve substantive responses by drafting from request-specific context packets rather than the full raw context blob.
- Keep mandatory objection rules deterministic.
- Keep objection text controlled by the rule library; the model selects rules and fills controlled placeholders.
- Add lightweight `needs_review` warnings for weak context, possible privilege concerns, conflicting facts, or other high-risk drafting conditions.
- Preserve user edits and final assembly behavior.

## Non-Goals

- Do not redesign the existing review screen.
- Do not add source-citation panels by default.
- Do not modify the advanced **Discovery -> Respond** subtab in this pass.
- Do not let the model invent new objection language outside selected custom rule text.
- Do not make automated tests perform real LLM calls.

## Existing Context

The current wizard flow is centered around:

- `icharlotte_core/ui/wizard/pages/respond_discovery_page.py`
- `icharlotte_core/discovery/response_generation_engine.py`
- `icharlotte_core/discovery/response_rule_library.py`
- `icharlotte_core/discovery/response_review_state.py`

The existing generation path already separates proposed objections from proposed substantive responses. That separation should be preserved. The final reviewed text should still flow through the current review state and response assembly path.

## Recommended Approach

Use a hidden, request-scoped proposal pipeline:

1. Parse incoming discovery into structured requests.
2. Read selected context files.
3. Build a lightweight chunk index from the context.
4. For each request, select the most relevant chunks.
5. Generate a structured proposal for that request.
6. Convert the structured proposal into the existing `RequestReview` fields.
7. Show only the proposed objections, proposed substantive response, and short warning flags in the UI.

This keeps the wizard fast while improving the quality of the generated draft.

## Internal Proposal Contract

Each request should produce a structured object with these fields:

```text
request_number
conditional_objection_rule_ids
applied_custom_rule_ids
applied_instruction_rule_ids
ambiguous_term
proposed_objections
proposed_substantive_response
needs_review
review_reason
```

Field behavior:

- `request_number`: must match a parsed request number.
- `conditional_objection_rule_ids`: selected built-in conditional objection rules that apply.
- `applied_custom_rule_ids`: selected custom objection or instruction rules that apply.
- `applied_instruction_rule_ids`: selected substantive instruction rules used for drafting.
- `ambiguous_term`: controlled placeholder value for undefined-term objections.
- `proposed_objections`: final objection text assembled by the application from selected rules.
- `proposed_substantive_response`: substantive answer only; no objections, waiver language, or reservation language.
- `needs_review`: boolean.
- `review_reason`: short user-facing reason when `needs_review` is true.

The model may choose conditional rule IDs, custom rule IDs, instruction rule IDs, and a placeholder term. The application must assemble objection text from the active `ResponseRule` objects. Mandatory rules bypass model selection and always apply when selected.

## Context Chunking

After context files are read, build a compact in-memory chunk list. A chunk should include:

- source path;
- local sequence number;
- text;
- optional heading or nearby heading text when available.

Initial chunking can be simple and deterministic:

- split text by headings and paragraph boundaries where practical;
- keep chunks large enough to contain useful context but small enough to fit several chunks per request;
- deduplicate empty or repeated chunks;
- avoid persisting full context text in wizard state.

The chunk selector should score chunks using:

- exact request terms;
- party names;
- dates;
- incident terms;
- injury, damages, document, and witness terms;
- discovery-type cues such as `identify`, `admit`, `produce`, `documents`, and `communications`;
- nearby headings.

The top chunks become the request-specific context packet. If no useful chunk is found, the packet is empty and the proposal gets a review warning.

## Prompting Rules

The hidden prompt should ask for structured JSON only. It should include:

- discovery type;
- request number and text;
- selected conditional objection rules with IDs and descriptions;
- selected custom rules with IDs and descriptions;
- request-specific context packet;
- response posture for the discovery type;
- output schema.

Hard prompt requirements:

- Do not include objections in the substantive response.
- Do not include waiver or reservation language.
- Do not invent facts not supported by the context packet.
- If context is weak, use cautious default language and set `needs_review`.
- For RFA, admit only when the context clearly supports admission.
- For RPD, say will comply only when context indicates responsive non-privileged documents exist or will be produced.
- For SI and FI, answer narrowly and avoid volunteering extra facts.
- Select only objection rules that are applicable to the request.
- Do not draft new objection text.

## Discovery-Type Behavior

### Form Interrogatories

Fixed FI mode remains the fast path. Known fixed responses and known inapplicable FI ranges should continue to use existing deterministic behavior.

For FI items that still require drafting, use the request-scoped context packet and the narrow-answer posture. If the context packet is empty, generate a cautious response and flag it for review.

### Special Interrogatories

Use the selected objection rules plus the minimal-answer instruction. Draft from the request-specific context packet. If the answer depends on facts not found in context, draft a narrow placeholder-style response and flag it for review.

### Requests For Admission

Use selected objection rules, then choose one of the accepted substantive response forms. The default remains cautious:

- admit only if clearly undisputed and context-supported;
- deny when the matter is disputed or unsupported;
- use insufficient-information language when context is inadequate after reasonable inquiry.

### Requests For Production

Use selected objection rules, then draft a production response based on context:

- will comply when responsive non-privileged documents are identified or expected;
- unable to comply when context indicates no responsive documents are available;
- cautious default when context is weak.

## UI Behavior

The review UI should remain fast:

- continue showing editable objections and substantive response boxes;
- continue quick objection and quick response controls;
- do not show source chunks by default;
- show a short warning label only when `needs_review` is true.

Example warnings:

- `Needs review: no specific context found.`
- `Needs review: possible privilege issue.`
- `Needs review: context appears conflicting.`
- `Needs review: response uses cautious default language.`

The warning should not block approval. The existing approval gate remains the control that prevents accidental final assembly.

## Error Handling

- Invalid model JSON: retry once with a repair prompt, then fall back to current deterministic defaults and flag the request for review.
- Unknown rule ID returned by model: ignore it and add a non-blocking review warning.
- Missing context: allow generation, use cautious defaults, and flag affected requests.
- Overlarge context: use chunk selection rather than truncating the entire context blob.
- Model returns objection text: ignore model objection text and assemble from rule IDs.

## Persistence

Persist only the final review state and lightweight warning fields needed to restore the review screen. Do not persist full context packets or full context file text.

If warning fields are added to `RequestReview`, they should be optional and backward-compatible when loading older saved state.

## Testing

Unit tests should cover:

- mandatory objection rules applying without model approval;
- conditional objection IDs from structured model output;
- unknown model rule IDs being ignored;
- undefined-term placeholder substitution;
- context chunk scoring and request-specific packet selection;
- empty context producing a `needs_review` warning;
- invalid model JSON fallback behavior;
- RFA cautious defaults;
- RPD will-comply versus unable-to-comply defaults;
- SI responses using relevant context instead of unrelated context.

Integration-style tests should use mocked model JSON and verify:

- parsed discovery plus selected rules produces review rows;
- generated objections and substantive responses populate the existing review UI;
- `needs_review` warnings appear without blocking approval;
- user edits override generated proposals;
- final assembly receives approved edited text only;
- no automated test performs a real LLM call.

## Rollout

Roll out only in Wizard Mode first. Leave the advanced **Discovery -> Respond** subtab unchanged.

Implementation should be narrow:

1. Add request-scoped context chunking helpers.
2. Add structured proposal dataclasses and parsing/repair helpers.
3. Update the wizard proposal worker to generate structured per-request proposals.
4. Map proposals into existing review state.
5. Add optional warning display in the review screen.
6. Add focused tests around generation, warnings, and existing review behavior.

This gives the wizard better default output while preserving the existing fast review and final assembly workflow.
