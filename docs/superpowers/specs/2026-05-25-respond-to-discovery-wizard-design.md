# Respond to Discovery Wizard Task - Design

**Date:** 2026-05-25
**Status:** Approved in brainstorming; pending written-spec review

---

## Overview

Add a new **Respond to Discovery** task to Wizard Mode. The task guides the user through selecting one incoming discovery PDF, choosing generation rules, selecting context files, reviewing each proposed response one by one, and assembling the final Word response document.

The existing **Discovery -> Respond** subtab stays intact for advanced/manual use. The wizard task reuses shared discovery-response services where practical, but presents a guided workflow built for review and approval.

Supported incoming discovery types:

- Form Interrogatories (FI)
- Special Interrogatories (SI)
- Requests for Admission (RFA)
- Requests for Production of Documents (RPD/RFP)

---

## Goals

- Add a Wizard card titled **Respond to Discovery**.
- Prompt the user to select the discovery PDF after choosing the task.
- Detect the discovery type quickly from filename and, if needed, first-page text.
- Run the current Phase 1 parser for incoming discovery PDFs.
- Show a Rules screen while parsing runs.
- Let the user select/deselect rule checkboxes that control objection and response generation.
- Let the user add custom rules, optionally saved globally.
- Let the user select context files after selecting rules.
- Generate proposed objections and proposed substantive responses.
- Require the user to review and approve every request before final assembly.
- Use the existing response assembly behavior for the final Word output: caption template, preliminary statements, general objections where applicable, waiver language, reservation language, verification, and existing save location.

## Non-goals

- Do not remove or visually redesign the existing Discovery -> Respond subtab.
- Do not make the existing subtab use the new wizard review UI in this pass.
- Do not make RFA/RPD expose visible substantive-response rules yet; keep their substantive defaults in the background.
- Do not make real LLM calls in automated tests.

---

## Architecture

Use **Approach B: shared response engine plus wizard UI**.

The new task should be a guided Wizard Mode front end over shared, non-UI discovery-response services. The existing tab can continue using its current UI while the underlying logic is gradually shared.

Planned new or extracted modules:

| Module | Purpose |
|---|---|
| `icharlotte_core/discovery/response_type_detector.py` | Filename and first-page discovery-type detection. |
| `icharlotte_core/discovery/response_rule_library.py` | Built-in and custom wizard rules. |
| `icharlotte_core/discovery/response_generation_engine.py` | Structured proposal generation from parsed requests, selected rules, and context. |
| `icharlotte_core/discovery/response_review_state.py` | Request-by-request review state, edits, approvals, and final text export. |
| `icharlotte_core/ui/wizard/pages/respond_discovery_page.py` | Wizard UI pages for rules, context selection, and request review. |

The existing modules remain important:

- `response_parser.py` continues to parse incoming discovery PDFs into `ParsedDiscovery`.
- `response_drafter.py` and `objection_selector.py` provide current drafting and objection behavior where it is still appropriate.
- `response_assembler.py` remains the final Word assembly path.
- `word_validator.py` must validate the generated `.docx` output.

The key boundary:

> The shared engine prepares structured proposals; the wizard owns user review and approval; the assembler produces the final Word document only after every request is approved.

---

## Wizard Registration And Task Launch

Add a new task to `TASK_REGISTRY`:

- `task_id`: `respond_to_discovery`
- `title`: `Respond to Discovery`
- `description`: `Draft objections and responses to incoming written discovery.`
- `default_folders`: prefer `DISCOVERY/PROPOUNDED`, then `DISCOVERY`, then case root
- `script_name`: empty; launch through a custom in-process wizard task builder

Task routing should treat `respond_to_discovery` as a special guided task:

1. User selects **Respond to Discovery** from the Wizard tab.
2. App opens a file picker immediately.
3. File picker accepts one PDF file.
4. If the user cancels, no task tab is created.
5. If a file is selected, the wizard opens a task tab and starts discovery-type detection and parsing.

---

## Discovery Type Detection

Detection order:

1. Filename hints.
2. First-page PDF text.
3. User selection if unclear or conflicting.

Filename hints:

| Type | Filename/text hints |
|---|---|
| FI | `frog`, `frogg`, `form interrog`, `form interrogatories` |
| SI | `srog`, `srogg`, `special interrog`, `special interrogatories` |
| RFA | `rfa`, `request for admission`, `requests for admission` |
| RPD/RFP | `rfp`, `rpd`, `request for production`, `requests for production` |

If filename and first-page text conflict, ask the user to choose FI, SI, RFA, or RPD before continuing.

---

## Main Data Flow

1. User selects **Respond to Discovery**.
2. App prompts for one incoming discovery PDF.
3. App detects the discovery type.
4. Phase 1 parser runs and extracts structured requests.
5. While parsing runs, the Rules screen displays rules for the detected discovery type.
6. User selects rules and may add custom rules.
7. User selects context files.
8. Engine generates proposed objections and proposed substantive responses.
9. Wizard displays one request at a time.
10. User edits objections/response text and approves each request.
11. Final assembly is enabled only when every request has been approved.
12. App writes the Word output using the existing response assembler and validates it.

---

## Rule Model

Each rule is a structured object:

| Field | Meaning |
|---|---|
| `id` | Stable internal key. |
| `name` | Exact visible rule wording. |
| `category` | `objection` or `substantive`. |
| `mode` | `mandatory`, `conditional`, or `instruction`. |
| `applies_to` | List of discovery types/modes where the rule appears. |
| `description` | Plain-English explanation shown under the rule. |
| `output_text` | Exact objection text or drafting instruction. |
| `enabled_by_default` | Whether the rule starts checked. |
| `is_global` | Whether a custom rule is saved for reuse across cases/sessions. |

Rule modes:

- `mandatory`: if selected, apply to every request. The LLM does not decide whether it applies.
- `conditional`: if selected, evaluate each request and apply only when the condition is met.
- `instruction`: if selected, pass as a drafting instruction for substantive responses.

The Rules screen must display the user-defined rule wording as the checkbox label. Descriptions and exact inserted text appear under the rule, not instead of the rule name.

---

## Built-In Objection Rules

These rules appear for SI, RFA, RPD/RFP, and FI custom mode.

### 1. ALWAYS include "vague, ambiguous, overbroad" objections

- Category: `objection`
- Mode: `mandatory`
- Default: checked
- Behavior: include for every request when selected, regardless of request text.
- Output text:

```text
Responding Party objects to this Request on the grounds that it calls for speculation and is vague, ambiguous, uncertain and overbroad.
```

### 2. Include ambiguous term objection when the discovery request contains word(s) that are confusing or not obviously clear

- Category: `objection`
- Mode: `conditional`
- Default: checked
- Behavior: include when a request contains a confusing, undefined, or not-obviously-clear term. The agent must identify the term.
- Output text:

```text
Responding Party specifically objects to this Interrogatory on the grounds that the term "[insert unclear term]" is undefined and therefore vague, ambiguous, uncertain, confusing, unintelligible and overbroad. (Code Civ. Proc., section 2030.060, subd. (e).)
```

### 3. Include relevance and privacy objections when Discovery asks for information not related to the case

- Category: `objection`
- Mode: `conditional`
- Default: checked
- Behavior: include when the request seeks information not related to the case or invades privacy.
- Output text:

```text
Responding Party objects to this request on the grounds that it is irrelevant and not reasonably calculated to lead to the discovery of admissible evidence and seeks to invade Responding Party's privacy.
```

### 4. Include burdensome objections when the discovery asks for a lot of information

- Category: `objection`
- Mode: `conditional`
- Default: checked
- Behavior: include when the request asks for a lot of information, lacks reasonable limits, or would impose undue burden or expense.
- Output text:

```text
Responding Party objects to this Request on the grounds that it is unduly burdensome and so overly broad and unlimited as to time and scope as to be an unwarranted annoyance, embarrassment, and is oppressive; to comply with the Request would be an undue burden and expense on Responding Party and is calculated to annoy and harass Responding Party. (See Code of Civ. Proc., section 2030.090, subd. (b); and Columbia Broadcasting System, Inc. v. Super. Ct. (1968) 263 Cal.App.2d 12, 19.).
```

### 5. Include privilege objections when the discovery asks for potentially privileged information

- Category: `objection`
- Mode: `conditional`
- Default: checked
- Behavior: include when the request could seek attorney-client communications, attorney work product, investigation strategy, or other protected information.
- Output text:

```text
Responding Party objects to this request to the extent that it seeks to invade attorney client privilege and/or attorney work product privilege.
```

### 6. Include Expert Opinion or Legal Conclusion objections when Request calls for potential legal conclusion or expert opinion

- Category: `objection`
- Mode: `conditional`
- Default: checked
- Behavior: include when the request calls for expert analysis, expert opinion, legal opinion, or legal characterization.
- Output text:

```text
Responding Party further objects to this request on the grounds that it calls for an expert opinion and a legal conclusion.
```

---

## Built-In Substantive Response Rule

This rule appears for SI and FI custom mode.

### Answer only the question being asked using as few words as possible.

- Category: `substantive`
- Mode: `instruction`
- Default: checked
- Behavior: applies to every drafted substantive answer.
- Instruction text:

```text
Answer only the question being asked using as few words as possible.
```

---

## Discovery-Type-Specific Rules Screen

### Form Interrogatories

The Rules screen first shows:

- `Use Fixed Objections/Responses`
- `Use Custom Objections/Responses`

If `Use Fixed Objections/Responses` is selected:

- Use the fixed FI objection text from `ResponseRules.fi_objections`.
- Use the fixed substantive response for FI 1.1.
- Use the fixed substantive response for FI 15.1.
- Use the fixed substantive response for all FI 16.x.
- Use the FI 17.1 placeholder.
- Other applicable FIs are drafted by the LLM.
- No other rules are displayed on the Rules screen.

If `Use Custom Objections/Responses` is selected:

- Display the same objection rules used for SI.
- Display the same substantive response rule used for SI.

### Special Interrogatories

Display:

- Objection Rules section with all six built-in objection rules.
- Substantive Response Rules section with the minimal-answer rule.
- Custom rule creation controls.

### Requests for Admission

Display only:

- Objection Rules section with the same six built-in objection rules.
- Custom rule creation controls for objection rules.

Substantive response posture remains hidden and uses the current cautious behavior in the background:

- lean toward `Deny` or insufficient information unless the matter is clearly undisputed.

### Requests for Production

Display only:

- Objection Rules section with the same six built-in objection rules.
- Custom rule creation controls for objection rules.

Substantive response posture remains hidden and uses the current context-dependent behavior in the background:

- will comply when documents are available and non-privileged;
- unable to comply when documents are unavailable or the request is too broad.

---

## Custom Rules

The Rules screen includes an **Add Custom Rule** control.

Custom rule fields:

- rule name
- category: objection or substantive
- mode: mandatory, conditional, or instruction
- description: what the rule does
- output text or instruction text
- applies-to discovery types
- save globally checkbox

If the user saves globally, the rule becomes available in future cases/sessions. If global save fails, the rule remains available for the current run and the UI tells the user it was not saved globally.

Initial implementation should wire the custom-rule UI in the wizard only. The shared rule library should be designed so the existing Discovery -> Respond subtab can reuse the same global rules later.

---

## Context File Selection

After the user clicks Next on the Rules screen, show a popup/page to select context files.

Allowed context file types should match the practical current response flow where possible:

- PDF
- DOCX
- TXT

If no context files are selected, allow generation but warn:

```text
No context files selected. Substantive responses may be generic or may need human input.
```

---

## Proposal Generation

For each parsed request, the engine produces:

- `request_number`
- `request_text`
- `proposed_objections`
- `proposed_substantive_response`
- `selected_rule_ids`
- `approval_state`

Generation rules:

- Mandatory objection rules apply deterministically to every request when selected.
- Conditional objection rules are evaluated request by request.
- Substantive instruction rules are passed to the drafting prompt.
- Objections and substantive responses remain separate fields.
- The LLM should not insert waiver language or reservation language; the assembler handles that.

---

## Review Screen

The review screen displays one request at a time.

For each request:

- show request number and full request text;
- show editable proposed objections text box;
- show editable proposed substantive response text box;
- show quick objection checkboxes on the right;
- show Back and Approve + Next controls;
- show progress: approved count and remaining count.

The final assembly button stays disabled until every request is approved.

User actions:

- edit objections;
- edit substantive response;
- toggle quick objections;
- approve current request;
- go back to earlier approved requests and revise them;
- re-approve if a previously approved request is edited.

---

## Quick Objections

The review screen includes these quick objection checkboxes.

### Vague / Ambiguous / Overbroad

```text
Responding Party objects to this Request on the grounds that it calls for speculation and is vague, ambiguous, uncertain and overbroad.
```

### Relevance / Privacy

```text
Responding Party objects to this Request on the grounds that it is not relevant and not reasonably calculated to lead to the discovery of admissible evidence and seeks to invade Responding Party's privacy.
```

### Expert Opinion / Legal Conclusion

```text
Responding Party further objects to this request on the grounds that it calls for an expert opinion and a legal conclusion.
```

### Privilege

```text
Responding Party objects to this request to the extent that it seeks to invade attorney client privilege and/or attorney work product privilege.
```

### Burdensome

```text
Responding Party objects to this Request on the grounds that it is unduly burdensome and so overly broad and unlimited as to time and scope as to be an unwarranted annoyance, embarrassment, and is oppressive; to comply with the Request would be an undue burden and expense on Responding Party and is calculated to annoy and harass Responding Party. (See Code of Civ. Proc., section 2030.090, subd. (b); and Columbia Broadcasting System, Inc. v. Super. Ct. (1968) 263 Cal.App.2d 12, 19.).
```

### Argumentative

```text
Responding Party further objects to this Request on the grounds that it, as phrased, is argumentative and requires the adoption of an assumption, which is improper; the question assumes facts which may or may not be true, but the form of the question requires that the answer adopt the assumption.
```

### Compound

```text
Responding Party objects to this Interrogatory on the grounds that it is compound in form.
```

### Undefined Term

```text
Responding Party specifically objects to this Request on the grounds that the term "{term}" is undefined and therefore vague, ambiguous, uncertain, confusing, unintelligible and overbroad.
```

If the Undefined Term quick objection is toggled and no term is already known for the request, prompt the user for the term before inserting the objection. If the user cancels the prompt, do not add the objection.

Toggling a quick objection on inserts its exact text into the objection text box if it is not already present. Toggling it off removes that exact text if it was inserted by the quick toggle.

---

## Final Assembly

After every request is approved:

1. Convert the review state into the plain-text format expected by `ResponseAssembler`.
2. Use the existing caption document as the template.
3. Include party block, intro, preliminary statement, and RFA/RPD general objections where applicable.
4. Include each request and response.
5. Insert waiver language and reservation language using the existing assembly rules.
6. Insert verification page.
7. Set footer.
8. Save under the current discovery response output location:

```text
NOTES\AI OUTPUT\DISCOVERY RESPONSES
```

9. Validate the generated `.docx` using `icharlotte_core/word_validator.py`.

---

## Persistence

The wizard should persist enough state to recover if the task tab is restored:

- selected discovery PDF
- detected discovery type
- selected rule IDs
- custom run-local rules
- selected context files
- parsed request data
- proposed objections and responses
- user edits
- approval status
- output path after final assembly

Use existing Wizard Mode persistence patterns where possible. Avoid storing full context document text in persistent state; store paths and regenerate/read as needed.

---

## Error Handling

- **No case loaded:** use existing Wizard-style "open a case first" message.
- **No file selected:** cancel without opening a task tab.
- **Unsupported file type:** require a PDF for incoming discovery.
- **Type unclear/conflicting:** ask user to select FI, SI, RFA, or RPD.
- **Parser failure:** keep task tab open and show retry/change-file option.
- **No requests parsed:** show a clear message and allow changing the discovery file.
- **No context files:** allow generation with warning.
- **Unapproved requests:** disable final assembly.
- **Caption template missing:** use the existing save error behavior from the response assembler.
- **Word validation failure:** show validation summary and do not claim success until reviewed/fixed.
- **Custom global rule save failure:** keep the rule for the current run and tell the user it was not saved globally.

---

## Testing

### Unit Tests

- Discovery type detection from filenames.
- Discovery type detection from first-page text.
- Conflict/unclear detection requiring user choice.
- Rule serialization and global/local rule behavior.
- Mandatory rule application to every request.
- Conditional rule application to matching requests only.
- Instruction rule inclusion in drafting prompts.
- Quick objection insertion/removal.
- Undefined-term quick objection prompt behavior.
- Approval gating.

### UI Tests

- Wizard registry shows **Respond to Discovery** task.
- Selecting task opens a single-PDF picker route.
- Rules screen changes by detected type.
- FI fixed mode hides other rules.
- FI custom mode shows SI-style rules.
- RFA/RPD show objection rules only.
- Custom rule creation captures description and global-save option.
- Context selection follows Rules screen.
- Review screen navigates request-by-request.
- Final assembly button remains disabled until all requests are approved.

### Integration-Style Tests

Use mocked LLM responses:

- parse -> rules -> proposals -> review -> assembly input conversion.
- mandatory vague/ambiguous/overbroad rule appears on every request.
- edited review text is what reaches final assembly.
- generated `.docx` output calls `word_validator.py`.

No automated test should make a real LLM call.

---

## Implementation Notes

- Keep generated proposal data structured until final assembly; avoid parsing rendered text until the final Word assembly step.
- Prefer deterministic rule handling for mandatory rules and quick objections.
- Avoid using the LLM to decide whether mandatory rules apply.
- Be careful with truncation: context limits should be treated as guardrails and surfaced when they omit material context.
- Do not close Microsoft Word windows during any manual verification or output-opening flow.
