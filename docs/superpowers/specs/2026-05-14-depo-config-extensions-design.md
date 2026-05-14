# Deposition Summary Config Dialog — Extensions Design

**Date:** 2026-05-14
**Status:** Approved for implementation
**Owner:** iCharlotte deposition workflow
**Builds on:** `2026-05-14-interactive-depo-summary-design.md`

## Goal

Extend `DepoSummaryConfigDialog` with three user-controlled inputs:

1. **Drag-and-drop topic reordering** — the user can rearrange the topic list to control the order topics appear in the final summary.
2. **Summary bias control** — a dropdown with presets (Neutral, Most favorable to plaintiff, Most favorable to defense, Custom…) plus a free-text field for the Custom… option.
3. **Context documents drop zone** — the user can drop `.pdf`, `.doc`, `.docx` files into the popup; Phase 2 reads them and injects their full text into the summary prompt to inform the LLM's bias judgments and content selection.

## Motivation

The first round of the interactive flow is working ("worked pretty good"), but the user has identified three concrete extensions they want:

- The agent's rank order isn't always the order the attorney wants in the final summary; drag-reorder gives full control.
- "Custom rules" is a generic catch-all; an explicit bias control communicates intent to the LLM more clearly and gives the user a fast preset for common slants.
- Sometimes a deposition summary needs to be informed by the complaint, prior medical records, or another related document. Dropping these into the popup makes the agent aware of them without manual prompt engineering.

## Out of scope (v1 of this extension)

- Persistence across runs. The popup still resets to defaults (Neutral bias, no context docs) every time, matching the original v1 rule.
- A "preview the merged prompt" feature. The user trusts the resulting summary; debugging the rendered prompt is a developer concern.
- Per-doc weighting or per-doc instructions. All context docs are concatenated equally into the prompt.
- Smart token-budgeting beyond a hard per-doc character cap.

## Architecture

The dialog stays a single modal `QDialog`. New fields land in the existing sidecar session JSON's `user_config` block. Phase 2 extracts the context documents lazily (at the start of its run, not in the dialog) so the popup stays snappy regardless of doc size.

## Session JSON additions

The `user_config` block grows by three fields:

```json
"user_config": {
  "selected_topics": ["Topic A", "Topic B"],
  "added_topics": ["Custom topic"],
  "bullets_per_topic": 5,
  "deponent_label": "Plaintiff",
  "custom_rules": "Use past tense.",
  "cross_check_enabled": true,
  "bias": "neutral",
  "bias_custom": "",
  "context_doc_paths": [
    "Z:\\...\\Smith_complaint.pdf",
    "C:\\...\\medical_record_summary.docx"
  ]
}
```

- `bias` is one of: `"neutral"`, `"pro_plaintiff"`, `"pro_defense"`, `"custom"`. Always present. Defaults to `"neutral"`.
- `bias_custom` is the free-text directive when `bias == "custom"`. Empty string otherwise.
- `context_doc_paths` is a list of absolute paths to dropped or browsed files. Empty list by default. Paths only — extraction is Phase 2's job.

`selected_topics` order semantics change: previously the agent's rank order, now whatever order the user dragged them into.

## UI changes — `DepoSummaryConfigDialog`

### Topics panel becomes a `QListWidget`

Replace the existing `QVBoxLayout` of `_TopicRow` widgets with a `QListWidget` (`self.topics_list`) configured with:

- `setDragDropMode(QListWidget.InternalMove)`
- `setSelectionMode(QListWidget.SingleSelection)`
- `setDefaultDropAction(Qt.MoveAction)`

Each topic stays a `_TopicRow` (checkbox + editable QLineEdit) inserted via `QListWidget.setItemWidget(item, row_widget)`. Tests change from iterating `self.topic_rows` (a list) to iterating via a new helper `self.topic_rows_in_order()` that yields each row widget in current visual order.

### Summary bias row

New horizontal row placed between the topics list and the additional-topics field:

```
Summary bias: [ Neutral ▾ ]  [ ← custom field, hidden until Custom… is selected ]
```

- `QComboBox` (`self.bias_combo`) items: *Neutral*, *Most favorable to plaintiff*, *Most favorable to defense*, *Custom…*. Internal data values: `"neutral"`, `"pro_plaintiff"`, `"pro_defense"`, `"custom"`.
- `QLineEdit` (`self.bias_custom_edit`) sits next to the combo, hidden by default, with placeholder text "Describe the editorial lens (e.g., 'Highlight any inconsistencies in injury testimony')".
- `currentIndexChanged` handler toggles the line edit's visibility. Switching AWAY from *Custom…* clears `bias_custom_edit.setText("")` so leftover text doesn't accidentally ship.

### Context documents panel

New panel placed between the bias row and the bullets/deponent/cross-check settings row.

- `QListWidget` (`self.context_docs_list`), ~80px tall, with `setAcceptDrops(True)` and overrides on `dragEnterEvent` and `dropEvent`.
- Drop event accepts URLs with `.pdf`, `.doc`, `.docx` extensions only; other extensions get a status label "Unsupported file type — skipped" that auto-clears after 3 seconds.
- Each accepted drop appends one `QListWidgetItem`. The item widget is a small horizontal layout with `<basename>` label and an `×` `QPushButton` that removes the row.
- `QTimer.singleShot(0, ...)` defers the drop handling so the OS shell isn't blocked (per existing memory note about Explorer drops).
- An "Add files…" `QPushButton` next to the list opens a `QFileDialog.getOpenFileNames(filter="Documents (*.pdf *.doc *.docx)")` for users who prefer click-to-browse.
- Internal state: `self._context_doc_paths: list[Path]` — populated and pruned as the user adds/removes; the QListWidget is purely a view.

### Dialog default size

Bumped from 700×600 to 800×750 to fit the new sections without cramping the topics scroll.

## Validation at dialog Accept

- `selected_topics + added_topics` must be non-empty (existing rule).
- `bias` always set (defaults to `"neutral"` if combobox somehow lands at the initial state).
- `bias_custom` is non-empty only when `bias == "custom"`; otherwise stored as empty string.
- `context_doc_paths` is filtered at Accept time — paths whose files have been moved or deleted between drop and Accept are silently dropped from the list. (No popup error; user can re-add if they want.)

## Phase 2 changes

### New helper: `_extract_context_documents(paths, logger) -> list[dict]`

Reads each context document as text. Returns `[{filename, text}, ...]` in the same order as the input paths.

- `.pdf` → `DocumentProcessor(ocr_config=OCRConfig(adaptive=True)).extract_with_dynamic_ocr(path)`
- `.docx` → `extract_docx_text(path)` from `icharlotte_core.document_processor` (the helper that walks `doc.element.body` in document order, per memory note)
- `.doc` → Word COM via a new `_extract_doc_via_word_com(path, logger)` helper that mirrors `ChatTab._extract_doc_text`: `Dispatch("Word.Application")`, open with `ReadOnly=True, AddToRecentFiles=False, ConfirmConversions=False`, **never** call `word.Quit()`, **never** set `word.Visible`, only close the `Document` instance the helper opened.

Per-doc cap: 100,000 characters. Anything longer is truncated with the suffix `[...truncated at 100000 chars]`.

Failures (missing file, extraction error, `.doc` with no Word installed, empty extracted text) log a warning and skip the doc. The run continues.

### New helper: `_resolve_bias_directive(cfg) -> str`

Maps `cfg.bias` to the directive string injected into the prompt's `{bias_directive}` slot:

| `bias` value | `bias_directive` string |
|---|---|
| `"neutral"` | `"Maintain a neutral, balanced tone. Include both favorable and unfavorable testimony equally."` |
| `"pro_plaintiff"` | `"Emphasize testimony most favorable to the plaintiff's case. Include testimony favorable to the defense only when it directly contradicts the plaintiff's claims."` |
| `"pro_defense"` | `"Emphasize testimony most favorable to the defense. Include testimony favorable to the plaintiff only when it directly bears on the defense's case."` |
| `"custom"` | The user's `cfg.bias_custom` text verbatim |

### Updated `_build_topic_locked_prompt`

Two new keyword args (`bias_directive`, `context_documents`) and two new substitutions:

```python
def _build_topic_locked_prompt(base_prompt, *, topic_list, bullets_per_topic,
                                deponent_label, custom_rules,
                                bias_directive, context_documents):
    def _strip_braces(s): return (s or "").replace("{", "").replace("}", "")

    rendered_topics = "\n".join(f"- {t}" for t in topic_list)

    if context_documents:
        ctx_blocks = []
        for doc in context_documents:
            ctx_blocks.append(
                f"=== CONTEXT DOC: {doc['filename']} ===\n{doc['text']}"
            )
        context_section = (
            "\n\nADDITIONAL CASE CONTEXT (read these in addition to the deposition transcript "
            "to better inform your summary):\n\n" + "\n\n".join(ctx_blocks)
        )
    else:
        context_section = ""

    return (base_prompt
            .replace("{deponent_label}", _strip_braces(deponent_label))
            .replace("{bullets_per_topic}", str(bullets_per_topic))
            .replace("{topic_list}", rendered_topics)
            .replace("{bias_directive}", _strip_braces(bias_directive))
            .replace("{context_section}", context_section)
            .replace("{custom_rules}", _strip_braces(custom_rules) or "(none)"))
```

`_strip_braces` is applied to all user-controlled strings (existing pattern from the prior fix) to prevent placeholder-leak across slots.

### `process_summary` orchestration

After loading session JSON, before building the prompt:

```python
context_docs = _extract_context_documents(cfg.get("context_doc_paths", []), logger)
bias_directive = _resolve_bias_directive(cfg)
```

Progress markers: context-doc extraction happens between the existing `5%` (load session) and `10%` (read cached transcript). If there are docs, log per-doc progress at intermediate percentages (e.g., `7%: "Extracting context: complaint.pdf (1/3)..."`). If none, jump straight from 5 to 10.

## Prompt template update

`Scripts/SUMMARIZE_DEPOSITION_PROMPT.txt` gets two new placeholders (`{bias_directive}`, `{context_section}`) and two new rules (8 — apply bias; 9 — context docs inform selection but aren't the source of testimony). Full updated template:

```
You are summarizing a deposition transcript using a fixed set of topic headings supplied by the attorney. You do not choose the topics. You do not add or omit topics. You write bullet points under the headings provided.

DEPONENT LABEL: {deponent_label}
BULLETS PER TOPIC: {bullets_per_topic}

SUMMARY BIAS / EDITORIAL LENS:
{bias_directive}
{context_section}

TOPIC HEADINGS (use exactly these, in this order):
{topic_list}

OUTPUT RULES:
1. Start the summary with this sentence: "The parties took the deposition of [name of deponent] on [date of deposition]. Below are the most salient portions of [name of deponent]'s testimony:"
2. Under each topic heading, write exactly {bullets_per_topic} bullet points summarizing what the deponent testified regarding that topic. If the transcript contains less testimony on a topic than requested, write fewer bullets — never invent testimony.
3. Each topic heading appears on its own line, in bold (use **Heading** markdown), title-cased, with no numbering and no trailing punctuation.
4. Each bullet is at least two complete sentences. Use markdown dashes ("- ") at the start of each bullet.
5. Refer to the deponent as "{deponent_label}" throughout. Do not use the deponent's full name except in the opening sentence.
6. Summarize testimony directly. Do not use introductory clauses like "Regarding," "Concerning," "With respect to," or "When asked about." Write "Plaintiff testified he broke his leg," not "Regarding his injuries, Plaintiff testified he broke his leg."
7. Avoid repeating phrases that flag content as testimony. Write "Joe texted Plaintiff photos," not "Plaintiff testified that Joe texted her photos."
8. Apply the SUMMARY BIAS instruction above when selecting which testimony to emphasize, while never inventing or distorting what the deponent actually said.
9. When ADDITIONAL CASE CONTEXT is provided, use it to inform your selection and phrasing of testimony, especially to identify contradictions or admissions that bear on the case. The deposition transcript itself remains the only source for direct testimony content — do not attribute statements from context documents to the deponent.
10. Do not add introductory or concluding paragraphs beyond rule 1's opening sentence. Do not add an exhibits section, impeachment section, or any section not in the topic list.

CUSTOM RULES FROM THE ATTORNEY (apply in addition to the rules above):
{custom_rules}
```

## Testing

### Dialog tests (`tests/test_deposition/test_depo_summary_config_dialog.py`)

Appended to existing file:

- `test_dialog_topic_drag_reorder_changes_selected_topics_order` — programmatically reorder QListWidget items, accept, verify `selected_topics` matches new order.
- `test_dialog_bias_combo_defaults_to_neutral_and_writes_neutral_to_session`.
- `test_dialog_bias_custom_reveals_text_field_and_round_trips`.
- `test_dialog_bias_switching_away_from_custom_clears_custom_field`.
- `test_dialog_context_docs_accept_via_add_button` — exercise the add-file handler with two temp files.
- `test_dialog_context_docs_reject_unsupported_extensions` — `.txt` file is not added; status label shows rejection.
- `test_dialog_context_docs_remove_button_drops_path`.
- `test_dialog_context_docs_drop_via_mime` — simulate a `QMimeData.setUrls(...)` drop on the QListWidget.
- `test_dialog_context_docs_missing_at_accept_are_silently_dropped` — add a file, delete it, accept, verify list is empty.

### Phase 2 tests (`tests/test_deposition/test_summarize_deposition_phases.py`)

Appended to existing file:

- `test_phase2_resolves_bias_directive_for_each_preset` — parametrized over the four bias values.
- `test_phase2_concatenates_context_documents_into_prompt` — verify the rendered prompt contains the `=== CONTEXT DOC: …` blocks.
- `test_phase2_per_doc_char_cap_truncates_long_docs` — mock extraction to return 200K chars, verify truncation marker.
- `test_phase2_missing_context_doc_logged_and_skipped`.
- `test_phase2_no_context_docs_leaves_context_section_empty`.
- `test_phase2_doc_extension_unsupported_skipped`.

### Smoke test update (`tests/test_deposition/test_full_flow_smoke.py`)

Add one `context_doc_paths` entry to the user_config update step, verify the run still completes end-to-end.

### Manual test plan

1. Drag-reorder three topics, generate, confirm docx headings match dragged order.
2. Bias = Most favorable to defense, generate, eyeball bullets for defense-friendly emphasis.
3. Bias = Custom… with directive "Highlight any inconsistencies in the deponent's testimony about injuries", generate, confirm prompt picked up the direction.
4. Drag a real complaint PDF and a med summary DOCX into the context-docs panel, generate, confirm the summary mentions or cross-references context.
5. Drop a `.txt`, confirm rejection.
6. Click `×` to remove a doc before Accept.

## Files touched

**Modified:**

- `icharlotte_core/ui/depo_summary_config_dialog.py` — QListWidget for topics, bias row, context-docs panel, validation in `accept()`.
- `Scripts/summarize_deposition.py` — `_extract_context_documents`, `_extract_doc_via_word_com`, `_resolve_bias_directive`, updated `_build_topic_locked_prompt`, updated `process_summary` orchestration with new progress markers.
- `Scripts/SUMMARIZE_DEPOSITION_PROMPT.txt` — two new placeholders and rules.
- `tests/test_deposition/test_depo_summary_config_dialog.py` — 9 new tests.
- `tests/test_deposition/test_summarize_deposition_phases.py` — 6 new tests.
- `tests/test_deposition/test_full_flow_smoke.py` — minor addition.

**Unchanged:**

- `icharlotte_core/deposition/session_manager.py` — schema additions are backward-compatible (new optional keys in `user_config`).
- `icharlotte_core/ui/widgets.py`, `iCharlotte.py` — no UI plumbing changes; the dialog is the only place that touches the new fields.

## Open questions

None. All decisions resolved during brainstorming.
