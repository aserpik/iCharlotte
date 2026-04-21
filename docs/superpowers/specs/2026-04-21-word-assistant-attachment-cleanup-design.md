# Word AI Assistant — Attachment Cleanup

## Problem

The AI Assistant popup in Word (invoked via Win+V) lets the user drag/drop
documents into an attachment area to supply context to the LLM. Two issues:

1. **Attachments persist across sessions.** `WordLLMPopup` is a singleton
   — constructed once and reused on every Win+V press. The reuse path calls
   `_reset_for_new_session()` (`icharlotte_core/word_hotkey.py:3419`), which
   clears every other piece of transient state (custom input, case detection,
   worker threads, redline flags, etc.) but forgets to clear the attachment
   area. Files attached in one invocation silently leak into every subsequent
   invocation until the host app is restarted.
2. **The per-chip remove button is hard to see.** Each attachment chip
   already has a `×` button (`word_hotkey.py:2284-2292`), but it is 16×16,
   transparent, and coloured `#a6adc8` against the chip's `#313244` ground —
   it blends into the chip and users don't notice it.

## Changes

### Fix 1 — clear attachments on session reset

In `_reset_for_new_session()` add one line inside the existing "Reset UI
to clean state" block:

```python
self.attachment_area.clear()
```

`AttachmentArea.clear()` is already defined (`word_hotkey.py:2367`) and
properly removes every chip and empties the `_attachments` dict. No new
method is needed.

### Fix 2 — make the remove button visible

Restyle the `remove_btn` `QPushButton` inside `AttachmentArea._add_chip`
(`word_hotkey.py:2284`). New spec:

| Property       | Current                          | New                                             |
| -------------- | -------------------------------- | ----------------------------------------------- |
| Size           | 16×16                            | 18×18                                           |
| Background     | transparent                      | `#45475a` with `border-radius: 9px` (circle)    |
| Foreground     | `#a6adc8`                        | `#cdd6f4`                                       |
| Hover bg       | (unchanged)                      | `#f38ba8`                                       |
| Hover fg       | `#f38ba8`                        | `#1e1e2e`                                       |
| Border         | none (implicit)                  | `border: none` (explicit)                       |
| Tooltip        | none                             | "Remove file"                                   |

The new hover state is a solid red pill with dark text — visually
unambiguous. Resting state is a grey pill that clearly reads as a
clickable control.

## Scope

- Single file touched: `icharlotte_core/word_hotkey.py`
- No data-model changes, no new imports, no persistence-layer changes.
- No tests required — the change is a lifecycle fix (line added to a
  reset method) plus pure stylesheet text. Manual verification: attach a
  file, dismiss popup, re-open via Win+V, confirm chip is gone.

## Non-goals

- No "clear all" button (single `×` per chip is enough once visible).
- No persistence of attachments across sessions (explicitly the opposite
  of what the user wants).
- No changes to the underlying extraction pipeline.
