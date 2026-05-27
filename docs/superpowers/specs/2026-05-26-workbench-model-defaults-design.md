# Workbench Model Defaults - Design

**Date:** 2026-05-26
**Status:** Approved in brainstorming; pending written-spec review

---

## Goal

Move default LLM model configuration out of the main window's **Settings > LLM Settings** dropdown and into the **Prompt Engineering Workbench**. The user should manage both prompt text and default model behavior from the Workbench, with no separate LLM Settings menu entry in the main app chrome.

The existing `LLMConfig` API and `config/llm_preferences.json` remain the source of truth.

---

## Current State

- `iCharlotte.py` imports `LLMSettingsDialog`, adds a **LLM Settings** action to the Settings menu, and opens the dialog through `open_settings_dialog()`.
- `PromptsDialog` in `icharlotte_core/ui/dialogs.py` already has a per-agent **Model Settings** panel for the currently selected Workbench agent.
- `LLMSettingsDialog` in `icharlotte_core/ui/dialogs.py` is still the only UI that exposes all default task profiles: `general`, `extraction`, `summary`, `cross_check`, `classification`, and `quick`.
- `LLMSettingsDialog` also exposes agent/function overrides grouped into document-processing agents, case agents, and UI functions.

---

## Selected Approach

Embed the existing LLM settings capabilities into `PromptsDialog` as a new Workbench tab, then remove the main-window menu entry.

This preserves all current model configuration behavior while moving the user-facing control surface to the Workbench. It avoids a new settings format and avoids changing runtime model lookup behavior.

---

## Workbench UI

Add a new top-level Workbench tab named **Model Defaults**.

The tab contains the existing LLM settings sections:

- API key/provider status header.
- **Agents & Functions** tab for agent/function-specific overrides.
- **Default Profiles** tab for task-type model sequences, retry counts, and timeout values.
- Existing save and reset behavior.

The existing selected-agent **Model Settings** panel stays in place as a convenience for the agent currently selected in the Workbench header. It continues to edit that agent's override. After saving from **Model Defaults**, the selected-agent panel refreshes so it reflects the latest inherited/default model sequence.

---

## Main Window

Remove the **LLM Settings** item from the Settings dropdown.

Remove the main-window method and import that only exist to launch `LLMSettingsDialog`. The Settings dropdown remains for the other existing options, including email monitoring and docket refresh controls.

The Workbench remains available through the existing **Prompts** button.

---

## Data Flow

No data migration is required.

All reads and writes continue through `LLMConfig`:

- Agent/function overrides use `update_agent_config()`.
- Default task profiles use `update_task_config()`.
- Reset behavior restores defaults by clearing/reloading the existing config file path.

Because `LLMConfig` is a singleton, the embedded settings tab must either share the same instance already used by `PromptsDialog` or refresh the Workbench's instance after reset/save operations.

---

## Implementation Shape

Prefer extracting the reusable settings UI from `LLMSettingsDialog` into a widget hosted by `PromptsDialog`.

Recommended structure:

- Add `LLMSettingsWidget(QWidget)` in `icharlotte_core/ui/dialogs.py`.
- Move the old dialog's setup, load, save, reset, and helper methods into that widget.
- Keep `LLMSettingsDialog` only as a thin wrapper if needed for compatibility, or remove main-window usage entirely while leaving the class harmless.
- Add the widget as the **Model Defaults** tab inside `PromptsDialog`.
- Emit or call a callback after save/reset so `PromptsDialog._load_agent_model_settings()` refreshes the selected-agent panel.

This keeps the behavior close to the existing implementation and limits the change to `iCharlotte.py`, `icharlotte_core/ui/dialogs.py`, and focused tests.

---

## Testing

Targeted automated checks:

- Instantiate `PromptsDialog` offscreen and verify it contains a **Model Defaults** tab.
- Verify the embedded settings widget can load task profile rows from `LLMConfig`.
- Verify saving a task profile calls through to `LLMConfig` and does not require the removed main-window menu action.
- Verify `iCharlotte.py` no longer imports or references `LLMSettingsDialog` from the main window.
- Run Python compile checks for `iCharlotte.py` and `icharlotte_core/ui/dialogs.py`.

Manual verification:

- Open iCharlotte.
- Confirm the main Settings dropdown no longer shows **LLM Settings**.
- Open **Prompts**.
- Confirm **Model Defaults** is available in the Workbench.
- Change a default profile, save, close/reopen Workbench, and confirm the saved value persists.

---

## Out of Scope

- Changing model fallback semantics.
- Changing `config/llm_preferences.json` schema.
- Redesigning the Workbench visual style beyond the necessary embedded tab.
- Changing prompt versioning, A/B testing, or prompt registry behavior.
