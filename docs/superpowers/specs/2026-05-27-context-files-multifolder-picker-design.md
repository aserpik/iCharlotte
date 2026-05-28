# Multi-folder context file picker for wizard tasks

**Date:** 2026-05-27
**Status:** Approved (design)

## Problem

When a wizard task asks the user to select context files, the user must be able
to select multiple files from **different folders / subfolders** in one session.
Two wizard tasks currently fall short because they use a single
`QFileDialog.getOpenFileNames` call, which only allows selecting multiple files
**within a single folder**:

- Oppose Motion — context files chosen in `build_oppose_motion_tab`
  (`icharlotte_core/ui/wizard/pages/oppose_motion_page.py`).
- Respond to Discovery — context files chosen in `_on_select_context_files`
  (`icharlotte_core/ui/wizard/pages/respond_discovery_page.py`).

### Already correct (no work)

- Generic `SettingsPage` base (covers all generic wizard tasks) already uses an
  additive "Add Files…" + `QListWidget` pattern that accumulates across folders.
  (`icharlotte_core/ui/wizard/pages/settings_page.py`)
- Depo Prep already uses the additive "+ Add files" pattern for both its
  deponent-materials and case-context lists.
  (`icharlotte_core/ui/wizard/pages/depo_prep_settings_page.py`)

### Intentionally out of scope

- Med Chron is deliberately one-file-per-tab: picking a new file *replaces* the
  current selection and may restart Phase 1. It is not a multi-context-file task
  and is left unchanged.
  (`icharlotte_core/ui/wizard/pages/med_chron_settings_page.py`)

## Solution

Introduce a reusable modal dialog that lets the user accumulate files across
multiple folder visits, then drop it into the two single-folder call sites.

### New component: `ContextFilesDialog`

Location: `icharlotte_core/ui/context_files_dialog.py`

UI:
- `QListWidget` in extended-selection mode. Each row shows the file basename
  with a dimmed parent-folder hint (e.g. `Smith Depo.pdf — DISCOVERY/Depos`).
  Full absolute path is stored in the item's `UserRole` and shown as the tooltip.
- **"Add files…"** button — opens `QFileDialog.getOpenFileNames` and appends the
  result to the list, de-duplicating by normalized absolute path.
- **"Remove selected"** button — removes the selected rows.
- **OK / Cancel** via `QDialogButtonBox`.

Behavior:
- The start folder for the first "Add files…" is the caller-supplied `start_dir`.
- After each "Add", the next "Add" defaults to the folder of the most recently
  added file, so hopping between subfolders is quick.
- De-duplication is by `os.path.normcase(os.path.abspath(path))`.

API (mirrors `getOpenFileNames` ergonomics):

```python
class ContextFilesDialog(QDialog):
    def __init__(
        self,
        parent=None,
        *,
        title: str = "Select context files",
        start_dir: str = "",
        file_filter: str = "All files (*.*)",
        initial: list[str] | None = None,
    ) -> None: ...

    def selected_files(self) -> list[str]:
        """Return accumulated full paths in display order."""

    @classmethod
    def get_files(
        cls,
        parent=None,
        *,
        title: str = "Select context files",
        start_dir: str = "",
        file_filter: str = "All files (*.*)",
        initial: list[str] | None = None,
    ) -> list[str] | None:
        """Show the dialog modally.

        Returns the accumulated list on OK (possibly empty), or None on Cancel,
        so callers can distinguish "cancelled" from "OK with empty list".
        """
```

### Integration

**Oppose Motion** (`build_oppose_motion_tab`):
Replace the `getOpenFileNames` block with:

```python
context_files = ContextFilesDialog.get_files(
    parent,
    title="Select context document(s)",
    start_dir=os.path.dirname(motion_file) or case_path,
    file_filter="Context files (*.pdf *.docx *.txt *.msg);;All files (*.*)",
)
context_files = [
    p for p in (context_files or []) if is_supported_context_file(p)
]
```

Cancel (`None`) → `or []` → the tab still builds with no context (context is
optional, matching today's behavior where a cancelled dialog returned `[]`).

**Respond to Discovery** (`_on_select_context_files`):

```python
def _on_select_context_files(self) -> None:
    paths = ContextFilesDialog.get_files(
        self,
        title="Select context file(s)",
        start_dir=context_file_start_dir(self.discovery_file, self.case_root),
        file_filter="Context files (*.pdf *.docx *.txt);;All files (*.*)",
    )
    if paths is None:
        return  # user cancelled — stay on the rules screen, do not generate
    self.context_files = list(paths)
    self._generate_proposals()
```

This is a small, intentional behavior improvement: today a cancel still launches
an expensive empty-context proposal generation; with the new dialog, cancel
backs out cleanly and the user can click "Next: Context Files" again.

The existing extension-filtering (`is_supported_context_file`) and start-dir
helpers at each site are preserved exactly.

## Testing

`pytest-qt` unit tests for `ContextFilesDialog`
(`tests/test_ui/test_context_files_dialog.py` or nearest existing UI test home):

- Monkeypatch `QFileDialog.getOpenFileNames` to return files from **different
  folders** on successive "Add files…" invocations; assert all accumulate.
- Assert de-duplication when the same path is added twice (including
  case-insensitive / path-normalization equivalence on Windows).
- Assert "Remove selected" removes the right rows.
- Assert `get_files` returns `None` when the dialog is rejected and a `list`
  (possibly empty) when accepted.
- Assert `selected_files()` returns full absolute paths in display order.

Integration tests (extend existing):

- `tests/test_wizard/test_oppose_motion_page.py`: monkeypatch
  `ContextFilesDialog.get_files` to return a multi-folder list; assert the built
  tab's `context_files` is populated correctly, and that a `None` return yields
  an empty context list without error.
- `tests/test_wizard/test_respond_discovery_page.py`: monkeypatch
  `ContextFilesDialog.get_files`; assert `context_files` is set and
  `_generate_proposals` runs on OK, and that a `None` return aborts without
  calling `_generate_proposals`.

## Files touched

- `icharlotte_core/ui/context_files_dialog.py` (new)
- `icharlotte_core/ui/wizard/pages/oppose_motion_page.py` (1 call site)
- `icharlotte_core/ui/wizard/pages/respond_discovery_page.py` (1 call site)
- tests as above
