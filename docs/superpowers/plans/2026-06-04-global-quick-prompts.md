# Global Quick Prompts Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Make custom chat quick prompts (templates) persist globally — across all cases and sessions — instead of per-case.

**Architecture:** A dedicated `GlobalQuickPromptStore` persists custom prompts to a single JSON file (`<.gemini>/global_quick_prompts.json`). The "Manage Templates" dialog and the Chat tab's template dropdown read/write custom prompts through this store instead of the per-case `ChatPersistence`. Built-in prompts stay hardcoded. A one-time, non-destructive migration imports any existing per-case custom prompts.

**Tech Stack:** Python 3, PySide6, `unittest`, JSON.

**Spec:** `docs/superpowers/specs/2026-06-04-global-quick-prompts-design.md`

**Working directory:** This plan is executed inside the `feature/global-quick-prompts` worktree at
`C:\geminiterminal2\.claude\worktrees\global-quick-prompts`. All file paths below are relative to
that worktree root, and all `git`/`python` commands run from there.

**Verification env note:** The iCharlotte app runs from the MAIN checkout (`C:\geminiterminal2`),
not this worktree. Unit tests and the offscreen UI smoke test run in the worktree. The final
real-app check (Task 4) ports the changes into the main checkout and restarts the app.

---

### Task 1: `GlobalQuickPromptStore` + tests

**Files:**
- Create: `tests/test_global_quick_prompts.py`
- Create: `icharlotte_core/chat/global_prompts.py`

The store mirrors the existing `ChatPersistence` quick-prompt method names
(`get_quick_prompts` / `add_quick_prompt` / `update_quick_prompt` / `delete_quick_prompt`) so the UI
callers change minimally. It reads fresh on every call and writes atomically. Migration runs once on
construction, guarded by a flag in the JSON.

- [ ] **Step 1: Write the failing test**

Create `tests/test_global_quick_prompts.py`:

```python
"""Tests for global quick prompt storage (shared across cases/sessions)."""
import json
import os
import shutil
import tempfile
import unittest

from icharlotte_core.chat.global_prompts import GlobalQuickPromptStore
from icharlotte_core.chat.models import QuickPrompt


class TestGlobalQuickPromptStore(unittest.TestCase):

    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        self.path = os.path.join(self.tmp, "global_quick_prompts.json")
        self.case_dir = os.path.join(self.tmp, "case_data")
        os.makedirs(self.case_dir)

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)

    def _store(self, auto_migrate=False):
        return GlobalQuickPromptStore(path=self.path, case_data_dir=self.case_dir,
                                      auto_migrate=auto_migrate)

    # --- CRUD ---

    def test_add_and_get(self):
        store = self._store()
        tid = store.add_quick_prompt("My Template", "Do the thing", "Custom")
        prompts = store.get_quick_prompts()
        self.assertEqual(len(prompts), 1)
        self.assertEqual(prompts[0].id, tid)
        self.assertEqual(prompts[0].name, "My Template")
        self.assertEqual(prompts[0].prompt, "Do the thing")
        self.assertEqual(prompts[0].category, "Custom")
        self.assertFalse(prompts[0].is_builtin)

    def test_get_empty(self):
        self.assertEqual(self._store().get_quick_prompts(), [])

    def test_update_all_fields(self):
        store = self._store()
        tid = store.add_quick_prompt("Old", "old text", "Custom")
        store.update_quick_prompt(tid, name="New", prompt="new text", category="Summary")
        p = store.get_quick_prompts()[0]
        self.assertEqual(p.name, "New")
        self.assertEqual(p.prompt, "new text")
        self.assertEqual(p.category, "Summary")

    def test_update_partial_leaves_others(self):
        store = self._store()
        tid = store.add_quick_prompt("Name", "text", "Custom")
        store.update_quick_prompt(tid, name="Renamed")
        p = store.get_quick_prompts()[0]
        self.assertEqual(p.name, "Renamed")
        self.assertEqual(p.prompt, "text")
        self.assertEqual(p.category, "Custom")

    def test_delete(self):
        store = self._store()
        tid = store.add_quick_prompt("Temp", "text", "Custom")
        self.assertEqual(len(store.get_quick_prompts()), 1)
        store.delete_quick_prompt(tid)
        self.assertEqual(len(store.get_quick_prompts()), 0)

    # --- persistence / robustness ---

    def test_persists_across_instances(self):
        """A new store instance reads what a previous one wrote (simulates a restart)."""
        store1 = self._store()
        store1.add_quick_prompt("Persisted", "stays around", "Custom")
        store2 = GlobalQuickPromptStore(path=self.path, case_data_dir=self.case_dir,
                                        auto_migrate=False)
        names = [p.name for p in store2.get_quick_prompts()]
        self.assertIn("Persisted", names)

    def test_file_is_valid_json(self):
        store = self._store()
        store.add_quick_prompt("X", "y", "Custom")
        with open(self.path, "r", encoding="utf-8") as f:
            data = json.load(f)
        self.assertIn("quick_prompts", data)
        self.assertEqual(len(data["quick_prompts"]), 1)

    def test_corrupt_file_recovers_to_empty(self):
        with open(self.path, "w", encoding="utf-8") as f:
            f.write("{not valid json")
        store = GlobalQuickPromptStore(path=self.path, case_data_dir=self.case_dir,
                                       auto_migrate=False)
        self.assertEqual(store.get_quick_prompts(), [])

    # --- migration ---

    def _write_case_file(self, file_number, quick_prompts):
        path = os.path.join(self.case_dir, f"{file_number}_chat.json")
        with open(path, "w", encoding="utf-8") as f:
            json.dump({"version": "1.0", "quick_prompts": quick_prompts}, f)

    def test_migration_imports_custom_skips_builtin(self):
        self._write_case_file("1000.001", [
            QuickPrompt(name="Builtin", prompt="b", is_builtin=True).to_dict(),
            QuickPrompt(name="Custom A", prompt="a text", category="Summary").to_dict(),
        ])
        store = self._store(auto_migrate=True)
        names = {p.name for p in store.get_quick_prompts()}
        self.assertEqual(names, {"Custom A"})
        self.assertEqual(store.get_quick_prompts()[0].category, "Summary")

    def test_migration_dedupes_by_name(self):
        self._write_case_file("1000.001", [QuickPrompt(name="Dupe", prompt="first").to_dict()])
        self._write_case_file("1000.002", [QuickPrompt(name="Dupe", prompt="second").to_dict()])
        store = self._store(auto_migrate=True)
        dupes = [p for p in store.get_quick_prompts() if p.name == "Dupe"]
        self.assertEqual(len(dupes), 1)

    def test_migration_runs_once(self):
        self._write_case_file("1000.001", [QuickPrompt(name="Once", prompt="text").to_dict()])
        store1 = self._store(auto_migrate=True)
        self.assertEqual(len([p for p in store1.get_quick_prompts() if p.name == "Once"]), 1)
        # User deletes it; a later launch must NOT resurrect it.
        for p in store1.get_quick_prompts():
            if p.name == "Once":
                store1.delete_quick_prompt(p.id)
        store2 = self._store(auto_migrate=True)
        self.assertEqual(len([p for p in store2.get_quick_prompts() if p.name == "Once"]), 0)

    def test_migration_non_destructive(self):
        """Per-case files are not modified by migration."""
        self._write_case_file("1000.001", [QuickPrompt(name="KeepMe", prompt="text").to_dict()])
        case_path = os.path.join(self.case_dir, "1000.001_chat.json")
        with open(case_path, encoding="utf-8") as f:
            before = f.read()
        self._store(auto_migrate=True)
        with open(case_path, encoding="utf-8") as f:
            after = f.read()
        self.assertEqual(before, after)


if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `python -m pytest tests/test_global_quick_prompts.py -q`
Expected: collection/import error — `ModuleNotFoundError: No module named 'icharlotte_core.chat.global_prompts'`.

- [ ] **Step 3: Create the store module**

Create `icharlotte_core/chat/global_prompts.py`:

```python
"""
Global quick prompt (chat template) storage.

Custom chat quick prompts are shared across every case and session. They live in a single JSON
file outside the per-case ``case_data/`` folder. Built-in prompts (``BUILTIN_PROMPTS``) are NOT
stored here — callers add those separately.
"""
import os
import json
import tempfile
from typing import List, Optional

from ..config import GEMINI_DATA_DIR
from .models import QuickPrompt


# Global store lives one level ABOVE the per-case case_data/ folder so it is clearly app-global
# and is never picked up by the case_data/*_chat.json migration glob.
DEFAULT_GLOBAL_PROMPTS_PATH = os.path.join(
    os.path.dirname(GEMINI_DATA_DIR), "global_quick_prompts.json"
)


class GlobalQuickPromptStore:
    """Stores custom quick prompts in a single global JSON file shared by all cases."""

    VERSION = "1.0"

    def __init__(self, path: Optional[str] = None,
                 case_data_dir: Optional[str] = None,
                 auto_migrate: bool = True):
        self.path = path or DEFAULT_GLOBAL_PROMPTS_PATH
        self.case_data_dir = case_data_dir if case_data_dir is not None else GEMINI_DATA_DIR
        if auto_migrate:
            self._migrate_once()

    # --- internal helpers ---

    def _default_data(self) -> dict:
        return {"version": self.VERSION, "migrated_from_per_case": False, "quick_prompts": []}

    def _load(self) -> dict:
        """Read the store from disk. Returns a default structure if missing or corrupt."""
        if not os.path.exists(self.path):
            return self._default_data()
        try:
            with open(self.path, "r", encoding="utf-8") as f:
                data = json.load(f)
        except Exception:
            return self._default_data()
        if not isinstance(data, dict):
            return self._default_data()
        if not isinstance(data.get("quick_prompts"), list):
            data["quick_prompts"] = []
        return data

    def _save(self, data: dict):
        """Atomically write the store to disk (temp file + os.replace)."""
        directory = os.path.dirname(self.path) or "."
        if not os.path.exists(directory):
            os.makedirs(directory)
        fd, tmp_path = tempfile.mkstemp(dir=directory, suffix=".tmp")
        try:
            with os.fdopen(fd, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=2, ensure_ascii=False)
            os.replace(tmp_path, self.path)
        except Exception:
            try:
                os.remove(tmp_path)
            except OSError:
                pass
            raise

    # --- public API (mirrors ChatPersistence quick-prompt methods) ---

    def get_quick_prompts(self) -> List[QuickPrompt]:
        """Return all custom quick prompts (never includes built-ins)."""
        prompts = []
        for p in self._load().get("quick_prompts", []):
            qp = QuickPrompt.from_dict(p)
            qp.is_builtin = False
            prompts.append(qp)
        return prompts

    def add_quick_prompt(self, name: str, prompt: str, category: str = "Custom") -> str:
        """Add a custom quick prompt. Returns the new prompt's ID."""
        qp = QuickPrompt(name=name, prompt=prompt, category=category, is_builtin=False)
        data = self._load()
        data["quick_prompts"].append(qp.to_dict())
        self._save(data)
        return qp.id

    def update_quick_prompt(self, prompt_id: str, name: str = None,
                            prompt: str = None, category: str = None):
        """Update an existing custom quick prompt in place."""
        data = self._load()
        for p in data.get("quick_prompts", []):
            if p.get("id") == prompt_id:
                if name is not None:
                    p["name"] = name
                if prompt is not None:
                    p["prompt"] = prompt
                if category is not None:
                    p["category"] = category
                break
        self._save(data)

    def delete_quick_prompt(self, prompt_id: str):
        """Delete a custom quick prompt by ID."""
        data = self._load()
        data["quick_prompts"] = [
            p for p in data.get("quick_prompts", []) if p.get("id") != prompt_id
        ]
        self._save(data)

    # --- one-time migration ---

    def _migrate_once(self):
        """Non-destructive, idempotent import of per-case custom prompts."""
        data = self._load()
        if data.get("migrated_from_per_case"):
            return
        existing_names = {p.get("name", "") for p in data.get("quick_prompts", [])}
        migrated = 0
        if self.case_data_dir and os.path.isdir(self.case_data_dir):
            for filename in sorted(os.listdir(self.case_data_dir)):
                if not filename.endswith("_chat.json"):
                    continue
                try:
                    with open(os.path.join(self.case_data_dir, filename),
                              "r", encoding="utf-8") as f:
                        case_data = json.load(f)
                except Exception:
                    continue
                for qp in case_data.get("quick_prompts", []):
                    if qp.get("is_builtin", False):
                        continue
                    name = (qp.get("name") or "").strip()
                    prompt = (qp.get("prompt") or "").strip()
                    if not name or not prompt or name in existing_names:
                        continue
                    new_qp = QuickPrompt(name=name, prompt=prompt,
                                         category=qp.get("category", "Custom"),
                                         is_builtin=False)
                    data["quick_prompts"].append(new_qp.to_dict())
                    existing_names.add(name)
                    migrated += 1
        data["migrated_from_per_case"] = True
        self._save(data)
        if migrated:
            print(f"[GlobalQuickPromptStore] Migrated {migrated} per-case quick prompts to global storage")


_global_store: Optional[GlobalQuickPromptStore] = None


def get_global_quick_prompt_store() -> GlobalQuickPromptStore:
    """Return the shared global quick prompt store (singleton)."""
    global _global_store
    if _global_store is None:
        _global_store = GlobalQuickPromptStore()
    return _global_store
```

- [ ] **Step 4: Run the tests to verify they pass**

Run: `python -m pytest tests/test_global_quick_prompts.py -q`
Expected: all tests pass (13 passed).

- [ ] **Step 5: Commit**

```bash
git add tests/test_global_quick_prompts.py icharlotte_core/chat/global_prompts.py
git commit -m "feat(chat): add GlobalQuickPromptStore for cross-case quick prompts" -m "Co-Authored-By: Claude Opus 4.8 <noreply@anthropic.com>"
```

---

### Task 2: Point the "Manage Templates" dialog at the global store

**Files:**
- Modify: `icharlotte_core/ui/chat_dialogs.py` (`PromptTemplateDialog`)

`PromptTemplateDialog` currently does custom-template CRUD through the per-case `self.persistence`.
Switch it to `get_global_quick_prompt_store()`. Built-in prompts still come from `BUILTIN_PROMPTS`.
Remove the "No case loaded" guard so saving works with no case open.

- [ ] **Step 1: Add the store to the constructor**

In `icharlotte_core/ui/chat_dialogs.py`, replace:

```python
    def __init__(self, persistence, parent=None):
        super().__init__(parent)
        self.persistence = persistence
        self.setWindowTitle("Manage Quick Prompts")
```

with:

```python
    def __init__(self, persistence=None, parent=None):
        super().__init__(parent)
        self.persistence = persistence  # kept for backward compatibility; not used for templates
        from ..chat.global_prompts import get_global_quick_prompt_store
        self.store = get_global_quick_prompt_store()
        self.setWindowTitle("Manage Quick Prompts")
```

- [ ] **Step 2: Load custom prompts from the global store**

Replace the `load_prompts` method:

```python
    def load_prompts(self):
        """Load all prompts into the list."""
        self.prompt_list.clear()
        self.all_prompts = []

        # Add builtin prompts
        for prompt in BUILTIN_PROMPTS:
            self.all_prompts.append(prompt)

        # Add custom prompts from persistence
        if self.persistence:
            for prompt in self.persistence.get_quick_prompts():
                if not prompt.is_builtin:
                    self.all_prompts.append(prompt)

        self.filter_prompts(self.category_filter.currentText())
```

with:

```python
    def load_prompts(self):
        """Load all prompts into the list."""
        self.prompt_list.clear()
        self.all_prompts = []

        # Add builtin prompts
        for prompt in BUILTIN_PROMPTS:
            self.all_prompts.append(prompt)

        # Add custom prompts from the global store (shared across all cases)
        for prompt in self.store.get_quick_prompts():
            self.all_prompts.append(prompt)

        self.filter_prompts(self.category_filter.currentText())
```

- [ ] **Step 3: Save to the global store (and drop the "No case loaded" guard)**

Replace the `save_prompt` method:

```python
    def save_prompt(self):
        """Save the current prompt."""
        name = self.name_edit.text().strip()
        category = self.category_edit.currentText().strip()
        prompt_text = self.prompt_edit.toPlainText().strip()

        if not name:
            QMessageBox.warning(self, "Error", "Please enter a prompt name.")
            return

        if not prompt_text:
            QMessageBox.warning(self, "Error", "Please enter the prompt text.")
            return

        if not self.persistence:
            QMessageBox.warning(self, "Error", "No case loaded. Cannot save prompts.")
            return

        if self.current_prompt_id and not self.is_builtin:
            # Update existing
            self.persistence.update_quick_prompt(
                self.current_prompt_id,
                name=name,
                prompt=prompt_text,
                category=category
            )
        else:
            # Create new
            self.persistence.add_quick_prompt(name, prompt_text, category)

        self.load_prompts()
        QMessageBox.information(self, "Saved", "Prompt saved successfully.")
```

with:

```python
    def save_prompt(self):
        """Save the current prompt."""
        name = self.name_edit.text().strip()
        category = self.category_edit.currentText().strip()
        prompt_text = self.prompt_edit.toPlainText().strip()

        if not name:
            QMessageBox.warning(self, "Error", "Please enter a prompt name.")
            return

        if not prompt_text:
            QMessageBox.warning(self, "Error", "Please enter the prompt text.")
            return

        if self.current_prompt_id and not self.is_builtin:
            # Update existing
            self.store.update_quick_prompt(
                self.current_prompt_id,
                name=name,
                prompt=prompt_text,
                category=category
            )
        else:
            # Create new
            self.store.add_quick_prompt(name, prompt_text, category)

        self.load_prompts()
        QMessageBox.information(self, "Saved", "Prompt saved successfully.")
```

- [ ] **Step 4: Delete via the global store**

Replace the `delete_prompt` method:

```python
        if reply == QMessageBox.StandardButton.Yes:
            if self.persistence:
                self.persistence.delete_quick_prompt(self.current_prompt_id)
            self.load_prompts()
            self.clear_editor()
```

with:

```python
        if reply == QMessageBox.StandardButton.Yes:
            self.store.delete_quick_prompt(self.current_prompt_id)
            self.load_prompts()
            self.clear_editor()
```

- [ ] **Step 5: Verify syntax**

Run: `python -m py_compile icharlotte_core/ui/chat_dialogs.py && echo OK`
Expected: `OK`

- [ ] **Step 6: Offscreen UI smoke test**

Create `tests/smoke_global_prompts_dialog.py` (NOT named `test_*`, so pytest will not collect it and
trip the known PySide6 app-running collection issue):

```python
"""Offscreen smoke test: PromptTemplateDialog writes to the global store. Run directly:
   QT_QPA_PLATFORM=offscreen python tests/smoke_global_prompts_dialog.py
"""
import os
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
import sys
import tempfile

from PySide6.QtWidgets import QApplication, QMessageBox

app = QApplication.instance() or QApplication(sys.argv)

# Silence modal popups so the headless run does not block.
QMessageBox.information = staticmethod(lambda *a, **k: QMessageBox.StandardButton.Ok)
QMessageBox.warning = staticmethod(lambda *a, **k: QMessageBox.StandardButton.Ok)
QMessageBox.question = staticmethod(lambda *a, **k: QMessageBox.StandardButton.Yes)

# Point the global singleton at a temp store.
import icharlotte_core.chat.global_prompts as gp
tmp = tempfile.mkdtemp()
gp._global_store = gp.GlobalQuickPromptStore(
    path=os.path.join(tmp, "g.json"),
    case_data_dir=os.path.join(tmp, "case_data"),
    auto_migrate=False,
)

from icharlotte_core.ui.chat_dialogs import PromptTemplateDialog

dlg = PromptTemplateDialog(persistence=None)
dlg.add_prompt()
dlg.name_edit.setText("Smoke Template")
dlg.prompt_edit.setPlainText("Smoke prompt body")
dlg.category_edit.setCurrentText("Custom")
dlg.save_prompt()

stored = gp.get_global_quick_prompt_store().get_quick_prompts()
assert any(p.name == "Smoke Template" and p.prompt == "Smoke prompt body" for p in stored), stored

# Survives a fresh store instance (cross-session).
fresh = gp.GlobalQuickPromptStore(path=gp._global_store.path,
                                  case_data_dir=gp._global_store.case_data_dir,
                                  auto_migrate=False)
assert any(p.name == "Smoke Template" for p in fresh.get_quick_prompts())
print("UI SMOKE OK")
```

Run: `python tests/smoke_global_prompts_dialog.py`
Expected: prints `UI SMOKE OK` (no assertion error, no hang).

- [ ] **Step 7: Commit**

```bash
git add icharlotte_core/ui/chat_dialogs.py tests/smoke_global_prompts_dialog.py
git commit -m "feat(chat-dialogs): manage quick prompts via global store" -m "Co-Authored-By: Claude Opus 4.8 <noreply@anthropic.com>"
```

---

### Task 3: Point the Chat tab template menu at the global store

**Files:**
- Modify: `icharlotte_core/ui/tabs.py` (`ChatTab.update_template_menu`)

The dropdown currently reads custom prompts from `self.persistence`. Switch to the global store, and
make the menu rebuild on `aboutToShow` so a template added anywhere is immediately available in every
open chat tab without a restart. The menu becomes persistent (created once, repopulated in place).

- [ ] **Step 1: Rewrite `update_template_menu`**

In `icharlotte_core/ui/tabs.py`, replace the entire `update_template_menu` method:

```python
    def update_template_menu(self):
        """Update the quick prompts template menu."""
        menu = QMenu(self)

        # Built-in prompts
        for prompt in BUILTIN_PROMPTS:
            if prompt.id == 'builtin_mediation_brief':
                continue  # Handled separately below
            action = QAction(prompt.name, self)
            action.triggered.connect(lambda checked, p=prompt: self.insert_template(p.prompt))
            menu.addAction(action)

        # Mediation Brief (special — triggers generation, not text insert)
        menu.addSeparator()
        med_brief_action = QAction("Mediation Brief", self)
        med_brief_action.triggered.connect(self._on_mediation_brief_selected)
        menu.addAction(med_brief_action)

        # Custom prompts (if persistence available)
        if self.persistence:
            prompts = self.persistence.get_quick_prompts()
            custom_prompts = [p for p in prompts if not p.is_builtin]
            if custom_prompts:
                menu.addSeparator()
                for prompt in custom_prompts:
                    action = QAction(prompt.name, self)
                    action.triggered.connect(lambda checked, p=prompt: self.insert_template(p.prompt))
                    menu.addAction(action)

        menu.addSeparator()
        manage_action = QAction("Manage Templates...", self)
        manage_action.triggered.connect(self.open_template_manager)
        menu.addAction(manage_action)

        self.template_btn.setMenu(menu)
```

with:

```python
    def update_template_menu(self):
        """(Re)populate the quick prompts template menu from built-ins + the global store."""
        # Reuse one persistent menu so we can refresh it on aboutToShow without re-wiring.
        menu = self.template_btn.menu()
        if menu is None:
            menu = QMenu(self)
            self.template_btn.setMenu(menu)
            # Rebuild every time the menu opens so newly added/edited/deleted global
            # templates appear in all chat tabs without needing a restart.
            menu.aboutToShow.connect(self.update_template_menu)
        menu.clear()

        # Built-in prompts
        for prompt in BUILTIN_PROMPTS:
            if prompt.id == 'builtin_mediation_brief':
                continue  # Handled separately below
            action = QAction(prompt.name, self)
            action.triggered.connect(lambda checked, p=prompt: self.insert_template(p.prompt))
            menu.addAction(action)

        # Mediation Brief (special — triggers generation, not text insert)
        menu.addSeparator()
        med_brief_action = QAction("Mediation Brief", self)
        med_brief_action.triggered.connect(self._on_mediation_brief_selected)
        menu.addAction(med_brief_action)

        # Custom prompts from the global store (shared across all cases/sessions)
        try:
            from ..chat.global_prompts import get_global_quick_prompt_store
            custom_prompts = get_global_quick_prompt_store().get_quick_prompts()
        except Exception as e:
            custom_prompts = []
            print(f"[ChatTab] Could not load global quick prompts: {e}")
        if custom_prompts:
            menu.addSeparator()
            for prompt in custom_prompts:
                action = QAction(prompt.name, self)
                action.triggered.connect(lambda checked, p=prompt: self.insert_template(p.prompt))
                menu.addAction(action)

        menu.addSeparator()
        manage_action = QAction("Manage Templates...", self)
        manage_action.triggered.connect(self.open_template_manager)
        menu.addAction(manage_action)
```

Note: `open_template_manager` is unchanged — it may keep passing `self.persistence` to
`PromptTemplateDialog` (the dialog ignores it for templates) and still calls
`self.update_template_menu()` after the dialog closes.

- [ ] **Step 2: Verify syntax**

Run: `python -m py_compile icharlotte_core/ui/tabs.py && echo OK`
Expected: `OK`

- [ ] **Step 3: Commit**

```bash
git add icharlotte_core/ui/tabs.py
git commit -m "feat(chat-tab): load quick prompts from global store, refresh on open" -m "Co-Authored-By: Claude Opus 4.8 <noreply@anthropic.com>"
```

---

### Task 4: End-to-end verification + deploy to the main checkout

**Files:** none (verification + deployment only)

- [ ] **Step 1: Run the full new test + the existing chat tests**

Run: `python -m pytest tests/test_global_quick_prompts.py tests/test_chat/test_persistence.py -q`
Expected: all pass (the per-case `ChatPersistence` behavior is untouched, so its tests still pass).

- [ ] **Step 2: Re-run the offscreen UI smoke test**

Run: `python tests/smoke_global_prompts_dialog.py`
Expected: `UI SMOKE OK`

- [ ] **Step 3: Deploy to the main checkout so the running app picks it up**

The app runs from `C:\geminiterminal2`. Copy the new files and re-apply the two edited methods there.

```bash
# New file (conflict-free): copy into the main checkout
cp icharlotte_core/chat/global_prompts.py C:/geminiterminal2/icharlotte_core/chat/global_prompts.py
```

`chat_dialogs.py` is identical between this worktree's base and the main checkout (no uncommitted
changes there), so it can be copied directly:

```bash
cp icharlotte_core/ui/chat_dialogs.py C:/geminiterminal2/icharlotte_core/ui/chat_dialogs.py
```

`tabs.py` HAS uncommitted generate-motion changes in the main checkout, so DO NOT copy it. Instead,
in `C:\geminiterminal2\icharlotte_core\ui\tabs.py`, apply the SAME `update_template_menu` replacement
from Task 3, Step 1 (the target methods are untouched by the generate-motion work, so the edit
applies cleanly).

Then verify both edited files compile in the main checkout:

Run: `python -c "import py_compile; py_compile.compile(r'C:/geminiterminal2/icharlotte_core/ui/tabs.py', doraise=True); py_compile.compile(r'C:/geminiterminal2/icharlotte_core/ui/chat_dialogs.py', doraise=True); print('OK')"`
Expected: `OK`

- [ ] **Step 4: Manual real-app verification**

Launch iCharlotte from the main checkout and confirm cross-case persistence:

1. `cd C:/geminiterminal2 && python iCharlotte.py` (or restart the running instance).
2. Open a case, go to the Chat tab → Templates ▾ → Manage Templates.
3. Add a template (name + body), Save, Close.
4. Confirm it appears in the Templates ▾ dropdown.
5. Switch to a DIFFERENT case, open the Chat tab → Templates ▾.
6. Confirm the same template appears there too.
7. (Optional) Confirm `C:\geminiterminal2\.gemini\global_quick_prompts.json` exists and contains the
   template.

Expected: the template added in case A is visible in case B and survives an app restart.

- [ ] **Step 5: Final commit (if any deploy-only adjustments were tracked)**

No source changes beyond Tasks 1–3 should be needed. If `tabs.py`/`chat_dialogs.py` in the worktree
match what was deployed, there is nothing further to commit here.

---

## Self-Review

**Spec coverage:**
- Dedicated global store + JSON location → Task 1 (`GlobalQuickPromptStore`, `DEFAULT_GLOBAL_PROMPTS_PATH`).
- Public API mirroring `ChatPersistence` → Task 1.
- Fresh-read + atomic-write behavior → Task 1 (`_load`/`_save`, `test_file_is_valid_json`, `test_corrupt_file_recovers_to_empty`).
- One-time, non-destructive, idempotent migration → Task 1 (`_migrate_once`, 4 migration tests).
- Dialog uses the store + "No case loaded" guard removed → Task 2.
- Menu uses the store + `aboutToShow` refresh → Task 3.
- `ChatPersistence` per-case methods left intact → not modified by any task (verified by running its tests in Task 4, Step 1).
- Pure-Python unit tests + offscreen UI smoke + real-app check → Tasks 1, 2, 4.

**Placeholder scan:** No TBD/TODO/"handle edge cases"/"similar to Task N". All code blocks are complete.

**Type/name consistency:** `GlobalQuickPromptStore`, `get_global_quick_prompt_store`,
`get_quick_prompts`/`add_quick_prompt`/`update_quick_prompt`/`delete_quick_prompt`,
`_global_store`, and `self.store` are used identically across Tasks 1–3. The store's
`get_quick_prompts()` returns `QuickPrompt` objects, which is what both the dialog (`prompt.name`,
`prompt.category`, `prompt.prompt`, `prompt.id`, `prompt.is_builtin`) and the menu (`prompt.name`,
`prompt.prompt`) consume.
