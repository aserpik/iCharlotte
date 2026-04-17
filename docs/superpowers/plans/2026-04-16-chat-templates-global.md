# Chat Templates → Global via Prompt Workbench — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Move custom chat templates from per-case storage to global PromptManager, making them editable in the Prompt Engineering Workbench and shared across all cases.

**Architecture:** Custom templates get stored as PromptManager passes under agent `chat_templates`. The ChatTab template menu and PromptTemplateDialog read/write from PromptManager instead of ChatPersistence. Built-in prompts remain hardcoded and read-only.

**Tech Stack:** Python, PromptManager (existing), ChatPersistence, PyQt6

---

### Task 1: Add global template CRUD to PromptManager

**Files:**
- Modify: `icharlotte_core/prompt_manager.py`
- Test: `tests/test_chat_template_global.py` (create)

Custom templates need to be stored globally. PromptManager already handles versioned prompts per agent:pass. We'll use agent `chat_templates` with the template name as the pass. But templates also have a `category` field that PromptManager doesn't track. We'll store category in the version description field (lightweight, no schema changes needed) and add helper methods.

- [ ] **Step 1: Write the failing test**

Create `tests/test_chat_template_global.py`:

```python
"""Tests for global chat template storage in PromptManager."""
import os
import shutil
import tempfile
import unittest

from icharlotte_core.prompt_manager import PromptManager


class TestGlobalChatTemplates(unittest.TestCase):

    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        self.pm = PromptManager(prompts_dir=self.tmp)

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)

    def test_add_chat_template(self):
        self.pm.add_chat_template("My Template", "Do something useful", "Custom")
        result = self.pm.get_chat_templates()
        self.assertEqual(len(result), 1)
        self.assertEqual(result[0]["name"], "My Template")
        self.assertEqual(result[0]["prompt"], "Do something useful")
        self.assertEqual(result[0]["category"], "Custom")
        self.assertTrue(len(result[0]["id"]) > 0)

    def test_add_multiple_templates(self):
        self.pm.add_chat_template("Template A", "Prompt A", "Summary")
        self.pm.add_chat_template("Template B", "Prompt B", "Analysis")
        result = self.pm.get_chat_templates()
        self.assertEqual(len(result), 2)
        names = {t["name"] for t in result}
        self.assertEqual(names, {"Template A", "Template B"})

    def test_update_chat_template(self):
        tid = self.pm.add_chat_template("Old Name", "Old prompt", "Custom")
        self.pm.update_chat_template(tid, name="New Name", prompt="New prompt", category="Analysis")
        result = self.pm.get_chat_templates()
        self.assertEqual(len(result), 1)
        self.assertEqual(result[0]["name"], "New Name")
        self.assertEqual(result[0]["prompt"], "New prompt")
        self.assertEqual(result[0]["category"], "Analysis")

    def test_delete_chat_template(self):
        tid = self.pm.add_chat_template("Temp", "Will delete", "Custom")
        self.assertEqual(len(self.pm.get_chat_templates()), 1)
        self.pm.delete_chat_template(tid)
        self.assertEqual(len(self.pm.get_chat_templates()), 0)

    def test_get_templates_returns_empty_when_none(self):
        result = self.pm.get_chat_templates()
        self.assertEqual(result, [])

    def test_template_id_is_stable(self):
        tid = self.pm.add_chat_template("Stable", "Test", "Custom")
        result = self.pm.get_chat_templates()
        self.assertEqual(result[0]["id"], tid)

    def test_template_accessible_via_get_prompt(self):
        """Templates should also be loadable via standard get_prompt for Workbench."""
        tid = self.pm.add_chat_template("WB Test", "Workbench prompt text", "Custom")
        # The pass_name is the template ID
        text = self.pm.get_prompt("chat_templates", tid)
        self.assertEqual(text, "Workbench prompt text")


if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_chat_template_global.py -v`
Expected: FAIL with `AttributeError: 'PromptManager' object has no attribute 'add_chat_template'`

- [ ] **Step 3: Add chat_templates to _ensure_directory_structure**

In `icharlotte_core/prompt_manager.py`, find the agent list in `_ensure_directory_structure` (around line 102). Add `'chat_templates'` to the list:

```python
        for agent in ['summarize', 'discovery', 'deposition', 'timeline', 'contradiction',
                      'word_assistant', 'legal_research', 'mediation_brief', 'chat_templates']:
```

- [ ] **Step 4: Add chat template CRUD methods**

Add these methods to the `PromptManager` class, after `seed_pipeline_prompts` (around line 490):

```python
    # --- Chat Templates (global, shared across cases) ---

    def get_chat_templates(self) -> list:
        """Return all global chat templates as a list of dicts.

        Each dict has keys: id, name, prompt, category.
        """
        registry = self._registry.get("prompts", {})
        templates = []
        for key, entry in registry.items():
            if not key.startswith("chat_templates:"):
                continue
            tid = key.split(":", 1)[1]
            # Read the current prompt text from disk
            text = self.get_prompt("chat_templates", tid)
            if text is None:
                continue
            # Category is stored in the latest version's description field
            # Format: "category:<Category>"
            category = "Custom"
            for v in reversed(entry.get("versions", [])):
                desc = v.get("description", "")
                if desc.startswith("category:"):
                    category = desc[len("category:"):]
                    break
            # Name is stored in the entry's pass_name, but we store the
            # display name in a special way: the first version's author field.
            name = tid  # fallback
            for v in entry.get("versions", []):
                author = v.get("author", "")
                if author.startswith("name:"):
                    name = author[len("name:"):]
                    break
            templates.append({
                "id": tid,
                "name": name,
                "prompt": text,
                "category": category,
            })
        return templates

    def add_chat_template(self, name: str, prompt: str, category: str = "Custom") -> str:
        """Add a new global chat template. Returns the template ID."""
        import uuid as _uuid
        tid = str(_uuid.uuid4())[:12]
        self.create_version(
            "chat_templates", tid, prompt,
            version="v1",
            description=f"category:{category}",
            author=f"name:{name}",
            set_as_current=True,
        )
        return tid

    def update_chat_template(self, template_id: str, name: str = None,
                             prompt: str = None, category: str = None):
        """Update an existing global chat template."""
        key = self._get_prompt_key("chat_templates", template_id)
        entry = self._registry.get("prompts", {}).get(key)
        if not entry:
            return

        # Resolve current values for unchanged fields
        current_name = template_id
        current_category = "Custom"
        for v in entry.get("versions", []):
            author = v.get("author", "")
            if author.startswith("name:"):
                current_name = author[len("name:"):]
                break
        for v in reversed(entry.get("versions", [])):
            desc = v.get("description", "")
            if desc.startswith("category:"):
                current_category = desc[len("category:"):]
                break

        final_name = name if name is not None else current_name
        final_category = category if category is not None else current_category
        final_prompt = prompt if prompt is not None else (self.get_prompt("chat_templates", template_id) or "")

        self.create_version(
            "chat_templates", template_id, final_prompt,
            description=f"category:{final_category}",
            author=f"name:{final_name}",
            set_as_current=True,
        )

    def delete_chat_template(self, template_id: str):
        """Delete a global chat template."""
        key = self._get_prompt_key("chat_templates", template_id)
        if key in self._registry.get("prompts", {}):
            del self._registry["prompts"][key]
            self._save_registry()
        # Remove files from disk
        import glob
        pattern = os.path.join(self.prompts_dir, "chat_templates", f"{template_id}_*")
        for f in glob.glob(pattern):
            try:
                os.remove(f)
            except OSError:
                pass
```

- [ ] **Step 5: Run tests**

Run: `python -m pytest tests/test_chat_template_global.py -v`
Expected: All 7 tests pass

- [ ] **Step 6: Commit**

```bash
git add icharlotte_core/prompt_manager.py tests/test_chat_template_global.py
git commit -m "feat(prompt-manager): add global chat template CRUD methods"
```

---

### Task 2: Migrate PromptTemplateDialog to use PromptManager

**Files:**
- Modify: `icharlotte_core/ui/chat_dialogs.py:15-239`

The dialog currently takes a `ChatPersistence` instance and reads/writes per-case. Change it to use PromptManager for custom templates while keeping built-in prompts from the `BUILTIN_PROMPTS` constant.

- [ ] **Step 1: Update constructor**

In `icharlotte_core/ui/chat_dialogs.py`, change the `__init__` method of `PromptTemplateDialog` (line 18) from:

```python
    def __init__(self, persistence, parent=None):
        super().__init__(parent)
        self.persistence = persistence
```

to:

```python
    def __init__(self, persistence=None, parent=None):
        super().__init__(parent)
        self.persistence = persistence  # Kept for backward compat but not used for templates
        from icharlotte_core.prompt_manager import get_prompt_manager
        self.pm = get_prompt_manager()
```

- [ ] **Step 2: Update load_prompts to read from PromptManager**

Replace the `load_prompts` method (lines 106-121) with:

```python
    def load_prompts(self):
        """Load all prompts into the list."""
        self.prompt_list.clear()
        self.all_prompts = []

        # Add builtin prompts
        for prompt in BUILTIN_PROMPTS:
            self.all_prompts.append(prompt)

        # Add custom prompts from global PromptManager
        for t in self.pm.get_chat_templates():
            self.all_prompts.append(QuickPrompt(
                id=t["id"],
                name=t["name"],
                prompt=t["prompt"],
                category=t["category"],
                is_builtin=False,
            ))

        self.filter_prompts(self.category_filter.currentText())
```

- [ ] **Step 3: Update save_prompt to write to PromptManager**

Replace the `save_prompt` method (lines 190-221) with:

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
            self.pm.update_chat_template(
                self.current_prompt_id,
                name=name,
                prompt=prompt_text,
                category=category,
            )
        else:
            # Create new
            self.pm.add_chat_template(name, prompt_text, category)

        self.load_prompts()
        QMessageBox.information(self, "Saved", "Prompt saved successfully.")
```

- [ ] **Step 4: Update delete_prompt to use PromptManager**

Replace the `delete_prompt` method (lines 223-238) with:

```python
    def delete_prompt(self):
        """Delete the selected prompt."""
        if not self.current_prompt_id or self.is_builtin:
            return

        reply = QMessageBox.question(
            self, "Confirm Delete",
            "Are you sure you want to delete this prompt?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )

        if reply == QMessageBox.StandardButton.Yes:
            self.pm.delete_chat_template(self.current_prompt_id)
            self.load_prompts()
            self.clear_editor()
```

- [ ] **Step 5: Verify syntax**

Run: `python -m py_compile icharlotte_core/ui/chat_dialogs.py && echo OK`
Expected: `OK`

- [ ] **Step 6: Commit**

```bash
git add icharlotte_core/ui/chat_dialogs.py
git commit -m "feat(chat-dialogs): migrate PromptTemplateDialog to global PromptManager"
```

---

### Task 3: Migrate ChatTab template menu to use PromptManager

**Files:**
- Modify: `icharlotte_core/ui/tabs.py:1840-1889`

The `update_template_menu` method currently reads custom prompts from `self.persistence` (per-case). Change it to read from PromptManager.

- [ ] **Step 1: Update update_template_menu**

In `icharlotte_core/ui/tabs.py`, replace the `update_template_menu` method (lines 1840-1874) with:

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

        # Custom prompts from global PromptManager
        try:
            from icharlotte_core.prompt_manager import get_prompt_manager
            pm = get_prompt_manager()
            custom_templates = pm.get_chat_templates()
            if custom_templates:
                menu.addSeparator()
                for t in custom_templates:
                    action = QAction(t["name"], self)
                    action.triggered.connect(lambda checked, text=t["prompt"]: self.insert_template(text))
                    menu.addAction(action)
        except Exception as e:
            print(f"[ChatTab] Could not load global templates: {e}")

        menu.addSeparator()
        manage_action = QAction("Manage Templates...", self)
        manage_action.triggered.connect(self.open_template_manager)
        menu.addAction(manage_action)

        self.template_btn.setMenu(menu)
```

- [ ] **Step 2: Update open_template_manager**

The `PromptTemplateDialog` constructor now has `persistence` as optional. The dialog still works without it. No change needed here — the existing call `PromptTemplateDialog(self.persistence, self)` still works since `persistence` is accepted but not used for templates anymore. But let me make the intent clear:

In `open_template_manager` (line 1884), no code change needed. The method already calls `PromptTemplateDialog(self.persistence, self)` which passes persistence for backward compatibility. The dialog ignores it for template operations.

- [ ] **Step 3: Verify syntax**

Run: `python -m py_compile icharlotte_core/ui/tabs.py && echo OK`
Expected: `OK`

- [ ] **Step 4: Commit**

```bash
git add icharlotte_core/ui/tabs.py
git commit -m "feat(chat-tab): load custom templates from global PromptManager"
```

---

### Task 4: Register chat_templates in the Workbench

**Files:**
- Modify: `icharlotte_core/ui/dialogs.py`

Add `chat_templates` to the Workbench agent list so templates are visible and editable there.

- [ ] **Step 1: Add to WORKBENCH_TO_AGENT_ID**

In `icharlotte_core/ui/dialogs.py`, find `WORKBENCH_TO_AGENT_ID` (around line 388). Add:

```python
    "chat_templates": "func_chat",
```

- [ ] **Step 2: Add to predefined agents**

Find the predefined agents list in `_populate_agents` (around line 1444). Add `'chat_templates'` to the list:

```python
        for agent in ['summarize', 'discovery', 'deposition',
                      'liability', 'exposure', 'med_record', 'med_chron', 'extraction',
                      'email_update', 'chat',
                      'word_assistant', 'legal_research', 'mediation_brief',
                      'chat_templates']:
```

- [ ] **Step 3: Add to TAB_TO_AGENT_MAP for auto-selection**

Find `TAB_TO_AGENT_MAP` (around line 552). Add an entry so opening the Workbench from the Chat tab auto-selects chat_templates:

```python
    TAB_TO_AGENT_MAP = {
        "Liability & Exposure": "liability",
        "Email Update": "email_update",
        "Chat": "chat_templates",
        "Index": "summarize",
        "Report": "summarize",
    }
```

Note: this changes "Chat" from mapping to `"chat"` to `"chat_templates"`. The `"chat"` agent still exists for the chat system prompt, but when coming from the Chat tab the user most likely wants to edit templates.

- [ ] **Step 4: Verify syntax**

Run: `python -m py_compile icharlotte_core/ui/dialogs.py && echo OK`
Expected: `OK`

- [ ] **Step 5: Commit**

```bash
git add icharlotte_core/ui/dialogs.py
git commit -m "feat(workbench): register chat_templates agent in Workbench UI"
```

---

### Task 5: Migrate existing per-case custom templates to global

**Files:**
- Modify: `icharlotte_core/prompt_manager.py`

Add a migration method that scans existing per-case chat JSON files, collects unique custom templates, and seeds them into PromptManager. Called once from the Workbench migration path.

- [ ] **Step 1: Add migration method to PromptManager**

Add to `PromptManager` class, after the `delete_chat_template` method:

```python
    def migrate_per_case_chat_templates(self) -> int:
        """Migrate custom templates from per-case chat JSON files to global storage.

        Scans all {file_number}_chat.json files in GEMINI_DATA_DIR, extracts
        custom (non-builtin) templates, and adds unique ones to global storage.
        Deduplicates by template name (first occurrence wins).
        Returns the number of templates migrated.
        """
        try:
            from ..config import GEMINI_DATA_DIR
        except (ImportError, ValueError):
            try:
                from icharlotte_core.config import GEMINI_DATA_DIR
            except ImportError:
                return 0

        if not os.path.isdir(GEMINI_DATA_DIR):
            return 0

        # Collect existing global template names to avoid duplicates
        existing_names = {t["name"] for t in self.get_chat_templates()}

        migrated = 0
        for filename in os.listdir(GEMINI_DATA_DIR):
            if not filename.endswith("_chat.json"):
                continue
            filepath = os.path.join(GEMINI_DATA_DIR, filename)
            try:
                with open(filepath, "r", encoding="utf-8") as f:
                    data = json.load(f)
            except Exception:
                continue
            for qp in data.get("quick_prompts", []):
                if qp.get("is_builtin", False):
                    continue
                name = qp.get("name", "").strip()
                prompt = qp.get("prompt", "").strip()
                if not name or not prompt:
                    continue
                if name in existing_names:
                    continue
                category = qp.get("category", "Custom")
                self.add_chat_template(name, prompt, category)
                existing_names.add(name)
                migrated += 1

        if migrated:
            print(f"[PromptManager] Migrated {migrated} per-case chat templates to global storage")
        return migrated
```

- [ ] **Step 2: Call migration from Workbench _migrate_if_needed**

In `icharlotte_core/ui/dialogs.py`, find `_migrate_if_needed` method. After the `seed_pipeline_prompts()` call, add:

```python
        # Migrate per-case chat templates to global storage
        try:
            migrated_templates = self.prompt_manager.migrate_per_case_chat_templates()
            if migrated_templates and migrated_templates > 0:
                log_event(f"Migrated {migrated_templates} per-case chat templates to global storage")
        except Exception as e:
            log_event(f"Error migrating chat templates: {e}", "error")
```

- [ ] **Step 3: Verify syntax**

Run: `python -m py_compile icharlotte_core/prompt_manager.py && python -m py_compile icharlotte_core/ui/dialogs.py && echo OK`
Expected: `OK`

- [ ] **Step 4: Commit**

```bash
git add icharlotte_core/prompt_manager.py icharlotte_core/ui/dialogs.py
git commit -m "feat(prompt-manager): migrate per-case chat templates to global storage"
```

---

### Task 6: End-to-end verification

**Files:** None (verification only)

- [ ] **Step 1: Run all tests**

Run: `python -m pytest tests/test_chat_template_global.py tests/test_prompt_manager_seed.py tests/test_quote_insertion.py tests/test_mediation_brief_live.py -v`
Expected: All pass

- [ ] **Step 2: Verify full compile chain**

Run: `python -m py_compile icharlotte_core/prompt_manager.py && python -m py_compile icharlotte_core/ui/chat_dialogs.py && python -m py_compile icharlotte_core/ui/tabs.py && python -m py_compile icharlotte_core/ui/dialogs.py && echo ALL OK`
Expected: `ALL OK`

- [ ] **Step 3: Verify template round-trip**

Run:
```python
python -c "
from icharlotte_core.prompt_manager import get_prompt_manager
pm = get_prompt_manager()
# Add a template
tid = pm.add_chat_template('Test Template', 'Summarize this document in 3 sentences.', 'Summary')
print(f'Added template: {tid}')
# Read it back
templates = pm.get_chat_templates()
t = [x for x in templates if x['id'] == tid][0]
print(f'Name: {t[\"name\"]}')
print(f'Prompt: {t[\"prompt\"]}')
print(f'Category: {t[\"category\"]}')
# Verify it's accessible via standard get_prompt (for Workbench editor)
text = pm.get_prompt('chat_templates', tid)
print(f'Via get_prompt: {len(text)} chars')
# Update it
pm.update_chat_template(tid, name='Updated Template', prompt='New prompt text')
t2 = [x for x in pm.get_chat_templates() if x['id'] == tid][0]
print(f'Updated name: {t2[\"name\"]}')
# Delete it
pm.delete_chat_template(tid)
remaining = [x for x in pm.get_chat_templates() if x['id'] == tid]
print(f'After delete: {len(remaining)} matches')
print('Round-trip OK!')
"
```
Expected: All operations succeed, final line prints `Round-trip OK!`
