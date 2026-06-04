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
