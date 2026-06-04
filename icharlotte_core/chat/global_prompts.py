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
