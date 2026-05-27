"""WizardStatePersistence — per-case JSON store for wizard open tabs + history.

Stored at `<case_root>/NOTES/AI OUTPUT/.icharlotte/wizard_state.json`. Atomic
writes via .tmp + os.replace. Recent tasks capped at 20 (newest first).
"""
import json
import os
from typing import Any


SCHEMA_VERSION = 1
_RECENT_CAP = 20

_README_TEXT = (
    "This folder stores iCharlotte app state for this case.\n"
    "Files here are managed by the application — do not edit manually.\n"
)


class WizardStatePersistence:
    def __init__(self, case_root: str):
        self.case_root = case_root
        self._data: dict[str, Any] | None = None

    # ---- Path helpers ----

    @property
    def folder(self) -> str:
        return os.path.join(self.case_root, "NOTES", "AI OUTPUT", ".icharlotte")

    @property
    def state_path(self) -> str:
        return os.path.join(self.folder, "wizard_state.json")

    @property
    def readme_path(self) -> str:
        return os.path.join(self.folder, "README.txt")

    # ---- Load / save ----

    def load(self) -> dict:
        if self._data is not None:
            return self._data
        if not os.path.isfile(self.state_path):
            self._data = self._default()
            return self._data
        try:
            with open(self.state_path, "r", encoding="utf-8") as f:
                raw = json.load(f)
        except (json.JSONDecodeError, OSError):
            self._data = self._default()
            return self._data
        # Tolerate missing keys.
        self._data = self._default()
        if isinstance(raw, dict):
            if isinstance(raw.get("open_tabs"), list):
                self._data["open_tabs"] = raw["open_tabs"]
            if isinstance(raw.get("recent_tasks"), list):
                self._data["recent_tasks"] = raw["recent_tasks"][:_RECENT_CAP]
        return self._data

    def save(self) -> None:
        data = self.load()
        os.makedirs(self.folder, exist_ok=True)
        if not os.path.exists(self.readme_path):
            try:
                with open(self.readme_path, "w", encoding="utf-8") as f:
                    f.write(_README_TEXT)
            except OSError:
                pass
        tmp_path = self.state_path + ".tmp"
        with open(tmp_path, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2)
        os.replace(tmp_path, self.state_path)

    # ---- Public API ----

    def set_open_tabs(self, tabs: list[dict]) -> None:
        d = self.load()
        d["open_tabs"] = list(tabs)

    def get_open_tabs(self) -> list[dict]:
        return list(self.load().get("open_tabs", []))

    def add_recent_task(self, entry: dict) -> None:
        d = self.load()
        d.setdefault("recent_tasks", []).insert(0, entry)
        if len(d["recent_tasks"]) > _RECENT_CAP:
            d["recent_tasks"] = d["recent_tasks"][:_RECENT_CAP]

    def get_recent_tasks(self) -> list[dict]:
        return list(self.load().get("recent_tasks", []))

    # ---- Internals ----

    def _default(self) -> dict:
        return {"version": SCHEMA_VERSION, "open_tabs": [], "recent_tasks": []}
