"""Global persistence for user-added custom Med-Cron analyses.

When the user adds a custom analysis row in the wizard picker, the entry
is auto-saved here on commit. Next time the picker opens, the saved
entries are loaded as pre-populated rows so the user does not have to
re-type repeating requests.

Storage is a single JSON file under the project's ``config/`` directory.
The file is per-user / per-checkout (not committed to git). The on-disk
shape is a list of ``{label, instruction}`` dicts.
"""

import json
import os
from pathlib import Path

_PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent
_STORE_PATH = _PROJECT_ROOT / "config" / "med_chron_custom_analyses.json"


def store_path() -> Path:
    """Return the absolute path to the global custom-analyses JSON file."""
    return _STORE_PATH


def load() -> list[dict]:
    """Return the saved list, or [] if the file is missing or unreadable.

    Each entry is ``{"label": str, "instruction": str}``. Entries with
    missing fields or non-string values are silently dropped.
    """
    try:
        raw = _STORE_PATH.read_text(encoding="utf-8")
    except FileNotFoundError:
        return []
    except OSError:
        return []
    try:
        data = json.loads(raw)
    except json.JSONDecodeError:
        return []
    if not isinstance(data, list):
        return []
    cleaned: list[dict] = []
    for entry in data:
        if not isinstance(entry, dict):
            continue
        label = entry.get("label")
        instruction = entry.get("instruction")
        if not isinstance(label, str) or not isinstance(instruction, str):
            continue
        label = label.strip()
        instruction = instruction.strip()
        if not label or not instruction:
            continue
        cleaned.append({"label": label, "instruction": instruction})
    return cleaned


def save(analyses: list[dict]) -> None:
    """Overwrite the global file with ``analyses``.

    Entries with empty label or instruction are silently dropped. Writes
    atomically (tmp + ``os.replace``) so a crash mid-write leaves the
    previous file intact.
    """
    cleaned: list[dict] = []
    for entry in analyses or []:
        label = (entry.get("label") or "").strip()
        instruction = (entry.get("instruction") or "").strip()
        if not label or not instruction:
            continue
        cleaned.append({"label": label, "instruction": instruction})

    _STORE_PATH.parent.mkdir(parents=True, exist_ok=True)
    tmp = _STORE_PATH.with_suffix(_STORE_PATH.suffix + ".tmp")
    tmp.write_text(json.dumps(cleaned, indent=2), encoding="utf-8")
    os.replace(tmp, _STORE_PATH)
