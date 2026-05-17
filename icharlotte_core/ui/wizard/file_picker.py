"""file_picker — default-folder resolution + (later) multi-file QFileDialog launch."""
import os
from typing import Iterable


def resolve_default_folder(case_root: str, prefs: Iterable[str]) -> str:
    """Return the first existing folder under `case_root` matching any pref.

    `prefs` is a list of relative paths like "DISCOVERY/RESPONSES". Matching is
    case-insensitive (Windows preserves on-disk case but users vary). Returns
    `case_root` itself if no pref matches or `prefs` is empty. Never raises.
    """
    if not os.path.isdir(case_root):
        return case_root

    for pref in prefs:
        candidate = _resolve_case_insensitive(case_root, pref)
        if candidate is not None and os.path.isdir(candidate):
            return candidate
    return case_root


def _resolve_case_insensitive(root: str, rel_path: str) -> str | None:
    """Walk `rel_path` segments under `root`, matching each segment case-insensitively.

    Returns the actual on-disk path if all segments resolve, else None.
    """
    parts = [p for p in rel_path.replace("\\", "/").split("/") if p]
    current = root
    for part in parts:
        try:
            entries = os.listdir(current)
        except OSError:
            return None
        target_lc = part.lower()
        match = next((e for e in entries if e.lower() == target_lc), None)
        if match is None:
            return None
        current = os.path.join(current, match)
    return current
