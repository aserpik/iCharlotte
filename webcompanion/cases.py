"""Case listing (master DB) + traversal-safe browsing inside a case folder."""
import os
from pathlib import Path
from typing import List, Optional, Tuple


def list_cases(query: str = "") -> List[dict]:
    from icharlotte_core.master_db import MasterCaseDatabase
    rows = MasterCaseDatabase().get_all_cases()
    q = (query or "").strip().lower()
    if not q:
        return rows
    return [
        r for r in rows
        if q in str(r.get("file_number", "")).lower()
        or q in str(r.get("plaintiff_last_name", "")).lower()
    ]


def get_case(file_number: str) -> Optional[dict]:
    from icharlotte_core.master_db import MasterCaseDatabase
    return MasterCaseDatabase().get_case(file_number)


def safe_resolve(case_root: str, rel: str) -> Path:
    """Resolve rel under case_root; raise ValueError if it escapes the root."""
    root = Path(case_root).resolve()
    target = (root / rel).resolve() if rel else root
    if target != root and root not in target.parents:
        raise ValueError("Path escapes the case folder.")
    return target


def browse(case_root: str, rel: str, exts: tuple) -> Tuple[List[str], List[str]]:
    """Return (subdir_names, file_names) under case_root/rel, ext-filtered."""
    target = safe_resolve(case_root, rel)
    dirs: List[str] = []
    files: List[str] = []
    if not target.is_dir():
        return dirs, files
    for entry in sorted(os.listdir(target), key=str.lower):
        p = target / entry
        if p.is_dir():
            dirs.append(entry)
        elif p.suffix.lower() in exts:
            files.append(entry)
    return dirs, files


def resolve_start_folder(case_root: str, default_folders: tuple) -> str:
    """First default folder that exists under the case root, else '' (root)."""
    for rel in default_folders:
        try:
            if safe_resolve(case_root, rel).is_dir():
                return rel
        except ValueError:
            continue
    return ""
