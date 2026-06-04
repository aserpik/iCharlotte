# icharlotte_core/firm_briefs/path_meta.py
"""Map a sorted-library file path to (motion_type, side).

The library folder layout (built by the organize step) encodes both the motion
type and the procedural side, so ingestion needs no manual tagging.
Top-level folders: "Motion - X", "Motions - Other", "Ex Parte Applications",
"Oppositions" (+ per-type subfolders), "Replies" (+ subfolders),
"Pleadings - X". Anything under "_Support*" / "_Other" is not a brief.
"""
from __future__ import annotations

import os
from typing import Optional, Tuple

# Canonical type ids. Keys are lowercased folder labels (with the leading
# "motion - " / "pleadings - " prefix already stripped) → type id.
_TYPE_ALIASES = {
    "summary judgment": "msj",
    "msj-msa": "msj",
    "msj": "msj",
    "demurrer": "demurrer",
    "strike": "strike",
    "motion to strike": "strike",
    "compel": "compel",
    "motion to compel": "compel",
    "in limine": "in_limine",
    "quash": "quash",
    "motion to quash": "quash",
    "sanctions": "sanctions",
    "relieve counsel": "relieve_counsel",
    "continue trial": "continue_trial",
    "continue trial & preference": "continue_trial",
    "other": "other",
    "motions - other": "other",
    # pleadings
    "answer": "answer",
    "complaint": "complaint",
    "amended complaint": "amended_complaint",
    "cross-complaint": "cross_complaint",
    # leave/dismiss seen as opp/reply subfolders
    "motion for leave": "leave",
    "motion to dismiss": "dismiss",
    "set aside default": "set_aside_default",
    "protective order": "protective_order",
}


def _canon_type(label: str) -> str:
    key = (label or "").strip().lower()
    if key in _TYPE_ALIASES:
        return _TYPE_ALIASES[key]
    # Fall back to a slug so unknown-but-real types still group consistently.
    return key.replace(" ", "_").replace("&", "and").replace("--", "-").strip("_") or "other"


def meta_for_path(path: str, root: str) -> Optional[Tuple[str, str]]:
    try:
        rel = os.path.relpath(path, root)
    except ValueError:
        return None
    parts = [p for p in rel.replace("\\", "/").split("/") if p]
    if len(parts) < 2:
        return None
    top = parts[0]
    low = top.lower()
    if low.startswith("_support") or low == "_other":
        return None

    sub = parts[1] if len(parts) >= 3 else ""  # subfolder when file is nested

    if low == "oppositions":
        return (_canon_type(sub), "opposition")
    if low == "replies":
        return (_canon_type(sub), "reply")
    if low == "ex parte applications":
        return ("ex_parte", "moving")
    if low.startswith("motion - "):
        return (_canon_type(top[len("motion - "):]), "moving")
    if low == "motions - other":
        return ("other", "moving")
    if low.startswith("pleadings - "):
        return (_canon_type(top[len("pleadings - "):]), "pleading")
    return None
