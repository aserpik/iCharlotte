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

from .motion_taxonomy import normalize_motion_type


def _canon_type(label: str) -> str:
    return normalize_motion_type(label)


def _type_from_filename(path: str) -> str:
    """Subtype for the generic 'Other' buckets, read from the filename
    (strip the 'Matter__' prefix the library uses)."""
    name = os.path.basename(path)
    name = name.split("__", 1)[-1]
    return normalize_motion_type(os.path.splitext(name)[0])


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
        t = _canon_type(sub)
        if t == "other":
            t = _type_from_filename(path)
        return (t, "opposition")
    if low == "replies":
        t = _canon_type(sub)
        if t == "other":
            t = _type_from_filename(path)
        return (t, "reply")
    if low == "ex parte applications":
        return ("ex_parte", "moving")
    if low.startswith("motion - "):
        return (_canon_type(top[len("motion - "):]), "moving")
    if low == "motions - other":
        return (_type_from_filename(path), "moving")
    if low.startswith("pleadings - "):
        return (_canon_type(top[len("pleadings - "):]), "pleading")
    return None
