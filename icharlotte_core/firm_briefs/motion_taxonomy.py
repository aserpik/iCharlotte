"""Canonical motion-type vocabulary + a freeform->id normalizer.

Single source of truth shared by ingest (path_meta), Oppose-a-Motion (analyzer
output), and Generate-a-Motion (intake). Matching is keyed on these ids, so all
three must agree. Order is MOST-SPECIFIC FIRST: the first pattern that matches
the lowercased text wins. Unknown -> "other".
"""
from __future__ import annotations

import re
from typing import List, Tuple

# (id, display_name, [regex patterns]) — ORDER MATTERS (specific before generic).
CANONICAL_TYPES: List[Tuple[str, str, List[str]]] = [
    ("ime",              "Motion for Leave to Conduct IME",
        [r"\bime\b", r"independent medical exam", r"medical examination",
         r"physical examination", r"\bdme\b", r"conduct.{0,20}examination"]),
    ("gfs",              "Motion for Good Faith Settlement Determination",
        [r"good faith settlement", r"\bgfs\b"]),
    ("set_aside_default","Motion to Set Aside Default",
        [r"set[\s-]?aside", r"vacate.{0,15}default", r"relief from default"]),
    ("protective_order", "Motion for Protective Order",
        [r"protective order"]),
    ("in_limine",        "Motion in Limine",
        [r"in limine", r"\bmil\b"]),
    ("summary_judgment_alias", "", []),  # placeholder removed below (see note)
    ("msj",              "Motion for Summary Judgment/Adjudication",
        [r"summary judgment", r"summary adjudication", r"\bmsj\b", r"\bmsa\b"]),
    ("demurrer",         "Demurrer",
        [r"demurrer", r"\bdemur"]),
    ("strike",           "Motion to Strike",
        [r"motion to strike", r"\bmts\b", r"strike.{0,20}(punitive|portions|answer|complaint)"]),
    ("compel",           "Motion to Compel",
        [r"compel", r"\bmtc\b", r"\bmtca\b", r"\bmtcf\b"]),
    ("quash",            "Motion to Quash",
        [r"quash"]),
    ("sanctions",        "Motion for Sanctions",
        [r"sanction"]),
    ("relieve_counsel",  "Motion to be Relieved as Counsel",
        [r"relieved? as counsel", r"be relieved", r"withdraw as counsel", r"motion to withdraw"]),
    ("ex_parte",         "Ex Parte Application",
        [r"ex[\s-]?parte", r"\bepa\b"]),
    ("consolidate",      "Motion to Consolidate",
        [r"consolidat"]),
    ("reconsider",       "Motion for Reconsideration",
        [r"reconsider"]),
    ("dismiss",          "Motion to Dismiss",
        [r"dismiss"]),
    ("leave",            "Motion for Leave",
        [r"leave to amend", r"leave to file", r"motion for leave", r"\bleave\b"]),
    ("continue_trial",   "Motion to Continue Trial",
        [r"continue trial", r"cont(?:inuance)?.{0,12}trial", r"trial continuance",
         r"trial preference", r"preferential", r"\bpreference\b", r"specially set"]),
]

# Drop the placeholder row (kept above only to make the ordering intent explicit
# in review diffs); real lookups skip empty-pattern rows anyway.
CANONICAL_TYPES = [t for t in CANONICAL_TYPES if t[2]]

_DISPLAY = {tid: name for tid, name, _ in CANONICAL_TYPES}


def normalize_motion_type(text: str) -> str:
    """Return the canonical motion-type id for a freeform label, else 'other'."""
    low = (text or "").strip().lower()
    if not low:
        return "other"
    for tid, _name, patterns in CANONICAL_TYPES:
        for pat in patterns:
            if re.search(pat, low):
                return tid
    return "other"


def display_name(type_id: str) -> str:
    return _DISPLAY.get(type_id, type_id)
