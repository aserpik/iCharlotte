"""Harvest case citations (with the proposition each supports) from brief text.

Thin wrapper over the opposition citation parser. Phase 1 keeps only case
cites; statutes are reused elsewhere (leginfo) and out of scope here.
"""
from __future__ import annotations

import re as _re
from dataclasses import dataclass
from typing import List

_SIGNAL_PREFIX = _re.compile(
    r"^\s*(?:see,?\s+e\.?g\.?,?|see\s+also|see|accord|cf\.?|but\s+see|in\s+re|in)\s+",
    _re.IGNORECASE,
)


def clean_case_name(name: str) -> str:
    """Strip leading citation signal words and collapse whitespace/newlines."""
    s = _re.sub(r"\s+", " ", (name or "").strip())
    prev = None
    while s and s != prev:           # strip stacked signals ("See, e.g., ...")
        prev = s
        s = _SIGNAL_PREFIX.sub("", s).strip()
    return s


@dataclass
class HarvestedCite:
    case_name: str = ""
    reporter_citation: str = ""
    year: str = ""
    norm_cite: str = ""        # reporter cite, spaces removed, lowercased
    proposition: str = ""      # sentence-window the cite supports
    quoted_passage: str = ""   # what to show if only the brief vouches for it


def _norm(reporter_citation: str) -> str:
    return (reporter_citation or "").replace(" ", "").lower()


def harvest_cites(text: str) -> List[HarvestedCite]:
    from icharlotte_core.opposition.citation_parser import extract_citations

    out: List[HarvestedCite] = []
    for c in extract_citations(text or ""):
        if getattr(c, "kind", "") != "case":
            continue
        reporter = getattr(c, "reporter_citation", "") or ""
        if not reporter:
            continue
        proposition = getattr(c, "proposition", "") or ""
        out.append(
            HarvestedCite(
                case_name=clean_case_name(getattr(c, "case_name", "") or ""),
                reporter_citation=reporter,
                year=str(getattr(c, "year", "") or ""),
                norm_cite=_norm(reporter),
                proposition=proposition,
                # Phase 1: the proposition sentence is the verify-only passage
                # used when neither the corpus nor CourtListener can confirm.
                quoted_passage=proposition,
            )
        )
    return out
