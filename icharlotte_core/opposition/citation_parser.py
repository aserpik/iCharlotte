"""Parse case and statute citations out of a drafted opposition body.

The parser identifies citation kinds (case / statute / rule / unknown), extracts
the surrounding 1-2 sentences as the brief's proposition, and computes a stable
``normalized`` form for use as a cache key.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field


@dataclass
class Citation:
    kind: str = "unknown"
    raw_text: str = ""
    normalized: str = ""
    proposition: str = ""
    body_offset: int = 0

    # Case-specific.
    case_name: str = ""
    reporter_citation: str = ""
    year: str = ""

    # Statute-specific.
    law_code: str = ""
    section_num: str = ""


# ---------------------------------------------------------------------------
# Case-cite regex
# ---------------------------------------------------------------------------

# California reporter tokens. Order matters: longer / more specific first.
_REPORTER_PATTERN = (
    r"Cal\.\s*App\.\s*(?:2d|3d|4th|5th|6th)?"
    r"|Cal\.\s*Rptr\.\s*(?:2d|3d)?"
    r"|Cal\.\s*(?:2d|3d|4th|5th|6th)?"
    r"|P\.\s*(?:2d|3d)"
)

# Case name: Two capitalized phrases separated by " v. ". May be wrapped in
# *...* or _..._ italic markers. We capture the inner name without markers.
# Allows hyphens, apostrophes, ampersands inside the names.
_CASE_NAME_FRAGMENT = (
    r"(?:[\*_])?"                                # optional italic open
    r"([A-Z][A-Za-z0-9&'.\-]*"                   # first word
    r"(?:\s+(?:de|del|la|of|the|von|van))?"      # optional connector
    r"(?:\s+[A-Z][A-Za-z0-9&'.\-]*){0,4}"        # 0-4 more capitalized words
    r"\s+v\.\s+"                                 # required " v. "
    r"[A-Z][A-Za-z0-9&'.\-]*"
    r"(?:\s+[A-Z][A-Za-z0-9&'.\-]*){0,4}"
    r")"
    r"(?:[\*_])?"                                # optional italic close
)

_YEAR = r"(\d{4})"
_VOL = r"(\d+)"
_PAGE = r"(\d+)"
_PINCITE = r"(?:\s*,\s*\d+(?:-\d+)?)?"          # optional pincite ", 415" / ", 415-17"

_CASE_CITE_RE = re.compile(
    rf"{_CASE_NAME_FRAGMENT}\s*\({_YEAR}\)\s+{_VOL}\s+({_REPORTER_PATTERN})\s+{_PAGE}{_PINCITE}",
)


def _strip_italic_markers(s: str) -> str:
    return s.strip().strip("*_").strip()


def _normalize_case(case_name: str, vol: str, reporter: str, page: str) -> str:
    name = _strip_italic_markers(case_name)
    return f"{name} {vol} {reporter} {page}".strip()


def extract_citations(body_text: str) -> list[Citation]:
    """Extract case + statute + rule citations from a draft body."""
    citations: list[Citation] = []
    if not body_text:
        return citations

    for m in _CASE_CITE_RE.finditer(body_text):
        case_name_raw, year, vol, reporter, page = m.group(1, 2, 3, 4, 5)
        raw_text = m.group(0)
        case_name = _strip_italic_markers(case_name_raw)
        citations.append(
            Citation(
                kind="case",
                raw_text=raw_text,
                normalized=_normalize_case(case_name, vol, reporter, page),
                proposition="",
                body_offset=m.start(),
                case_name=case_name,
                year=year,
                reporter_citation=f"{vol} {reporter} {page}",
            )
        )

    citations.sort(key=lambda c: c.body_offset)
    return citations
