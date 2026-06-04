"""Join verified citations to the firm RetrievedAuthority pool by normalized cite,
copying provenance (source / source_brief / verification / alternatives) onto the
CitationVerification records the output panel renders."""
from __future__ import annotations

import re
from typing import List


def _norm(cite: str) -> str:
    return re.sub(r"\s+", "", (cite or "")).lower()


def attach_firm_provenance(citations: List, retrieved: List) -> None:
    """Mutates each CitationVerification in-place when a firm authority matches."""
    by_cite = {}
    for ra in (retrieved or []):
        if getattr(ra, "source", "") == "firm":
            by_cite[_norm(getattr(ra, "citation", ""))] = ra
    for c in (citations or []):
        key = _norm(getattr(c, "normalized_citation", "") or getattr(c, "citation_text", ""))
        ra = by_cite.get(key)
        if not ra:
            continue
        c.source = "firm"
        c.source_brief = getattr(ra, "source_brief", "") or ""
        c.firm_verification = getattr(ra, "verification", "") or ""
        alts = []
        for a in (getattr(ra, "alternatives", []) or []):
            alts.append({"case_name": getattr(a, "case_name", ""),
                         "citation": getattr(a, "citation", ""),
                         "year": getattr(a, "year", "")})
        c.alternatives = alts
