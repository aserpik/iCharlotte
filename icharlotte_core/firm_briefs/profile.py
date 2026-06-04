"""Compose a short 'issue profile' string for a brief, for embedding.

We embed this distilled profile (relief + argument headings + propositions),
NOT the raw OCR text, so similarity reflects the legal issues rather than
caption/boilerplate noise.
"""
from __future__ import annotations

import re
from typing import List

# A heading line: mostly uppercase letters, a few words, optional roman/numeric
# prefix. Excludes long sentences (those are prose, not captions).
_HEADING_RE = re.compile(r"^\s*(?:[IVXLC]+\.|\d+\.)?\s*([A-Z][A-Z0-9 ,'&\-\.]{6,90})\s*$")


def extract_headings(text: str, *, limit: int = 12) -> List[str]:
    heads: List[str] = []
    for line in (text or "").splitlines():
        m = _HEADING_RE.match(line)
        if not m:
            continue
        cap = m.group(1).strip()
        letters = [ch for ch in cap if ch.isalpha()]
        if len(letters) < 4:
            continue
        upper = sum(1 for ch in letters if ch.isupper())
        if upper / max(1, len(letters)) < 0.85:  # require near-all-caps
            continue
        if cap not in heads:
            heads.append(cap)
        if len(heads) >= limit:
            break
    return heads


def compose_profile(relief: str, headings: List[str], propositions: List[str]) -> str:
    parts: List[str] = []
    if relief:
        parts.append(relief.strip())
    parts.extend(h.strip() for h in headings if h.strip())
    parts.extend(p.strip() for p in propositions if p.strip())
    text = " \n".join(parts)
    return re.sub(r"\s+", " ", text).strip()


def profile_from_text(text: str, *, propositions: List[str] | None = None) -> str:
    return compose_profile("", extract_headings(text), propositions or [])
