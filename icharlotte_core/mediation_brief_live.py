"""Live-document utilities for mediation briefs.

Parses an open Word document into a :class:`LiveBrief` and provides helpers
for locating section ranges and inserting formatted quote blocks.  Used by
the Word AI assistant popup (Win+V) so refinement and quote insertion can
run against the live document without depending on in-memory generator
state.

Only the parser and range helpers live here.  Quote insertion is added in
a separate task.
"""

from __future__ import annotations

import logging
import re
from dataclasses import dataclass, field
from typing import Dict, List, Optional

from icharlotte_core.mediation_brief import (
    _HEADING_PATTERN,
    _HEADING_TO_SECTION,
)

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Dataclasses
# ---------------------------------------------------------------------------


@dataclass
class LiveSection:
    """A single section of a mediation brief parsed out of a live Word doc.

    Attributes:
        name: Canonical section name (e.g. ``"liability"``).
        heading_title: Heading text as it appears in the document (e.g.
            ``"IV. LIABILITY"``).
        text: Body text of the section — excludes the heading paragraph.
        start_para_index: 1-based index of the first body paragraph in
            ``doc.Paragraphs``.
        end_para_index: 1-based index of the last body paragraph in
            ``doc.Paragraphs``.
    """

    name: str
    heading_title: str
    text: str
    start_para_index: int
    end_para_index: int


@dataclass
class LiveBrief:
    """The parsed result of walking a live Word document for a brief."""

    doc_path: str
    sections: Dict[str, LiveSection] = field(default_factory=dict)


# ---------------------------------------------------------------------------
# Heading detection
# ---------------------------------------------------------------------------


def _match_heading(text: str) -> Optional[str]:
    """If *text* is a brief section heading, return its canonical name.

    Returns None if the paragraph text is not a recognised heading.
    """
    stripped = text.strip()
    if not stripped:
        return None
    m = _HEADING_PATTERN.match(stripped)
    if not m:
        return None
    heading_title = m.group(2).strip()
    canonical = _HEADING_TO_SECTION.get(heading_title)
    if canonical is not None:
        return canonical
    # Partial match fallback — the existing parser in mediation_brief.py
    # does the same thing.
    for variant, name in _HEADING_TO_SECTION.items():
        if heading_title.startswith(variant) or variant in heading_title:
            return name
    return None


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------


def parse_brief_from_word_doc(doc_com) -> LiveBrief:
    """Walk *doc_com*'s paragraphs and return a :class:`LiveBrief`.

    Recognises roman-numeral section headings via the same pattern and
    heading map that :mod:`icharlotte_core.mediation_brief` uses.
    Unrecognised headings are skipped silently.

    The returned ``sections`` dict is keyed by canonical section name; for
    each section, ``text`` is the concatenation of all body paragraphs
    between that heading and the next recognised heading (or end of doc).
    """
    doc_path = getattr(doc_com, "FullName", "") or ""
    live = LiveBrief(doc_path=doc_path)

    # Walk paragraphs once, tracking the current section.
    current_name: Optional[str] = None
    current_heading_title: str = ""
    current_body: List[str] = []
    current_start: int = 0
    current_end: int = 0

    def _commit():
        if current_name and current_name not in live.sections:
            live.sections[current_name] = LiveSection(
                name=current_name,
                heading_title=current_heading_title,
                text="\n".join(current_body).strip(),
                start_para_index=current_start,
                end_para_index=current_end,
            )

    paragraphs = doc_com.Paragraphs
    count = paragraphs.Count
    for idx in range(1, count + 1):
        para = paragraphs(idx)
        raw = para.Range.Text or ""
        # Word COM appends \r (and sometimes \x07 for table markers). Strip.
        text = raw.rstrip("\r\n\x07 \t")

        canonical = _match_heading(text)
        if canonical is not None:
            _commit()
            current_name = canonical
            current_heading_title = text.strip()
            current_body = []
            current_start = idx + 1
            current_end = idx  # will be updated when body paragraphs arrive
            continue

        if current_name is not None:
            if text.strip():
                current_body.append(text)
            current_end = idx

    _commit()
    return live


def is_mediation_brief(doc_com) -> bool:
    """Return True if *doc_com* contains at least 3 recognised brief sections.

    Used by the Word popup to gate the "Mediation Brief" template entries so
    they only appear when the active document looks like a brief.
    """
    try:
        live = parse_brief_from_word_doc(doc_com)
    except Exception as e:
        logger.debug("is_mediation_brief: parse failed: %s", e)
        return False
    return len(live.sections) >= 3
