"""Map CAP HTML page-label anchors to plain-text character offsets.

CAP opinion HTML marks reporter page breaks inline as
``<a ... class="page-label">*56</a>``. We reconstruct the plain text the same
way a tag-stripper would (so offsets line up with stored passage text) and
record, for each page-label anchor, the plain-text offset at which that page
begins.
"""
from __future__ import annotations

import re
from bisect import bisect_right

_TAG = re.compile(r"<[^>]+>")
_PAGE_ANCHOR = re.compile(
    r'<a\b[^>]*class="[^"]*page-label[^"]*"[^>]*>\s*\*?(?P<label>\d+)\s*</a>',
    re.I,
)


def page_label_map(html: str) -> list[tuple[int, str]]:
    """Return ordered [(plaintext_offset, page_label), ...] for page breaks."""
    if not html:
        return []
    breaks: list[tuple[int, str]] = []
    out_chars = 0
    last = 0
    for m in _PAGE_ANCHOR.finditer(html):
        between = _TAG.sub("", html[last:m.start()])
        out_chars += len(between)
        breaks.append((out_chars, m.group("label")))
        last = m.end()   # the anchor itself contributes no plain text
    return breaks


def page_label_for_offset(breaks: list[tuple[int, str]], offset: int) -> str:
    """Page label whose break is at or before `offset`; '' if before the first."""
    if not breaks:
        return ""
    offs = [b[0] for b in breaks]
    idx = bisect_right(offs, offset) - 1
    if idx < 0:
        return ""
    return breaks[idx][1]


def plain_text_from_html(html: str) -> str:
    """Strip tags to recover the plain text offsets are measured against."""
    return _TAG.sub("", html or "")
