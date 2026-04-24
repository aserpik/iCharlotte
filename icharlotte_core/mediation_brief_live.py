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
import os
import tempfile
from dataclasses import dataclass, field
from typing import Dict, List, Optional

from docx import Document as DocxDocument

from icharlotte_core.mediation_brief import (
    _HEADING_PATTERN,
    _HEADING_TO_SECTION,
    MediationBriefGenerator,
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
            ``doc.Paragraphs``.  For sections with no body paragraphs
            (a heading immediately followed by the next heading),
            ``end_para_index`` will be LESS than ``start_para_index`` —
            callers can use ``end_para_index < start_para_index`` as the
            empty-body signal.  The range helper in Task 3 relies on this
            contract.
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

    Fast path: reads the full document text in a single ``doc.Content.Text``
    COM call and counts matching headings, short-circuiting once 3 are found.
    Avoids the per-paragraph COM round-trips of ``parse_brief_from_word_doc``,
    which were making the Win+V popup take 10+ seconds to open on long briefs.
    """
    try:
        text = doc_com.Content.Text or ""
    except Exception as e:
        logger.debug("is_mediation_brief: Content.Text failed: %s", e)
        return False

    seen = set()
    for line in text.splitlines():
        canonical = _match_heading(line)
        if canonical is None or canonical in seen:
            continue
        seen.add(canonical)
        if len(seen) >= 3:
            return True
    return False


def get_word_range_for_section(doc_com, section: LiveSection):
    """Return a Word ``Range`` object covering *section*'s body paragraphs.

    The range runs from the start of the first body paragraph to the end of
    the last body paragraph (inclusive of its trailing paragraph mark).

    If the section has no body paragraphs (heading followed directly by the
    next heading), returns a zero-length range at the end of the heading
    paragraph — suitable as an insertion point.
    """
    if section.end_para_index < section.start_para_index:
        # Empty body — caret at the end of the heading paragraph.
        heading_idx = section.start_para_index - 1
        if heading_idx < 1:
            heading_idx = 1
        heading_para = doc_com.Paragraphs(heading_idx)
        pos = heading_para.Range.End
        return doc_com.Range(pos, pos)

    first = doc_com.Paragraphs(section.start_para_index)
    last = doc_com.Paragraphs(section.end_para_index)
    return doc_com.Range(first.Range.Start, last.Range.End)


def _format_quote_block_text(quote: Dict) -> str:
    """Assemble a single quote block in the same format as the chat-tab flow.

    Matches the format produced by
    :meth:`MediationBriefGenerator.insert_quotes_quick` — a Q&A block
    followed by the citation on the next line, without the
    ``DEPO_QUOTE_START``/``DEPO_QUOTE_END`` markers (which are consumed by
    the section-text parser, not rendered).
    """
    qa = (quote.get("qa_text") or "").strip()
    deponent = (quote.get("deponent") or "").strip()
    page_line = (quote.get("page_line") or "").strip()
    citation = f"({deponent} Depo Trns., at p. {page_line}.)"
    return f"{qa}\n{citation}"


def insert_formatted_quotes_at_range(doc_com, range_com, quotes: List[Dict]) -> None:
    """Insert *quotes* as formatted Q&A blocks at *range_com*.

    Builds a temporary .docx containing the formatted quote paragraphs using
    :meth:`MediationBriefGenerator._add_depo_quote` — the same formatter the
    chat-tab flow uses — then calls Word COM ``Range.InsertFile`` to splice
    that content into the live document at the given range.

    The temporary file is always deleted, even if ``InsertFile`` raises.

    Args:
        doc_com: The Word COM ``Document`` (currently unused — kept for
            future hook points and to make the call site explicit).
        range_com: A Word COM ``Range`` — the insertion point / replacement
            target.
        quotes: List of quote dicts as produced by
            :meth:`MediationBriefGenerator.search_quotes`.
    """
    if not quotes:
        return

    # Build the temp docx.
    tmp = tempfile.NamedTemporaryFile(suffix=".docx", delete=False)
    tmp_path = tmp.name
    tmp.close()
    try:
        docx_doc = DocxDocument()
        generator = MediationBriefGenerator()
        for quote in quotes:
            block_text = _format_quote_block_text(quote)
            generator._add_depo_quote(docx_doc, block_text)
        docx_doc.save(tmp_path)

        # Insert into the live Word document.
        range_com.InsertFile(FileName=str(tmp_path), ConfirmConversions=False)
    finally:
        try:
            if os.path.isfile(tmp_path):
                os.unlink(tmp_path)
        except OSError as e:
            logger.warning("Failed to delete temp quote docx %s: %s", tmp_path, e)
