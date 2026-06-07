"""Parse medical chronology summary documents for Med Record Extractor."""
from __future__ import annotations

import hashlib
import os
import re
from dataclasses import dataclass, field
from typing import Literal

from docx import Document

from icharlotte_core.med_record_extractor import _parse_page_no


MatchStatus = Literal["confident", "ambiguous", "none"]


@dataclass(frozen=True)
class SynopsisParagraph:
    id: str
    order: int
    text: str
    warning: str = ""


@dataclass(frozen=True)
class SelectableChronologyRow:
    id: str
    order: int
    date: str
    page_no: str
    provider: str
    description: str
    flags: str
    record_filename: str = ""
    page_start: int = 0
    page_end: int = 0
    warning: str = ""

    @property
    def extractable(self) -> bool:
        return bool(
            self.record_filename
            and self.page_start > 0
            and self.page_end >= self.page_start
        )


@dataclass(frozen=True)
class ChronologyDocument:
    source_path: str
    synopsis_paragraphs: list[SynopsisParagraph] = field(default_factory=list)
    rows: list[SelectableChronologyRow] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    blocking_errors: list[str] = field(default_factory=list)


@dataclass(frozen=True)
class MatchResult:
    status: MatchStatus
    row_ids: tuple[str, ...] = ()
    candidate_row_ids: tuple[str, ...] = ()
    reason: str = ""


_SYNOPSIS_HEADING_RE = re.compile(
    r"^BRIEF\s+SYNOPSIS\s+OF\s+POST[-\s]INJURY\s+MEDICAL\s+RECORD:?\s*$",
    re.I,
)
_MAJOR_HEADING_RE = re.compile(r"^[A-Z0-9][A-Z0-9\s/&.,'()-]{4,}:?$")
_DATE_RE = re.compile(r"\b(\d{1,2}/\d{1,2}/\d{2,4})\b")
_CHRON_HEADERS = ("date", "pageno", "provider", "description", "redflagscomments")


def parse_chronology_document(path: str) -> ChronologyDocument:
    warnings: list[str] = []
    blocking_errors: list[str] = []
    synopsis = _parse_synopsis(path)
    rows = _parse_rows(path)
    if not synopsis:
        warnings.append("No Brief Synopsis section found.")
    if not rows:
        blocking_errors.append("No usable 5-column chronology table found.")
    return ChronologyDocument(
        source_path=os.path.normpath(path),
        synopsis_paragraphs=synopsis,
        rows=rows,
        warnings=warnings,
        blocking_errors=blocking_errors,
    )


def _parse_synopsis(path: str) -> list[SynopsisParagraph]:
    doc = Document(path)
    in_synopsis = False
    paragraphs: list[SynopsisParagraph] = []
    for para in doc.paragraphs:
        text = _collapse(para.text)
        if not text:
            continue
        if _SYNOPSIS_HEADING_RE.match(text):
            in_synopsis = True
            continue
        if not in_synopsis:
            continue
        if _MAJOR_HEADING_RE.match(text) and not _DATE_RE.search(text):
            break
        if _DATE_RE.search(text):
            order = len(paragraphs)
            paragraphs.append(
                SynopsisParagraph(
                    id=_stable_id("syn", order, text),
                    order=order,
                    text=text,
                )
            )
    return paragraphs


def _parse_rows(path: str) -> list[SelectableChronologyRow]:
    doc = Document(path)
    for table in doc.tables:
        if not table.rows or len(table.rows[0].cells) != 5:
            continue
        headers = [_normalize_header(cell.text) for cell in table.rows[0].cells]
        if not all(expected in headers[index] for index, expected in enumerate(_CHRON_HEADERS)):
            continue

        rows: list[SelectableChronologyRow] = []
        for raw_row in table.rows[1:]:
            cells = [_collapse(cell.text) for cell in raw_row.cells]
            if not cells[0]:
                continue
            if len(set(cells)) == 1:
                continue
            record_filename, page_start, page_end = _parse_page_no(raw_row.cells[1].text)
            warning = ""
            if not record_filename or page_start <= 0:
                warning = f"Could not parse record/pages from PAGE NO: {cells[1][:80]}"
            order = len(rows)
            rows.append(
                SelectableChronologyRow(
                    id=_stable_id("row", order, "|".join(cells[:4])),
                    order=order,
                    date=cells[0],
                    page_no=raw_row.cells[1].text.strip(),
                    provider=cells[2],
                    description=cells[3],
                    flags=cells[4],
                    record_filename=record_filename,
                    page_start=page_start,
                    page_end=page_end,
                    warning=warning,
                )
            )
        return rows
    return []


def _collapse(text: str) -> str:
    return re.sub(r"\s+", " ", text or "").strip()


def _normalize_header(text: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", _collapse(text).lower())


def _stable_id(prefix: str, order: int, text: str) -> str:
    digest = hashlib.sha1(text.encode("utf-8", errors="ignore")).hexdigest()[:12]
    return f"{prefix}-{order}-{digest}"
