"""Harvard CAP volume ZIP -> normalized CaseRecord + PassageRecord.

Each ZIP holds json/NNNN-01.json (metadata + opinion text + cites_to) and a
paired html/NNNN-01.html (page-label anchors for pin-cites). We index CA cases
only; the page-label map lets each passage carry the reporter page it begins on.
"""
from __future__ import annotations

import io
import json
import logging
import zipfile
from typing import Iterator

from icharlotte_core.legal_research.local_corpus.models import CaseRecord, PassageRecord
from icharlotte_core.legal_research.local_corpus.pincite import (
    page_label_for_offset, page_label_map,
)
from icharlotte_core.legal_research.local_corpus.textproc import chunk_passages, normalize_text

logger = logging.getLogger(__name__)


def _preferred_citation(citations: list[dict]) -> tuple[str, list[str]]:
    official = [c.get("cite", "") for c in citations if c.get("type") == "official" and c.get("cite")]
    others = [c.get("cite", "") for c in citations if c.get("type") != "official" and c.get("cite")]
    primary = official[0] if official else (others[0] if others else "")
    parallel = [c for c in (official[1:] + others) if c and c != primary]
    return primary, parallel


def _is_citable_citation(citation: str) -> bool:
    return "unrep" not in (citation or "").lower()


def _is_california(case: dict) -> bool:
    j = case.get("jurisdiction") or {}
    name = (j.get("name") or "") + (j.get("name_long") or "")
    return "cal" in name.lower()


def iter_cases_from_zip(zip_bytes: bytes) -> Iterator[tuple[CaseRecord, list[PassageRecord]]]:
    zf = zipfile.ZipFile(io.BytesIO(zip_bytes))
    json_names = sorted(n for n in zf.namelist() if n.startswith("json/") and n.endswith(".json"))
    for jname in json_names:
        try:
            case = json.loads(zf.read(jname).decode("utf-8"))
        except (ValueError, KeyError):
            logger.warning("CAP: bad json %s", jname, exc_info=True)
            continue
        if not _is_california(case):
            continue

        cid = case.get("id")
        if cid is None:
            continue
        case_uid = f"cap:{cid}"
        primary, parallel = _preferred_citation(case.get("citations") or [])

        opinions = ((case.get("casebody") or {}).get("opinions")) or []
        opinion_text = "\n\n".join(normalize_text(o.get("text", "")) for o in opinions if o.get("text"))

        # Page-label map from the paired HTML (best-effort; absent -> no pincites).
        hname = jname.replace("json/", "html/").replace(".json", ".html")
        breaks: list[tuple[int, str]] = []
        if hname in zf.namelist():
            try:
                breaks = page_label_map(zf.read(hname).decode("utf-8"))
            except Exception:
                logger.warning("CAP: page-label parse failed for %s", hname, exc_info=True)

        court = (case.get("court") or {}).get("name_abbreviation") or (case.get("court") or {}).get("name") or ""
        date = case.get("decision_date") or ""
        cites_to = [c.get("cite", "") for c in (case.get("cites_to") or []) if c.get("cite")]

        rec = CaseRecord(
            case_uid=case_uid, source="cap",
            name=case.get("name", ""), name_abbreviation=case.get("name_abbreviation", ""),
            citation=primary, parallel_citations=parallel, court=court,
            decision_date=date, year=(date[:4] if len(date) >= 4 else ""),
            docket_number=case.get("docket_number", ""),
            url=f"https://static.case.law/{case_uid}",  # informational
            full_text=opinion_text, cites_to=cites_to,
            published_status=("published" if _is_citable_citation(primary) else "unreported"),
            citable=_is_citable_citation(primary),
        )

        passages: list[PassageRecord] = []
        cursor = 0
        for i, chunk in enumerate(chunk_passages(opinion_text)):
            # Map this chunk's start offset (approx, via search) to a page label.
            start = opinion_text.find(chunk[:40], cursor) if chunk else -1
            if start < 0:
                start = cursor
            cursor = start + len(chunk)
            label = page_label_for_offset(breaks, start) if breaks else ""
            passages.append(PassageRecord(
                passage_uid=f"{case_uid}#{i}", case_uid=case_uid,
                ordinal=i, text=chunk, page_label=label,
            ))
        yield rec, passages
