# CA Case Law Local Corpus — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Replace the rate-limited CourtListener live-API dependency in the Oppose-a-Motion research + verification pipeline with a local, source-agnostic CA case-law corpus (Harvard CAP backbone + CourtListener bulk 2020+ gap), serving retrieval and verification entirely offline.

**Architecture:** Bulk loaders normalize CAP ZIPs and CourtListener bulk CSVs into one SQLite schema (cases + passages + FTS5) plus a `vectors.f16` memmap sidecar. `LocalCaseCorpus` does hybrid retrieval (FTS5 BM25 + exact-cosine semantic rerank via fastembed, fused by Reciprocal Rank Fusion) and exposes the same interface as `CourtListenerClient`, so it drops into `argument_research` unchanged. A `LocalCaseVerifier` reuses the existing `OppositionVerifier` orchestration but verifies case cites against local text. The whole pipeline becomes API-free.

**Tech Stack:** Python 3.x, SQLite FTS5 (stdlib `sqlite3`), `fastembed` (ONNX BGE-small, no torch), `numpy` (memmap cosine), existing `icharlotte_core.opposition` + `icharlotte_core.legal_research` packages.

**Spec:** `docs/superpowers/specs/2026-05-29-ca-caselaw-local-corpus-design.md`

**Conventions for every task below:**
- Run tests with the venv: `C:\geminiterminal2\.venv\Scripts\python.exe -m pytest <path> -v`
- New tests live under `tests/test_legal_research/test_local_corpus/` unless noted.
- All new source lives under `icharlotte_core/legal_research/local_corpus/`.
- Commit after each task with the message shown in its final step.

---

## Task 0: Scaffolding — package, dependencies, config, gitignore

**Files:**
- Create: `icharlotte_core/legal_research/local_corpus/__init__.py`
- Create: `icharlotte_core/legal_research/local_corpus/loaders/__init__.py`
- Create: `tests/test_legal_research/test_local_corpus/__init__.py`
- Modify: `requirements.txt` (append corpus deps)
- Modify: `icharlotte_core/config.py` (add `CASELAW_DATA_DIR`)
- Modify: `.gitignore` (ignore the data dir)

- [ ] **Step 1: Create the package `__init__.py` files**

`icharlotte_core/legal_research/local_corpus/__init__.py`:
```python
"""Local CA case-law corpus: offline retrieval + verification from bulk data.

Replaces the rate-limited CourtListener live API in the Oppose-a-Motion
pipeline. See docs/superpowers/specs/2026-05-29-ca-caselaw-local-corpus-design.md.
"""
```

`icharlotte_core/legal_research/local_corpus/loaders/__init__.py`:
```python
"""Source-specific bulk loaders that normalize into the corpus schema."""
```

`tests/test_legal_research/test_local_corpus/__init__.py`:
```python
```

- [ ] **Step 2: Append dependencies to `requirements.txt`**

Add after the `beautifulsoup4` line:
```
# Local case-law corpus (offline retrieval) — ONNX embeddings, no torch
fastembed>=0.3.0  # ONNX BGE-small embeddings for semantic rerank
numpy>=1.24.0     # memmap cosine similarity over passage vectors
```

- [ ] **Step 3: Add `CASELAW_DATA_DIR` to `config.py`**

In `icharlotte_core/config.py`, after the `TEMP_DIR = ...` line, add:
```python
# Local CA case-law corpus storage (SQLite DB + vectors.f16 memmap).
# Relocatable: set CASELAW_DATA_DIR env var to point at a roomier drive.
CASELAW_DATA_DIR = os.environ.get(
    "CASELAW_DATA_DIR", os.path.join(os.getcwd(), ".gemini", "caselaw")
)
```

- [ ] **Step 4: Add the data dir to `.gitignore`**

Append to `.gitignore`:
```
.gemini/caselaw/
```

- [ ] **Step 5: Verify imports resolve**

Run: `C:\geminiterminal2\.venv\Scripts\python.exe -c "import icharlotte_core.legal_research.local_corpus; from icharlotte_core.config import CASELAW_DATA_DIR; print('OK', CASELAW_DATA_DIR)"`
Expected: `OK <path>\.gemini\caselaw`

- [ ] **Step 6: Install the new deps into the venv**

Run: `C:\geminiterminal2\.venv\Scripts\python.exe -m pip install "fastembed>=0.3.0" "numpy>=1.24.0"`
Expected: installs fastembed + onnxruntime + numpy (no torch). If fastembed fails to build on Windows, STOP and report — do not substitute torch.

- [ ] **Step 7: Commit**
```bash
git add icharlotte_core/legal_research/local_corpus tests/test_legal_research/test_local_corpus requirements.txt icharlotte_core/config.py .gitignore
git commit -m "feat(corpus): scaffold local case-law package, deps, config"
```

---

## Task 1: Normalized data models + SQLite schema

**Files:**
- Create: `icharlotte_core/legal_research/local_corpus/models.py`
- Create: `icharlotte_core/legal_research/local_corpus/schema.py`
- Test: `tests/test_legal_research/test_local_corpus/test_schema.py`

- [ ] **Step 1: Write the failing test**

`tests/test_legal_research/test_local_corpus/test_schema.py`:
```python
import sqlite3
from icharlotte_core.legal_research.local_corpus.models import CaseRecord, PassageRecord
from icharlotte_core.legal_research.local_corpus import schema


def test_caserecord_roundtrips_to_row():
    rec = CaseRecord(
        case_uid="cap:1", source="cap", name="People v. Snow",
        name_abbreviation="People v. Snow", citation="30 Cal. 4th 43",
        parallel_citations=["131 Cal. Rptr. 2d 1"], court="Cal.",
        decision_date="2003-04-03", year="2003", docket_number="S018033",
        url="https://x", full_text="full text body",
        citation_count=12, latest_citing_year="2019", cites_to=["536 U.S. 584"],
    )
    row = rec.to_row()
    assert row["case_uid"] == "cap:1"
    assert row["parallel_citations"] == '["131 Cal. Rptr. 2d 1"]'  # JSON-encoded
    back = CaseRecord.from_row(row)
    assert back.parallel_citations == ["131 Cal. Rptr. 2d 1"]
    assert back.cites_to == ["536 U.S. 584"]


def test_schema_creates_tables_and_fts():
    con = sqlite3.connect(":memory:")
    schema.create_schema(con)
    names = {r[0] for r in con.execute("SELECT name FROM sqlite_master WHERE type IN ('table','view')")}
    assert {"cases", "passages", "citation_edges"}.issubset(names)
    # FTS5 virtual table is queryable
    con.execute("INSERT INTO passages_fts(rowid, text) VALUES (1, 'duty of care negligence')")
    hits = con.execute("SELECT rowid FROM passages_fts WHERE passages_fts MATCH 'negligence'").fetchall()
    assert hits == [(1,)]
```

- [ ] **Step 2: Run test to verify it fails**

Run: `C:\geminiterminal2\.venv\Scripts\python.exe -m pytest tests/test_legal_research/test_local_corpus/test_schema.py -v`
Expected: FAIL with `ModuleNotFoundError: ...models`

- [ ] **Step 3: Write `models.py`**

`icharlotte_core/legal_research/local_corpus/models.py`:
```python
"""Normalized, source-agnostic records for the local case-law corpus."""
from __future__ import annotations

import json
from dataclasses import dataclass, field
from typing import Any


@dataclass
class CaseRecord:
    case_uid: str                       # source-prefixed, e.g. "cap:269732", "cl:4408734"
    source: str                         # "cap" | "cl"
    name: str = ""                      # full caption
    name_abbreviation: str = ""         # short name for display
    citation: str = ""                  # preferred CA reporter cite
    parallel_citations: list[str] = field(default_factory=list)
    court: str = ""
    decision_date: str = ""             # ISO yyyy-mm-dd
    year: str = ""
    docket_number: str = ""
    url: str = ""
    full_text: str = ""
    citation_count: int | None = None   # inbound count (good-law soft signal)
    latest_citing_year: str = ""
    cites_to: list[str] = field(default_factory=list)  # outbound reporter cites

    def to_row(self) -> dict[str, Any]:
        return {
            "case_uid": self.case_uid,
            "source": self.source,
            "name": self.name,
            "name_abbreviation": self.name_abbreviation,
            "citation": self.citation,
            "parallel_citations": json.dumps(self.parallel_citations),
            "court": self.court,
            "decision_date": self.decision_date,
            "year": self.year,
            "docket_number": self.docket_number,
            "url": self.url,
            "full_text": self.full_text,
            "citation_count": self.citation_count,
            "latest_citing_year": self.latest_citing_year,
            "cites_to": json.dumps(self.cites_to),
        }

    @classmethod
    def from_row(cls, row: dict[str, Any]) -> "CaseRecord":
        def _loads(v: Any) -> list[str]:
            if not v:
                return []
            try:
                out = json.loads(v)
                return [str(x) for x in out] if isinstance(out, list) else []
            except (TypeError, ValueError):
                return []
        return cls(
            case_uid=row["case_uid"],
            source=row.get("source", ""),
            name=row.get("name", ""),
            name_abbreviation=row.get("name_abbreviation", ""),
            citation=row.get("citation", ""),
            parallel_citations=_loads(row.get("parallel_citations")),
            court=row.get("court", ""),
            decision_date=row.get("decision_date", ""),
            year=row.get("year", ""),
            docket_number=row.get("docket_number", ""),
            url=row.get("url", ""),
            full_text=row.get("full_text", ""),
            citation_count=row.get("citation_count"),
            latest_citing_year=row.get("latest_citing_year", ""),
            cites_to=_loads(row.get("cites_to")),
        )


@dataclass
class PassageRecord:
    passage_uid: str          # f"{case_uid}#{ordinal}"
    case_uid: str
    ordinal: int
    text: str
    page_label: str = ""      # reporter page this passage starts on (pin-cite)
    vec_row: int | None = None  # row index into vectors.f16 (set by indexer)
```

- [ ] **Step 4: Write `schema.py`**

`icharlotte_core/legal_research/local_corpus/schema.py`:
```python
"""SQLite schema + connection helper for the local case-law corpus."""
from __future__ import annotations

import os
import sqlite3

_DDL = """
CREATE TABLE IF NOT EXISTS cases (
    case_uid            TEXT PRIMARY KEY,
    source              TEXT NOT NULL,
    name                TEXT,
    name_abbreviation   TEXT,
    citation            TEXT,
    parallel_citations  TEXT,
    court               TEXT,
    decision_date       TEXT,
    year                TEXT,
    docket_number       TEXT,
    url                 TEXT,
    full_text           TEXT,
    citation_count      INTEGER,
    latest_citing_year  TEXT,
    cites_to            TEXT
);
CREATE INDEX IF NOT EXISTS idx_cases_citation ON cases(citation);

CREATE TABLE IF NOT EXISTS passages (
    passage_uid  TEXT PRIMARY KEY,
    case_uid     TEXT NOT NULL,
    ordinal      INTEGER NOT NULL,
    text         TEXT NOT NULL,
    page_label   TEXT,
    vec_row      INTEGER
);
CREATE INDEX IF NOT EXISTS idx_passages_case ON passages(case_uid);
CREATE INDEX IF NOT EXISTS idx_passages_vec  ON passages(vec_row);

CREATE TABLE IF NOT EXISTS citation_edges (
    from_case_uid  TEXT NOT NULL,
    to_citation    TEXT NOT NULL,
    weight         INTEGER DEFAULT 1
);
CREATE INDEX IF NOT EXISTS idx_edges_to ON citation_edges(to_citation);

CREATE VIRTUAL TABLE IF NOT EXISTS passages_fts USING fts5(
    text,
    content=''        -- external-content-less; we store text here directly
);
"""


def connect(db_path: str) -> sqlite3.Connection:
    """Open (creating parent dirs) a corpus DB with row dict access."""
    if db_path != ":memory:":
        os.makedirs(os.path.dirname(os.path.abspath(db_path)), exist_ok=True)
    con = sqlite3.connect(db_path)
    con.row_factory = sqlite3.Row
    con.execute("PRAGMA journal_mode=WAL")
    return con


def create_schema(con: sqlite3.Connection) -> None:
    con.executescript(_DDL)
    con.commit()
```

- [ ] **Step 5: Run tests to verify they pass**

Run: `C:\geminiterminal2\.venv\Scripts\python.exe -m pytest tests/test_legal_research/test_local_corpus/test_schema.py -v`
Expected: PASS (2 tests)

- [ ] **Step 6: Commit**
```bash
git add icharlotte_core/legal_research/local_corpus/models.py icharlotte_core/legal_research/local_corpus/schema.py tests/test_legal_research/test_local_corpus/test_schema.py
git commit -m "feat(corpus): normalized records + SQLite schema with FTS5"
```

---

## Task 2: Text normalization + passage chunking

**Files:**
- Create: `icharlotte_core/legal_research/local_corpus/textproc.py`
- Test: `tests/test_legal_research/test_local_corpus/test_textproc.py`

- [ ] **Step 1: Write the failing test**

`tests/test_legal_research/test_local_corpus/test_textproc.py`:
```python
from icharlotte_core.legal_research.local_corpus.textproc import (
    normalize_text, chunk_passages,
)


def test_normalize_fixes_section_and_whitespace():
    raw = "Pen. Code, � 187;  multiple   spaces\n\n\n\nand  breaks"
    out = normalize_text(raw)
    assert "�" not in out
    assert "§ 187" in out          # replacement char -> section symbol
    assert "multiple spaces" in out      # runs collapsed
    assert "\n\n\n" not in out           # >2 newlines collapsed


def test_chunk_passages_splits_on_paragraphs_within_budget():
    para = "Sentence one. " * 30          # ~ a chunk-ish paragraph
    text = "\n\n".join([para, para, para])
    chunks = chunk_passages(text, target_tokens=120)
    assert len(chunks) >= 2
    assert all(c.strip() for c in chunks)
    # No chunk wildly exceeds the budget (token ~ chars/4 heuristic, allow 2x)
    assert all(len(c) <= 120 * 4 * 2 for c in chunks)


def test_chunk_passages_empty_returns_empty():
    assert chunk_passages("   ") == []
```

- [ ] **Step 2: Run test to verify it fails**

Run: `C:\geminiterminal2\.venv\Scripts\python.exe -m pytest tests/test_legal_research/test_local_corpus/test_textproc.py -v`
Expected: FAIL with `ModuleNotFoundError: ...textproc`

- [ ] **Step 3: Write `textproc.py`**

`icharlotte_core/legal_research/local_corpus/textproc.py`:
```python
"""Text normalization + passage chunking for the corpus."""
from __future__ import annotations

import re

# CAP JSON occasionally carries U+FFFD where the section symbol belongs.
_REPLACEMENT = "�"


def normalize_text(text: str) -> str:
    if not text:
        return ""
    t = text.replace("\r\n", "\n").replace("\r", "\n")
    # The most common mojibake in CAP CA opinions is the section symbol.
    t = t.replace(_REPLACEMENT, "§")
    t = t.replace(" ", " ")                 # NBSP -> space
    t = re.sub(r"[ \t]+", " ", t)                # collapse intra-line runs
    t = re.sub(r"\n{3,}", "\n\n", t)             # collapse blank-line runs
    t = re.sub(r" *\n *", "\n", t)               # trim around newlines
    return t.strip()


def _approx_tokens(s: str) -> int:
    # Cheap heuristic: ~4 chars/token. Good enough for chunk budgeting.
    return max(1, len(s) // 4)


def chunk_passages(text: str, *, target_tokens: int = 512) -> list[str]:
    """Split normalized text into ~target_tokens passages on paragraph
    boundaries. Oversized paragraphs are hard-split by sentence."""
    text = normalize_text(text)
    if not text:
        return []
    paras = [p.strip() for p in text.split("\n\n") if p.strip()]
    chunks: list[str] = []
    buf: list[str] = []
    buf_tok = 0
    for para in paras:
        ptok = _approx_tokens(para)
        if ptok > target_tokens:
            # Flush buffer, then hard-split the big paragraph by sentences.
            if buf:
                chunks.append("\n\n".join(buf)); buf, buf_tok = [], 0
            chunks.extend(_split_sentences(para, target_tokens))
            continue
        if buf_tok + ptok > target_tokens and buf:
            chunks.append("\n\n".join(buf)); buf, buf_tok = [], 0
        buf.append(para); buf_tok += ptok
    if buf:
        chunks.append("\n\n".join(buf))
    return chunks


def _split_sentences(para: str, target_tokens: int) -> list[str]:
    sentences = re.split(r"(?<=[.;:])\s+", para)
    out: list[str] = []
    buf: list[str] = []
    buf_tok = 0
    for s in sentences:
        stok = _approx_tokens(s)
        if buf_tok + stok > target_tokens and buf:
            out.append(" ".join(buf)); buf, buf_tok = [], 0
        buf.append(s); buf_tok += stok
    if buf:
        out.append(" ".join(buf))
    return out
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `C:\geminiterminal2\.venv\Scripts\python.exe -m pytest tests/test_legal_research/test_local_corpus/test_textproc.py -v`
Expected: PASS (3 tests)

- [ ] **Step 5: Commit**
```bash
git add icharlotte_core/legal_research/local_corpus/textproc.py tests/test_legal_research/test_local_corpus/test_textproc.py
git commit -m "feat(corpus): text normalization + paragraph-aware passage chunking"
```

---

## Task 3: CAP HTML page-label parser (pin-cite anchors)

**Files:**
- Create: `icharlotte_core/legal_research/local_corpus/pincite.py`
- Test: `tests/test_legal_research/test_local_corpus/test_pincite.py`

Goal: from a CAP `html/NNNN-01.html` file, build an ordered list of
`(char_offset_in_plaintext, page_label)` so a passage's starting offset can be
mapped to the reporter page it begins on. CAP marks page breaks as
`<a ... class="page-label">*56</a>` inline in the opinion HTML.

- [ ] **Step 1: Write the failing test**

`tests/test_legal_research/test_local_corpus/test_pincite.py`:
```python
from icharlotte_core.legal_research.local_corpus.pincite import (
    page_label_map, page_label_for_offset,
)

SAMPLE_HTML = """
<section class="casebody">
  <p id="b1">Opening text on page fifty-five.
     <a data-label="56" class="page-label">*56</a>Now we are on page 56 and continue.
     <a data-label="57" class="page-label">*57</a>Page 57 content here.</p>
</section>
"""


def test_page_label_map_extracts_ordered_breaks():
    breaks = page_label_map(SAMPLE_HTML)
    labels = [lbl for _off, lbl in breaks]
    assert labels == ["56", "57"]
    # offsets strictly increasing
    offs = [off for off, _lbl in breaks]
    assert offs == sorted(offs)


def test_page_label_for_offset_returns_enclosing_page():
    breaks = page_label_map(SAMPLE_HTML)
    first_break_off = breaks[0][0]
    # Just before the first *56 marker -> page is the volume's first_page-ish
    # (unknown here) -> empty string (no preceding label).
    assert page_label_for_offset(breaks, max(0, first_break_off - 5)) == ""
    # At/after the *56 marker -> "56"
    assert page_label_for_offset(breaks, first_break_off + 1) == "56"
    # After the *57 marker -> "57"
    assert page_label_for_offset(breaks, breaks[1][0] + 1) == "57"
```

- [ ] **Step 2: Run test to verify it fails**

Run: `C:\geminiterminal2\.venv\Scripts\python.exe -m pytest tests/test_legal_research/test_local_corpus/test_pincite.py -v`
Expected: FAIL with `ModuleNotFoundError: ...pincite`

- [ ] **Step 3: Write `pincite.py`**

`icharlotte_core/legal_research/local_corpus/pincite.py`:
```python
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
    out_len = 0          # running length of plain text emitted so far
    pos = 0              # cursor in html
    for m in re.finditer(r"<[^>]+>", html):
        # Emit the plain text between the previous tag and this one.
        out_len += len(html[pos:m.start()])
        tag = m.group(0)
        anchor = _PAGE_ANCHOR.match(tag + _following_close(html, m))
        # Detect page-label anchor by scanning the opening <a ...> + its inner.
        pos = m.end()
        # Handle the full <a ...>*NN</a> via a dedicated scan below instead.
    # The streaming approach above is fiddly; use a robust two-pass instead:
    return _two_pass(html)


def _following_close(html: str, m) -> str:  # pragma: no cover - helper stub
    return ""


def _two_pass(html: str) -> list[tuple[int, str]]:
    breaks: list[tuple[int, str]] = []
    out_chars = 0
    last = 0
    for m in _PAGE_ANCHOR.finditer(html):
        # plain text emitted from `last` up to the anchor start
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
```

> Note: `page_label_map` delegates to the robust `_two_pass` implementation;
> the streaming scaffolding above it is intentionally short-circuited. A later
> simplification pass may delete the dead streaming branch — keep behavior
> identical (tests guard it).

- [ ] **Step 4: Run tests to verify they pass**

Run: `C:\geminiterminal2\.venv\Scripts\python.exe -m pytest tests/test_legal_research/test_local_corpus/test_pincite.py -v`
Expected: PASS (2 tests)

- [ ] **Step 5: Simplify — remove the dead streaming branch**

Replace the body of `page_label_map` so it IS `_two_pass` directly, and delete `_following_close` and the unused streaming loop. Re-run the test to confirm still PASS.

Final `page_label_map`:
```python
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
        last = m.end()
    return breaks
```

Run: `C:\geminiterminal2\.venv\Scripts\python.exe -m pytest tests/test_legal_research/test_local_corpus/test_pincite.py -v`
Expected: PASS (2 tests)

- [ ] **Step 6: Commit**
```bash
git add icharlotte_core/legal_research/local_corpus/pincite.py tests/test_legal_research/test_local_corpus/test_pincite.py
git commit -m "feat(corpus): CAP HTML page-label parser for pin-cites"
```

---

## Task 4: CAP loader (ZIP → normalized records)

**Files:**
- Create: `icharlotte_core/legal_research/local_corpus/loaders/cap_loader.py`
- Test: `tests/test_legal_research/test_local_corpus/test_cap_loader.py`

`cap_loader.iter_cases_from_zip(zip_bytes)` yields `(CaseRecord, list[PassageRecord])`
for each CA case in a CAP volume ZIP. Download orchestration (which volumes, parallel,
idempotent skip) lives in `build.py` (Task 11) and just calls this per ZIP.

- [ ] **Step 1: Write the failing test (with a synthetic fixture ZIP)**

`tests/test_legal_research/test_local_corpus/test_cap_loader.py`:
```python
import io
import json
import zipfile

from icharlotte_core.legal_research.local_corpus.loaders import cap_loader


def _make_zip() -> bytes:
    case = {
        "id": 269732,
        "name": "THE PEOPLE v. SNOW",
        "name_abbreviation": "People v. Snow",
        "decision_date": "2003-04-03",
        "docket_number": "S018033",
        "citations": [{"type": "official", "cite": "30 Cal. 4th 43"},
                      {"type": "parallel", "cite": "131 Cal. Rptr. 2d 1"}],
        "court": {"name": "Supreme Court of California", "name_abbreviation": "Cal."},
        "jurisdiction": {"name": "Cal.", "name_long": "California"},
        "cites_to": [{"cite": "536 U.S. 584"}, {"cite": "30 Cal. 4th 1"}],
        "casebody": {"opinions": [
            {"type": "majority", "author": "THE COURT",
             "text": "Para one about duty. \n\nPara two about the page break and privacy."}
        ]},
    }
    html = ('<section class="casebody"><p>Para one about duty. '
            '<a data-label="56" class="page-label">*56</a>'
            'Para two about the page break and privacy.</p></section>')
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as zf:
        zf.writestr("json/0043-01.json", json.dumps(case))
        zf.writestr("html/0043-01.html", html)
    return buf.getvalue()


def test_iter_cases_yields_normalized_case_and_passages():
    results = list(cap_loader.iter_cases_from_zip(_make_zip()))
    assert len(results) == 1
    case, passages = results[0]
    assert case.case_uid == "cap:269732"
    assert case.source == "cap"
    assert case.citation == "30 Cal. 4th 43"
    assert case.parallel_citations == ["131 Cal. Rptr. 2d 1"]
    assert case.court == "Cal."
    assert case.year == "2003"
    assert "536 U.S. 584" in case.cites_to
    assert "duty" in case.full_text
    assert passages, "expected at least one passage"
    assert passages[0].case_uid == "cap:269732"
    # Page-label propagated to at least one passage near the *56 break
    assert any(p.page_label == "56" for p in passages) or passages[0].page_label == ""


def test_non_ca_case_is_skipped():
    case = {"id": 1, "jurisdiction": {"name_long": "New York"},
            "citations": [{"type": "official", "cite": "1 N.Y. 1"}],
            "casebody": {"opinions": [{"type": "majority", "text": "x"}]}}
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as zf:
        zf.writestr("json/0001-01.json", json.dumps(case))
        zf.writestr("html/0001-01.html", "<p>x</p>")
    assert list(cap_loader.iter_cases_from_zip(buf.getvalue())) == []
```

- [ ] **Step 2: Run test to verify it fails**

Run: `C:\geminiterminal2\.venv\Scripts\python.exe -m pytest tests/test_legal_research/test_local_corpus/test_cap_loader.py -v`
Expected: FAIL with `ModuleNotFoundError: ...cap_loader`

- [ ] **Step 3: Write `cap_loader.py`**

`icharlotte_core/legal_research/local_corpus/loaders/cap_loader.py`:
```python
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
```

> Note: passage→page mapping uses the plain-text offset; CAP's HTML and JSON
> texts differ slightly in whitespace, so the label is a close approximation
> (good enough to suggest a pin-cite page; the attorney confirms exact lines).

- [ ] **Step 4: Run tests to verify they pass**

Run: `C:\geminiterminal2\.venv\Scripts\python.exe -m pytest tests/test_legal_research/test_local_corpus/test_cap_loader.py -v`
Expected: PASS (2 tests)

- [ ] **Step 5: Commit**
```bash
git add icharlotte_core/legal_research/local_corpus/loaders/cap_loader.py tests/test_legal_research/test_local_corpus/test_cap_loader.py
git commit -m "feat(corpus): CAP volume loader -> normalized records + pincite passages"
```

---

## Task 5: CourtListener bulk loader (stream-filter CA + 2020+)

**Files:**
- Create: `icharlotte_core/legal_research/local_corpus/loaders/cl_bulk_loader.py`
- Test: `tests/test_legal_research/test_local_corpus/test_cl_bulk_loader.py`

The loader takes three already-opened text streams (courts, clusters, opinions) so
tests feed `io.StringIO` and `build.py` feeds bz2-decompressed file streams. It
yields CA cases whose decision_date >= cutoff. Cross-source dedup (CAP wins) happens
at write time in `build.py`, not here.

- [ ] **Step 1: Write the failing test**

`tests/test_legal_research/test_local_corpus/test_cl_bulk_loader.py`:
```python
import csv
import io

from icharlotte_core.legal_research.local_corpus.loaders import cl_bulk_loader


def _csv(rows: list[dict]) -> io.StringIO:
    buf = io.StringIO()
    w = csv.DictWriter(buf, fieldnames=list(rows[0].keys()))
    w.writeheader()
    for r in rows:
        w.writerow(r)
    buf.seek(0)
    return buf


def test_streams_only_ca_recent_cases():
    courts = _csv([
        {"id": "cal", "full_name": "Supreme Court of California"},
        {"id": "ny", "full_name": "New York Court of Appeals"},
    ])
    clusters = _csv([
        {"id": "100", "court_id": "cal", "case_name": "Recent v. CA",
         "date_filed": "2023-06-01", "citation": "15 Cal. 5th 1"},
        {"id": "101", "court_id": "cal", "case_name": "Old v. CA",
         "date_filed": "1990-01-01", "citation": "50 Cal. 3d 1"},  # before cutoff
        {"id": "102", "court_id": "ny", "case_name": "NY thing",
         "date_filed": "2024-01-01", "citation": "1 N.Y.3d 1"},     # wrong court
    ])
    opinions = _csv([
        {"cluster_id": "100", "plain_text": "Recent CA opinion about privacy.", "html": ""},
        {"cluster_id": "101", "plain_text": "Old opinion.", "html": ""},
        {"cluster_id": "102", "plain_text": "NY opinion.", "html": ""},
    ])

    out = list(cl_bulk_loader.iter_recent_ca_cases(
        courts_stream=courts, clusters_stream=clusters, opinions_stream=opinions,
        cutoff_date="2020-01-01",
    ))
    assert len(out) == 1
    case, passages = out[0]
    assert case.case_uid == "cl:100"
    assert case.source == "cl"
    assert case.citation == "15 Cal. 5th 1"
    assert case.year == "2023"
    assert "privacy" in case.full_text
    assert passages and passages[0].case_uid == "cl:100"
```

- [ ] **Step 2: Run test to verify it fails**

Run: `C:\geminiterminal2\.venv\Scripts\python.exe -m pytest tests/test_legal_research/test_local_corpus/test_cl_bulk_loader.py -v`
Expected: FAIL with `ModuleNotFoundError: ...cl_bulk_loader`

- [ ] **Step 3: Write `cl_bulk_loader.py`**

`icharlotte_core/legal_research/local_corpus/loaders/cl_bulk_loader.py`:
```python
"""CourtListener bulk CSV stream-filter -> normalized recent CA CaseRecords.

CL bulk is full-corpus, single-format CSV. We never store the 50 GB opinions
file: we stream it, keep only rows whose cluster is CA + post-cutoff, and
discard the rest. Callers pass decompressed text streams (build.py wraps bz2).
"""
from __future__ import annotations

import csv
import logging
from typing import Iterator, TextIO

from icharlotte_core.legal_research.local_corpus.models import CaseRecord, PassageRecord
from icharlotte_core.legal_research.local_corpus.textproc import chunk_passages, normalize_text

logger = logging.getLogger(__name__)

# Court ids on CourtListener that are California state courts.
CA_COURT_IDS = {
    "cal", "calctapp", "calag", "calapp", "calsuperct",
    "calapp1st", "calapp2nd", "calapp3rd", "calapp4th", "calapp5th", "calapp6th",
}

# Opinion text columns in priority order (mirror courtlistener.py field priority).
_TEXT_COLS = ("plain_text", "html_with_citations", "html", "html_columbia",
              "html_lawbox", "xml_harvard")


def _ca_court_ids_from_courts(courts_stream: TextIO) -> set[str]:
    found: set[str] = set()
    for row in csv.DictReader(courts_stream):
        cid = (row.get("id") or "").strip()
        if cid in CA_COURT_IDS:
            found.add(cid)
    # Always include the known set even if the courts file is sparse.
    return found or set(CA_COURT_IDS)


def _recent_ca_clusters(clusters_stream: TextIO, ca_courts: set[str], cutoff: str) -> dict[str, dict]:
    keep: dict[str, dict] = {}
    for row in csv.DictReader(clusters_stream):
        court = (row.get("court_id") or "").strip()
        date = (row.get("date_filed") or "").strip()
        if court in ca_courts and date >= cutoff:
            cid = (row.get("id") or "").strip()
            if cid:
                keep[cid] = row
    return keep


def _strip_html(s: str) -> str:
    import re
    return re.sub(r"<[^>]+>", "", s or "")


def iter_recent_ca_cases(
    *,
    courts_stream: TextIO,
    clusters_stream: TextIO,
    opinions_stream: TextIO,
    cutoff_date: str,
) -> Iterator[tuple[CaseRecord, list[PassageRecord]]]:
    ca_courts = _ca_court_ids_from_courts(courts_stream)
    clusters = _recent_ca_clusters(clusters_stream, ca_courts, cutoff_date)
    if not clusters:
        return

    for row in csv.DictReader(opinions_stream):
        cid = (row.get("cluster_id") or "").strip()
        meta = clusters.get(cid)
        if not meta:
            continue
        text = ""
        for col in _TEXT_COLS:
            raw = row.get(col) or ""
            if raw:
                text = normalize_text(_strip_html(raw))
                if text:
                    break
        if not text:
            continue
        case_uid = f"cl:{cid}"
        date = (meta.get("date_filed") or "").strip()
        rec = CaseRecord(
            case_uid=case_uid, source="cl",
            name=meta.get("case_name", ""), name_abbreviation=meta.get("case_name", ""),
            citation=(meta.get("citation") or "").strip(),
            court=(meta.get("court_id") or "").strip(),
            decision_date=date, year=(date[:4] if len(date) >= 4 else ""),
            url=f"https://www.courtlistener.com/opinion/{cid}/",
            full_text=text,
        )
        passages = [
            PassageRecord(passage_uid=f"{case_uid}#{i}", case_uid=case_uid,
                          ordinal=i, text=chunk, page_label="")
            for i, chunk in enumerate(chunk_passages(text))
        ]
        yield rec, passages
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `C:\geminiterminal2\.venv\Scripts\python.exe -m pytest tests/test_legal_research/test_local_corpus/test_cl_bulk_loader.py -v`
Expected: PASS (1 test)

- [ ] **Step 5: Commit**
```bash
git add icharlotte_core/legal_research/local_corpus/loaders/cl_bulk_loader.py tests/test_legal_research/test_local_corpus/test_cl_bulk_loader.py
git commit -m "feat(corpus): CourtListener bulk stream-filter loader (CA, post-cutoff)"
```

---

## Task 6: Embedder (swappable ONNX) + deterministic fake

**Files:**
- Create: `icharlotte_core/legal_research/local_corpus/embedder.py`
- Test: `tests/test_legal_research/test_local_corpus/test_embedder.py`

- [ ] **Step 1: Write the failing test**

`tests/test_legal_research/test_local_corpus/test_embedder.py`:
```python
import numpy as np
from icharlotte_core.legal_research.local_corpus.embedder import (
    Embedder, FakeEmbedder, cosine_topk,
)


def test_fake_embedder_is_deterministic_and_normalized():
    emb = FakeEmbedder(dim=16)
    a = emb.encode(["duty of care"])
    b = emb.encode(["duty of care"])
    assert a.shape == (1, 16)
    np.testing.assert_allclose(a, b)               # deterministic
    np.testing.assert_allclose(np.linalg.norm(a[0]), 1.0, atol=1e-5)  # unit norm


def test_fake_embedder_satisfies_protocol():
    assert isinstance(FakeEmbedder(dim=8), Embedder)


def test_cosine_topk_ranks_by_similarity():
    mat = np.array([[1.0, 0.0], [0.0, 1.0], [0.7, 0.7]], dtype=np.float32)
    q = np.array([1.0, 0.0], dtype=np.float32)
    idx, scores = cosine_topk(q, mat, k=2)
    assert idx[0] == 0           # identical vector ranks first
    assert idx[1] == 2           # the 45-degree one beats the orthogonal one
```

- [ ] **Step 2: Run test to verify it fails**

Run: `C:\geminiterminal2\.venv\Scripts\python.exe -m pytest tests/test_legal_research/test_local_corpus/test_embedder.py -v`
Expected: FAIL with `ModuleNotFoundError: ...embedder`

- [ ] **Step 3: Write `embedder.py`**

`icharlotte_core/legal_research/local_corpus/embedder.py`:
```python
"""Swappable text embedder. Default = ONNX BGE-small via fastembed (no torch).

`Embedder` is a runtime-checkable Protocol so a stronger model can be dropped
in later without touching the corpus/index code. `FakeEmbedder` gives CI a
deterministic, dependency-free embedder.
"""
from __future__ import annotations

import hashlib
from typing import Protocol, runtime_checkable

import numpy as np


@runtime_checkable
class Embedder(Protocol):
    dim: int
    def encode(self, texts: list[str]) -> np.ndarray:  # (n, dim) float32, unit-norm
        ...


class FakeEmbedder:
    """Deterministic hash-based embedder for tests (no model download)."""
    def __init__(self, dim: int = 384) -> None:
        self.dim = dim

    def encode(self, texts: list[str]) -> np.ndarray:
        out = np.zeros((len(texts), self.dim), dtype=np.float32)
        for i, t in enumerate(texts):
            for tok in (t or "").lower().split():
                h = int(hashlib.md5(tok.encode()).hexdigest(), 16)
                out[i, h % self.dim] += 1.0
            n = np.linalg.norm(out[i])
            if n > 0:
                out[i] /= n
        return out


class OnnxEmbedder:
    """fastembed BGE-small (ONNX, CPU). Lazy-loads the model on first encode."""
    def __init__(self, model_name: str = "BAAI/bge-small-en-v1.5", dim: int = 384) -> None:
        self.model_name = model_name
        self.dim = dim
        self._model = None

    def _ensure(self):
        if self._model is None:
            from fastembed import TextEmbedding  # lazy import; heavy
            self._model = TextEmbedding(model_name=self.model_name)
        return self._model

    def encode(self, texts: list[str]) -> np.ndarray:
        if not texts:
            return np.zeros((0, self.dim), dtype=np.float32)
        model = self._ensure()
        vecs = np.array(list(model.embed(texts)), dtype=np.float32)
        norms = np.linalg.norm(vecs, axis=1, keepdims=True)
        norms[norms == 0] = 1.0
        return vecs / norms


def cosine_topk(query: np.ndarray, matrix: np.ndarray, k: int) -> tuple[np.ndarray, np.ndarray]:
    """Top-k rows of `matrix` by cosine similarity to `query` (both assumed
    finite; query need not be unit-norm). Returns (indices, scores) desc."""
    if matrix.shape[0] == 0:
        return np.array([], dtype=int), np.array([], dtype=np.float32)
    q = query.astype(np.float32)
    qn = np.linalg.norm(q)
    if qn > 0:
        q = q / qn
    sims = matrix @ q                       # matrix rows assumed unit-norm
    k = min(k, sims.shape[0])
    idx = np.argpartition(-sims, k - 1)[:k]
    idx = idx[np.argsort(-sims[idx])]
    return idx, sims[idx]
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `C:\geminiterminal2\.venv\Scripts\python.exe -m pytest tests/test_legal_research/test_local_corpus/test_embedder.py -v`
Expected: PASS (3 tests)

- [ ] **Step 5: Commit**
```bash
git add icharlotte_core/legal_research/local_corpus/embedder.py tests/test_legal_research/test_local_corpus/test_embedder.py
git commit -m "feat(corpus): swappable ONNX embedder + fake embedder + cosine topk"
```

---

## Task 7: Indexer (write cases, passages, FTS5, vectors.f16)

**Files:**
- Create: `icharlotte_core/legal_research/local_corpus/indexer.py`
- Test: `tests/test_legal_research/test_local_corpus/test_indexer.py`

- [ ] **Step 1: Write the failing test**

`tests/test_legal_research/test_local_corpus/test_indexer.py`:
```python
import os
import numpy as np

from icharlotte_core.legal_research.local_corpus import schema
from icharlotte_core.legal_research.local_corpus.models import CaseRecord, PassageRecord
from icharlotte_core.legal_research.local_corpus.embedder import FakeEmbedder
from icharlotte_core.legal_research.local_corpus.indexer import CorpusIndexer


def _case(uid, cite, text):
    rec = CaseRecord(case_uid=uid, source="cap", name=uid, citation=cite,
                     decision_date="2003-01-01", year="2003", full_text=text)
    passages = [PassageRecord(passage_uid=f"{uid}#0", case_uid=uid, ordinal=0, text=text)]
    return rec, passages


def test_indexer_writes_rows_fts_and_vectors(tmp_path):
    db = str(tmp_path / "corpus.db")
    vec = str(tmp_path / "vectors.f16")
    con = schema.connect(db)
    schema.create_schema(con)
    idx = CorpusIndexer(con, vectors_path=vec, embedder=FakeEmbedder(dim=32))
    idx.add(*_case("cap:1", "30 Cal. 4th 43", "duty of care and negligence"))
    idx.add(*_case("cap:2", "10 Cal. 5th 1", "privacy and discovery limits"))
    idx.finalize()

    assert con.execute("SELECT COUNT(*) FROM cases").fetchone()[0] == 2
    assert con.execute("SELECT COUNT(*) FROM passages").fetchone()[0] == 2
    fts = con.execute("SELECT COUNT(*) FROM passages_fts WHERE passages_fts MATCH 'privacy'").fetchone()[0]
    assert fts == 1
    # vectors.f16 has 2 rows of dim 32
    arr = np.memmap(vec, dtype=np.float16, mode="r").reshape(-1, 32)
    assert arr.shape == (2, 32)
    # vec_row assigned on passages
    rows = {r["passage_uid"]: r["vec_row"] for r in con.execute("SELECT passage_uid, vec_row FROM passages")}
    assert set(rows.values()) == {0, 1}
```

- [ ] **Step 2: Run test to verify it fails**

Run: `C:\geminiterminal2\.venv\Scripts\python.exe -m pytest tests/test_legal_research/test_local_corpus/test_indexer.py -v`
Expected: FAIL with `ModuleNotFoundError: ...indexer`

- [ ] **Step 3: Write `indexer.py`**

`icharlotte_core/legal_research/local_corpus/indexer.py`:
```python
"""Write normalized records into the corpus DB + vectors.f16 memmap.

Usage: create, .add(case, passages) per case, then .finalize(). Embeddings are
batched and appended to a float16 sidecar; each passage row records its vec_row.
"""
from __future__ import annotations

import sqlite3
from typing import Iterable

import numpy as np

from icharlotte_core.legal_research.local_corpus.embedder import Embedder
from icharlotte_core.legal_research.local_corpus.models import CaseRecord, PassageRecord

_BATCH = 256


class CorpusIndexer:
    def __init__(self, con: sqlite3.Connection, *, vectors_path: str, embedder: Embedder) -> None:
        self.con = con
        self.vectors_path = vectors_path
        self.embedder = embedder
        self._pending: list[PassageRecord] = []
        self._vec_blocks: list[np.ndarray] = []
        self._next_vec_row = 0
        self._seen_citation: set[str] = set()

    def add(self, case: CaseRecord, passages: Iterable[PassageRecord]) -> bool:
        """Insert one case + its passages. Returns False if deduped (skipped)."""
        # Cross-source dedup by normalized citation (first writer wins).
        norm = (case.citation or "").replace(" ", "").lower()
        if norm and norm in self._seen_citation:
            return False
        if norm:
            self._seen_citation.add(norm)
        self.con.execute(
            "INSERT OR REPLACE INTO cases (%s) VALUES (%s)" % (
                ",".join(case.to_row().keys()),
                ",".join(["?"] * len(case.to_row())),
            ),
            list(case.to_row().values()),
        )
        for p in passages:
            self._pending.append(p)
            if len(self._pending) >= _BATCH:
                self._flush()
        return True

    def _flush(self) -> None:
        if not self._pending:
            return
        vecs = self.embedder.encode([p.text for p in self._pending]).astype(np.float16)
        self._vec_blocks.append(vecs)
        for p in self._pending:
            vec_row = self._next_vec_row
            self._next_vec_row += 1
            self.con.execute(
                "INSERT OR REPLACE INTO passages (passage_uid, case_uid, ordinal, text, page_label, vec_row) "
                "VALUES (?,?,?,?,?,?)",
                (p.passage_uid, p.case_uid, p.ordinal, p.text, p.page_label, vec_row),
            )
            self.con.execute(
                "INSERT INTO passages_fts (rowid, text) VALUES (?, ?)",
                (vec_row + 1, p.text),   # fts rowid aligned to vec_row+1 (1-based)
            )
        self._pending.clear()

    def finalize(self) -> None:
        self._flush()
        self.con.commit()
        if self._vec_blocks:
            allvecs = np.concatenate(self._vec_blocks, axis=0).astype(np.float16)
        else:
            allvecs = np.zeros((0, self.embedder.dim), dtype=np.float16)
        mm = np.memmap(self.vectors_path, dtype=np.float16, mode="w+", shape=allvecs.shape)
        mm[:] = allvecs[:]
        mm.flush()
        del mm
```

> Note on FTS rowid: we set `passages_fts.rowid = vec_row + 1` so FTS hits map
> back to `passages.vec_row` (and thus the vector row) via `rowid - 1`. The
> corpus search relies on this alignment.

- [ ] **Step 4: Run tests to verify they pass**

Run: `C:\geminiterminal2\.venv\Scripts\python.exe -m pytest tests/test_legal_research/test_local_corpus/test_indexer.py -v`
Expected: PASS (1 test)

- [ ] **Step 5: Commit**
```bash
git add icharlotte_core/legal_research/local_corpus/indexer.py tests/test_legal_research/test_local_corpus/test_indexer.py
git commit -m "feat(corpus): indexer writes cases/passages/FTS5/vectors.f16"
```

---

## Task 8: Good-law soft-signal builder

**Files:**
- Create: `icharlotte_core/legal_research/local_corpus/authority_signals.py`
- Test: `tests/test_legal_research/test_local_corpus/test_authority_signals.py`

- [ ] **Step 1: Write the failing test**

`tests/test_legal_research/test_local_corpus/test_authority_signals.py`:
```python
from icharlotte_core.legal_research.local_corpus import schema
from icharlotte_core.legal_research.local_corpus.models import CaseRecord
from icharlotte_core.legal_research.local_corpus.authority_signals import build_signals


def _insert(con, rec: CaseRecord):
    con.execute(
        "INSERT INTO cases (%s) VALUES (%s)" % (
            ",".join(rec.to_row().keys()), ",".join(["?"] * len(rec.to_row()))),
        list(rec.to_row().values()),
    )


def test_build_signals_counts_inbound_and_latest_year():
    con = schema.connect(":memory:")
    schema.create_schema(con)
    # case A is cited by B (2023) and C (2019)
    _insert(con, CaseRecord(case_uid="cap:A", source="cap", citation="30 Cal. 4th 43",
                            decision_date="2003-01-01", year="2003"))
    _insert(con, CaseRecord(case_uid="cap:B", source="cap", citation="40 Cal. 4th 1",
                            decision_date="2023-01-01", year="2023", cites_to=["30 Cal. 4th 43"]))
    _insert(con, CaseRecord(case_uid="cap:C", source="cap", citation="35 Cal. 4th 9",
                            decision_date="2019-01-01", year="2019", cites_to=["30 Cal. 4th 43"]))
    con.commit()

    build_signals(con)

    row = con.execute("SELECT citation_count, latest_citing_year FROM cases WHERE case_uid='cap:A'").fetchone()
    assert row["citation_count"] == 2
    assert row["latest_citing_year"] == "2023"
```

- [ ] **Step 2: Run test to verify it fails**

Run: `C:\geminiterminal2\.venv\Scripts\python.exe -m pytest tests/test_legal_research/test_local_corpus/test_authority_signals.py -v`
Expected: FAIL with `ModuleNotFoundError: ...authority_signals`

- [ ] **Step 3: Write `authority_signals.py`**

`icharlotte_core/legal_research/local_corpus/authority_signals.py`:
```python
"""Build soft good-law signals: inbound citation count + latest citing year.

NOT a Shepard's/KeyCite check — only a staleness hint. Inverts each case's
`cites_to` across the corpus, matching by normalized reporter citation.
"""
from __future__ import annotations

import json
import sqlite3


def _norm(cite: str) -> str:
    return (cite or "").replace(" ", "").lower()


def build_signals(con: sqlite3.Connection) -> None:
    # Map normalized citation -> case_uid for resolvable targets.
    cite_to_uid: dict[str, str] = {}
    for row in con.execute("SELECT case_uid, citation FROM cases"):
        n = _norm(row["citation"])
        if n:
            cite_to_uid.setdefault(n, row["case_uid"])

    counts: dict[str, int] = {}
    latest: dict[str, str] = {}
    for row in con.execute("SELECT case_uid, year, cites_to FROM cases"):
        citing_year = row["year"] or ""
        try:
            targets = json.loads(row["cites_to"] or "[]")
        except (TypeError, ValueError):
            targets = []
        for tgt in targets:
            uid = cite_to_uid.get(_norm(tgt))
            if not uid:
                continue
            counts[uid] = counts.get(uid, 0) + 1
            if citing_year > latest.get(uid, ""):
                latest[uid] = citing_year

    for uid, n in counts.items():
        con.execute(
            "UPDATE cases SET citation_count=?, latest_citing_year=? WHERE case_uid=?",
            (n, latest.get(uid, ""), uid),
        )
    con.commit()
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `C:\geminiterminal2\.venv\Scripts\python.exe -m pytest tests/test_legal_research/test_local_corpus/test_authority_signals.py -v`
Expected: PASS (1 test)

- [ ] **Step 5: Commit**
```bash
git add icharlotte_core/legal_research/local_corpus/authority_signals.py tests/test_legal_research/test_local_corpus/test_authority_signals.py
git commit -m "feat(corpus): good-law soft-signal builder (inbound citation graph)"
```

---

## Task 9: LocalCaseCorpus (hybrid search + opinion text + signals)

**Files:**
- Create: `icharlotte_core/legal_research/local_corpus/corpus.py`
- Test: `tests/test_legal_research/test_local_corpus/test_corpus.py`

`LocalCaseCorpus` mirrors `CourtListenerClient`: `search_opinions`, `get_opinion_text`,
`get_authority_signals`, plus `lookup_by_citation` (used by the verifier). Returns
`CaseResult` objects (the existing `legal_research.models.CaseResult`).

- [ ] **Step 1: Write the failing test**

`tests/test_legal_research/test_local_corpus/test_corpus.py`:
```python
import numpy as np

from icharlotte_core.legal_research.local_corpus import schema
from icharlotte_core.legal_research.local_corpus.models import CaseRecord, PassageRecord
from icharlotte_core.legal_research.local_corpus.embedder import FakeEmbedder
from icharlotte_core.legal_research.local_corpus.indexer import CorpusIndexer
from icharlotte_core.legal_research.local_corpus.corpus import LocalCaseCorpus
from icharlotte_core.legal_research.models import CaseResult


def _build(tmp_path):
    db = str(tmp_path / "c.db"); vec = str(tmp_path / "v.f16")
    con = schema.connect(db); schema.create_schema(con)
    emb = FakeEmbedder(dim=64)
    idx = CorpusIndexer(con, vectors_path=vec, embedder=emb)
    idx.add(
        CaseRecord(case_uid="cap:1", source="cap", name="Duty v. Care",
                   citation="30 Cal. 4th 43", decision_date="2003-01-01", year="2003",
                   full_text="The duty of care in negligence is well established."),
        [PassageRecord(passage_uid="cap:1#0", case_uid="cap:1", ordinal=0, page_label="44",
                       text="The duty of care in negligence is well established.")],
    )
    idx.add(
        CaseRecord(case_uid="cap:2", source="cap", name="Privacy v. Discovery",
                   citation="10 Cal. 5th 1", decision_date="2020-01-01", year="2020",
                   full_text="Constitutional privacy limits civil discovery scope."),
        [PassageRecord(passage_uid="cap:2#0", case_uid="cap:2", ordinal=0, page_label="2",
                       text="Constitutional privacy limits civil discovery scope.")],
    )
    idx.finalize()
    con.close()
    return db, vec, emb


def test_search_returns_caseresults_for_keyword(tmp_path):
    db, vec, emb = _build(tmp_path)
    corpus = LocalCaseCorpus(db_path=db, vectors_path=vec, embedder=emb)
    results = corpus.search_opinions("privacy discovery", semantic=False, max_results=5)
    assert results and isinstance(results[0], CaseResult)
    assert results[0].cluster_id == "cap:2"
    assert results[0].citation == "10 Cal. 5th 1"


def test_search_semantic_path_runs(tmp_path):
    db, vec, emb = _build(tmp_path)
    corpus = LocalCaseCorpus(db_path=db, vectors_path=vec, embedder=emb)
    results = corpus.search_opinions("negligence duty", semantic=True, max_results=5)
    assert any(r.cluster_id == "cap:1" for r in results)


def test_get_opinion_text_and_lookup(tmp_path):
    db, vec, emb = _build(tmp_path)
    corpus = LocalCaseCorpus(db_path=db, vectors_path=vec, embedder=emb)
    assert "duty of care" in corpus.get_opinion_text("cap:1")
    hit = corpus.lookup_by_citation("30 Cal. 4th 43")
    assert hit is not None and hit["case_uid"] == "cap:1"
    assert "negligence" in hit["full_text"]
```

- [ ] **Step 2: Run test to verify it fails**

Run: `C:\geminiterminal2\.venv\Scripts\python.exe -m pytest tests/test_legal_research/test_local_corpus/test_corpus.py -v`
Expected: FAIL with `ModuleNotFoundError: ...corpus`

- [ ] **Step 3: Write `corpus.py`**

`icharlotte_core/legal_research/local_corpus/corpus.py`:
```python
"""LocalCaseCorpus: offline retrieval mirroring CourtListenerClient's interface.

search_opinions runs FTS5 BM25 + exact-cosine semantic over memmap'd vectors,
fused by Reciprocal Rank Fusion. get_opinion_text / get_authority_signals /
lookup_by_citation serve the drafter + verifier without any network call.
"""
from __future__ import annotations

import logging
import re
import sqlite3
from typing import Any

import numpy as np

from icharlotte_core.legal_research.local_corpus import schema
from icharlotte_core.legal_research.local_corpus.embedder import Embedder, OnnxEmbedder
from icharlotte_core.legal_research.models import CaseResult

logger = logging.getLogger(__name__)

_RRF_K = 60          # standard reciprocal-rank-fusion constant
_CANDIDATES = 100    # passages pulled per retrieval arm before fusion


def _fts_query(q: str) -> str:
    # OR the bare terms so partial overlaps still match; quote to neutralize FTS syntax.
    terms = re.findall(r"[A-Za-z0-9]+", q or "")
    return " OR ".join(f'"{t}"' for t in terms) if terms else '""'


class LocalCaseCorpus:
    def __init__(self, *, db_path: str, vectors_path: str, embedder: Embedder | None = None) -> None:
        self.db_path = db_path
        self.vectors_path = vectors_path
        self.embedder = embedder or OnnxEmbedder()
        self._con: sqlite3.Connection | None = None
        self._vectors: np.ndarray | None = None

    # ---- lazy resources -------------------------------------------------
    def _conn(self) -> sqlite3.Connection:
        if self._con is None:
            self._con = schema.connect(self.db_path)
        return self._con

    def _vecs(self) -> np.ndarray:
        if self._vectors is None:
            self._vectors = np.memmap(
                self.vectors_path, dtype=np.float16, mode="r"
            ).reshape(-1, self.embedder.dim)
        return self._vectors

    # ---- retrieval arms -------------------------------------------------
    def _bm25_case_ranking(self, query: str, limit: int) -> list[str]:
        con = self._conn()
        rows = con.execute(
            "SELECT p.case_uid AS uid, bm25(passages_fts) AS score "
            "FROM passages_fts JOIN passages p ON p.vec_row = passages_fts.rowid - 1 "
            "WHERE passages_fts MATCH ? ORDER BY score LIMIT ?",
            (_fts_query(query), limit),
        ).fetchall()
        seen, order = set(), []
        for r in rows:                       # bm25() ascending = best first
            if r["uid"] not in seen:
                seen.add(r["uid"]); order.append(r["uid"])
        return order

    def _semantic_case_ranking(self, query: str, limit: int) -> list[str]:
        vecs = self._vecs()
        if vecs.shape[0] == 0:
            return []
        from icharlotte_core.legal_research.local_corpus.embedder import cosine_topk
        qv = self.embedder.encode([query])[0].astype(np.float32)
        idx, _scores = cosine_topk(qv, vecs.astype(np.float32), k=limit)
        con = self._conn()
        order, seen = [], set()
        for vec_row in idx.tolist():
            row = con.execute("SELECT case_uid FROM passages WHERE vec_row=?", (int(vec_row),)).fetchone()
            if row and row["case_uid"] not in seen:
                seen.add(row["case_uid"]); order.append(row["case_uid"])
        return order

    @staticmethod
    def _rrf(*rankings: list[str]) -> list[str]:
        scores: dict[str, float] = {}
        for ranking in rankings:
            for rank, uid in enumerate(ranking):
                scores[uid] = scores.get(uid, 0.0) + 1.0 / (_RRF_K + rank + 1)
        return [uid for uid, _ in sorted(scores.items(), key=lambda kv: -kv[1])]

    # ---- public interface (mirrors CourtListenerClient) -----------------
    def search_opinions(self, query: str, *, semantic: bool = False,
                        max_results: int = 15, published_only: bool = True) -> list[CaseResult]:
        bm25 = self._bm25_case_ranking(query, _CANDIDATES)
        rankings = [bm25]
        if semantic:
            try:
                rankings.append(self._semantic_case_ranking(query, _CANDIDATES))
            except Exception:
                logger.warning("semantic ranking failed; BM25 only", exc_info=True)
        fused = self._rrf(*rankings)[:max_results]
        return [self._case_result(uid, query) for uid in fused]

    def _case_result(self, case_uid: str, query: str) -> CaseResult:
        con = self._conn()
        c = con.execute("SELECT * FROM cases WHERE case_uid=?", (case_uid,)).fetchone()
        snippet = ""
        if c:
            p = con.execute(
                "SELECT text FROM passages WHERE case_uid=? ORDER BY ordinal LIMIT 1", (case_uid,)
            ).fetchone()
            snippet = (p["text"][:400] if p else (c["full_text"] or "")[:400])
        return CaseResult(
            name=c["name"] if c else "", citation=c["citation"] if c else "",
            date=c["decision_date"] if c else "", court=c["court"] if c else "",
            snippet=snippet, url=c["url"] if c else "", cluster_id=case_uid,
        )

    def get_opinion_text(self, case_uid: str | int) -> str | None:
        row = self._conn().execute(
            "SELECT full_text FROM cases WHERE case_uid=?", (str(case_uid),)
        ).fetchone()
        return (row["full_text"] if row else None) or None

    def get_authority_signals(self, case_uid: str | int) -> dict[str, Any]:
        row = self._conn().execute(
            "SELECT citation_count, latest_citing_year FROM cases WHERE case_uid=?", (str(case_uid),)
        ).fetchone()
        if not row:
            return {"citation_count": None, "latest_citing_year": ""}
        return {"citation_count": row["citation_count"], "latest_citing_year": row["latest_citing_year"] or ""}

    def lookup_by_citation(self, citation: str) -> dict[str, Any] | None:
        norm = (citation or "").replace(" ", "").lower()
        if not norm:
            return None
        for row in self._conn().execute("SELECT * FROM cases"):
            if (row["citation"] or "").replace(" ", "").lower() == norm:
                return dict(row)
        return None
```

> Note: `lookup_by_citation` does a normalized scan for v1 simplicity; for large
> corpora replace with a normalized-citation index column. Tracked in README.

- [ ] **Step 4: Run tests to verify they pass**

Run: `C:\geminiterminal2\.venv\Scripts\python.exe -m pytest tests/test_legal_research/test_local_corpus/test_corpus.py -v`
Expected: PASS (3 tests)

- [ ] **Step 5: Commit**
```bash
git add icharlotte_core/legal_research/local_corpus/corpus.py tests/test_legal_research/test_local_corpus/test_corpus.py
git commit -m "feat(corpus): LocalCaseCorpus hybrid search + opinion text + signals"
```

---

## Task 10: Make `argument_research._opinion_text` id-type-agnostic

**Files:**
- Modify: `icharlotte_core/opposition/argument_research.py` (the `_opinion_text` function, ~lines 155-168)
- Test: `tests/test_opposition/test_argument_research_idtype.py`

- [ ] **Step 1: Write the failing test**

`tests/test_opposition/test_argument_research_idtype.py`:
```python
from icharlotte_core.opposition import argument_research


class _StrIdClient:
    """Mimics LocalCaseCorpus: get_opinion_text takes a STRING uid."""
    def __init__(self):
        self.calls = []
    def get_opinion_text(self, uid):
        self.calls.append(uid)
        assert isinstance(uid, str)        # must NOT be int()-cast
        return "opinion text" if uid == "cap:1" else ""


def test_opinion_text_passes_string_uid_through(tmp_path):
    client = _StrIdClient()
    text = argument_research._opinion_text(client, str(tmp_path), "cap:1")
    assert text == "opinion text"
    assert client.calls == ["cap:1"]
```

- [ ] **Step 2: Run test to verify it fails**

Run: `C:\geminiterminal2\.venv\Scripts\python.exe -m pytest tests/test_opposition/test_argument_research_idtype.py -v`
Expected: FAIL — current code does `int(cluster_id)` → `ValueError` → returns "" (assertion `text == "opinion text"` fails)

- [ ] **Step 3: Make the change**

In `icharlotte_core/opposition/argument_research.py`, replace the body of `_opinion_text`'s fetch block:
```python
def _opinion_text(cl_client, cache_dir: str | None, cluster_id: str) -> str:
    cached = _load_cached_opinion(cache_dir, cluster_id)
    if cached is not None:
        return cached
    try:
        text = cl_client.get_opinion_text(cluster_id) or ""
    except Exception:
        logger.warning("opinion fetch failed for %s", cluster_id, exc_info=True)
        text = ""
    if text:
        _save_cached_opinion(cache_dir, cluster_id, text)
    return text
```

> The CourtListener client keeps its own internal `int()` cast, so passing the
> raw id through still works for both clients. The previous `(TypeError, ValueError)`
> branch (which only existed to catch the `int()` cast) is removed.

- [ ] **Step 4: Run tests to verify they pass (and no CL regression)**

Run: `C:\geminiterminal2\.venv\Scripts\python.exe -m pytest tests/test_opposition/test_argument_research_idtype.py tests/test_opposition -v`
Expected: PASS (new test + existing argument_research tests unchanged)

- [ ] **Step 5: Commit**
```bash
git add icharlotte_core/opposition/argument_research.py tests/test_opposition/test_argument_research_idtype.py
git commit -m "refactor(opposition): id-type-agnostic _opinion_text for local corpus uids"
```

---

## Task 11: LocalCaseVerifier + local verifier builder

**Files:**
- Create: `icharlotte_core/opposition/local_case_verifier.py`
- Modify: `icharlotte_core/opposition/verifier.py` (add `build_local_opposition_verifier`)
- Test: `tests/test_opposition/test_local_case_verifier.py`

`LocalCaseVerifier.verify(citation)` mirrors `CaseVerifier.verify`: look the cite up in
the corpus; NOT_FOUND if absent; otherwise run the existing `verify_citation` prompt
against local text. `build_local_opposition_verifier` wires it + the existing
`StatuteVerifier` into the existing `OppositionVerifier` (reusing all its orchestration).

- [ ] **Step 1: Write the failing test**

`tests/test_opposition/test_local_case_verifier.py`:
```python
from icharlotte_core.opposition.citation_parser import Citation
from icharlotte_core.opposition.local_case_verifier import LocalCaseVerifier


class _FakeCorpus:
    def __init__(self, found):
        self._found = found
    def lookup_by_citation(self, cite):
        return self._found


def _llm_supported(_sys, _user):
    return "VERDICT: SUPPORTED\nEVIDENCE: duty of care is established\nNOTE: on point"


def test_not_found_when_citation_absent():
    v = LocalCaseVerifier(corpus=_FakeCorpus(None), llm_callback=_llm_supported)
    c = Citation(kind="case", raw_text="30 Cal. 4th 43", normalized="30 Cal. 4th 43",
                 reporter_citation="30 Cal. 4th 43", proposition="duty exists")
    cv = v.verify(c)
    assert cv.verdict == "NOT_FOUND"


def test_supported_when_text_supports(monkeypatch):
    import icharlotte_core.opposition.local_case_verifier as mod
    monkeypatch.setattr(mod, "get_prompt", lambda *_a, **_k: "{proposition}|{citation_text}|{authority_text}")
    found = {"case_uid": "cap:1", "full_text": "The duty of care is established.",
             "name": "Duty v. Care", "url": "u", "court": "Cal.", "decision_date": "2003-01-01",
             "citation_count": 9, "latest_citing_year": "2019"}
    v = LocalCaseVerifier(corpus=_FakeCorpus(found), llm_callback=_llm_supported)
    c = Citation(kind="case", raw_text="30 Cal. 4th 43", normalized="30 Cal. 4th 43",
                 reporter_citation="30 Cal. 4th 43", proposition="duty exists")
    cv = v.verify(c)
    assert cv.verdict == "SUPPORTED"
    assert cv.cluster_id == "cap:1"
    assert cv.citation_count == 9
```

- [ ] **Step 2: Run test to verify it fails**

Run: `C:\geminiterminal2\.venv\Scripts\python.exe -m pytest tests/test_opposition/test_local_case_verifier.py -v`
Expected: FAIL with `ModuleNotFoundError: ...local_case_verifier`

- [ ] **Step 3: Write `local_case_verifier.py`**

`icharlotte_core/opposition/local_case_verifier.py`:
```python
"""Verify case citations against the LocalCaseCorpus (no network).

Mirrors CaseVerifier.verify: corpus lookup by reporter citation -> NOT_FOUND if
absent, else run the shared verify_citation prompt against local full text.
Carries the corpus's good-law soft signal onto the verdict.
"""
from __future__ import annotations

import logging
from typing import Callable

from icharlotte_core.opposition.citation_parser import Citation
from icharlotte_core.opposition.models import CitationVerification
from icharlotte_core.opposition.statute_verifier import _parse_verdict_response
from icharlotte_core.prompt_manager import get_prompt

logger = logging.getLogger(__name__)

LLMCallback = Callable[[str, str], str]
_VALID = {"SUPPORTED", "PARTIAL", "NOT_SUPPORTED"}


class LocalCaseVerifier:
    def __init__(self, *, corpus, llm_callback: LLMCallback) -> None:
        self.corpus = corpus
        self.llm = llm_callback

    def verify(self, citation: Citation) -> CitationVerification:
        cv = CitationVerification(
            citation_text=citation.raw_text, normalized_citation=citation.normalized,
            kind="case", case_name=citation.case_name, date=citation.year,
            proposition=citation.proposition, body_offset=citation.body_offset,
        )
        cite = citation.reporter_citation or citation.normalized
        rec = None
        try:
            rec = self.corpus.lookup_by_citation(cite)
        except Exception:
            logger.warning("local corpus lookup failed for %s", cite, exc_info=True)
        if not rec:
            cv.verdict = "NOT_FOUND"
            cv.note = ("This citation is not in the local CA corpus (CAP through "
                       "~2020 + CourtListener recent); it may be invented, mis-cited, "
                       "out-of-state, or newer than the latest corpus build.")
            return cv

        cv.cluster_id = rec.get("case_uid", "")
        cv.opinion_url = rec.get("url", "")
        cv.court = rec.get("court", "")
        cv.date = rec.get("decision_date", "") or cv.date
        if rec.get("name") and not cv.case_name:
            cv.case_name = rec["name"]
        cv.citation_count = rec.get("citation_count")
        cv.latest_citing_year = rec.get("latest_citing_year", "") or ""

        text = rec.get("full_text") or ""
        template = get_prompt("oppose_motion", "verify_citation") or ""
        if not text or not template:
            cv.verdict = "UNVERIFIED"
            cv.note = "Corpus hit but no text/prompt available; verify manually."
            return cv

        user_prompt = template.format(
            proposition=citation.proposition or "(no proposition extracted)",
            citation_text=citation.raw_text, authority_text=text,
        )
        try:
            response = self.llm("", user_prompt) or ""
        except Exception:
            logger.warning("local verifier LLM call failed", exc_info=True)
            cv.verdict = "UNVERIFIED"; cv.note = "Verifier LLM call failed; verify manually."
            return cv

        verdict, evidence, note = _parse_verdict_response(response)
        if verdict not in _VALID:
            cv.verdict = "UNVERIFIED"; cv.note = "Could not parse verifier response; verify manually."
            return cv
        cv.verdict = verdict; cv.evidence = evidence; cv.note = note
        return cv
```

- [ ] **Step 4: Add `build_local_opposition_verifier` to `verifier.py`**

Append to `icharlotte_core/opposition/verifier.py`:
```python
def build_local_opposition_verifier(
    *,
    corpus,
    llm_callback: Callable[[str, str], str],
    max_workers: int = 4,
    cache_root: str | None = None,
) -> "OppositionVerifier":
    """OppositionVerifier whose case path is the local corpus (no network).

    Statute path keeps the existing leginfo verifier (not rate-limited).
    """
    from icharlotte_core.opposition.local_case_verifier import LocalCaseVerifier
    if cache_root is None:
        repo_root = _os.path.dirname(_os.path.dirname(_os.path.dirname(__file__)))
        cache_root = _os.path.join(repo_root, "Scripts", "prompts", "oppose_motion", ".cache")
    statute_v = StatuteVerifier(
        leginfo_client=_CALeg(), llm_callback=llm_callback,
        cache_dir=_os.path.join(cache_root, "statutes"),
    )
    return OppositionVerifier(
        case_verifier=LocalCaseVerifier(corpus=corpus, llm_callback=llm_callback),
        statute_verifier=statute_v,
        max_workers=max_workers,
    )
```

> `OppositionVerifier` calls `self.case.verify(c)` for case cites — `LocalCaseVerifier`
> satisfies that interface, so all dedup/threadpool/progress orchestration is reused.

- [ ] **Step 5: Run tests to verify they pass**

Run: `C:\geminiterminal2\.venv\Scripts\python.exe -m pytest tests/test_opposition/test_local_case_verifier.py -v`
Expected: PASS (2 tests)

- [ ] **Step 6: Commit**
```bash
git add icharlotte_core/opposition/local_case_verifier.py icharlotte_core/opposition/verifier.py tests/test_opposition/test_local_case_verifier.py
git commit -m "feat(opposition): LocalCaseVerifier + local verifier builder (no network)"
```

---

## Task 12: Build CLI orchestrator

**Files:**
- Create: `icharlotte_core/legal_research/local_corpus/build.py`
- Test: `tests/test_legal_research/test_local_corpus/test_build.py`

`build.py` exposes `build_from_cap_zips(zip_paths, ...)`, `build_from_cl_streams(...)`,
and a `main()` CLI. CAP download (which volumes, parallel, idempotent skip) and CL bz2
streaming live here. Tests exercise the in-process builders with fixtures; the network
download path is a thin documented wrapper not unit-tested.

- [ ] **Step 1: Write the failing test**

`tests/test_legal_research/test_local_corpus/test_build.py`:
```python
import io, json, zipfile
import numpy as np

from icharlotte_core.legal_research.local_corpus import build, schema
from icharlotte_core.legal_research.local_corpus.embedder import FakeEmbedder
from icharlotte_core.legal_research.local_corpus.corpus import LocalCaseCorpus


def _cap_zip():
    case = {"id": 1, "name": "Duty v. Care", "name_abbreviation": "Duty v. Care",
            "decision_date": "2003-01-01", "citations": [{"type": "official", "cite": "30 Cal. 4th 43"}],
            "court": {"name_abbreviation": "Cal."}, "jurisdiction": {"name_long": "California"},
            "cites_to": [], "casebody": {"opinions": [{"type": "majority",
            "text": "The duty of care in negligence is established."}]}}
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as zf:
        zf.writestr("json/0043-01.json", json.dumps(case))
        zf.writestr("html/0043-01.html", '<p>The duty of care in negligence is established.</p>')
    return buf.getvalue()


def test_build_from_cap_then_search(tmp_path):
    db = str(tmp_path / "c.db"); vec = str(tmp_path / "v.f16")
    emb = FakeEmbedder(dim=48)
    zpath = tmp_path / "cal-4th-1.zip"; zpath.write_bytes(_cap_zip())

    summary = build.build_from_cap_zips([str(zpath)], db_path=db, vectors_path=vec, embedder=emb)
    assert summary["cases"] == 1

    corpus = LocalCaseCorpus(db_path=db, vectors_path=vec, embedder=emb)
    results = corpus.search_opinions("negligence duty", semantic=True, max_results=5)
    assert results and results[0].cluster_id == "cap:1"
```

- [ ] **Step 2: Run test to verify it fails**

Run: `C:\geminiterminal2\.venv\Scripts\python.exe -m pytest tests/test_legal_research/test_local_corpus/test_build.py -v`
Expected: FAIL with `ModuleNotFoundError: ...build`

- [ ] **Step 3: Write `build.py`**

`icharlotte_core/legal_research/local_corpus/build.py`:
```python
"""Build/refresh the local CA case-law corpus from bulk data.

CLI: python -m icharlotte_core.legal_research.local_corpus.build --source {cap|cl|all}

CAP volumes are downloaded from static.case.law; CL bulk is streamed from S3 and
filtered to CA + post-cutoff. Both feed the same DB + vectors.f16 via CorpusIndexer.
"""
from __future__ import annotations

import argparse
import logging
import os
from typing import Any

from icharlotte_core.legal_research.local_corpus import schema
from icharlotte_core.legal_research.local_corpus.authority_signals import build_signals
from icharlotte_core.legal_research.local_corpus.embedder import Embedder, OnnxEmbedder
from icharlotte_core.legal_research.local_corpus.indexer import CorpusIndexer
from icharlotte_core.legal_research.local_corpus.loaders import cap_loader, cl_bulk_loader

logger = logging.getLogger(__name__)

CAP_CUTOFF_DATE = "2018-01-01"   # overlap buffer; CAP wins dedup so CL only fills the gap

# CA reporter series on static.case.law and their volume counts (see spec).
CAP_REPORTERS = {
    "cal": 219, "cal-2d": 71, "cal-3d": 54, "cal-4th": 63, "cal-5th": 1,
    "cal-app": 140, "cal-app-2d": 276, "cal-app-3d": 235, "cal-app-4th": 248,
    "cal-app-5th": 11, "cal-rptr-3d": 56, "cal-unrep": 7,
}


def _default_paths() -> tuple[str, str]:
    from icharlotte_core.config import CASELAW_DATA_DIR
    return (os.path.join(CASELAW_DATA_DIR, "corpus.db"),
            os.path.join(CASELAW_DATA_DIR, "vectors.f16"))


def build_from_cap_zips(zip_paths: list[str], *, db_path: str, vectors_path: str,
                        embedder: Embedder) -> dict[str, Any]:
    con = schema.connect(db_path)
    schema.create_schema(con)
    idx = CorpusIndexer(con, vectors_path=vectors_path, embedder=embedder)
    n_cases = 0
    for zp in zip_paths:
        with open(zp, "rb") as f:
            data = f.read()
        for case, passages in cap_loader.iter_cases_from_zip(data):
            if idx.add(case, passages):
                n_cases += 1
    idx.finalize()
    build_signals(con)
    con.close()
    return {"cases": n_cases}


def build_from_cl_streams(*, courts_stream, clusters_stream, opinions_stream,
                          db_path: str, vectors_path: str, embedder: Embedder,
                          cutoff_date: str = CAP_CUTOFF_DATE) -> dict[str, Any]:
    con = schema.connect(db_path)
    schema.create_schema(con)
    idx = CorpusIndexer(con, vectors_path=vectors_path, embedder=embedder)
    n_cases = 0
    for case, passages in cl_bulk_loader.iter_recent_ca_cases(
        courts_stream=courts_stream, clusters_stream=clusters_stream,
        opinions_stream=opinions_stream, cutoff_date=cutoff_date,
    ):
        if idx.add(case, passages):
            n_cases += 1
    idx.finalize()
    build_signals(con)
    con.close()
    return {"cases": n_cases}


def _download_cap_volumes(scratch_dir: str) -> list[str]:  # pragma: no cover - network
    """Download every CA reporter volume ZIP to scratch_dir; skip existing."""
    import urllib.request
    os.makedirs(scratch_dir, exist_ok=True)
    paths: list[str] = []
    for rep, count in CAP_REPORTERS.items():
        for vol in range(1, count + 1):
            dest = os.path.join(scratch_dir, f"{rep}-{vol}.zip")
            if not os.path.exists(dest):
                url = f"https://static.case.law/{rep}/{vol}.zip"
                try:
                    req = urllib.request.Request(url, headers={"User-Agent": "iCharlotte/1.0"})
                    with urllib.request.urlopen(req, timeout=60) as r, open(dest, "wb") as out:
                        out.write(r.read())
                except Exception:
                    logger.warning("CAP download failed: %s", url, exc_info=True)
                    continue
            paths.append(dest)
    return paths


def main() -> None:  # pragma: no cover - CLI wrapper
    ap = argparse.ArgumentParser(description="Build the local CA case-law corpus")
    ap.add_argument("--source", choices=["cap", "cl", "all"], default="all")
    ap.add_argument("--data-dir", default=None)
    args = ap.parse_args()
    logging.basicConfig(level=logging.INFO)

    if args.data_dir:
        db_path = os.path.join(args.data_dir, "corpus.db")
        vectors_path = os.path.join(args.data_dir, "vectors.f16")
    else:
        db_path, vectors_path = _default_paths()
    embedder = OnnxEmbedder()

    if args.source in ("cap", "all"):
        scratch = os.path.join(os.path.dirname(db_path), "_cap_scratch")
        zips = _download_cap_volumes(scratch)
        summary = build_from_cap_zips(zips, db_path=db_path, vectors_path=vectors_path, embedder=embedder)
        logger.info("CAP ingest: %s cases", summary["cases"])
    if args.source in ("cl", "all"):
        logger.info("CL bulk ingest: stream CourtListener bulk CSVs into "
                    "build_from_cl_streams (see README for the exact stream wiring).")


if __name__ == "__main__":  # pragma: no cover
    main()
```

> The CL streaming wrapper (bz2-decompress the S3 bulk files into
> `build_from_cl_streams`) is documented in the README and intentionally not
> unit-tested (it is a ~52 GB network operation). `build_from_cl_streams` itself
> is fully tested via `cl_bulk_loader` fixtures.

- [ ] **Step 4: Run tests to verify they pass**

Run: `C:\geminiterminal2\.venv\Scripts\python.exe -m pytest tests/test_legal_research/test_local_corpus/test_build.py -v`
Expected: PASS (1 test)

- [ ] **Step 5: Commit**
```bash
git add icharlotte_core/legal_research/local_corpus/build.py tests/test_legal_research/test_local_corpus/test_build.py
git commit -m "feat(corpus): build CLI orchestrator (CAP + CL ingest)"
```

---

## Task 13: Wire the corpus into the Oppose-a-Motion worker

**Files:**
- Modify: `icharlotte_core/ui/wizard/pages/oppose_motion_page.py` (research block ~1112-1138 and verifier block ~1169-1179)
- Test: `tests/test_wizard/test_oppose_motion_local_corpus.py`

Replace the live `CourtListenerClient` + `build_opposition_verifier` with the local
corpus + `build_local_opposition_verifier`, guarded by corpus availability (DB exists).
Falls back to the existing CL path only if the corpus DB is absent (so nothing breaks
before the first build).

- [ ] **Step 1: Write the failing test**

`tests/test_wizard/test_oppose_motion_local_corpus.py`:
```python
import os
from icharlotte_core.ui.wizard.pages import oppose_motion_page as omp


def test_corpus_available_true_when_db_exists(tmp_path, monkeypatch):
    db = tmp_path / "corpus.db"; db.write_text("x")
    monkeypatch.setattr(omp, "CASELAW_DATA_DIR", str(tmp_path))
    assert omp._corpus_available() is True


def test_corpus_available_false_when_missing(tmp_path, monkeypatch):
    monkeypatch.setattr(omp, "CASELAW_DATA_DIR", str(tmp_path))
    assert omp._corpus_available() is False


def test_make_local_corpus_returns_corpus(tmp_path, monkeypatch):
    # Build a tiny real corpus so the constructor path is exercised.
    from icharlotte_core.legal_research.local_corpus import build
    from icharlotte_core.legal_research.local_corpus.embedder import FakeEmbedder
    import io, json, zipfile
    case = {"id": 1, "name": "A v. B", "decision_date": "2003-01-01",
            "citations": [{"type": "official", "cite": "30 Cal. 4th 43"}],
            "jurisdiction": {"name_long": "California"}, "cites_to": [],
            "casebody": {"opinions": [{"type": "majority", "text": "duty of care."}]}}
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as zf:
        zf.writestr("json/0043-01.json", json.dumps(case))
        zf.writestr("html/0043-01.html", "<p>duty of care.</p>")
    z = tmp_path / "v.zip"; z.write_bytes(buf.getvalue())
    build.build_from_cap_zips([str(z)], db_path=str(tmp_path / "corpus.db"),
                              vectors_path=str(tmp_path / "vectors.f16"), embedder=FakeEmbedder(dim=32))

    monkeypatch.setattr(omp, "CASELAW_DATA_DIR", str(tmp_path))
    monkeypatch.setattr(omp, "_corpus_embedder", lambda: FakeEmbedder(dim=32))
    corpus = omp._make_local_corpus()
    assert corpus is not None
    assert corpus.get_opinion_text("cap:1")
```

- [ ] **Step 2: Run test to verify it fails**

Run: `C:\geminiterminal2\.venv\Scripts\python.exe -m pytest tests/test_wizard/test_oppose_motion_local_corpus.py -v`
Expected: FAIL — `_corpus_available` / `_make_local_corpus` / `CASELAW_DATA_DIR` not defined in module

- [ ] **Step 3: Add corpus helpers + imports near the top of `oppose_motion_page.py`**

After the existing `from icharlotte_core.opposition.argument_research import research_arguments` import, add:
```python
import os as _os_corpus
from icharlotte_core.config import CASELAW_DATA_DIR


def _corpus_paths() -> tuple[str, str]:
    return (_os_corpus.path.join(CASELAW_DATA_DIR, "corpus.db"),
            _os_corpus.path.join(CASELAW_DATA_DIR, "vectors.f16"))


def _corpus_available() -> bool:
    db, _vec = _corpus_paths()
    return _os_corpus.path.exists(db)


def _corpus_embedder():
    from icharlotte_core.legal_research.local_corpus.embedder import OnnxEmbedder
    return OnnxEmbedder()


def _make_local_corpus():
    if not _corpus_available():
        return None
    from icharlotte_core.legal_research.local_corpus.corpus import LocalCaseCorpus
    db, vec = _corpus_paths()
    return LocalCaseCorpus(db_path=db, vectors_path=vec, embedder=_corpus_embedder())
```

- [ ] **Step 4: Swap the research client in the worker**

In the research block (currently builds `CourtListenerClient(token)` and calls
`research_arguments(... cl_client=CourtListenerClient(token) ...)`), replace with:
```python
            # Retrieval-first grounding: prefer the local CA corpus (offline,
            # unlimited). Fall back to the live CourtListener API only if the
            # corpus has not been built yet.
            corpus = _make_local_corpus()
            token = os.environ.get("COURTLISTENER_API_TOKEN", "").strip()
            retrieved = []
            if corpus is not None and metadata.principal_arguments:
                opinion_cache = os.path.join(os.path.dirname(registry_path), ".cache", "opinions")
                self.progress.emit(
                    f"Researching authorities locally ({len(metadata.principal_arguments)} arguments)..."
                )
                retrieved = research_arguments(
                    metadata.principal_arguments,
                    cl_client=corpus,
                    query_llm=make_llm("research_queries"),
                    rerank_llm=make_llm("rerank_select"),
                    max_workers=4,
                    on_progress=self.progress.emit,
                    cache_dir=opinion_cache,
                )
                self.progress.emit(f"Retrieved {len(retrieved)} grounded authorities.")
            elif corpus is None and token and metadata.principal_arguments:
                from icharlotte_core.legal_research.sources.courtlistener import CourtListenerClient
                opinion_cache = os.path.join(os.path.dirname(registry_path), ".cache", "opinions")
                self.progress.emit(
                    "Local corpus not built; falling back to CourtListener API "
                    f"({len(metadata.principal_arguments)} arguments)..."
                )
                retrieved = research_arguments(
                    metadata.principal_arguments,
                    cl_client=CourtListenerClient(token),
                    query_llm=make_llm("research_queries"),
                    rerank_llm=make_llm("rerank_select"),
                    max_workers=4, on_progress=self.progress.emit, cache_dir=opinion_cache,
                )
                self.progress.emit(f"Retrieved {len(retrieved)} grounded authorities.")
            else:
                self.progress.emit(
                    "WARNING: no local corpus and no COURTLISTENER_API_TOKEN; "
                    "drafting without grounded research."
                )
```

- [ ] **Step 5: Swap the verifier in the worker**

In the verification block (currently `verifier = build_opposition_verifier(...)`), replace with:
```python
                if corpus is not None:
                    from icharlotte_core.opposition.verifier import build_local_opposition_verifier
                    verifier = build_local_opposition_verifier(
                        corpus=corpus, llm_callback=llm, max_workers=4,
                    )
                else:
                    verifier = build_opposition_verifier(
                        courtlistener_token=token, llm_callback=llm, max_workers=4,
                    )
```

Ensure `build_local_opposition_verifier` is importable where `build_opposition_verifier`
is already imported (top of file) — add it to that import line.

- [ ] **Step 6: Run tests to verify they pass (+ existing wizard opposition tests)**

Run: `C:\geminiterminal2\.venv\Scripts\python.exe -m pytest tests/test_wizard/test_oppose_motion_local_corpus.py tests/test_wizard/test_oppose_motion_page.py -v`
Expected: PASS (new tests + existing page tests unchanged)

- [ ] **Step 7: Commit**
```bash
git add icharlotte_core/ui/wizard/pages/oppose_motion_page.py tests/test_wizard/test_oppose_motion_local_corpus.py
git commit -m "feat(opposition): wire local corpus + local verifier into the wizard worker"
```

---

## Task 14: Module README + DEVELOPMENT_LOG

**Files:**
- Create: `icharlotte_core/legal_research/local_corpus/README.md`
- Modify: `DEVELOPMENT_LOG.md` (prepend an entry)

- [ ] **Step 1: Write the README**

Create `icharlotte_core/legal_research/local_corpus/README.md` covering:
- What it is + why (rate-limit escape; spec link).
- Build commands: `python -m icharlotte_core.legal_research.local_corpus.build --source all`.
- The CL streaming wiring snippet (how to bz2-decompress the S3 bulk files into
  `build_from_cl_streams`) with exact S3 URLs from the spec.
- Quarterly refresh instructions.
- How to swap the embedder (implement `Embedder`, pass to `LocalCaseCorpus`).
- Known limits (copy the spec's Risks list): no good-law, recency gap, CA-only,
  semantic < hosted, recent-slice text quality, ~52 GB one-time stream, quarterly rebuild.
- The `lookup_by_citation` linear-scan TODO (add a normalized-citation index column for scale).

- [ ] **Step 2: Prepend a DEVELOPMENT_LOG entry**

Add a dated entry summarizing the feature, its modules, and the build command.

- [ ] **Step 3: Commit**
```bash
git add icharlotte_core/legal_research/local_corpus/README.md DEVELOPMENT_LOG.md
git commit -m "docs(corpus): README + development log entry"
```

---

## Task 15: Full-suite regression + gated real-embedder integration test

**Files:**
- Create: `tests/test_legal_research/test_local_corpus/test_integration_real_embedder.py`

- [ ] **Step 1: Write the gated integration test**

`tests/test_legal_research/test_local_corpus/test_integration_real_embedder.py`:
```python
import io, json, zipfile
import pytest

fastembed = pytest.importorskip("fastembed")  # skip if model deps absent


def test_real_embedder_end_to_end(tmp_path):
    from icharlotte_core.legal_research.local_corpus import build
    from icharlotte_core.legal_research.local_corpus.embedder import OnnxEmbedder
    from icharlotte_core.legal_research.local_corpus.corpus import LocalCaseCorpus

    def _zip(cid, cite, text):
        case = {"id": cid, "name": f"Case {cid}", "decision_date": "2010-01-01",
                "citations": [{"type": "official", "cite": cite}],
                "jurisdiction": {"name_long": "California"}, "cites_to": [],
                "casebody": {"opinions": [{"type": "majority", "text": text}]}}
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w") as zf:
            zf.writestr(f"json/{cid:04d}-01.json", json.dumps(case))
            zf.writestr(f"html/{cid:04d}-01.html", f"<p>{text}</p>")
        p = tmp_path / f"z{cid}.zip"; p.write_bytes(buf.getvalue()); return str(p)

    zips = [
        _zip(1, "30 Cal. 4th 43", "The duty of care in negligence requires foreseeability."),
        _zip(2, "10 Cal. 5th 1", "Constitutional privacy constrains the scope of civil discovery."),
    ]
    db = str(tmp_path / "c.db"); vec = str(tmp_path / "v.f16")
    build.build_from_cap_zips(zips, db_path=db, vectors_path=vec, embedder=OnnxEmbedder())

    corpus = LocalCaseCorpus(db_path=db, vectors_path=vec, embedder=OnnxEmbedder())
    # Semantic query phrased differently from the opinion text should still hit case 2.
    results = corpus.search_opinions("can the other side inspect private records",
                                     semantic=True, max_results=2)
    assert any(r.cluster_id == "cap:2" for r in results)
```

- [ ] **Step 2: Run the gated test (downloads the model on first run)**

Run: `C:\geminiterminal2\.venv\Scripts\python.exe -m pytest tests/test_legal_research/test_local_corpus/test_integration_real_embedder.py -v`
Expected: PASS (or SKIP if fastembed unavailable)

- [ ] **Step 3: Run the full corpus + opposition + wizard suites**

Run: `C:\geminiterminal2\.venv\Scripts\python.exe -m pytest tests/test_legal_research tests/test_opposition tests/test_wizard -q`
Expected: all PASS (pre-existing skips unchanged)

- [ ] **Step 4: Commit**
```bash
git add tests/test_legal_research/test_local_corpus/test_integration_real_embedder.py
git commit -m "test(corpus): gated real-embedder end-to-end + full-suite regression"
```

---

## Post-implementation (not tasks — do after the plan)

- **Run the real build:** `python -m icharlotte_core.legal_research.local_corpus.build --source cap`
  first (smaller), confirm a real opposition run grounds + verifies offline, then run
  `--source cl` for the recent slice.
- **Manual verification in the app** (per CLAUDE.md "always test after developing"): launch
  iCharlotte, run Oppose-a-Motion on a real motion, confirm the research step completes in
  seconds with grounded cites and the draft contains verified case citations.

