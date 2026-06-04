# Firm Brief Library — Phase 1 (Authority Core) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build an incremental, refreshable index of the firm's sorted brief PDFs and reuse the case law they cite — preferred over, but merged with, the existing local-corpus research — in the `oppose_motion` and `generate_motion` drafting pipelines.

**Architecture:** New headless package `icharlotte_core/firm_briefs/` ingests the sorted PDF library (folder path → `(motion_type, side)`), harvests citations + propositions, embeds an issue-profile per brief, and stores everything in a SQLite DB + float16 vector sidecar. A `FirmAuthorityProvider` resolves harvested cites to opinion text (local corpus → live CourtListener fallback → unverified flag) and injects them as **preferred** candidates into the existing `research_argument()` rerank+verify path. Purely additive: absent the index, behavior is unchanged.

**Tech Stack:** Python 3, SQLite (FTS5), numpy memmap, fastembed BGE-small (reused from `legal_research/local_corpus/embedder.py`), pytest. Tests run with `C:\geminiterminal2\.venv\Scripts\python.exe`.

**Scope note:** Phase 1 is the headless authority core + ingestion. Style auto-selection (Phase 2) and the citation-panel/Workbench UI (Phase 3) are separate plans. Authority matching here is FTS5 keyword over stored propositions; per-proposition vector rerank is deferred.

---

### Task 1: Config + package skeleton

**Files:**
- Modify: `icharlotte_core/config.py` (after the `CASELAW_DATA_DIR` block, ~line 56)
- Create: `icharlotte_core/firm_briefs/__init__.py`
- Test: `tests/test_firm_briefs/test_config.py`

- [ ] **Step 1: Write the failing test**

```python
# tests/test_firm_briefs/test_config.py
import os
import importlib


def test_firm_briefs_config_present():
    config = importlib.import_module("icharlotte_core.config")
    assert isinstance(config.FIRM_BRIEFS_DATA_DIR, str)
    assert config.FIRM_BRIEFS_DATA_DIR  # non-empty
    assert isinstance(config.FIRM_BRIEFS_ROOTS, list)
    # Default seeds the 5800 library if present in cwd, but the value must be a list.
    assert all(isinstance(p, str) for p in config.FIRM_BRIEFS_ROOTS)


def test_firm_briefs_data_dir_env_override(monkeypatch):
    monkeypatch.setenv("FIRM_BRIEFS_DATA_DIR", os.path.join("X:", "fb"))
    config = importlib.reload(importlib.import_module("icharlotte_core.config"))
    assert config.FIRM_BRIEFS_DATA_DIR == os.path.join("X:", "fb")
```

- [ ] **Step 2: Run test to verify it fails**

Run: `.venv\Scripts\python.exe -m pytest tests/test_firm_briefs/test_config.py -v`
Expected: FAIL with `AttributeError: module ... has no attribute 'FIRM_BRIEFS_DATA_DIR'`

- [ ] **Step 3: Add config**

```python
# icharlotte_core/config.py  (insert after the CASELAW_DATA_DIR assignment)

# Firm brief sample library (sorted PDF folders) — authority reuse + style.
# Relocatable via FIRM_BRIEFS_DATA_DIR. Roots are firm-wide brief libraries.
FIRM_BRIEFS_DATA_DIR = os.environ.get(
    "FIRM_BRIEFS_DATA_DIR", os.path.join(os.getcwd(), ".gemini", "firm_briefs")
)


def _default_firm_briefs_roots() -> list:
    candidates = [
        os.path.join(os.getcwd(), "5800_AMTRUST_Pleadings_PDFs"),
    ]
    return [p for p in candidates if os.path.isdir(p)]


FIRM_BRIEFS_ROOTS = _default_firm_briefs_roots()
```

- [ ] **Step 4: Create the package marker**

```python
# icharlotte_core/firm_briefs/__init__.py
"""Firm brief sample library: authority reuse + style selection."""
```

Also create empty `tests/test_firm_briefs/__init__.py`.

- [ ] **Step 5: Run tests to verify they pass**

Run: `.venv\Scripts\python.exe -m pytest tests/test_firm_briefs/test_config.py -v`
Expected: PASS (2 passed)

- [ ] **Step 6: Commit**

```bash
git add icharlotte_core/config.py icharlotte_core/firm_briefs/__init__.py tests/test_firm_briefs/
git commit -m "feat(firm_briefs): config (data dir + roots) and package skeleton"
```

---

### Task 2: Folder → (motion_type, side) mapping

**Files:**
- Create: `icharlotte_core/firm_briefs/path_meta.py`
- Test: `tests/test_firm_briefs/test_path_meta.py`

- [ ] **Step 1: Write the failing test**

```python
# tests/test_firm_briefs/test_path_meta.py
from icharlotte_core.firm_briefs.path_meta import meta_for_path

ROOT = r"C:\lib\5800_AMTRUST_Pleadings_PDFs"


def test_moving_motion_folder():
    p = ROOT + r"\Motion - Summary Judgment\013 - Hall__msj.pdf"
    assert meta_for_path(p, ROOT) == ("msj", "moving")


def test_opposition_subfolder():
    p = ROOT + r"\Oppositions\Motion to Compel\008 - Rosas__opp.pdf"
    assert meta_for_path(p, ROOT) == ("compel", "opposition")


def test_reply_subfolder():
    p = ROOT + r"\Replies\Demurrer\072 - Forney__reply.pdf"
    assert meta_for_path(p, ROOT) == ("demurrer", "reply")


def test_pleading_folder():
    p = ROOT + r"\Pleadings - Answer\002 - Campos__answer.pdf"
    assert meta_for_path(p, ROOT) == ("answer", "pleading")


def test_ex_parte_is_moving():
    p = ROOT + r"\Ex Parte Applications\Continue Trial\x.pdf"
    assert meta_for_path(p, ROOT) == ("ex_parte", "moving")


def test_support_and_other_return_none():
    assert meta_for_path(ROOT + r"\_Support - Notices\x.pdf", ROOT) is None
    assert meta_for_path(ROOT + r"\_Other\x.pdf", ROOT) is None
```

- [ ] **Step 2: Run test to verify it fails**

Run: `.venv\Scripts\python.exe -m pytest tests/test_firm_briefs/test_path_meta.py -v`
Expected: FAIL with `ModuleNotFoundError: ... path_meta`

- [ ] **Step 3: Implement**

```python
# icharlotte_core/firm_briefs/path_meta.py
"""Map a sorted-library file path to (motion_type, side).

The library folder layout (built by the organize step) encodes both the motion
type and the procedural side, so ingestion needs no manual tagging.
Top-level folders: "Motion - X", "Motions - Other", "Ex Parte Applications",
"Oppositions" (+ per-type subfolders), "Replies" (+ subfolders),
"Pleadings - X". Anything under "_Support*" / "_Other" is not a brief.
"""
from __future__ import annotations

import os
from typing import Optional, Tuple

# Canonical type ids. Keys are lowercased folder labels (with the leading
# "motion - " / "pleadings - " prefix already stripped) → type id.
_TYPE_ALIASES = {
    "summary judgment": "msj",
    "msj-msa": "msj",
    "msj": "msj",
    "demurrer": "demurrer",
    "strike": "strike",
    "motion to strike": "strike",
    "compel": "compel",
    "motion to compel": "compel",
    "in limine": "in_limine",
    "quash": "quash",
    "motion to quash": "quash",
    "sanctions": "sanctions",
    "relieve counsel": "relieve_counsel",
    "continue trial": "continue_trial",
    "continue trial & preference": "continue_trial",
    "other": "other",
    "motions - other": "other",
    # pleadings
    "answer": "answer",
    "complaint": "complaint",
    "amended complaint": "amended_complaint",
    "cross-complaint": "cross_complaint",
    # leave/dismiss seen as opp/reply subfolders
    "motion for leave": "leave",
    "motion to dismiss": "dismiss",
    "set aside default": "set_aside_default",
    "protective order": "protective_order",
}


def _canon_type(label: str) -> str:
    key = (label or "").strip().lower()
    if key in _TYPE_ALIASES:
        return _TYPE_ALIASES[key]
    # Fall back to a slug so unknown-but-real types still group consistently.
    return key.replace(" ", "_").replace("&", "and").replace("--", "-").strip("_") or "other"


def meta_for_path(path: str, root: str) -> Optional[Tuple[str, str]]:
    try:
        rel = os.path.relpath(path, root)
    except ValueError:
        return None
    parts = [p for p in rel.replace("\\", "/").split("/") if p]
    if len(parts) < 2:
        return None
    top = parts[0]
    low = top.lower()
    if low.startswith("_support") or low == "_other":
        return None

    sub = parts[1] if len(parts) >= 3 else ""  # subfolder when file is nested

    if low == "oppositions":
        return (_canon_type(sub), "opposition")
    if low == "replies":
        return (_canon_type(sub), "reply")
    if low == "ex parte applications":
        return ("ex_parte", "moving")
    if low.startswith("motion - "):
        return (_canon_type(top[len("motion - "):]), "moving")
    if low == "motions - other":
        return ("other", "moving")
    if low.startswith("pleadings - "):
        return (_canon_type(top[len("pleadings - "):]), "pleading")
    return None
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `.venv\Scripts\python.exe -m pytest tests/test_firm_briefs/test_path_meta.py -v`
Expected: PASS (6 passed)

- [ ] **Step 5: Commit**

```bash
git add icharlotte_core/firm_briefs/path_meta.py tests/test_firm_briefs/test_path_meta.py
git commit -m "feat(firm_briefs): folder path -> (motion_type, side) mapping"
```

---

### Task 3: Citation harvesting

**Files:**
- Create: `icharlotte_core/firm_briefs/citation_harvest.py`
- Test: `tests/test_firm_briefs/test_citation_harvest.py`

Reuses `icharlotte_core.opposition.citation_parser.extract_citations`, which returns `Citation` objects with fields `kind` ("case"/"statute"/...), `case_name`, `reporter_citation`, `year`, `proposition`.

- [ ] **Step 1: Write the failing test**

```python
# tests/test_firm_briefs/test_citation_harvest.py
from icharlotte_core.firm_briefs.citation_harvest import harvest_cites, HarvestedCite

TEXT = (
    "Plaintiff failed to meet and confer before moving. "
    "A party must engage in a reasonable and good faith effort. "
    "Townsend v. Superior Court (1998) 61 Cal.App.4th 1431, 1438. "
    "The motion is therefore procedurally improper."
)


def test_harvests_case_cite_with_norm_and_proposition():
    cites = harvest_cites(TEXT)
    assert len(cites) == 1
    c = cites[0]
    assert isinstance(c, HarvestedCite)
    assert c.case_name.startswith("Townsend")
    assert c.reporter_citation == "61 Cal.App.4th 1431"
    assert c.year == "1998"
    assert c.norm_cite == "61cal.app.4th1431"
    assert "meet and confer" in c.proposition.lower() or "good faith" in c.proposition.lower()
    assert c.quoted_passage  # non-empty


def test_skips_statutes_in_phase1():
    cites = harvest_cites("See Code Civ. Proc. section 2031.310. Nothing else.")
    assert cites == []
```

- [ ] **Step 2: Run test to verify it fails**

Run: `.venv\Scripts\python.exe -m pytest tests/test_firm_briefs/test_citation_harvest.py -v`
Expected: FAIL with `ModuleNotFoundError`

- [ ] **Step 3: Implement**

```python
# icharlotte_core/firm_briefs/citation_harvest.py
"""Harvest case citations (with the proposition each supports) from brief text.

Thin wrapper over the opposition citation parser. Phase 1 keeps only case
cites; statutes are reused elsewhere (leginfo) and out of scope here.
"""
from __future__ import annotations

from dataclasses import dataclass
from typing import List


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
                case_name=getattr(c, "case_name", "") or "",
                reporter_citation=reporter,
                year=getattr(c, "year", "") or "",
                norm_cite=_norm(reporter),
                proposition=proposition,
                # Phase 1: the proposition sentence is the verify-only passage
                # used when neither the corpus nor CourtListener can confirm.
                quoted_passage=proposition,
            )
        )
    return out
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `.venv\Scripts\python.exe -m pytest tests/test_firm_briefs/test_citation_harvest.py -v`
Expected: PASS (2 passed). If `reporter_citation`/`norm_cite` differ from the parser's exact output, adjust the expected strings to match `extract_citations` (do not change the parser).

- [ ] **Step 5: Commit**

```bash
git add icharlotte_core/firm_briefs/citation_harvest.py tests/test_firm_briefs/test_citation_harvest.py
git commit -m "feat(firm_briefs): harvest case cites + propositions from brief text"
```

---

### Task 4: Issue-profile composition

**Files:**
- Create: `icharlotte_core/firm_briefs/profile.py`
- Test: `tests/test_firm_briefs/test_profile.py`

- [ ] **Step 1: Write the failing test**

```python
# tests/test_firm_briefs/test_profile.py
from icharlotte_core.firm_briefs.profile import extract_headings, compose_profile, profile_from_text

DOC = (
    "NOTICE OF MOTION\n"
    "I. PLAINTIFF FAILED TO MEET AND CONFER\n"
    "Some argument prose here that is not a heading.\n"
    "II. THE DISCOVERY CUTOFF HAS PASSED\n"
    "More prose.\n"
)


def test_extract_headings_picks_caps_lines():
    heads = extract_headings(DOC)
    assert any("MEET AND CONFER" in h for h in heads)
    assert any("DISCOVERY CUTOFF" in h for h in heads)
    assert "Some argument prose here that is not a heading." not in heads


def test_compose_profile_concatenates():
    prof = compose_profile("compel further responses", ["FAILED TO MEET AND CONFER"], ["cutoff passed"])
    assert "compel further responses" in prof
    assert "MEET AND CONFER" in prof
    assert "cutoff passed" in prof


def test_profile_from_text_nonempty():
    assert profile_from_text(DOC).strip()
```

- [ ] **Step 2: Run test to verify it fails**

Run: `.venv\Scripts\python.exe -m pytest tests/test_firm_briefs/test_profile.py -v`
Expected: FAIL with `ModuleNotFoundError`

- [ ] **Step 3: Implement**

```python
# icharlotte_core/firm_briefs/profile.py
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
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `.venv\Scripts\python.exe -m pytest tests/test_firm_briefs/test_profile.py -v`
Expected: PASS (3 passed)

- [ ] **Step 5: Commit**

```bash
git add icharlotte_core/firm_briefs/profile.py tests/test_firm_briefs/test_profile.py
git commit -m "feat(firm_briefs): issue-profile composition for embedding"
```

---

### Task 5: Embedding wrapper (reuse corpus embedder)

**Files:**
- Create: `icharlotte_core/firm_briefs/embedding.py`
- Test: `tests/test_firm_briefs/test_embedding.py`

- [ ] **Step 1: Write the failing test**

```python
# tests/test_firm_briefs/test_embedding.py
import numpy as np
from icharlotte_core.firm_briefs.embedding import get_embedder, EMBED_DIM
from icharlotte_core.legal_research.local_corpus.embedder import FakeEmbedder


def test_fake_embedder_shape():
    emb = FakeEmbedder(dim=EMBED_DIM)
    vecs = emb.encode(["meet and confer", "discovery cutoff"])
    assert vecs.shape == (2, EMBED_DIM)
    assert vecs.dtype == np.float32


def test_get_embedder_returns_object_with_encode():
    emb = get_embedder(fake=True)
    assert hasattr(emb, "encode")
    assert emb.dim == EMBED_DIM
```

- [ ] **Step 2: Run test to verify it fails**

Run: `.venv\Scripts\python.exe -m pytest tests/test_firm_briefs/test_embedding.py -v`
Expected: FAIL with `ModuleNotFoundError`

- [ ] **Step 3: Implement**

```python
# icharlotte_core/firm_briefs/embedding.py
"""Reuse the local-corpus fastembed embedder (BGE-small, 384-dim, no torch)."""
from __future__ import annotations

from icharlotte_core.legal_research.local_corpus.embedder import (
    FakeEmbedder,
    OnnxEmbedder,
    cosine_topk,
)

EMBED_DIM = 384


def get_embedder(*, fake: bool = False):
    """Return an embedder. ``fake=True`` for tests (deterministic, no model)."""
    if fake:
        return FakeEmbedder(dim=EMBED_DIM)
    return OnnxEmbedder(dim=EMBED_DIM)


__all__ = ["get_embedder", "cosine_topk", "FakeEmbedder", "EMBED_DIM"]
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `.venv\Scripts\python.exe -m pytest tests/test_firm_briefs/test_embedding.py -v`
Expected: PASS (2 passed)

- [ ] **Step 5: Commit**

```bash
git add icharlotte_core/firm_briefs/embedding.py tests/test_firm_briefs/test_embedding.py
git commit -m "feat(firm_briefs): embedding wrapper reusing corpus fastembed"
```

---

### Task 6: FirmBriefIndex — schema, upsert, thread-local connections

**Files:**
- Create: `icharlotte_core/firm_briefs/index.py`
- Test: `tests/test_firm_briefs/test_index.py`

The index uses one SQLite DB plus a float16 vector sidecar appended one row per
brief. **Connections are thread-local** — the research pipeline fans out over a
ThreadPoolExecutor and a single shared `sqlite3.Connection` raises cross-thread
errors that get swallowed into empty results.

- [ ] **Step 1: Write the failing test**

```python
# tests/test_firm_briefs/test_index.py
import os
import concurrent.futures
import numpy as np
import pytest

from icharlotte_core.firm_briefs.index import FirmBriefIndex
from icharlotte_core.firm_briefs.citation_harvest import HarvestedCite


def _vec(seed: int) -> np.ndarray:
    rng = np.random.default_rng(seed)
    v = rng.standard_normal(384).astype(np.float32)
    return v / np.linalg.norm(v)


@pytest.fixture
def idx(tmp_path):
    db = str(tmp_path / "fb.db")
    vec = str(tmp_path / "profiles.f16")
    index = FirmBriefIndex(db_path=db, vectors_path=vec)
    index.create_schema()
    return index


def test_upsert_and_has_current(idx):
    cites = [HarvestedCite(case_name="Townsend v. Superior Court",
                           reporter_citation="61 Cal.App.4th 1431", year="1998",
                           norm_cite="61cal.app.4th1431",
                           proposition="meet and confer is required",
                           quoted_passage="reasonable and good faith effort")]
    bid = idx.upsert_brief(path="p1.pdf", content_hash="h1", motion_type="compel",
                           side="opposition", heading="MEET AND CONFER",
                           profile="compel meet and confer", profile_vec=_vec(1),
                           char_len=5000, ocr_ratio=0.1, cites=cites)
    assert bid > 0
    assert idx.has_current("p1.pdf", "h1") is True
    assert idx.has_current("p1.pdf", "DIFFERENT") is False


def test_authority_candidates_keyword(idx):
    idx.upsert_brief(path="p1.pdf", content_hash="h1", motion_type="compel",
                     side="opposition", heading="", profile="x", profile_vec=_vec(1),
                     char_len=10, ocr_ratio=0.0,
                     cites=[HarvestedCite(case_name="Townsend v. Superior Court",
                                          reporter_citation="61 Cal.App.4th 1431",
                                          year="1998", norm_cite="61cal.app.4th1431",
                                          proposition="a party must meet and confer in good faith",
                                          quoted_passage="good faith")])
    hits = idx.authority_candidates("meet and confer good faith", motion_type="compel", limit=5)
    assert any("Townsend" in h["case_name"] for h in hits)
    # type filter excludes other types
    assert idx.authority_candidates("meet and confer", motion_type="msj", limit=5) == []


def test_upsert_replaces_on_rehash(idx):
    idx.upsert_brief(path="p1.pdf", content_hash="h1", motion_type="compel", side="moving",
                     heading="", profile="x", profile_vec=_vec(1), char_len=1, ocr_ratio=0.0,
                     cites=[HarvestedCite(reporter_citation="1 Cal.5th 1", norm_cite="1cal.5th1",
                                          proposition="p1")])
    idx.upsert_brief(path="p1.pdf", content_hash="h2", motion_type="compel", side="moving",
                     heading="", profile="x", profile_vec=_vec(2), char_len=1, ocr_ratio=0.0,
                     cites=[HarvestedCite(reporter_citation="2 Cal.5th 2", norm_cite="2cal.5th2",
                                          proposition="p2")])
    assert idx.has_current("p1.pdf", "h2")
    hits = idx.authority_candidates("p2", motion_type="compel", limit=5)
    assert any(h["norm_cite"] == "2cal.5th2" for h in hits)
    assert idx.authority_candidates("p1", motion_type="compel", limit=5) == []  # old cite gone


def test_mark_stale_excludes(idx):
    idx.upsert_brief(path="p1.pdf", content_hash="h1", motion_type="compel", side="moving",
                     heading="", profile="x", profile_vec=_vec(1), char_len=1, ocr_ratio=0.0,
                     cites=[HarvestedCite(reporter_citation="1 Cal.5th 1", norm_cite="1cal.5th1",
                                          proposition="stale point")])
    idx.mark_stale("p1.pdf")
    assert idx.authority_candidates("stale point", motion_type="compel", limit=5) == []


def test_thread_local_connections(idx):
    idx.upsert_brief(path="p1.pdf", content_hash="h1", motion_type="compel", side="moving",
                     heading="", profile="x", profile_vec=_vec(1), char_len=1, ocr_ratio=0.0,
                     cites=[HarvestedCite(reporter_citation="1 Cal.5th 1", norm_cite="1cal.5th1",
                                          proposition="threaded meet and confer")])

    def _q(_):
        return idx.authority_candidates("threaded meet and confer", motion_type="compel", limit=5)

    with concurrent.futures.ThreadPoolExecutor(max_workers=4) as pool:
        results = list(pool.map(_q, range(8)))
    assert all(len(r) >= 1 for r in results)  # no swallowed cross-thread errors
```

- [ ] **Step 2: Run test to verify it fails**

Run: `.venv\Scripts\python.exe -m pytest tests/test_firm_briefs/test_index.py -v`
Expected: FAIL with `ModuleNotFoundError`

- [ ] **Step 3: Implement**

```python
# icharlotte_core/firm_briefs/index.py
"""SQLite + float16 vector sidecar index over the firm brief library.

Connections are thread-local (WAL); the vector sidecar is an append-only
float16 file, one row per brief, addressed by briefs.vec_row.
"""
from __future__ import annotations

import os
import sqlite3
import threading
from typing import Any, List, Optional

import numpy as np

from .embedding import EMBED_DIM

_SCHEMA = """
CREATE TABLE IF NOT EXISTS briefs(
  id INTEGER PRIMARY KEY,
  path TEXT UNIQUE, content_hash TEXT,
  motion_type TEXT, side TEXT,
  heading TEXT, profile TEXT,
  vec_row INTEGER DEFAULT -1,
  char_len INTEGER DEFAULT 0, ocr_ratio REAL DEFAULT 0.0,
  ingested_at TEXT DEFAULT (datetime('now')),
  status TEXT DEFAULT 'ok'
);
CREATE TABLE IF NOT EXISTS citations(
  id INTEGER PRIMARY KEY,
  brief_id INTEGER REFERENCES briefs(id) ON DELETE CASCADE,
  case_name TEXT, reporter_cite TEXT, year TEXT, norm_cite TEXT,
  proposition TEXT, quoted_passage TEXT
);
CREATE INDEX IF NOT EXISTS ix_cit_norm ON citations(norm_cite);
CREATE INDEX IF NOT EXISTS ix_cit_brief ON citations(brief_id);
CREATE VIRTUAL TABLE IF NOT EXISTS citations_fts USING fts5(
  proposition, content='citations', content_rowid='id'
);
"""


class FirmBriefIndex:
    def __init__(self, *, db_path: str, vectors_path: str, embedder=None) -> None:
        self.db_path = db_path
        self.vectors_path = vectors_path
        self.embedder = embedder
        self._local = threading.local()
        os.makedirs(os.path.dirname(db_path) or ".", exist_ok=True)

    # -- connection -------------------------------------------------------
    def _conn(self) -> sqlite3.Connection:
        con = getattr(self._local, "con", None)
        if con is None:
            con = sqlite3.connect(self.db_path)
            con.row_factory = sqlite3.Row
            con.execute("PRAGMA journal_mode=WAL")
            con.execute("PRAGMA foreign_keys=ON")
            self._local.con = con
        return con

    def create_schema(self) -> None:
        con = self._conn()
        con.executescript(_SCHEMA)
        con.commit()

    # -- vector sidecar ---------------------------------------------------
    def _append_vector(self, vec: np.ndarray) -> int:
        v = np.asarray(vec, dtype=np.float16).reshape(EMBED_DIM)
        row = 0
        if os.path.exists(self.vectors_path):
            row = os.path.getsize(self.vectors_path) // (EMBED_DIM * 2)
        with open(self.vectors_path, "ab") as f:
            f.write(v.tobytes())
            f.flush()
            os.fsync(f.fileno())
        return int(row)

    def load_vectors(self) -> np.ndarray:
        if not os.path.exists(self.vectors_path) or os.path.getsize(self.vectors_path) == 0:
            return np.zeros((0, EMBED_DIM), dtype=np.float16)
        return np.memmap(self.vectors_path, dtype=np.float16, mode="r").reshape(-1, EMBED_DIM)

    # -- writes -----------------------------------------------------------
    def upsert_brief(self, *, path: str, content_hash: str, motion_type: str, side: str,
                     heading: str, profile: str, profile_vec, char_len: int, ocr_ratio: float,
                     cites: List[Any]) -> int:
        con = self._conn()
        # Append a fresh vector row (sidecar is append-only; old rows orphaned
        # until --compact). fsync sidecar BEFORE the DB commit (crash ordering).
        vec_row = self._append_vector(profile_vec)
        existing = con.execute("SELECT id FROM briefs WHERE path=?", (path,)).fetchone()
        if existing:
            bid = int(existing["id"])
            con.execute("DELETE FROM citations WHERE brief_id=?", (bid,))
            con.execute(
                "UPDATE briefs SET content_hash=?, motion_type=?, side=?, heading=?, "
                "profile=?, vec_row=?, char_len=?, ocr_ratio=?, status='ok', "
                "ingested_at=datetime('now') WHERE id=?",
                (content_hash, motion_type, side, heading, profile, vec_row,
                 char_len, ocr_ratio, bid),
            )
        else:
            cur = con.execute(
                "INSERT INTO briefs(path, content_hash, motion_type, side, heading, "
                "profile, vec_row, char_len, ocr_ratio) VALUES(?,?,?,?,?,?,?,?,?)",
                (path, content_hash, motion_type, side, heading, profile, vec_row,
                 char_len, ocr_ratio),
            )
            bid = int(cur.lastrowid)
        for c in cites:
            cur = con.execute(
                "INSERT INTO citations(brief_id, case_name, reporter_cite, year, "
                "norm_cite, proposition, quoted_passage) VALUES(?,?,?,?,?,?,?)",
                (bid, getattr(c, "case_name", ""), getattr(c, "reporter_citation", ""),
                 getattr(c, "year", ""), getattr(c, "norm_cite", ""),
                 getattr(c, "proposition", ""), getattr(c, "quoted_passage", "")),
            )
            con.execute("INSERT INTO citations_fts(rowid, proposition) VALUES(?,?)",
                        (cur.lastrowid, getattr(c, "proposition", "")))
        con.commit()
        return bid

    def mark_stale(self, path: str) -> None:
        con = self._conn()
        row = con.execute("SELECT id FROM briefs WHERE path=?", (path,)).fetchone()
        if row:
            con.execute("DELETE FROM citations WHERE brief_id=?", (int(row["id"]),))
            con.execute("UPDATE briefs SET status='stale' WHERE id=?", (int(row["id"]),))
            con.commit()

    # -- reads ------------------------------------------------------------
    def has_current(self, path: str, content_hash: str) -> bool:
        con = self._conn()
        row = con.execute(
            "SELECT 1 FROM briefs WHERE path=? AND content_hash=? AND status='ok'",
            (path, content_hash),
        ).fetchone()
        return row is not None

    def authority_candidates(self, proposition: str, *, motion_type: str,
                             limit: int = 8) -> List[dict]:
        con = self._conn()
        q = _fts_query(proposition)
        if not q:
            return []
        rows = con.execute(
            "SELECT c.case_name, c.reporter_cite, c.year, c.norm_cite, c.proposition, "
            "c.quoted_passage, b.path AS source_brief "
            "FROM citations_fts f "
            "JOIN citations c ON c.id = f.rowid "
            "JOIN briefs b ON b.id = c.brief_id "
            "WHERE citations_fts MATCH ? AND b.status='ok' AND b.motion_type=? "
            "ORDER BY bm25(citations_fts) LIMIT ?",
            (q, motion_type, limit),
        ).fetchall()
        return [dict(r) for r in rows]

    def stats(self) -> dict:
        con = self._conn()
        b = con.execute("SELECT COUNT(*) n FROM briefs WHERE status='ok'").fetchone()["n"]
        c = con.execute("SELECT COUNT(*) n FROM citations").fetchone()["n"]
        return {"briefs": b, "citations": c}


def _fts_query(text: str) -> str:
    import re
    toks = re.findall(r"[A-Za-z0-9]+", (text or "").lower())
    toks = [t for t in toks if len(t) > 2][:12]
    return " OR ".join(toks)
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `.venv\Scripts\python.exe -m pytest tests/test_firm_briefs/test_index.py -v`
Expected: PASS (5 passed)

- [ ] **Step 5: Commit**

```bash
git add icharlotte_core/firm_briefs/index.py tests/test_firm_briefs/test_index.py
git commit -m "feat(firm_briefs): FirmBriefIndex schema, upsert, FTS5 authority lookup, thread-local conns"
```

---

### Task 7: Incremental ingestion

**Files:**
- Create: `icharlotte_core/firm_briefs/ingest.py`
- Test: `tests/test_firm_briefs/test_ingest.py`

PDF text extraction reuses `icharlotte_core.document_processor`. To keep the test
hermetic (no real PDFs), `ingest_root` takes an injectable `extract_fn(path)->str`
(defaulting to the real extractor).

- [ ] **Step 1: Write the failing test**

```python
# tests/test_firm_briefs/test_ingest.py
import os
import numpy as np
from icharlotte_core.firm_briefs.index import FirmBriefIndex
from icharlotte_core.firm_briefs.embedding import get_embedder
from icharlotte_core.firm_briefs.ingest import ingest_root

ROOT_NAME = "5800_AMTRUST_Pleadings_PDFs"
SAMPLE = ("I. PLAINTIFF FAILED TO MEET AND CONFER\n"
          "A party must meet and confer in good faith. "
          "Townsend v. Superior Court (1998) 61 Cal.App.4th 1431, 1438.\n")


def _make_lib(tmp_path):
    root = tmp_path / ROOT_NAME / "Oppositions" / "Motion to Compel"
    root.mkdir(parents=True)
    f = root / "008 - Rosas__opp.pdf"
    f.write_text("placeholder")  # content irrelevant; extract_fn is injected
    return str(tmp_path / ROOT_NAME), str(f)


def test_ingest_indexes_one_brief(tmp_path):
    root, fpath = _make_lib(tmp_path)
    idx = FirmBriefIndex(db_path=str(tmp_path / "fb.db"),
                         vectors_path=str(tmp_path / "v.f16"))
    idx.create_schema()
    emb = get_embedder(fake=True)
    res = ingest_root(root, idx, emb, extract_fn=lambda p: SAMPLE)
    assert res["added"] == 1
    hits = idx.authority_candidates("meet and confer good faith", motion_type="compel", limit=5)
    assert any("Townsend" in h["case_name"] for h in hits)


def test_ingest_is_incremental(tmp_path):
    root, fpath = _make_lib(tmp_path)
    idx = FirmBriefIndex(db_path=str(tmp_path / "fb.db"),
                         vectors_path=str(tmp_path / "v.f16"))
    idx.create_schema()
    emb = get_embedder(fake=True)
    ingest_root(root, idx, emb, extract_fn=lambda p: SAMPLE)
    res2 = ingest_root(root, idx, emb, extract_fn=lambda p: SAMPLE)
    assert res2["added"] == 0 and res2["skipped"] == 1  # unchanged → skipped


def test_ingest_marks_removed_stale(tmp_path):
    root, fpath = _make_lib(tmp_path)
    idx = FirmBriefIndex(db_path=str(tmp_path / "fb.db"),
                         vectors_path=str(tmp_path / "v.f16"))
    idx.create_schema()
    emb = get_embedder(fake=True)
    ingest_root(root, idx, emb, extract_fn=lambda p: SAMPLE)
    os.remove(fpath)
    res = ingest_root(root, idx, emb, extract_fn=lambda p: SAMPLE)
    assert res["staled"] == 1
    assert idx.authority_candidates("meet and confer", motion_type="compel", limit=5) == []
```

- [ ] **Step 2: Run test to verify it fails**

Run: `.venv\Scripts\python.exe -m pytest tests/test_firm_briefs/test_ingest.py -v`
Expected: FAIL with `ModuleNotFoundError`

- [ ] **Step 3: Implement**

```python
# icharlotte_core/firm_briefs/ingest.py
"""Incrementally ingest a sorted firm-brief library root into a FirmBriefIndex."""
from __future__ import annotations

import hashlib
import os
from typing import Callable, Optional

from .citation_harvest import harvest_cites
from .path_meta import meta_for_path
from .profile import extract_headings, compose_profile, profile_from_text


def content_hash(path: str) -> str:
    try:
        st = os.stat(path)
    except OSError:
        return ""
    h = hashlib.sha1()
    h.update(f"{os.path.abspath(path)}|{st.st_mtime_ns}|{st.st_size}".encode("utf-8"))
    return h.hexdigest()


def _default_extract(path: str) -> str:
    from icharlotte_core.document_processor import DocumentProcessor
    try:
        return DocumentProcessor().extract_text(path) or ""
    except Exception:
        return ""


def _ocr_ratio(text: str) -> float:
    # Crude noise signal: share of non-ASCII / replacement chars in the text.
    if not text:
        return 1.0
    bad = sum(1 for ch in text if ord(ch) > 0x2122 or ch == "�")
    return bad / len(text)


def ingest_root(root: str, index, embedder, *,
                extract_fn: Optional[Callable[[str], str]] = None,
                on_progress: Optional[Callable[[str], None]] = None) -> dict:
    extract_fn = extract_fn or _default_extract
    added = updated = skipped = failed = 0
    seen_paths: set[str] = set()

    for dirpath, _dirs, files in os.walk(root):
        for name in files:
            if not name.lower().endswith(".pdf"):
                continue
            path = os.path.join(dirpath, name)
            meta = meta_for_path(path, root)
            if meta is None:
                continue  # _Support / _Other / unrecognized
            motion_type, side = meta
            seen_paths.add(os.path.abspath(path))
            h = content_hash(path)
            if index.has_current(path, h):
                skipped += 1
                continue
            text = extract_fn(path)
            if not text.strip():
                failed += 1
                continue
            cites = harvest_cites(text)
            headings = extract_headings(text)
            profile = compose_profile("", headings, [c.proposition for c in cites]) \
                or profile_from_text(text)
            vec = embedder.encode([profile])[0]
            existed = index.has_current(path, "")  # path present (any hash)?
            index.upsert_brief(
                path=path, content_hash=h, motion_type=motion_type, side=side,
                heading=headings[0] if headings else "", profile=profile,
                profile_vec=vec, char_len=len(text), ocr_ratio=_ocr_ratio(text),
                cites=cites,
            )
            updated += 1 if existed else 0
            added += 0 if existed else 1
            if on_progress:
                on_progress(f"  indexed {name} ({motion_type}/{side}, {len(cites)} cites)")

    # Mark DB briefs under this root that no longer exist on disk as stale.
    staled = 0
    con = index._conn()
    rows = con.execute(
        "SELECT path FROM briefs WHERE status='ok' AND path LIKE ?",
        (os.path.join(os.path.abspath(root), "") + "%",),
    ).fetchall()
    for r in rows:
        if os.path.abspath(r["path"]) not in seen_paths and not os.path.exists(r["path"]):
            index.mark_stale(r["path"])
            staled += 1

    return {"added": added, "updated": updated, "skipped": skipped,
            "failed": failed, "staled": staled}
```

Note: `has_current(path, "")` is used as a path-existence probe; since no real
hash is `""`, it returns False, so `existed` reflects only the same-hash case.
To detect a true update vs add, query the DB directly instead:

```python
            existed = bool(index._conn().execute(
                "SELECT 1 FROM briefs WHERE path=?", (path,)).fetchone())
```
Use this line in place of the `has_current(path, "")` line above.

- [ ] **Step 4: Run tests to verify they pass**

Run: `.venv\Scripts\python.exe -m pytest tests/test_firm_briefs/test_ingest.py -v`
Expected: PASS (3 passed)

- [ ] **Step 5: Commit**

```bash
git add icharlotte_core/firm_briefs/ingest.py tests/test_firm_briefs/test_ingest.py
git commit -m "feat(firm_briefs): incremental ingestion (hash-skip, upsert, stale-on-remove)"
```

---

### Task 8: RetrievedAuthority — provenance fields

**Files:**
- Modify: `icharlotte_core/opposition/models.py` (RetrievedAuthority dataclass, ~line 49-60)
- Test: `tests/test_firm_briefs/test_retrieved_authority_fields.py`

- [ ] **Step 1: Write the failing test**

```python
# tests/test_firm_briefs/test_retrieved_authority_fields.py
from icharlotte_core.opposition.models import RetrievedAuthority


def test_new_provenance_fields_default_safely():
    ra = RetrievedAuthority()
    assert ra.source == "corpus"
    assert ra.verification == "local"
    assert ra.source_brief == ""
    assert ra.alternatives == []


def test_alternatives_are_independent_lists():
    a, b = RetrievedAuthority(), RetrievedAuthority()
    a.alternatives.append(b)
    assert b.alternatives == []  # no shared default
```

- [ ] **Step 2: Run test to verify it fails**

Run: `.venv\Scripts\python.exe -m pytest tests/test_firm_briefs/test_retrieved_authority_fields.py -v`
Expected: FAIL with `AttributeError: ... 'source'`

- [ ] **Step 3: Implement** — add fields to the `RetrievedAuthority` dataclass (after `latest_citing_year`):

```python
    # Provenance (firm-brief authority reuse). Defaults keep corpus-only behavior.
    source: str = "corpus"            # "firm" | "corpus"
    verification: str = "local"       # "local" | "courtlistener" | "unverified_firm"
    source_brief: str = ""            # path/label of the firm brief this came from
    alternatives: list = field(default_factory=list)  # corpus options for same point
```

Ensure `from dataclasses import field` is imported at the top of `models.py` (it
already is — `MotionMetadata` uses `field`).

- [ ] **Step 4: Run tests to verify they pass**

Run: `.venv\Scripts\python.exe -m pytest tests/test_firm_briefs/test_retrieved_authority_fields.py -v`
Expected: PASS (2 passed)

- [ ] **Step 5: Run the opposition models/parser regression**

Run: `.venv\Scripts\python.exe -m pytest tests/test_opposition -q` (or the opposition test dir)
Expected: PASS (no regressions from the additive fields)

- [ ] **Step 6: Commit**

```bash
git add icharlotte_core/opposition/models.py tests/test_firm_briefs/test_retrieved_authority_fields.py
git commit -m "feat(opposition): RetrievedAuthority provenance fields (source/verification/alternatives)"
```

---

### Task 9: FirmAuthorityProvider — resolve cites to candidates

**Files:**
- Create: `icharlotte_core/firm_briefs/provider.py`
- Test: `tests/test_firm_briefs/test_provider.py`

The provider turns harvested cites into **candidate dicts** shaped exactly like
the ones `argument_research._run` builds (`cluster_id, case_name, citation, year,
text, opinion_url`) plus provenance keys (`source, verification, source_brief,
quoted_passage`). Resolution order: local corpus opinion text → live CL verify →
unverified flag.

- [ ] **Step 1: Write the failing test**

```python
# tests/test_firm_briefs/test_provider.py
from icharlotte_core.firm_briefs.provider import FirmAuthorityProvider


class FakeIndex:
    def __init__(self, rows): self._rows = rows
    def authority_candidates(self, proposition, *, motion_type, limit=8):
        return self._rows


class FakeCorpus:
    def __init__(self, texts): self._texts = texts  # norm_cite -> (uid, text)
    def lookup_by_citation(self, citation):
        norm = (citation or "").replace(" ", "").lower()
        if norm in self._texts:
            uid, _ = self._texts[norm]
            return {"case_uid": uid, "citation": citation}
        return None
    def get_opinion_text(self, uid):
        for _norm, (u, text) in self._texts.items():
            if u == uid:
                return text
        return None


ROW = {"case_name": "Townsend v. Superior Court", "reporter_cite": "61 Cal.App.4th 1431",
       "year": "1998", "norm_cite": "61cal.app.4th1431",
       "proposition": "meet and confer required", "quoted_passage": "good faith effort",
       "source_brief": "Oppositions/Motion to Compel/x.pdf"}


def test_resolves_via_corpus_local():
    idx = FakeIndex([ROW])
    corpus = FakeCorpus({"61cal.app.4th1431": ("cap:1", "... reasonable and good faith effort ...")})
    prov = FirmAuthorityProvider(idx, corpus, cl_client=None)
    cands = prov.candidates_for("meet and confer", motion_type="compel", side="opposition")
    assert len(cands) == 1
    c = cands[0]
    assert c["source"] == "firm"
    assert c["verification"] == "local"
    assert c["cluster_id"] == "cap:1"
    assert "good faith" in c["text"]
    assert c["source_brief"].endswith("x.pdf")


def test_unverified_when_not_in_corpus_and_no_cl():
    idx = FakeIndex([ROW])
    corpus = FakeCorpus({})  # not found
    prov = FirmAuthorityProvider(idx, corpus, cl_client=None)
    cands = prov.candidates_for("meet and confer", motion_type="compel", side="opposition")
    assert len(cands) == 1
    c = cands[0]
    assert c["verification"] == "unverified_firm"
    assert c["text"] == ""  # no opinion text → handled specially downstream
    assert c["passage"] == "good faith effort"


def test_cl_fallback_verifies():
    idx = FakeIndex([ROW])
    corpus = FakeCorpus({})

    class FakeCL:
        def get_opinion_text(self, cite):
            return "court text mentioning good faith effort"
        def lookup_by_citation(self, cite):
            return {"case_uid": "cl:9"}
    prov = FirmAuthorityProvider(idx, corpus, cl_client=FakeCL())
    cands = prov.candidates_for("meet and confer", motion_type="compel", side="opposition")
    assert cands[0]["verification"] == "courtlistener"
    assert cands[0]["text"]
```

- [ ] **Step 2: Run test to verify it fails**

Run: `.venv\Scripts\python.exe -m pytest tests/test_firm_briefs/test_provider.py -v`
Expected: FAIL with `ModuleNotFoundError`

- [ ] **Step 3: Implement**

```python
# icharlotte_core/firm_briefs/provider.py
"""Turn harvested firm cites into preferred research candidates.

Resolution: local corpus opinion text -> live CourtListener -> unverified flag.
Returned dicts match argument_research._run's candidate shape plus provenance.
"""
from __future__ import annotations

import logging
from typing import Any, List, Optional

logger = logging.getLogger(__name__)


class FirmAuthorityProvider:
    def __init__(self, index, corpus, cl_client: Optional[Any] = None) -> None:
        self.index = index
        self.corpus = corpus
        self.cl_client = cl_client

    def candidates_for(self, proposition: str, *, motion_type: str, side: str,
                       limit: int = 6) -> List[dict]:
        try:
            rows = self.index.authority_candidates(
                proposition, motion_type=motion_type, limit=limit)
        except Exception:
            logger.warning("firm authority lookup failed", exc_info=True)
            return []
        out: List[dict] = []
        for r in rows:
            out.append(self._resolve(r))
        return out

    def _resolve(self, r: dict) -> dict:
        cite = r.get("reporter_cite", "")
        base = {
            "case_name": r.get("case_name", ""),
            "citation": cite,
            "year": r.get("year", ""),
            "opinion_url": "",
            "source": "firm",
            "source_brief": r.get("source_brief", ""),
            "passage": r.get("quoted_passage", ""),
            "proposition": r.get("proposition", ""),
        }
        # 1) local corpus
        try:
            hit = self.corpus.lookup_by_citation(cite) if self.corpus else None
        except Exception:
            hit = None
        if hit:
            uid = str(hit.get("case_uid") or hit.get("cluster_id") or "")
            text = ""
            try:
                text = self.corpus.get_opinion_text(uid) or ""
            except Exception:
                text = ""
            if text:
                base.update({"cluster_id": uid, "text": text, "verification": "local"})
                return base
        # 2) live CourtListener fallback
        if self.cl_client is not None:
            try:
                hit2 = self.cl_client.lookup_by_citation(cite)
                uid2 = str((hit2 or {}).get("case_uid") or "")
                text2 = self.cl_client.get_opinion_text(cite) or ""
            except Exception:
                uid2, text2 = "", ""
            if text2:
                base.update({"cluster_id": uid2 or ("cl:" + cite),
                             "text": text2, "verification": "courtlistener"})
                return base
        # 3) unverified — keep, flagged; no opinion text
        base.update({"cluster_id": "firm:" + (r.get("norm_cite", "") or cite),
                     "text": "", "verification": "unverified_firm"})
        return base
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `.venv\Scripts\python.exe -m pytest tests/test_firm_briefs/test_provider.py -v`
Expected: PASS (3 passed)

- [ ] **Step 5: Commit**

```bash
git add icharlotte_core/firm_briefs/provider.py tests/test_firm_briefs/test_provider.py
git commit -m "feat(firm_briefs): FirmAuthorityProvider (corpus -> CL -> unverified resolution)"
```

---

### Task 10: Inject firm candidates into research_argument (prefer-firm + flag-both)

**Files:**
- Modify: `icharlotte_core/opposition/argument_research.py`
  (`research_argument`, `research_arguments`, and `_run` inside `research_argument`)
- Test: `tests/test_firm_briefs/test_research_injection.py`

Add an optional `firm_provider` parameter. Inside `_run`, prepend firm candidates
(with `text`) to the corpus candidate dicts so the rerank sees them first; mark
each returned authority's `source`/`verification`/`source_brief`; attach corpus
selections for the same argument as `alternatives`; and append `unverified_firm`
firm cites (no text) directly as flagged authorities.

- [ ] **Step 1: Write the failing test**

```python
# tests/test_firm_briefs/test_research_injection.py
from icharlotte_core.opposition.argument_research import research_argument


class StubCorpusClient:
    """Minimal cl_client: returns no corpus search hits (isolate firm path)."""
    def search_opinions(self, q, *, semantic=False, max_results=20, published_only=True):
        return []
    def get_opinion_text(self, uid):
        return ""
    def get_authority_signals(self, uid):
        return {}


class StubProvider:
    def __init__(self, cands): self._c = cands
    def candidates_for(self, proposition, *, motion_type, side, limit=6):
        return self._c


def _llm(_sys, _user):
    # rerank reply selecting the firm candidate with a verbatim passage.
    return '{"selections":[{"id":"cap:1","passage":"good faith effort","supports":"meet and confer required"}]}'


def test_firm_local_candidate_selected_and_tagged():
    firm = [{"cluster_id": "cap:1", "case_name": "Townsend v. Superior Court",
             "citation": "61 Cal.App.4th 1431", "year": "1998",
             "text": "a reasonable and good faith effort", "opinion_url": "",
             "source": "firm", "verification": "local",
             "source_brief": "x.pdf", "passage": "good faith effort",
             "proposition": "meet and confer required"}]
    out = research_argument(
        "Plaintiff failed to meet and confer",
        cl_client=StubCorpusClient(), query_llm=lambda s, u: '{"queries":[]}',
        rerank_llm=_llm, motion_type="compel", side="opposition",
        firm_provider=StubProvider(firm),
    )
    assert len(out) == 1
    assert out[0].source == "firm"
    assert out[0].verification == "local"
    assert out[0].source_brief == "x.pdf"


def test_unverified_firm_cite_appended_flagged():
    firm = [{"cluster_id": "firm:1", "case_name": "Smith v. Jones",
             "citation": "999 F.3d 1", "year": "2024", "text": "",
             "opinion_url": "", "source": "firm", "verification": "unverified_firm",
             "source_brief": "y.pdf", "passage": "federal rule applies",
             "proposition": "federal rule applies"}]
    out = research_argument(
        "Federal standard governs",
        cl_client=StubCorpusClient(), query_llm=lambda s, u: '{"queries":[]}',
        rerank_llm=lambda s, u: '{"selections":[]}', motion_type="compel",
        side="opposition", firm_provider=StubProvider(firm),
    )
    assert len(out) == 1
    assert out[0].verification == "unverified_firm"
    assert out[0].case_name == "Smith v. Jones"
    assert out[0].passage == "federal rule applies"
```

- [ ] **Step 2: Run test to verify it fails**

Run: `.venv\Scripts\python.exe -m pytest tests/test_firm_briefs/test_research_injection.py -v`
Expected: FAIL with `TypeError: research_argument() got an unexpected keyword argument 'motion_type'`

- [ ] **Step 3: Implement** — update `research_argument` signature and body.

Change the signature to add (keep all existing params):
```python
def research_argument(
    argument: str,
    *,
    cl_client,
    query_llm: LLMCallback,
    rerank_llm: LLMCallback,
    argument_id: str = "",
    max_candidates: int = 20,
    fetch_top: int = 8,
    cache_dir: str | None = None,
    firm_provider=None,
    motion_type: str = "",
    side: str = "",
) -> list[RetrievedAuthority]:
```

Inside, fetch firm candidates once at the top of `research_argument` (before `_run`):
```python
    firm_cands: list[dict] = []
    if firm_provider is not None and (motion_type or side):
        try:
            firm_cands = firm_provider.candidates_for(
                argument, motion_type=motion_type, side=side) or []
        except Exception:
            logger.warning("firm provider failed", exc_info=True)
            firm_cands = []
    firm_with_text = [c for c in firm_cands if c.get("text")]
    firm_unverified = [c for c in firm_cands if not c.get("text")]
    firm_by_id = {str(c.get("cluster_id")): c for c in firm_with_text}
```

In `_run`, prepend firm candidates (with text) ahead of the corpus `cand_dicts`
before calling `select_authorities`:
```python
        cand_dicts = [
            {
                "cluster_id": c.get("cluster_id"), "case_name": c.get("case_name", ""),
                "citation": c.get("citation", ""), "year": c.get("year", ""),
                "text": c.get("text", ""), "opinion_url": c.get("opinion_url", ""),
            }
            for c in firm_with_text
        ] + cand_dicts
```

After `selected = _run(queries)` (and the broaden fallback), tag provenance and
attach alternatives:
```python
    firm_selected, corpus_selected = [], []
    for ra in selected:
        fc = firm_by_id.get(ra.cluster_id)
        if fc:
            ra.source = "firm"
            ra.verification = fc.get("verification", "local")
            ra.source_brief = fc.get("source_brief", "")
            firm_selected.append(ra)
        else:
            corpus_selected.append(ra)
    # prefer-firm + flag-both: firm leads; corpus picks become alternatives.
    for ra in firm_selected:
        ra.alternatives = list(corpus_selected)
    # Append unverified firm cites (no opinion text → bypass the verbatim gate).
    for c in firm_unverified:
        firm_selected.append(RetrievedAuthority(
            argument_id=argument_id, argument_text=argument,
            cluster_id=str(c.get("cluster_id", "")), case_name=c.get("case_name", ""),
            citation=c.get("citation", ""), year=str(c.get("year", "")),
            supports=c.get("proposition", ""), passage=c.get("passage", ""),
            opinion_url="", source="firm", verification="unverified_firm",
            source_brief=c.get("source_brief", ""),
        ))
    selected = firm_selected + corpus_selected
```

Insert this block BEFORE the `get_authority_signals` stamping loop (so signals are
stamped on the final set). Then thread `firm_provider`, `motion_type`, `side`
through `research_arguments`:
```python
def research_arguments(
    arguments: list[str],
    *,
    cl_client,
    query_llm: LLMCallback,
    rerank_llm: LLMCallback,
    max_workers: int = 4,
    on_progress: ProgressCallback | None = None,
    cache_dir: str | None = None,
    firm_provider=None,
    motion_type: str = "",
    side: str = "",
) -> list[RetrievedAuthority]:
```
and in `_one`, pass them through:
```python
        result = research_argument(
            arg, cl_client=cl_client, query_llm=query_llm, rerank_llm=rerank_llm,
            argument_id=f"arg-{idx}", cache_dir=cache_dir,
            firm_provider=firm_provider, motion_type=motion_type, side=side,
        )
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `.venv\Scripts\python.exe -m pytest tests/test_firm_briefs/test_research_injection.py -v`
Expected: PASS (2 passed)

- [ ] **Step 5: Run the opposition research regression**

Run: `.venv\Scripts\python.exe -m pytest tests/test_opposition -q -k "research or argument"`
Expected: PASS (existing calls omit `firm_provider`/`motion_type` → unchanged behavior)

- [ ] **Step 6: Commit**

```bash
git add icharlotte_core/opposition/argument_research.py tests/test_firm_briefs/test_research_injection.py
git commit -m "feat(opposition): inject preferred firm authorities into research_argument (prefer-firm, flag-both)"
```

---

### Task 11: CLI ingest entrypoint + factory

**Files:**
- Create: `icharlotte_core/firm_briefs/__main__.py`
- Create: `icharlotte_core/firm_briefs/factory.py`
- Test: `tests/test_firm_briefs/test_factory.py`

`factory.py` centralizes path/availability logic (mirrors the corpus's
`_make_local_corpus`) so the wizard pages and CLI share one place.

- [ ] **Step 1: Write the failing test**

```python
# tests/test_firm_briefs/test_factory.py
import os
from icharlotte_core.firm_briefs import factory


def test_paths_under_data_dir(monkeypatch, tmp_path):
    monkeypatch.setattr(factory, "DATA_DIR", str(tmp_path))
    db, vec = factory.index_paths()
    assert db.startswith(str(tmp_path))
    assert vec.startswith(str(tmp_path))


def test_available_false_when_absent(monkeypatch, tmp_path):
    monkeypatch.setattr(factory, "DATA_DIR", str(tmp_path))
    assert factory.index_available() is False


def test_make_index_none_when_unavailable(monkeypatch, tmp_path):
    monkeypatch.setattr(factory, "DATA_DIR", str(tmp_path))
    assert factory.make_index() is None
```

- [ ] **Step 2: Run test to verify it fails**

Run: `.venv\Scripts\python.exe -m pytest tests/test_firm_briefs/test_factory.py -v`
Expected: FAIL with `ModuleNotFoundError`

- [ ] **Step 3: Implement**

```python
# icharlotte_core/firm_briefs/factory.py
"""Shared paths + availability + index construction for firm-brief features."""
from __future__ import annotations

import os
from typing import Optional

from icharlotte_core import config

DATA_DIR = config.FIRM_BRIEFS_DATA_DIR


def index_paths():
    return (os.path.join(DATA_DIR, "firm_briefs.db"),
            os.path.join(DATA_DIR, "profiles.f16"))


def index_available() -> bool:
    db, vec = index_paths()
    return os.path.exists(db) and os.path.exists(vec)


def make_index(*, embedder=None):
    if not index_available():
        return None
    from .index import FirmBriefIndex
    db, vec = index_paths()
    return FirmBriefIndex(db_path=db, vectors_path=vec, embedder=embedder)
```

```python
# icharlotte_core/firm_briefs/__main__.py
"""CLI: python -m icharlotte_core.firm_briefs --root <path> [--root ...]"""
from __future__ import annotations

import argparse
import os

from icharlotte_core import config
from .factory import index_paths, DATA_DIR
from .index import FirmBriefIndex
from .embedding import get_embedder
from .ingest import ingest_root


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="Ingest firm brief libraries.")
    ap.add_argument("--root", action="append", default=[], help="library root (repeatable)")
    ap.add_argument("--fake-embed", action="store_true", help="deterministic embedder (tests)")
    args = ap.parse_args(argv)

    roots = args.root or config.FIRM_BRIEFS_ROOTS
    if not roots:
        print("No roots given and config.FIRM_BRIEFS_ROOTS is empty.")
        return 2
    os.makedirs(DATA_DIR, exist_ok=True)
    db, vec = index_paths()
    index = FirmBriefIndex(db_path=db, vectors_path=vec)
    index.create_schema()
    embedder = get_embedder(fake=args.fake_embed)
    totals = {"added": 0, "updated": 0, "skipped": 0, "failed": 0, "staled": 0}
    for root in roots:
        print(f"Ingesting {root} ...")
        res = ingest_root(root, index, embedder, on_progress=lambda m: print(m))
        for k in totals:
            totals[k] += res.get(k, 0)
        print(f"  {res}")
    print(f"DONE {totals}  stats={index.stats()}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `.venv\Scripts\python.exe -m pytest tests/test_firm_briefs/test_factory.py -v`
Expected: PASS (3 passed)

- [ ] **Step 5: Commit**

```bash
git add icharlotte_core/firm_briefs/factory.py icharlotte_core/firm_briefs/__main__.py tests/test_firm_briefs/test_factory.py
git commit -m "feat(firm_briefs): shared factory + CLI ingest entrypoint"
```

---

### Task 12: Wire the provider into oppose_motion (guarded, additive)

**Files:**
- Modify: `icharlotte_core/ui/wizard/pages/oppose_motion_page.py`
  (the worker block ~line 621-665 where `research_arguments(...)` is called)
- Test: `tests/test_firm_briefs/test_oppose_wiring.py`

Build a `FirmAuthorityProvider` from the factory (None when the index isn't
built), pass it plus `motion_type`/`side="opposition"` into `research_arguments`.
When the provider is None, behavior is identical to today.

- [ ] **Step 1: Write the failing test** (unit-tests a small helper so we don't run the whole Qt worker)

```python
# tests/test_firm_briefs/test_oppose_wiring.py
from icharlotte_core.ui.wizard.pages import oppose_motion_page as omp


def test_make_firm_provider_none_when_index_absent(monkeypatch):
    monkeypatch.setattr(omp, "_make_firm_provider", omp._make_firm_provider)
    # With no built index, factory.make_index() returns None → provider None.
    prov = omp._make_firm_provider(corpus=None)
    assert prov is None


def test_make_firm_provider_builds_when_index_present(monkeypatch):
    class FakeIndex: ...
    monkeypatch.setattr("icharlotte_core.firm_briefs.factory.make_index",
                        lambda **k: FakeIndex())
    prov = omp._make_firm_provider(corpus="CORPUS")
    assert prov is not None
    assert prov.corpus == "CORPUS"
```

- [ ] **Step 2: Run test to verify it fails**

Run: `.venv\Scripts\python.exe -m pytest tests/test_firm_briefs/test_oppose_wiring.py -v`
Expected: FAIL with `AttributeError: module ... has no attribute '_make_firm_provider'`

- [ ] **Step 3: Implement** — add a module-level helper near `_make_local_corpus` in `oppose_motion_page.py`:

```python
def _make_firm_provider(corpus):
    """Build a FirmAuthorityProvider if the firm-brief index is built, else None.

    cl_client is the live CourtListener fallback for firm cites not in the local
    corpus; reuse the same token the research path uses.
    """
    try:
        from icharlotte_core.firm_briefs import factory
        index = factory.make_index()
        if index is None:
            return None
        from icharlotte_core.firm_briefs.provider import FirmAuthorityProvider
        cl = None
        token = os.environ.get("COURTLISTENER_API_TOKEN", "").strip()
        if token:
            from icharlotte_core.legal_research.sources.courtlistener import CourtListenerClient
            cl = CourtListenerClient(token)
        return FirmAuthorityProvider(index, corpus, cl_client=cl)
    except Exception:
        return None
```

Then in the worker, where `research_arguments(... cl_client=corpus ...)` is called
(the local-corpus branch, ~line 640), add the provider + type/side:

```python
                firm_provider = _make_firm_provider(corpus)
                if firm_provider is not None:
                    self.progress.emit("  Firm brief library active (preferring your prior authorities).")
                retrieved = research_arguments(
                    research_targets,
                    cl_client=corpus,
                    query_llm=make_llm("research_queries"),
                    rerank_llm=make_llm("rerank_select"),
                    max_workers=2,
                    on_progress=self.progress.emit,
                    cache_dir=opinion_cache,
                    firm_provider=firm_provider,
                    motion_type=metadata.motion_type,
                    side="opposition",
                )
```

(Leave the CourtListener-only fallback branch unchanged — the firm index still
needs the local corpus to resolve cites, so it's only wired into the corpus branch.)

- [ ] **Step 4: Run tests to verify they pass**

Run: `.venv\Scripts\python.exe -m pytest tests/test_firm_briefs/test_oppose_wiring.py -v`
Expected: PASS (2 passed). Note: importing `oppose_motion_page` requires PySide6;
if collection errors on a running app instance, stop iCharlotte first (known
quirk) and re-run.

- [ ] **Step 5: Run the wizard regression**

Run: `.venv\Scripts\python.exe -m pytest tests/test_wizard -q`
Expected: PASS (additive change; existing flows unaffected).

- [ ] **Step 6: Commit**

```bash
git add icharlotte_core/ui/wizard/pages/oppose_motion_page.py tests/test_firm_briefs/test_oppose_wiring.py
git commit -m "feat(oppose_motion): wire firm authority provider into research (guarded, additive)"
```

---

### Task 13: Wire the provider into generate_motion + full regression

**Files:**
- Modify: `icharlotte_core/ui/wizard/pages/generate_motion_page.py` (its `research_arguments` call site; mirror Task 12)
- Test: reuse `tests/test_firm_briefs/test_oppose_wiring.py` pattern in `tests/test_firm_briefs/test_generate_wiring.py`

- [ ] **Step 1: Write the failing test**

```python
# tests/test_firm_briefs/test_generate_wiring.py
from icharlotte_core.ui.wizard.pages import generate_motion_page as gmp


def test_generate_make_firm_provider_present(monkeypatch):
    class FakeIndex: ...
    monkeypatch.setattr("icharlotte_core.firm_briefs.factory.make_index",
                        lambda **k: FakeIndex())
    prov = gmp._make_firm_provider(corpus="C")
    assert prov is not None and prov.corpus == "C"
```

- [ ] **Step 2: Run test to verify it fails**

Run: `.venv\Scripts\python.exe -m pytest tests/test_firm_briefs/test_generate_wiring.py -v`
Expected: FAIL with `AttributeError: ... '_make_firm_provider'`

- [ ] **Step 3: Implement** — add the same `_make_firm_provider` helper to
`generate_motion_page.py` (identical body to Task 12), and at its
`research_arguments(...)` call pass `firm_provider=_make_firm_provider(corpus)`,
`motion_type=metadata.motion_type`, `side="moving"`.

- [ ] **Step 4: Run test to verify it passes**

Run: `.venv\Scripts\python.exe -m pytest tests/test_firm_briefs/test_generate_wiring.py -v`
Expected: PASS

- [ ] **Step 5: Full feature + regression sweep**

Run: `.venv\Scripts\python.exe -m pytest tests/test_firm_briefs tests/test_opposition tests/test_wizard -q`
Expected: PASS (all firm_briefs tests + no opposition/wizard regressions).

- [ ] **Step 6: Commit**

```bash
git add icharlotte_core/ui/wizard/pages/generate_motion_page.py tests/test_firm_briefs/test_generate_wiring.py
git commit -m "feat(generate_motion): wire firm authority provider into research (side=moving)"
```

---

## Build the real index (manual, after Task 13)

Not a code task — run once to populate the index from the sorted 5800 library:

```bash
.venv\Scripts\python.exe -m icharlotte_core.firm_briefs --root "C:\geminiterminal2\5800_AMTRUST_Pleadings_PDFs"
```
Then **restart iCharlotte** (it runs from the main checkout; Python caches modules).
Oppose/Generate a Motion will now prefer your firm's previously-cited authority,
falling back to corpus-only when the index is absent.

---

## Self-review notes (for the implementer)

- **Provenance is additive:** every existing `research_arguments` caller omits the
  new kwargs, so corpus-only behavior is byte-for-byte unchanged. Verify by running
  the opposition research tests unmodified.
- **Thread-local connections (Task 6) are non-negotiable** — the research pipeline
  calls `authority_candidates` from a ThreadPoolExecutor.
- **OCR/citation noise:** harvested cites that don't resolve in the corpus and fail
  CL get `unverified_firm` and are flagged, never silently trusted.
- **Deferred to later phases:** style auto-selection (Phase 2), per-proposition
  vector rerank, the citation-panel source/alternatives/swap UI + Workbench
  "Sample Library" tab + `--compact`/`--rebuild` flags (Phase 3).
