# Firm Brief Library — Phase 3A (Backend Polish) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development. Steps use checkbox (`- [ ]`) syntax.

**Goal:** Backend half of Phase 3 — clean harvested case names, refactor the shared wizard helpers into one module, store full text so image-only briefs model style, add per-proposition semantic rerank for authority, and add `--rebuild`/`--compact` CLI flags.

**Architecture:** All changes are additive/backward-compatible to the merged firm_briefs package. New columns/sidecars default empty so old index rows still work; the style/authority paths fall back to current behavior when the new data is absent.

**Tech Stack:** Python, sqlite, numpy memmap, fastembed, pytest. Tests: `C:\geminiterminal2\.venv\Scripts\python.exe`. **Work in worktree `C:\firm-briefs-p3` on branch `feat/firm-briefs-phase3`.**

**Environment for implementers:** shell cwd RESETS to `C:\geminiterminal2` between commands — begin every PowerShell command with `Set-Location 'C:\firm-briefs-p3';`. Tests: `Set-Location 'C:\firm-briefs-p3'; & 'C:\geminiterminal2\.venv\Scripts\python.exe' -m pytest <path> -v`. PowerShell only; no bash compound `cd && `. git via `git -C "C:/firm-briefs-p3" ...`.

---

### Task 1 (3.7): Clean harvested case names

**Files:**
- Modify: `icharlotte_core/firm_briefs/citation_harvest.py`
- Test: `tests/test_firm_briefs/test_harvest_cleanup.py`

- [ ] **Step 1: Write the failing test**

```python
# tests/test_firm_briefs/test_harvest_cleanup.py
from icharlotte_core.firm_briefs.citation_harvest import clean_case_name

def test_strips_signal_words():
    assert clean_case_name("See Townsend v. Superior Court") == "Townsend v. Superior Court"
    assert clean_case_name("See, e.g., Blank v. Kirwan") == "Blank v. Kirwan"
    assert clean_case_name("In Beckstead v. Superior Court") == "Beckstead v. Superior Court"
    assert clean_case_name("Cf. Ellis v. Toshiba") == "Ellis v. Toshiba"
    assert clean_case_name("Accord Sangster v. Paetkau") == "Sangster v. Paetkau"

def test_collapses_whitespace_and_newlines():
    assert clean_case_name("North \nCoast Business Park v. Nielsen") == "North Coast Business Park v. Nielsen"

def test_leaves_clean_names_untouched():
    assert clean_case_name("Dore v. Arnold Worldwide, Inc.") == "Dore v. Arnold Worldwide, Inc."
    assert clean_case_name("") == ""
```

- [ ] **Step 2: Run test to verify it fails**

Run: `Set-Location 'C:\firm-briefs-p3'; & 'C:\geminiterminal2\.venv\Scripts\python.exe' -m pytest tests/test_firm_briefs/test_harvest_cleanup.py -v`
Expected: FAIL (`ImportError: cannot import name 'clean_case_name'`)

- [ ] **Step 3: Implement** — add to `citation_harvest.py` and apply it in `harvest_cites`:

```python
import re as _re

_SIGNAL_PREFIX = _re.compile(
    r"^\s*(?:see,?\s+e\.?g\.?,?|see\s+also|see|accord|cf\.?|but\s+see|in\s+re|in)\s+",
    _re.IGNORECASE,
)


def clean_case_name(name: str) -> str:
    """Strip leading citation signal words and collapse whitespace/newlines."""
    s = _re.sub(r"\s+", " ", (name or "").strip())
    prev = None
    while s and s != prev:           # strip stacked signals ("See, e.g., ...")
        prev = s
        s = _SIGNAL_PREFIX.sub("", s).strip()
    return s
```

Then in `harvest_cites`, wrap the case name: change `case_name=getattr(c, "case_name", "") or ""` to
`case_name=clean_case_name(getattr(c, "case_name", "") or "")`.

- [ ] **Step 4: Run test to verify it passes**

Run: same as Step 2. Expected: PASS (3 passed)

- [ ] **Step 5: Run harvest regression**

Run: `Set-Location 'C:\firm-briefs-p3'; & 'C:\geminiterminal2\.venv\Scripts\python.exe' -m pytest tests/test_firm_briefs/test_citation_harvest.py -v`
Expected: PASS (existing harvest tests unaffected; clean names don't change the Townsend example which has no signal prefix).

- [ ] **Step 6: Commit**

```bash
git -C "C:/firm-briefs-p3" add icharlotte_core/firm_briefs/citation_harvest.py tests/test_firm_briefs/test_harvest_cleanup.py
git -C "C:/firm-briefs-p3" commit -m "feat(firm_briefs): clean harvested case names (strip signal words, collapse whitespace)"
```

---

### Task 2 (3.7 backfill): one-off script to clean existing case names

**Files:**
- Create: `clean_firm_index_names.py` (worktree root)
- Test: `tests/test_firm_briefs/test_clean_names_backfill.py`

- [ ] **Step 1: Write the failing test**

```python
# tests/test_firm_briefs/test_clean_names_backfill.py
import numpy as np
from icharlotte_core.firm_briefs.index import FirmBriefIndex
from icharlotte_core.firm_briefs.citation_harvest import HarvestedCite
from clean_firm_index_names import clean_names

def _vec():
    v = np.ones(384, dtype=np.float32); return v/np.linalg.norm(v)

def test_backfill_cleans_existing(tmp_path):
    idx = FirmBriefIndex(db_path=str(tmp_path/"fb.db"), vectors_path=str(tmp_path/"v.f16"))
    idx.create_schema()
    idx.upsert_brief(path="p.pdf", content_hash="h", motion_type="compel", side="moving",
                     heading="", profile="p", profile_vec=_vec(), char_len=1, ocr_ratio=0.0,
                     cites=[HarvestedCite(case_name="See Townsend v. Superior Court",
                                          reporter_citation="61 Cal.App.4th 1431",
                                          norm_cite="61cal.app.4th1431", proposition="x")])
    changed = clean_names(idx)
    assert changed == 1
    con = idx._conn()
    name = con.execute("SELECT case_name FROM citations").fetchone()[0]
    assert name == "Townsend v. Superior Court"
    assert clean_names(idx) == 0   # idempotent
```

- [ ] **Step 2: Run test to verify it fails**

Run: `Set-Location 'C:\firm-briefs-p3'; & 'C:\geminiterminal2\.venv\Scripts\python.exe' -m pytest tests/test_firm_briefs/test_clean_names_backfill.py -v`
Expected: FAIL (`ModuleNotFoundError: clean_firm_index_names`)

- [ ] **Step 3: Implement**

```python
# clean_firm_index_names.py
"""One-off: clean existing citations.case_name in the firm index (strip signal
words / collapse whitespace). In-place UPDATE; idempotent."""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from icharlotte_core.firm_briefs.citation_harvest import clean_case_name


def clean_names(index) -> int:
    con = index._conn()
    rows = con.execute("SELECT id, case_name FROM citations").fetchall()
    changed = 0
    for r in rows:
        cleaned = clean_case_name(r["case_name"] or "")
        if cleaned != (r["case_name"] or ""):
            con.execute("UPDATE citations SET case_name=? WHERE id=?", (cleaned, r["id"]))
            changed += 1
    con.commit()
    return changed


def main() -> int:
    from icharlotte_core.firm_briefs import factory
    from icharlotte_core.firm_briefs.index import FirmBriefIndex
    if not factory.index_available():
        print("No index; nothing to clean."); return 1
    db, vec = factory.index_paths()
    idx = FirmBriefIndex(db_path=db, vectors_path=vec); idx.create_schema()
    print(f"Cleaned {clean_names(idx)} case names."); return 0


if __name__ == "__main__":
    raise SystemExit(main())
```

- [ ] **Step 4: Run test to verify it passes**

Run: same as Step 2. Expected: PASS (1 passed)

- [ ] **Step 5: Commit**

```bash
git -C "C:/firm-briefs-p3" add clean_firm_index_names.py tests/test_firm_briefs/test_clean_names_backfill.py
git -C "C:/firm-briefs-p3" commit -m "feat(firm_briefs): clean_firm_index_names backfill for existing case names"
```

---

### Task 3 (3.6): Extract shared wizard-research helpers into one module

**Files:**
- Create: `icharlotte_core/ui/wizard/pages/_motion_research_support.py`
- Modify: `icharlotte_core/ui/wizard/pages/oppose_motion_page.py`, `generate_motion_page.py`
- Test: `tests/test_firm_briefs/test_research_support_module.py`

- [ ] **Step 1: Write the failing test**

```python
# tests/test_firm_briefs/test_research_support_module.py
from icharlotte_core.ui.wizard.pages import _motion_research_support as mrs
from icharlotte_core.ui.wizard.pages import oppose_motion_page as omp
from icharlotte_core.ui.wizard.pages import generate_motion_page as gmp


def test_shared_module_exposes_helpers():
    for name in ("make_firm_provider", "make_local_corpus", "research_targets",
                 "firm_style_exemplars"):
        assert hasattr(mrs, name)


def test_pages_use_shared_helpers():
    # Both pages re-export / reference the shared helpers (no private cross-import).
    assert omp._make_firm_provider is mrs.make_firm_provider
    assert gmp._make_firm_provider is mrs.make_firm_provider
    assert gmp._firm_style_exemplars is mrs.firm_style_exemplars
```

- [ ] **Step 2: Run test to verify it fails**

Run: `Set-Location 'C:\firm-briefs-p3'; & 'C:\geminiterminal2\.venv\Scripts\python.exe' -m pytest tests/test_firm_briefs/test_research_support_module.py -v`
Expected: FAIL (`ModuleNotFoundError: _motion_research_support`)

- [ ] **Step 3: Implement**

First READ the current bodies of `_make_firm_provider`, `_make_local_corpus`, `_research_targets`, `_firm_style_exemplars` in `oppose_motion_page.py` (and the `_corpus_available`/`_corpus_paths`/`_corpus_embedder` helpers `_make_local_corpus` depends on). Move them verbatim into the new module:

```python
# icharlotte_core/ui/wizard/pages/_motion_research_support.py
"""Shared research/style helpers for the Oppose- and Generate-a-Motion pages.

Extracted so both pages depend on one module instead of generate importing
oppose's private helpers. Behavior is identical to the previous oppose_motion_page
definitions.
"""
from __future__ import annotations

import os

# (Move the bodies of the following from oppose_motion_page.py verbatim, renamed
# without the leading underscore; keep their internal imports.)
#   make_local_corpus()       <- _make_local_corpus
#   make_firm_provider(corpus)<- _make_firm_provider
#   research_targets(metadata, plan) <- _research_targets
#   firm_style_exemplars(motion_type, side, metadata) <- _firm_style_exemplars
# Also move any private corpus helpers they call (_corpus_available, _corpus_paths,
# _corpus_embedder) into this module if they are not used elsewhere; if they ARE
# used elsewhere in oppose_motion_page, leave them there and import them here.
```

Then in **oppose_motion_page.py**: delete the moved function bodies and replace with imports + backward-compat aliases (so existing references and tests keep working):
```python
from icharlotte_core.ui.wizard.pages._motion_research_support import (
    make_firm_provider as _make_firm_provider,
    make_local_corpus as _make_local_corpus,
    research_targets as _research_targets,
    firm_style_exemplars as _firm_style_exemplars,
)
```
In **generate_motion_page.py**: replace its `from ...oppose_motion_page import _make_firm_provider, _make_local_corpus, _research_targets` (and the `_firm_style_exemplars` import) with the same imports from `_motion_research_support`, keeping the `_make_firm_provider`/`_firm_style_exemplars` local aliases its code already uses.

Keep `normalize_motion_type` imports as they are (already from motion_taxonomy in both pages).

- [ ] **Step 4: Run test + regression**

Run: `Set-Location 'C:\firm-briefs-p3'; & 'C:\geminiterminal2\.venv\Scripts\python.exe' -m pytest tests/test_firm_briefs tests/test_opposition tests/test_motion_generation -q`
Expected: PASS (the shared-module test + all existing oppose/generate/firm tests — behavior preserved).

- [ ] **Step 5: Commit**

```bash
git -C "C:/firm-briefs-p3" add icharlotte_core/ui/wizard/pages/_motion_research_support.py icharlotte_core/ui/wizard/pages/oppose_motion_page.py icharlotte_core/ui/wizard/pages/generate_motion_page.py tests/test_firm_briefs/test_research_support_module.py
git -C "C:/firm-briefs-p3" commit -m "refactor(wizard): extract shared motion-research helpers into _motion_research_support"
```

---

### Task 4 (3.3): Store full_text in the index; style reads it

**Files:**
- Modify: `icharlotte_core/firm_briefs/index.py`, `ingest.py`, `style.py`
- Test: `tests/test_firm_briefs/test_full_text_style.py`

- [ ] **Step 1: Write the failing test**

```python
# tests/test_firm_briefs/test_full_text_style.py
import numpy as np
from icharlotte_core.firm_briefs.index import FirmBriefIndex
from icharlotte_core.firm_briefs.citation_harvest import HarvestedCite
from icharlotte_core.firm_briefs import style

def _vec(): v = np.ones(384, np.float32); return v/np.linalg.norm(v)

def test_index_stores_and_returns_full_text(tmp_path):
    idx = FirmBriefIndex(db_path=str(tmp_path/"fb.db"), vectors_path=str(tmp_path/"v.f16"))
    idx.create_schema()
    idx.upsert_brief(path="p.pdf", content_hash="h", motion_type="compel", side="opposition",
                     heading="", profile="p", profile_vec=_vec(), char_len=10, ocr_ratio=0.0,
                     cites=[HarvestedCite(reporter_citation="1 Cal.5th 1", norm_cite="1", proposition="x")],
                     full_text="ARGUMENT\nThe motion fails because the meet and confer was inadequate.")
    assert "meet and confer" in idx.get_full_text("p.pdf")

def test_style_uses_stored_full_text_no_extraction(tmp_path):
    idx = FirmBriefIndex(db_path=str(tmp_path/"fb.db"), vectors_path=str(tmp_path/"v.f16"))
    idx.create_schema()
    idx.upsert_brief(path="img.pdf", content_hash="h", motion_type="compel", side="opposition",
                     heading="", profile="meet confer", profile_vec=_vec(), char_len=10, ocr_ratio=0.0,
                     cites=[HarvestedCite(reporter_citation="1 Cal.5th 1", norm_cite="1", proposition="x")],
                     full_text="ARGUMENT\n" + "stored opposition body " * 20)
    class Emb:
        dim=384
        def encode(self, t): return np.ones((len(t),384), np.float32)
    from types import SimpleNamespace
    m = SimpleNamespace(relief_requested="compel", principal_arguments=["meet and confer"])
    # extract_fn raises -> proves it used stored full_text, not extraction
    def boom(p): raise AssertionError("should not extract")
    out = style.select_exemplars("compel", "opposition", m, index=idx, embedder=Emb(),
                                 extract_fn=boom, cache_dir=str(tmp_path), max_chars=200)
    assert len(out) == 1 and out[0].startswith("ARGUMENT")
```

- [ ] **Step 2: Run test to verify it fails**

Run: `Set-Location 'C:\firm-briefs-p3'; & 'C:\geminiterminal2\.venv\Scripts\python.exe' -m pytest tests/test_firm_briefs/test_full_text_style.py -v`
Expected: FAIL (`upsert_brief` has no `full_text` kwarg / `get_full_text` missing)

- [ ] **Step 3: Implement**

In `index.py`:
- Add `full_text TEXT DEFAULT ''` to the `briefs` CREATE TABLE.
- Add a migration in `create_schema()` after `executescript`: `try: con.execute("ALTER TABLE briefs ADD COLUMN full_text TEXT DEFAULT ''") except sqlite3.OperationalError: pass` (column may already exist).
- `upsert_brief(...)` gains `full_text: str = ""` kwarg; include it in both the INSERT and UPDATE column lists/values.
- Add:
```python
    def get_full_text(self, path: str) -> str:
        row = self._conn().execute("SELECT full_text FROM briefs WHERE path=?", (path,)).fetchone()
        return (row["full_text"] if row else "") or ""
    def get_full_text_by_id(self, brief_id: int) -> str:
        row = self._conn().execute("SELECT full_text FROM briefs WHERE id=?", (brief_id,)).fetchone()
        return (row["full_text"] if row else "") or ""
```

In `ingest.py`: pass the already-extracted `text` as `full_text=text` to `upsert_brief`.

In `style.py` `_excerpt(...)`: accept an optional `stored_text`; if provided and non-empty, trim it instead of calling `extract_fn`. In `select_exemplars`, look up `index.get_full_text(c["path"])` first; pass it as `stored_text`; only fall back to `extract_fn` when empty. Concretely:
```python
        stored = ""
        try:
            stored = index.get_full_text(c["path"])
        except Exception:
            stored = ""
        txt = _excerpt(c["path"], cache_dir=cache_dir, extract_fn=extract_fn,
                       max_chars=max_chars, stored_text=stored)
```
and `_excerpt`:
```python
def _excerpt(path, *, cache_dir, extract_fn, max_chars, stored_text=""):
    ... # cache check unchanged
    raw = stored_text if (stored_text and stored_text.strip()) else extract_fn(path)
    excerpt = _trim_to_argument(raw, max_chars)
    ...
```

- [ ] **Step 4: Run test + style regression**

Run: `Set-Location 'C:\firm-briefs-p3'; & 'C:\geminiterminal2\.venv\Scripts\python.exe' -m pytest tests/test_firm_briefs/test_full_text_style.py tests/test_firm_briefs/test_style_select.py tests/test_firm_briefs/test_index.py -v`
Expected: PASS (new + existing style/index tests; existing `extract_fn`-injected tests still pass because `full_text` is absent there → falls back to extract_fn).

- [ ] **Step 5: Commit**

```bash
git -C "C:/firm-briefs-p3" add icharlotte_core/firm_briefs/index.py icharlotte_core/firm_briefs/ingest.py icharlotte_core/firm_briefs/style.py tests/test_firm_briefs/test_full_text_style.py
git -C "C:/firm-briefs-p3" commit -m "feat(firm_briefs): store full_text at ingest; style reuses it (fixes image-only style gap)"
```

---

### Task 5 (3.4): Per-proposition vectors + semantic rerank in authority

**Files:**
- Modify: `icharlotte_core/firm_briefs/index.py`, `ingest.py`, `provider.py`
- Test: `tests/test_firm_briefs/test_authority_semantic.py`

- [ ] **Step 1: Write the failing test**

```python
# tests/test_firm_briefs/test_authority_semantic.py
import numpy as np
from icharlotte_core.firm_briefs.index import FirmBriefIndex
from icharlotte_core.firm_briefs.citation_harvest import HarvestedCite

def _v(i):
    v = np.zeros(384, np.float32); v[i] = 1.0; return v

def test_semantic_rerank_orders_paraphrase_first(tmp_path):
    idx = FirmBriefIndex(db_path=str(tmp_path/"fb.db"), vectors_path=str(tmp_path/"v.f16"))
    idx.create_schema()
    # two cites that BOTH keyword-match "discovery", but only one is semantically the query
    idx.upsert_brief(path="p.pdf", content_hash="h", motion_type="compel", side="opposition",
                     heading="", profile="p", profile_vec=_v(0), char_len=1, ocr_ratio=0.0,
                     cites=[
                         HarvestedCite(case_name="A", reporter_citation="1 Cal.5th 1", norm_cite="1",
                                       proposition="forensic discovery imaging is permitted on a showing"),
                         HarvestedCite(case_name="B", reporter_citation="2 Cal.5th 2", norm_cite="2",
                                       proposition="discovery is generally broad"),
                     ],
                     prop_vecs=[_v(5), _v(9)])  # A->dim5, B->dim9
    # query vector aligned with A (dim5)
    hits = idx.authority_candidates("discovery", motion_type="compel", limit=5, query_vec=_v(5))
    assert hits[0]["case_name"] == "A"   # semantic rerank floats A above B

def test_fts_only_fallback_when_no_query_vec(tmp_path):
    idx = FirmBriefIndex(db_path=str(tmp_path/"fb.db"), vectors_path=str(tmp_path/"v.f16"))
    idx.create_schema()
    idx.upsert_brief(path="p.pdf", content_hash="h", motion_type="compel", side="opposition",
                     heading="", profile="p", profile_vec=_v(0), char_len=1, ocr_ratio=0.0,
                     cites=[HarvestedCite(case_name="A", reporter_citation="1 Cal.5th 1", norm_cite="1",
                                          proposition="meet and confer required")])
    hits = idx.authority_candidates("meet and confer", motion_type="compel", limit=5)
    assert any(h["case_name"] == "A" for h in hits)   # works without query_vec
```

- [ ] **Step 2: Run test to verify it fails**

Run: `Set-Location 'C:\firm-briefs-p3'; & 'C:\geminiterminal2\.venv\Scripts\python.exe' -m pytest tests/test_firm_briefs/test_authority_semantic.py -v`
Expected: FAIL (`upsert_brief` has no `prop_vecs`; `authority_candidates` has no `query_vec`)

- [ ] **Step 3: Implement**

In `index.py`:
- Add `prop_vec_row INTEGER DEFAULT -1` to the `citations` CREATE TABLE + a `create_schema` migration `ALTER TABLE citations ADD COLUMN prop_vec_row ...` (guarded by OperationalError).
- Add a second sidecar mirroring the profile sidecar: `self.prop_vectors_path = vectors_path + ".prop"`; add `_append_prop_vector(vec)->row` and `load_prop_vectors()` (copy the `_append_vector`/`load_vectors` code, pointed at `prop_vectors_path`).
- `upsert_brief(...)` gains `prop_vecs: list | None = None`. When provided (len == len(cites)), append each to the prop sidecar and store the row in `citations.prop_vec_row` (set it in the per-cite INSERT). When None, store `-1`.
- `authority_candidates(self, proposition, *, motion_type, limit=8, query_vec=None)`: keep the existing FTS5 query to get candidate rows (also SELECT `c.prop_vec_row`). If `query_vec is not None` AND any candidate has `prop_vec_row >= 0`, compute cosine(query_vec, prop_vectors[row]) per candidate and **reorder** by a blend: final rank = fuse(fts_rank_position, semantic_cos) via reciprocal-rank fusion (rank by `1/(60+fts_pos) + cos`); candidates without a prop vector keep only their fts component. Return dicts unchanged in shape (plus existing keys). When `query_vec is None`, behavior is exactly today's FTS5 ordering.

In `ingest.py`: build `prop_vecs = embedder.encode([c.proposition for c in cites])` (a single batch) and pass `prop_vecs=list(prop_vecs)` to `upsert_brief`. Guard: if `cites` is empty, pass `prop_vecs=None`.

In `provider.py` `candidates_for(...)`: embed the proposition once (reuse the same embedder the index uses — accept an optional `embedder`, default `get_embedder()`), and pass `query_vec` to `index.authority_candidates`. Wrap in try/except so an embed failure degrades to FTS-only (`query_vec=None`).

- [ ] **Step 4: Run test + regression**

Run: `Set-Location 'C:\firm-briefs-p3'; & 'C:\geminiterminal2\.venv\Scripts\python.exe' -m pytest tests/test_firm_briefs -q`
Expected: PASS (new semantic tests + all existing firm_briefs tests; the `authority_candidates` calls without `query_vec` in older tests still work).

- [ ] **Step 5: Commit**

```bash
git -C "C:/firm-briefs-p3" add icharlotte_core/firm_briefs/index.py icharlotte_core/firm_briefs/ingest.py icharlotte_core/firm_briefs/provider.py tests/test_firm_briefs/test_authority_semantic.py
git -C "C:/firm-briefs-p3" commit -m "feat(firm_briefs): per-proposition vectors + semantic rerank for authority (RRF, FTS5 fallback)"
```

---

### Task 6 (3.5): `--rebuild` / `--compact` CLI flags

**Files:**
- Modify: `icharlotte_core/firm_briefs/__main__.py`, `icharlotte_core/firm_briefs/index.py`
- Test: `tests/test_firm_briefs/test_compact.py`

- [ ] **Step 1: Write the failing test**

```python
# tests/test_firm_briefs/test_compact.py
import os, numpy as np
from icharlotte_core.firm_briefs.index import FirmBriefIndex
from icharlotte_core.firm_briefs.citation_harvest import HarvestedCite

def _v(): v=np.ones(384,np.float32); return v/np.linalg.norm(v)

def test_compact_drops_stale_vector_rows(tmp_path):
    idx = FirmBriefIndex(db_path=str(tmp_path/"fb.db"), vectors_path=str(tmp_path/"v.f16"))
    idx.create_schema()
    for p in ("a.pdf", "b.pdf"):
        idx.upsert_brief(path=p, content_hash="h", motion_type="compel", side="moving",
                         heading="", profile="p", profile_vec=_v(), char_len=1, ocr_ratio=0.0,
                         cites=[HarvestedCite(reporter_citation="1 Cal.5th 1", norm_cite="1", proposition="x")])
    idx.mark_stale("a.pdf")            # a becomes stale -> its vector row is dead
    rows_before = idx.load_vectors().shape[0]
    idx.compact()
    rows_after = idx.load_vectors().shape[0]
    assert rows_after < rows_before    # dead row reclaimed
    # surviving brief still queryable + vec aligned
    hits = idx.style_candidates(_v(), motion_type="compel", side="moving", k=5)
    assert [h["path"] for h in hits] == ["b.pdf"]
```

- [ ] **Step 2: Run test to verify it fails**

Run: `Set-Location 'C:\firm-briefs-p3'; & 'C:\geminiterminal2\.venv\Scripts\python.exe' -m pytest tests/test_firm_briefs/test_compact.py -v`
Expected: FAIL (`FirmBriefIndex` has no `compact`)

- [ ] **Step 3: Implement**

In `index.py` add:
```python
    def compact(self) -> None:
        """Rewrite the profile + prop vector sidecars keeping only live rows,
        re-pointing vec_row/prop_vec_row. Reclaims space from upsert/stale churn."""
        import numpy as np
        con = self._conn()
        # profile vectors: rebuild for status='ok' briefs
        vecs = self.load_vectors()
        live = con.execute("SELECT id, vec_row FROM briefs WHERE status='ok' AND vec_row>=0 ORDER BY id").fetchall()
        new = []
        for new_row, r in enumerate(live):
            if 0 <= r["vec_row"] < vecs.shape[0]:
                new.append(np.asarray(vecs[r["vec_row"]], dtype=np.float16))
                con.execute("UPDATE briefs SET vec_row=? WHERE id=?", (new_row, r["id"]))
        with open(self.vectors_path, "wb") as f:
            if new:
                f.write(np.stack(new).astype(np.float16).tobytes())
            f.flush(); os.fsync(f.fileno())
        # prop vectors (if present)
        if os.path.exists(self.prop_vectors_path):
            pvecs = self.load_prop_vectors()
            plive = con.execute(
                "SELECT c.id, c.prop_vec_row FROM citations c JOIN briefs b ON b.id=c.brief_id "
                "WHERE b.status='ok' AND c.prop_vec_row>=0 ORDER BY c.id").fetchall()
            pnew = []
            for nr, r in enumerate(plive):
                if 0 <= r["prop_vec_row"] < pvecs.shape[0]:
                    pnew.append(np.asarray(pvecs[r["prop_vec_row"]], dtype=np.float16))
                    con.execute("UPDATE citations SET prop_vec_row=? WHERE id=?", (nr, r["id"]))
            with open(self.prop_vectors_path, "wb") as f:
                if pnew:
                    f.write(np.stack(pnew).astype(np.float16).tobytes())
                f.flush(); os.fsync(f.fileno())
        con.commit()
```
(`import os` is already at the top of index.py.)

In `__main__.py` add flags:
```python
    ap.add_argument("--rebuild", action="store_true", help="drop the index and rebuild from scratch")
    ap.add_argument("--compact", action="store_true", help="reclaim dead vector-sidecar rows")
```
After computing `db, vec` and before/at the right point in `main`:
- `--rebuild`: if set, delete `db`, `vec`, and `vec + ".prop"` if they exist, before creating the index.
- `--compact`: after (or instead of) ingest, open the index and call `idx.compact()`, print sidecar sizes before/after, and return.

- [ ] **Step 4: Run test + CLI smoke**

Run: `Set-Location 'C:\firm-briefs-p3'; & 'C:\geminiterminal2\.venv\Scripts\python.exe' -m pytest tests/test_firm_briefs/test_compact.py -v`
Expected: PASS. Then smoke the help: `Set-Location 'C:\firm-briefs-p3'; & 'C:\geminiterminal2\.venv\Scripts\python.exe' -m icharlotte_core.firm_briefs --help` shows `--rebuild` and `--compact`.

- [ ] **Step 5: Commit**

```bash
git -C "C:/firm-briefs-p3" add icharlotte_core/firm_briefs/index.py icharlotte_core/firm_briefs/__main__.py tests/test_firm_briefs/test_compact.py
git -C "C:/firm-briefs-p3" commit -m "feat(firm_briefs): --rebuild / --compact CLI flags + index.compact()"
```

---

### Task 7: Full backend regression

- [ ] **Step 1:** `Set-Location 'C:\firm-briefs-p3'; & 'C:\geminiterminal2\.venv\Scripts\python.exe' -m pytest tests/test_firm_briefs tests/test_opposition tests/test_motion_generation tests/test_wizard -q`
Expected: all firm_briefs pass; no NEW opposition/generate/wizard failures (pre-existing Qt-collection errors are not regressions). Report counts.

---

## Self-review notes
- All new columns/sidecars/kwargs default to empty/None → old index rows + existing tests behave as before.
- Re-index (`--rebuild`) is needed AFTER merge to populate `full_text` + prop vectors + clean names on the real library (a Phase 3 execution step, not a code task).
- Task 5 RRF: fuse FTS rank position with semantic cosine; candidates lacking a prop vector still rank by FTS. Keep `query_vec=None` path byte-identical to today.
- Phase 3B (citation panel provenance UI + Workbench Sample Library tab) is a separate plan.
