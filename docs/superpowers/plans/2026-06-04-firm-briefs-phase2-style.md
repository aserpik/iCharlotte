# Firm Brief Library — Phase 2 (Style Auto-Selection) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development. Steps use checkbox (`- [ ]`) syntax.

**Goal:** Make the firm-brief library function as **style guides** — at draft time, auto-select the 2-3 past briefs most similar to the current motion (by the issue-profile embeddings already stored in the index) and feed their argument-section text to the Oppose-a-Motion and Generate-a-Motion drafters.

**Architecture:** Phase 1 already embeds and stores a profile vector per brief (`profiles.f16`, addressed by `briefs.vec_row`) — nothing reads them yet. Phase 2 adds `FirmBriefIndex.style_candidates()` (cosine over those vectors, filtered by motion_type + side, with a quality/dedup guard), a `firm_briefs.style.select_exemplars()` that embeds the current motion's profile and returns trimmed excerpts of the top matches, and wires those excerpts into both drafters' existing `style_exemplars` channel. Purely additive: absent the index, the drafters use today's manual exemplars unchanged.

**Tech Stack:** numpy (cosine over the float16 memmap), fastembed (reused), pytest. Tests: `C:\geminiterminal2\.venv\Scripts\python.exe`. **Work happens in the worktree `C:\firm-briefs-p2` on branch `feat/firm-briefs-phase2-style`.**

**Environment for implementers:** shell cwd RESETS to `C:\geminiterminal2` between commands — begin every PowerShell command with `Set-Location 'C:\firm-briefs-p2';`. Run tests from the worktree: `Set-Location 'C:\firm-briefs-p2'; & 'C:\geminiterminal2\.venv\Scripts\python.exe' -m pytest <path> -v`. PowerShell only; no bash compound `cd && `. git from the worktree.

---

### Task 1: `profile_from_metadata` — current-motion issue profile

**Files:**
- Modify: `icharlotte_core/firm_briefs/profile.py`
- Test: `tests/test_firm_briefs/test_profile_from_metadata.py`

- [ ] **Step 1: Write the failing test**

```python
# tests/test_firm_briefs/test_profile_from_metadata.py
from types import SimpleNamespace
from icharlotte_core.firm_briefs.profile import profile_from_metadata


def test_profile_from_metadata_combines_relief_and_arguments():
    meta = SimpleNamespace(
        relief_requested="compel further responses to RFP set one",
        principal_arguments=["Plaintiff failed to meet and confer", "Responses are evasive"],
    )
    prof = profile_from_metadata(meta)
    assert "compel further responses" in prof
    assert "meet and confer" in prof
    assert "evasive" in prof


def test_profile_from_metadata_handles_missing_fields():
    assert profile_from_metadata(SimpleNamespace()) == ""
    assert profile_from_metadata(None) == ""
```

- [ ] **Step 2: Run test to verify it fails**

Run: `Set-Location 'C:\firm-briefs-p2'; & 'C:\geminiterminal2\.venv\Scripts\python.exe' -m pytest tests/test_firm_briefs/test_profile_from_metadata.py -v`
Expected: FAIL (`ImportError: cannot import name 'profile_from_metadata'`)

- [ ] **Step 3: Implement** — append to `icharlotte_core/firm_briefs/profile.py`:

```python
def profile_from_metadata(meta) -> str:
    """Compose the same shape of issue-profile string used at ingest, but from a
    live motion's analyzer metadata (duck-typed: relief_requested + principal_arguments)."""
    if meta is None:
        return ""
    relief = getattr(meta, "relief_requested", "") or ""
    args = getattr(meta, "principal_arguments", None) or []
    parts = [relief] + [str(a) for a in args]
    text = " \n".join(p for p in parts if p and str(p).strip())
    return re.sub(r"\s+", " ", text).strip()
```

(`re` is already imported at the top of profile.py.)

- [ ] **Step 4: Run test to verify it passes**

Run: same command as Step 2. Expected: PASS (2 passed)

- [ ] **Step 5: Commit**

```bash
git -C "C:/firm-briefs-p2" add icharlotte_core/firm_briefs/profile.py tests/test_firm_briefs/test_profile_from_metadata.py
git -C "C:/firm-briefs-p2" commit -m "feat(firm_briefs): profile_from_metadata for draft-time style query"
```

---

### Task 2: `FirmBriefIndex.style_candidates` — cosine over profile vectors

**Files:**
- Modify: `icharlotte_core/firm_briefs/index.py`
- Test: `tests/test_firm_briefs/test_style_candidates.py`

`load_vectors()` already returns the float16 memmap `(n, EMBED_DIM)`; each brief's row is `briefs.vec_row`. Style filters **strictly by side** (moving voice != opposition voice), unlike authority.

- [ ] **Step 1: Write the failing test**

```python
# tests/test_firm_briefs/test_style_candidates.py
import numpy as np
import pytest
from icharlotte_core.firm_briefs.index import FirmBriefIndex
from icharlotte_core.firm_briefs.citation_harvest import HarvestedCite


def _unit(*vals):
    v = np.zeros(384, dtype=np.float32)
    for i, x in vals:
        v[i] = x
    n = np.linalg.norm(v) or 1.0
    return v / n


@pytest.fixture
def idx(tmp_path):
    ix = FirmBriefIndex(db_path=str(tmp_path / "fb.db"), vectors_path=str(tmp_path / "v.f16"))
    ix.create_schema()
    return ix


def _add(idx, path, side, vec, char_len=5000, ocr=0.0):
    idx.upsert_brief(path=path, content_hash="h", motion_type="compel", side=side,
                     heading="", profile="p", profile_vec=vec, char_len=char_len,
                     ocr_ratio=ocr, cites=[HarvestedCite(reporter_citation="1 Cal.5th 1",
                                                         norm_cite="1cal.5th1", proposition="x")])


def test_picks_most_similar_same_side(idx):
    _add(idx, "a.pdf", "opposition", _unit((0, 1.0)))
    _add(idx, "b.pdf", "opposition", _unit((1, 1.0)))
    _add(idx, "c.pdf", "opposition", _unit((2, 1.0)))
    res = idx.style_candidates(_unit((1, 1.0)), motion_type="compel", side="opposition", k=2)
    assert res[0]["path"] == "b.pdf"            # exact match ranks first
    assert len(res) == 2


def test_side_filter_excludes_other_side(idx):
    _add(idx, "moving.pdf", "moving", _unit((1, 1.0)))
    _add(idx, "opp.pdf", "opposition", _unit((1, 1.0)))
    res = idx.style_candidates(_unit((1, 1.0)), motion_type="compel", side="opposition", k=5)
    assert [r["path"] for r in res] == ["opp.pdf"]


def test_quality_penalty_downranks_tiny_noisy(idx):
    # Both equally similar, but 'tiny' is short + high OCR noise -> should rank lower.
    _add(idx, "good.pdf", "opposition", _unit((1, 1.0)), char_len=8000, ocr=0.0)
    _add(idx, "tiny.pdf", "opposition", _unit((1, 1.0)), char_len=300, ocr=0.4)
    res = idx.style_candidates(_unit((1, 1.0)), motion_type="compel", side="opposition", k=2)
    assert res[0]["path"] == "good.pdf"


def test_dedupes_versioned_copies(idx):
    _add(idx, "Brief.pdf", "opposition", _unit((1, 1.0)))
    _add(idx, "Brief_1.pdf", "opposition", _unit((1, 1.0)))
    res = idx.style_candidates(_unit((1, 1.0)), motion_type="compel", side="opposition", k=5)
    assert len(res) == 1   # _1 versioned duplicate collapsed
```

- [ ] **Step 2: Run test to verify it fails**

Run: `Set-Location 'C:\firm-briefs-p2'; & 'C:\geminiterminal2\.venv\Scripts\python.exe' -m pytest tests/test_firm_briefs/test_style_candidates.py -v`
Expected: FAIL (`AttributeError: ... 'style_candidates'`)

- [ ] **Step 3: Implement** — add to `FirmBriefIndex` (after `authority_candidates`):

```python
    def style_candidates(self, query_vec, *, motion_type: str, side: str,
                         k: int = 3) -> list[dict]:
        """Top-k briefs of this motion_type AND side by cosine similarity of the
        stored profile vector to query_vec, with a quality penalty (short/noisy
        briefs make poor style models) and version dedup."""
        import os
        import re as _re
        con = self._conn()
        rows = con.execute(
            "SELECT id, path, vec_row, char_len, ocr_ratio FROM briefs "
            "WHERE status='ok' AND motion_type=? AND side=? AND vec_row>=0",
            (motion_type, side),
        ).fetchall()
        if not rows:
            return []
        vecs = self.load_vectors()
        if getattr(vecs, "shape", (0,))[0] == 0:
            return []
        q = np.asarray(query_vec, dtype=np.float32).reshape(-1)
        qn = float(np.linalg.norm(q)) or 1.0
        scored: list[tuple[float, dict]] = []
        for r in rows:
            vr = int(r["vec_row"])
            if vr < 0 or vr >= vecs.shape[0]:
                continue
            v = np.asarray(vecs[vr], dtype=np.float32)
            vn = float(np.linalg.norm(v)) or 1.0
            cos = float(np.dot(q, v) / (qn * vn))
            penalty = 0.0
            if (r["char_len"] or 0) < 1500:
                penalty += 0.15
            if (r["ocr_ratio"] or 0.0) > 0.10:
                penalty += 0.15
            scored.append((cos - penalty, dict(r)))
        scored.sort(key=lambda x: x[0], reverse=True)
        out: list[dict] = []
        seen: set[str] = set()
        for score, r in scored:
            base = _re.sub(r"_\d+$", "", os.path.splitext(os.path.basename(r["path"]))[0])
            if base in seen:
                continue
            seen.add(base)
            r["score"] = score
            out.append(r)
            if len(out) >= k:
                break
        return out
```

(`np` is already imported at the top of index.py.)

- [ ] **Step 4: Run test to verify it passes**

Run: same as Step 2. Expected: PASS (4 passed)

- [ ] **Step 5: Commit**

```bash
git -C "C:/firm-briefs-p2" add icharlotte_core/firm_briefs/index.py tests/test_firm_briefs/test_style_candidates.py
git -C "C:/firm-briefs-p2" commit -m "feat(firm_briefs): style_candidates cosine ranking (side-strict, quality+dedup guard)"
```

---

### Task 3: `firm_briefs/style.py` — select_exemplars + excerpt extraction

**Files:**
- Create: `icharlotte_core/firm_briefs/style.py`
- Test: `tests/test_firm_briefs/test_style_select.py`

- [ ] **Step 1: Write the failing test**

```python
# tests/test_firm_briefs/test_style_select.py
from types import SimpleNamespace
from icharlotte_core.firm_briefs import style


class FakeIndex:
    def __init__(self, cands): self._c = cands
    def style_candidates(self, qv, *, motion_type, side, k=3): return self._c


class FakeEmb:
    dim = 384
    def encode(self, texts):
        import numpy as np
        return np.ones((len(texts), 384), dtype=np.float32)


META = SimpleNamespace(relief_requested="compel further responses",
                       principal_arguments=["failed to meet and confer"])


def test_select_returns_trimmed_excerpts(tmp_path):
    idx = FakeIndex([{"path": "a.pdf"}, {"path": "b.pdf"}])
    texts = {"a.pdf": "CAPTION...\nARGUMENT\nThe motion fails because " + "x" * 50,
             "b.pdf": "no heading here just prose " + "y" * 50}
    out = style.select_exemplars("compel", "opposition", META, index=idx, embedder=FakeEmb(),
                                 extract_fn=lambda p: texts[p],
                                 cache_dir=str(tmp_path), max_chars=200)
    assert len(out) == 2
    # 'a' is trimmed to start at the ARGUMENT heading
    assert out[0].startswith("ARGUMENT")
    assert all(len(t) <= 210 for t in out)


def test_select_empty_when_no_index():
    assert style.select_exemplars("compel", "opposition", META, index=None) == []


def test_select_empty_when_no_profile(tmp_path):
    idx = FakeIndex([{"path": "a.pdf"}])
    blank = SimpleNamespace(relief_requested="", principal_arguments=[])
    out = style.select_exemplars("compel", "opposition", blank, index=idx, embedder=FakeEmb(),
                                 extract_fn=lambda p: "text", cache_dir=str(tmp_path))
    assert out == []   # no query profile -> no selection
```

- [ ] **Step 2: Run test to verify it fails**

Run: `Set-Location 'C:\firm-briefs-p2'; & 'C:\geminiterminal2\.venv\Scripts\python.exe' -m pytest tests/test_firm_briefs/test_style_select.py -v`
Expected: FAIL (`ModuleNotFoundError`)

- [ ] **Step 3: Implement**

```python
# icharlotte_core/firm_briefs/style.py
"""Draft-time style-exemplar selection from the firm brief library.

Embeds the current motion's issue profile, finds the most similar past briefs of
the same motion type AND side, and returns trimmed argument-section excerpts to
feed the drafter as style/voice models. Degrades to [] when the index is absent.
"""
from __future__ import annotations

import hashlib
import logging
import os
import re
from typing import Callable, List, Optional

logger = logging.getLogger(__name__)

_ARG_HEADING_RE = re.compile(r"(?im)^\s*(?:[IVXLC]+\.\s*)?(ARGUMENT|MEMORANDUM OF POINTS|LEGAL ARGUMENT|DISCUSSION)\b")


def _trim_to_argument(text: str, max_chars: int) -> str:
    """Start at the ARGUMENT/Memorandum heading if present (skips caption/TOC),
    then cap to max_chars."""
    text = text or ""
    m = _ARG_HEADING_RE.search(text)
    if m:
        text = text[m.start():]
    return text[:max_chars].strip()


def _default_excerpt_extract(path: str) -> str:
    from icharlotte_core.document_processor import DocumentProcessor
    try:
        return DocumentProcessor().extract_text(path, ocr_enabled=False).text or ""
    except Exception:
        logger.warning("style excerpt extract failed: %s", path, exc_info=True)
        return ""


def _cache_key(path: str) -> str:
    try:
        mtime = os.path.getmtime(path)
    except OSError:
        mtime = 0.0
    return hashlib.sha1(f"{os.path.abspath(path)}|{mtime}".encode("utf-8")).hexdigest()


def _excerpt(path: str, *, cache_dir: str, extract_fn: Callable[[str], str],
             max_chars: int) -> str:
    os.makedirs(cache_dir, exist_ok=True)
    cache_path = os.path.join(cache_dir, f"{_cache_key(path)}.txt")
    if os.path.exists(cache_path):
        try:
            with open(cache_path, "r", encoding="utf-8") as f:
                return f.read()
        except OSError:
            pass
    excerpt = _trim_to_argument(extract_fn(path), max_chars)
    if excerpt:
        try:
            with open(cache_path, "w", encoding="utf-8") as f:
                f.write(excerpt)
        except OSError:
            logger.warning("could not cache style excerpt: %s", cache_path, exc_info=True)
    return excerpt


def select_exemplars(
    motion_type: str,
    side: str,
    motion_metadata,
    *,
    k: int = 3,
    max_chars: int = 8000,
    index=None,
    embedder=None,
    extract_fn: Optional[Callable[[str], str]] = None,
    cache_dir: Optional[str] = None,
) -> List[str]:
    """Return up to k trimmed style excerpts most similar to the current motion."""
    from .factory import make_index, DATA_DIR
    from .profile import profile_from_metadata
    from .embedding import get_embedder

    if index is None:
        index = make_index()
    if index is None:
        return []
    prof = profile_from_metadata(motion_metadata)
    if not prof.strip():
        return []
    try:
        embedder = embedder or get_embedder()
        qv = embedder.encode([prof])[0]
        cands = index.style_candidates(qv, motion_type=motion_type, side=side, k=k)
    except Exception:
        logger.warning("style candidate selection failed", exc_info=True)
        return []
    extract_fn = extract_fn or _default_excerpt_extract
    cache_dir = cache_dir or os.path.join(DATA_DIR, ".cache", "style")
    out: List[str] = []
    for c in cands:
        txt = _excerpt(c["path"], cache_dir=cache_dir, extract_fn=extract_fn, max_chars=max_chars)
        if txt.strip():
            out.append(txt)
    return out
```

- [ ] **Step 4: Run test to verify it passes**

Run: same as Step 2. Expected: PASS (3 passed)

- [ ] **Step 5: Commit**

```bash
git -C "C:/firm-briefs-p2" add icharlotte_core/firm_briefs/style.py tests/test_firm_briefs/test_style_select.py
git -C "C:/firm-briefs-p2" commit -m "feat(firm_briefs): style.select_exemplars (embed motion profile, trim argument section)"
```

---

### Task 4: Wire firm style exemplars into oppose_motion

**Files:**
- Modify: `icharlotte_core/ui/wizard/pages/oppose_motion_page.py` (exemplar block ~lines 606-636; passed to drafter at ~711 as `style_exemplars=exemplar_texts`)
- Test: `tests/test_firm_briefs/test_oppose_style_wiring.py`

- [ ] **Step 1: Write the failing test**

```python
# tests/test_firm_briefs/test_oppose_style_wiring.py
from types import SimpleNamespace
from icharlotte_core.ui.wizard.pages import oppose_motion_page as omp


def test_firm_style_exemplars_helper_guarded(monkeypatch):
    monkeypatch.setattr("icharlotte_core.firm_briefs.style.select_exemplars",
                        lambda mt, side, meta, **k: ["FIRM STYLE EXCERPT"])
    meta = SimpleNamespace(motion_type="compel", relief_requested="x", principal_arguments=["y"])
    out = omp._firm_style_exemplars("compel", "opposition", meta)
    assert out == ["FIRM STYLE EXCERPT"]


def test_firm_style_exemplars_swallows_errors(monkeypatch):
    def boom(*a, **k): raise RuntimeError("no index")
    monkeypatch.setattr("icharlotte_core.firm_briefs.style.select_exemplars", boom)
    out = omp._firm_style_exemplars("compel", "opposition", SimpleNamespace())
    assert out == []
```

- [ ] **Step 2: Run test to verify it fails**

Run: `Set-Location 'C:\firm-briefs-p2'; & 'C:\geminiterminal2\.venv\Scripts\python.exe' -m pytest tests/test_firm_briefs/test_oppose_style_wiring.py -v`
Expected: FAIL (`AttributeError: ... '_firm_style_exemplars'`)

- [ ] **Step 3: Implement** — add a module-level helper near `_make_firm_provider` in `oppose_motion_page.py`:

```python
def _firm_style_exemplars(motion_type, side, metadata):
    """Firm-library style excerpts most similar to this motion; [] if no index."""
    try:
        from icharlotte_core.firm_briefs import style
        return style.select_exemplars(motion_type, side, metadata) or []
    except Exception:
        return []
```

Then in the worker's exemplar block, AFTER `exemplar_texts` is built from the registry and BEFORE it is passed to the drafter, prepend firm picks (cap total to 3 for the drafter's token budget):

```python
            firm_style = _firm_style_exemplars(metadata.motion_type, "opposition", metadata)
            if firm_style:
                self.progress.emit(f"  + {len(firm_style)} firm-library style sample(s).")
            exemplar_texts = (firm_style + exemplar_texts)[:3]
```

(Insert this right after the existing `if matches: ... else: ...` progress block, so `exemplar_texts` already holds any manual-registry picks. The later `style_exemplars=exemplar_texts` call is unchanged.)

- [ ] **Step 4: Run test to verify it passes**

Run: same as Step 2. Expected: PASS (2 passed). (PySide6 imports fine in this env; if collection errors due to a running app, stop it and re-run.)

- [ ] **Step 5: Commit**

```bash
git -C "C:/firm-briefs-p2" add icharlotte_core/ui/wizard/pages/oppose_motion_page.py tests/test_firm_briefs/test_oppose_style_wiring.py
git -C "C:/firm-briefs-p2" commit -m "feat(oppose_motion): feed firm-library style exemplars to the drafter (guarded)"
```

---

### Task 5: Wire firm style exemplars into generate_motion

**Files:**
- Modify: `icharlotte_core/ui/wizard/pages/generate_motion_page.py` (exemplar usage ~lines 559-572, `load_exemplars(...)` → `style_exemplars=exemplars`)
- Test: `tests/test_firm_briefs/test_generate_style_wiring.py`

- [ ] **Step 1: Write the failing test**

```python
# tests/test_firm_briefs/test_generate_style_wiring.py
from types import SimpleNamespace
from icharlotte_core.ui.wizard.pages import generate_motion_page as gmp


def test_generate_firm_style_helper(monkeypatch):
    monkeypatch.setattr("icharlotte_core.firm_briefs.style.select_exemplars",
                        lambda mt, side, meta, **k: ["MOVING STYLE EXCERPT"])
    out = gmp._firm_style_exemplars("msj", "moving", SimpleNamespace(relief_requested="x", principal_arguments=[]))
    assert out == ["MOVING STYLE EXCERPT"]
```

- [ ] **Step 2: Run test to verify it fails**

Run: `Set-Location 'C:\firm-briefs-p2'; & 'C:\geminiterminal2\.venv\Scripts\python.exe' -m pytest tests/test_firm_briefs/test_generate_style_wiring.py -v`
Expected: FAIL (`AttributeError: ... '_firm_style_exemplars'`)

- [ ] **Step 3: Implement** — in `generate_motion_page.py`, add the same helper (either define it identically or import it from `oppose_motion_page` alongside the existing `_make_firm_provider` import — match whatever pattern that file already uses for `_make_firm_provider`):

```python
def _firm_style_exemplars(motion_type, side, metadata):
    try:
        from icharlotte_core.firm_briefs import style
        return style.select_exemplars(motion_type, side, metadata) or []
    except Exception:
        return []
```

Then at the exemplar site (after `exemplars = load_exemplars(...)`), prepend firm picks with `side="moving"`, using the motion type id and the analyzer metadata available in that scope:

```python
            firm_style = _firm_style_exemplars(
                self.settings.get("motion_type_id") or getattr(metadata, "motion_type", ""),
                "moving", metadata)
            if firm_style:
                self.progress.emit(f"Using {len(firm_style)} firm-library style sample(s).")
            exemplars = (firm_style + exemplars)[:3]
```

If the metadata object isn't named `metadata` in that scope, use whatever the analyzer-result variable is (it has `relief_requested` + `principal_arguments`). Confirm the variable name by reading the surrounding code.

- [ ] **Step 4: Run test to verify it passes**

Run: same as Step 2. Expected: PASS

- [ ] **Step 5: Commit**

```bash
git -C "C:/firm-briefs-p2" add icharlotte_core/ui/wizard/pages/generate_motion_page.py tests/test_firm_briefs/test_generate_style_wiring.py
git -C "C:/firm-briefs-p2" commit -m "feat(generate_motion): feed firm-library style exemplars to the drafter (side=moving)"
```

---

### Task 6: Full regression sweep

**Files:** none (verification only)

- [ ] **Step 1: Run the firm_briefs + opposition + wizard suites**

Run:
```
Set-Location 'C:\firm-briefs-p2'; & 'C:\geminiterminal2\.venv\Scripts\python.exe' -m pytest tests/test_firm_briefs tests/test_opposition tests/test_wizard -q
```
Expected: all `tests/test_firm_briefs` pass; no NEW failures in opposition/wizard. Pre-existing Qt-fixture collection errors in unrelated wizard files are not regressions — classify any failure as "mine" (touches firm_briefs/style/oppose/generate) vs "pre-existing".

- [ ] **Step 2: Sanity-check live style selection against the real index** (the index is already built from Phase 1; no rebuild needed — the profile vectors are present):

Run:
```
Set-Location 'C:\firm-briefs-p2'; & 'C:\geminiterminal2\.venv\Scripts\python.exe' -c "import sys; sys.path.insert(0,r'C:\firm-briefs-p2'); from types import SimpleNamespace; from icharlotte_core.firm_briefs import style; m=SimpleNamespace(motion_type='compel', relief_requested='compel further responses to requests for production', principal_arguments=['plaintiff failed to meet and confer','responses are evasive and incomplete']); ex=style.select_exemplars('compel','opposition',m); print('exemplars:',len(ex)); [print('---',len(t),'chars; starts:',t[:80].replace(chr(10),' ')) for t in ex]"
```
Expected: 1-3 excerpts returned (non-empty), each a trimmed argument-section excerpt. Report the output.

- [ ] **Step 3: Commit** (if any test-only fixups were needed; otherwise skip)

---

## Self-review notes (for the implementer)
- **Additive guarantee:** both drafters keep their manual-registry exemplars; firm picks are prepended and the list is capped at 3. If the index is absent, `_firm_style_exemplars` returns [] and behavior is today's.
- **Side is strict for style** (moving vs opposition vs reply) — unlike authority, which ignores side.
- **No rebuild needed:** Phase 1 already wrote profile vectors to `profiles.f16`; Task 2 reads them.
- **Deferred to Phase 3:** citation-panel "flag both" UI + Workbench Sample Library tab; harvest case-name cleanup; per-proposition vectors.
