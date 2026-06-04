# Firm Brief Library — Phase 3B (UI Surfacing) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development. Steps use checkbox (`- [ ]`) syntax.

**Goal:** Surface provenance in the citation panel ("From your brief" badge + verification tier + swap-to-corpus-alternative) and add a Workbench "Sample Library" tab to manage roots / re-index / view stats.

**Architecture:** `CitationVerification` (what the panel renders) gains provenance fields, populated after verify by joining to the `RetrievedAuthority` pool by normalized cite. The panel renders the new fields and exposes alternatives as clickable anchors that swap the cite in the draft body. The Workbench tab mirrors the existing `StyleExamplesTab` pattern.

**Tech Stack:** PySide6 (pytest-qt), pytest. Tests: `C:\geminiterminal2\.venv\Scripts\python.exe`. **Work in worktree `C:\firm-briefs-p3` on branch `feat/firm-briefs-phase3` (continues Phase 3A).**

**Environment:** shell cwd RESETS each command — begin every PowerShell command with `Set-Location 'C:\firm-briefs-p3';`. Tests: `Set-Location 'C:\firm-briefs-p3'; & 'C:\geminiterminal2\.venv\Scripts\python.exe' -m pytest <path> -v`. Qt tests use `pytest.importorskip("PySide6")` and assert `not widget.isHidden()` (not isVisible). PowerShell only; no bash `cd && `. git via `git -C "C:/firm-briefs-p3" ...`.

---

### Task 1 (3.1): Provenance fields on CitationVerification

**Files:**
- Modify: `icharlotte_core/opposition/models.py` (`CitationVerification`)
- Test: `tests/test_firm_briefs/test_citation_provenance_fields.py`

- [ ] **Step 1: Write the failing test**

```python
# tests/test_firm_briefs/test_citation_provenance_fields.py
from icharlotte_core.opposition.models import CitationVerification

def test_provenance_fields_default():
    c = CitationVerification()
    assert c.source == ""
    assert c.source_brief == ""
    assert c.firm_verification == ""
    assert c.alternatives == []

def test_provenance_roundtrips_through_dict():
    c = CitationVerification(citation_text="Townsend v. Superior Court (1998) 61 Cal.App.4th 1431",
                             source="firm", source_brief=r"C:\lib\Oppositions\Motion to Compel\x.pdf",
                             firm_verification="local",
                             alternatives=[{"case_name": "Leko v. Cornerstone", "citation": "86 Cal.App.4th 1109"}])
    d = c.to_dict()
    c2 = CitationVerification.from_dict(d)
    assert c2.source == "firm"
    assert c2.source_brief.endswith("x.pdf")
    assert c2.firm_verification == "local"
    assert c2.alternatives[0]["citation"] == "86 Cal.App.4th 1109"
```

- [ ] **Step 2: Run to verify fail**

Run: `Set-Location 'C:\firm-briefs-p3'; & 'C:\geminiterminal2\.venv\Scripts\python.exe' -m pytest tests/test_firm_briefs/test_citation_provenance_fields.py -v`
Expected: FAIL (`source` etc. not attributes)

- [ ] **Step 3: Implement** — in `CitationVerification` (models.py) add fields after the existing ones:
```python
    # Firm-library provenance (Phase 3). Defaults keep corpus-only behavior.
    source: str = ""                  # "firm" | "corpus" | ""
    source_brief: str = ""            # path/label of the firm brief it came from
    firm_verification: str = ""       # "local" | "courtlistener" | "unverified_firm" | ""
    alternatives: list = field(default_factory=list)  # [{case_name, citation, year?}]
```
In `from_dict`, add: `source=data.get("source", ""), source_brief=data.get("source_brief", ""), firm_verification=data.get("firm_verification", ""), alternatives=list(data.get("alternatives", []) or []),`. If `to_dict` is a manual dict (not asdict), add these keys; if it uses `asdict`, they're automatic.

- [ ] **Step 4: Run to verify pass** — same command. Expected: PASS (2).

- [ ] **Step 5: Commit**
```bash
git -C "C:/firm-briefs-p3" add icharlotte_core/opposition/models.py tests/test_firm_briefs/test_citation_provenance_fields.py
git -C "C:/firm-briefs-p3" commit -m "feat(opposition): CitationVerification provenance fields (source/source_brief/firm_verification/alternatives)"
```

---

### Task 2 (3.1): Join verified citations to the firm authority pool

**Files:**
- Create: `icharlotte_core/firm_briefs/provenance.py`
- Test: `tests/test_firm_briefs/test_provenance_join.py`

- [ ] **Step 1: Write the failing test**

```python
# tests/test_firm_briefs/test_provenance_join.py
from icharlotte_core.opposition.models import CitationVerification, RetrievedAuthority
from icharlotte_core.firm_briefs.provenance import attach_firm_provenance

def test_attach_marks_firm_and_alternatives():
    cits = [CitationVerification(citation_text="Townsend v. Superior Court (1998) 61 Cal.App.4th 1431",
                                 normalized_citation="61 Cal.App.4th 1431")]
    pool = [RetrievedAuthority(case_name="Townsend v. Superior Court", citation="61 Cal.App.4th 1431",
                              source="firm", verification="local",
                              source_brief=r"C:\lib\x.pdf",
                              alternatives=[RetrievedAuthority(case_name="Leko v. Cornerstone",
                                                              citation="86 Cal.App.4th 1109", source="corpus")])]
    attach_firm_provenance(cits, pool)
    assert cits[0].source == "firm"
    assert cits[0].firm_verification == "local"
    assert cits[0].source_brief.endswith("x.pdf")
    assert cits[0].alternatives[0]["citation"] == "86 Cal.App.4th 1109"

def test_no_match_leaves_citation_untouched():
    cits = [CitationVerification(citation_text="Other v. Case", normalized_citation="1 Cal.5th 1")]
    attach_firm_provenance(cits, [])
    assert cits[0].source == ""
```

- [ ] **Step 2: Run to verify fail** — Expected: `ModuleNotFoundError`.

- [ ] **Step 3: Implement**
```python
# icharlotte_core/firm_briefs/provenance.py
"""Join verified citations to the firm RetrievedAuthority pool by normalized cite,
copying provenance (source / source_brief / verification / alternatives) onto the
CitationVerification records the output panel renders."""
from __future__ import annotations

import re
from typing import List


def _norm(cite: str) -> str:
    return re.sub(r"\s+", "", (cite or "")).lower()


def attach_firm_provenance(citations: List, retrieved: List) -> None:
    """Mutates each CitationVerification in-place when a firm authority matches."""
    by_cite = {}
    for ra in (retrieved or []):
        if getattr(ra, "source", "") == "firm":
            by_cite[_norm(getattr(ra, "citation", ""))] = ra
    for c in (citations or []):
        key = _norm(getattr(c, "normalized_citation", "") or getattr(c, "citation_text", ""))
        ra = by_cite.get(key)
        if not ra:
            continue
        c.source = "firm"
        c.source_brief = getattr(ra, "source_brief", "") or ""
        c.firm_verification = getattr(ra, "verification", "") or ""
        alts = []
        for a in (getattr(ra, "alternatives", []) or []):
            alts.append({"case_name": getattr(a, "case_name", ""),
                         "citation": getattr(a, "citation", ""),
                         "year": getattr(a, "year", "")})
        c.alternatives = alts
```

- [ ] **Step 4: Run to verify pass** — Expected: PASS (2).

- [ ] **Step 5: Commit**
```bash
git -C "C:/firm-briefs-p3" add icharlotte_core/firm_briefs/provenance.py tests/test_firm_briefs/test_provenance_join.py
git -C "C:/firm-briefs-p3" commit -m "feat(firm_briefs): attach_firm_provenance joins verified cites to the firm authority pool"
```

---

### Task 3 (3.1): Wire the join into both workers

**Files:**
- Modify: `icharlotte_core/ui/wizard/pages/oppose_motion_page.py`, `generate_motion_page.py`
- Test: `tests/test_firm_briefs/test_provenance_wiring.py`

- [ ] **Step 1: Write the failing test**

```python
# tests/test_firm_briefs/test_provenance_wiring.py
from icharlotte_core.ui.wizard.pages import oppose_motion_page as omp
from icharlotte_core.ui.wizard.pages import generate_motion_page as gmp

def test_pages_expose_attach_provenance():
    assert hasattr(omp, "attach_firm_provenance")
    assert hasattr(gmp, "attach_firm_provenance")
```

- [ ] **Step 2: Run to verify fail** — Expected: `AttributeError`.

- [ ] **Step 3: Implement** — in BOTH pages add the import:
```python
from icharlotte_core.firm_briefs.provenance import attach_firm_provenance
```
Then in each worker, after the citations are verified into the `DraftDocument` (i.e. after the verifier produces the `list[CitationVerification]` and before/at the point the `DraftDocument` is built or emitted) AND where the `retrieved` RetrievedAuthority pool is still in scope, call:
```python
            attach_firm_provenance(draft.citations, retrieved)
```
READ each worker to find where `draft.citations` (or the verified-citations list) and `retrieved` coexist; place the call there. If `retrieved` isn't in scope at the DraftDocument build site, thread it through (it's the `research_arguments(...)` result already held in a local). Guard with `try/except` so a join failure never breaks the draft.

- [ ] **Step 4: Run to verify pass + regression**

Run: `Set-Location 'C:\firm-briefs-p3'; & 'C:\geminiterminal2\.venv\Scripts\python.exe' -m pytest tests/test_firm_briefs tests/test_wizard -q`
Expected: PASS (new test + no wizard regressions).

- [ ] **Step 5: Commit**
```bash
git -C "C:/firm-briefs-p3" add icharlotte_core/ui/wizard/pages/oppose_motion_page.py icharlotte_core/ui/wizard/pages/generate_motion_page.py tests/test_firm_briefs/test_provenance_wiring.py
git -C "C:/firm-briefs-p3" commit -m "feat(wizard): attach firm provenance to verified citations in both drafters"
```

---

### Task 4 (3.1): Render provenance in the citation panel

**Files:**
- Modify: `icharlotte_core/ui/wizard/pages/citation_review.py` (`_citation_body_html`)
- Test: `tests/test_firm_briefs/test_panel_provenance_html.py`

- [ ] **Step 1: Write the failing test**

```python
# tests/test_firm_briefs/test_panel_provenance_html.py
from icharlotte_core.opposition.models import CitationVerification
from icharlotte_core.ui.wizard.pages.citation_review import _citation_body_html

def test_firm_badge_and_tier_rendered():
    c = CitationVerification(citation_text="Townsend (1998) 61 Cal.App.4th 1431", verdict="SUPPORTED",
                             source="firm", source_brief=r"C:\lib\Oppositions\Motion to Compel\Smith Opp.pdf",
                             firm_verification="local",
                             alternatives=[{"case_name": "Leko v. Cornerstone", "citation": "86 Cal.App.4th 1109"}])
    html = _citation_body_html(c, "SUPPORTED")
    assert "From your brief" in html
    assert "Smith Opp" in html            # basename of source_brief shown
    assert "Leko v. Cornerstone" in html  # alternative listed

def test_unverified_firm_amber_warning():
    c = CitationVerification(citation_text="Smith v. Jones (2024) 999 F.3d 1", verdict="UNVERIFIED",
                             source="firm", source_brief=r"C:\lib\x.pdf", firm_verification="unverified_firm")
    html = _citation_body_html(c, "UNVERIFIED")
    assert "from firm brief" in html.lower()
    assert "not independently verified" in html.lower()

def test_corpus_citation_no_firm_badge():
    c = CitationVerification(citation_text="X v. Y (2000) 1 Cal.5th 1", verdict="SUPPORTED")
    assert "From your brief" not in _citation_body_html(c, "SUPPORTED")
```

- [ ] **Step 2: Run to verify fail** — Expected: FAIL (no firm badge in HTML).

- [ ] **Step 3: Implement** — at the TOP of `_citation_body_html`, before the existing verdict blocks, prepend a provenance block built from the new fields:
```python
    import os as _os
    parts: list[str] = []
    if getattr(citation, "source", "") == "firm":
        label = _os.path.splitext(_os.path.basename(citation.source_brief or ""))[0] or "your brief"
        # strip the "Matter__" prefix the library uses, if present
        label = label.split("__", 1)[-1]
        fv = getattr(citation, "firm_verification", "")
        if fv == "unverified_firm":
            parts.append(
                f"<p style='color:#b06000;'>⚠ <b>From firm brief</b> "
                f"(<i>{html.escape(label)}</i>) — not independently verified.</p>"
            )
        else:
            tier = "verified locally" if fv == "local" else (
                   "verified via CourtListener" if fv == "courtlistener" else "")
            tail = f" — {tier}" if tier else ""
            parts.append(
                f"<p style='color:#188038;'>\U0001F4C1 <b>From your brief:</b> "
                f"<i>{html.escape(label)}</i>{tail}</p>"
            )
        alts = getattr(citation, "alternatives", []) or []
        if alts:
            links = []
            for i, a in enumerate(alts):
                nm = html.escape((a.get('case_name') or '').strip())
                ct = html.escape((a.get('citation') or '').strip())
                links.append(f"<li>{nm}, {ct} &nbsp;<a href='altswap:{i}'>Use this instead</a></li>")
            parts.append("<p><b>Corpus alternatives:</b></p><ul>" + "".join(links) + "</ul>")
```
Then make the existing function append its current `parts` to THIS list rather than starting a fresh one — i.e., keep the existing body-building code but have it extend the same `parts` list (the existing function already builds a `parts` list; merge: declare the provenance `parts` first, then let the rest of the function append to it, and keep the final `return "\n".join(parts)`).

- [ ] **Step 4: Run to verify pass** — Expected: PASS (3).

- [ ] **Step 5: Commit**
```bash
git -C "C:/firm-briefs-p3" add icharlotte_core/ui/wizard/pages/citation_review.py tests/test_firm_briefs/test_panel_provenance_html.py
git -C "C:/firm-briefs-p3" commit -m "feat(citation-panel): render firm-brief provenance badge, verification tier, alternatives"
```

---

### Task 5 (3.1): Make alternatives swap the cite in the draft body

**Files:**
- Modify: `icharlotte_core/ui/wizard/pages/citation_review.py` (panel anchor handling + a pure swap helper)
- Test: `tests/test_firm_briefs/test_alt_swap.py`

- [ ] **Step 1: Write the failing test**

```python
# tests/test_firm_briefs/test_alt_swap.py
from icharlotte_core.ui.wizard.pages.citation_review import apply_alternative_to_body

def test_swap_replaces_cite_text():
    body = "As held in Townsend v. Superior Court (1998) 61 Cal.App.4th 1431, the duty applies."
    new = apply_alternative_to_body(
        body, old_cite="61 Cal.App.4th 1431",
        alternative={"case_name": "Leko v. Cornerstone Building Inspection Service",
                     "citation": "86 Cal.App.4th 1109", "year": "2001"})
    assert "86 Cal.App.4th 1109" in new
    assert "Leko v. Cornerstone" in new
    assert "61 Cal.App.4th 1431" not in new

def test_swap_noop_when_cite_absent():
    body = "No citation here."
    assert apply_alternative_to_body(body, "1 Cal.5th 1", {"citation": "2 Cal.5th 2"}) == body
```

- [ ] **Step 2: Run to verify fail** — Expected: `ImportError`.

- [ ] **Step 3: Implement** — add the pure helper to `citation_review.py`:
```python
def apply_alternative_to_body(body_text: str, old_cite: str, alternative: dict) -> str:
    """Replace the firm authority's reporter cite (and the case name preceding it)
    in the draft body with the chosen corpus alternative. Best-effort string swap:
    replaces the reporter cite, and if the case name + cite appear together as
    'Name ... old_cite', swaps the whole 'Name (year) cite' phrase."""
    import re
    if not old_cite or old_cite not in (body_text or ""):
        return body_text
    name = (alternative.get("case_name") or "").strip()
    cite = (alternative.get("citation") or "").strip()
    year = (alternative.get("year") or "").strip()
    new_phrase = name
    if year:
        new_phrase = f"{name} ({year})" if name else f"({year})"
    new_phrase = f"{new_phrase} {cite}".strip() if new_phrase else cite
    # Swap "<Name up to ~10 words> <old_cite>" if a name precedes the cite; else just the cite.
    pat = re.compile(r"([A-Z][\w.,'&\- ]{0,80}?)\s*\(?\d{0,4}\)?\s*" + re.escape(old_cite))
    if pat.search(body_text):
        return pat.sub(new_phrase, body_text, count=1)
    return body_text.replace(old_cite, cite, 1)
```
Then in `CitationDetailPanel`: the body is shown in a `QTextBrowser`. Set `setOpenLinks(False)` and connect `anchorClicked` to a handler that, for an `altswap:<i>` URL, calls `apply_alternative_to_body(self._output_page.draft.body_text, self.citation.normalized_citation or <reporter cite from citation_text>, self.citation.alternatives[i])`, updates the draft body, re-renders (`self._output_page.editor.setHtml(_render_draft_html(...))`), and re-parses/re-shows. READ the panel class to find its handle to the output page / draft (it's constructed with a parent or reference); if the panel lacks a back-reference to the draft, add one (pass the output page in, or emit a `swap_requested(int)` signal the output page connects to). Keep it minimal: a `swap_requested = Signal(object, int)` emitting `(citation, alt_index)` that `CitationReviewOutputPage` connects to a method performing the body edit + re-render is the cleanest and most testable.

- [ ] **Step 4: Run to verify pass + qt smoke**

Run: `Set-Location 'C:\firm-briefs-p3'; & 'C:\geminiterminal2\.venv\Scripts\python.exe' -m pytest tests/test_firm_briefs/test_alt_swap.py tests/test_wizard -q`
Expected: PASS (swap helper + no wizard regressions).

- [ ] **Step 5: Commit**
```bash
git -C "C:/firm-briefs-p3" add icharlotte_core/ui/wizard/pages/citation_review.py tests/test_firm_briefs/test_alt_swap.py
git -C "C:/firm-briefs-p3" commit -m "feat(citation-panel): swap a cite to a corpus alternative from the panel"
```

---

### Task 6 (3.2): Workbench "Sample Library" tab

**Files:**
- Create: `icharlotte_core/ui/dialogs_sample_library.py`
- Modify: `icharlotte_core/ui/dialogs.py` (add `_refresh_sample_library_tab`, mirroring `_refresh_style_examples_tab`)
- Test: `tests/test_firm_briefs/test_sample_library_tab.py`

- [ ] **Step 1: Write the failing test**

```python
# tests/test_firm_briefs/test_sample_library_tab.py
import pytest
pytest.importorskip("PySide6")
from icharlotte_core.ui.dialogs_sample_library import SampleLibraryTab

def test_roots_add_remove_persist(tmp_path, qtbot):
    cfg = str(tmp_path / "roots.json")
    tab = SampleLibraryTab(roots_config_path=cfg)
    qtbot.addWidget(tab)
    tab.add_root_programmatic(r"C:\lib\5800")
    tab.add_root_programmatic(r"C:\lib\3800")
    assert set(tab.roots()) == {r"C:\lib\5800", r"C:\lib\3800"}
    # persisted -> a fresh tab reads them back
    tab2 = SampleLibraryTab(roots_config_path=cfg)
    qtbot.addWidget(tab2)
    assert set(tab2.roots()) == {r"C:\lib\5800", r"C:\lib\3800"}
    tab.remove_root_programmatic(r"C:\lib\3800")
    assert tab.roots() == [r"C:\lib\5800"]

def test_stats_text_from_index(tmp_path, qtbot):
    tab = SampleLibraryTab(roots_config_path=str(tmp_path / "roots.json"))
    qtbot.addWidget(tab)
    class FakeIdx:
        def stats(self): return {"briefs": 547, "citations": 3349}
    txt = tab.stats_text(FakeIdx())
    assert "547" in txt and "3349" in txt

def test_stats_text_no_index(tmp_path, qtbot):
    tab = SampleLibraryTab(roots_config_path=str(tmp_path / "roots.json"))
    qtbot.addWidget(tab)
    assert "not built" in tab.stats_text(None).lower()
```

- [ ] **Step 2: Run to verify fail** — Expected: `ModuleNotFoundError`.

- [ ] **Step 3: Implement** — create `dialogs_sample_library.py` modeled on `dialogs_style_examples.py`:
- `SampleLibraryTab(QWidget)` ctor takes `roots_config_path` (default = `Scripts/prompts/firm_briefs/roots.json`).
- `roots() -> list[str]`; `add_root_programmatic(path)`, `remove_root_programmatic(path)` (persist to the JSON: `{"roots": [...]}`; load on init, seeding from `config.FIRM_BRIEFS_ROOTS` when the file is absent).
- `stats_text(index) -> str`: returns "Index not built." when `index is None`, else `f"{s['briefs']} briefs, {s['citations']} citations"` (plus a per-type line if cheap).
- UI: a roots list with Add (QFileDialog.getExistingDirectory) / Remove buttons, a stats label (populated from `factory.make_index()`), and a **"Re-index" button** that runs ingestion on a `QThread` (worker calls `ingest_root` for each root with an `OnnxEmbedder`; emit progress to a small log label; an "OCR image-only" checkbox is recorded for a follow-up but may be a no-op stub if the OCR path is out of scope here — keep the button working for the native pass). Use the existing wizard QThread patterns. The programmatic methods above must NOT require the event loop.
- In `dialogs.py`: add `self._sample_library_tab` + `_refresh_sample_library_tab()` mirroring `_refresh_style_examples_tab()` (add the tab labeled "Sample Library"), and call it where `_refresh_style_examples_tab()`/`_refresh_motion_types_tab()` are called.

- [ ] **Step 4: Run to verify pass**

Run: `Set-Location 'C:\firm-briefs-p3'; & 'C:\geminiterminal2\.venv\Scripts\python.exe' -m pytest tests/test_firm_briefs/test_sample_library_tab.py -v`
Expected: PASS (3).

- [ ] **Step 5: Commit**
```bash
git -C "C:/firm-briefs-p3" add icharlotte_core/ui/dialogs_sample_library.py icharlotte_core/ui/dialogs.py tests/test_firm_briefs/test_sample_library_tab.py
git -C "C:/firm-briefs-p3" commit -m "feat(workbench): Sample Library tab (manage roots, re-index, stats)"
```

---

### Task 7: Full Phase 3 regression

- [ ] **Step 1:** `Set-Location 'C:\firm-briefs-p3'; & 'C:\geminiterminal2\.venv\Scripts\python.exe' -m pytest tests/test_firm_briefs tests/test_opposition tests/test_motion_generation tests/test_wizard -q`
Expected: all firm_briefs pass; no NEW opposition/generate/wizard failures. Report counts; classify any failure as mine vs pre-existing Qt-collection.

---

## Self-review notes
- Provenance fields default empty → corpus-only behavior unchanged; the join is guarded.
- The panel provenance block PREPENDS to the existing verdict body (don't drop the existing content — extend the same `parts` list).
- `apply_alternative_to_body` is a pure, tested string op; the panel wires it via a `swap_requested` signal the output page handles (keeps UI glue thin + testable).
- Workbench tab's programmatic API (roots/stats) is event-loop-free for tests; the Re-index QThread is exercised only in the live app.
- After merge: re-index the real library (`--rebuild`) so provenance/full_text/prop-vectors/clean-names are all populated; restart iCharlotte.
