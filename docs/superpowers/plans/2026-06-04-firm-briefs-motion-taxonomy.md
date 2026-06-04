# Firm Brief Library — Phase 2.5 (Motion-Type Taxonomy) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development. Steps use checkbox (`- [ ]`) syntax.

**Goal:** One canonical motion-type vocabulary + `normalize_motion_type()` used by ingest, Oppose-a-Motion, and Generate-a-Motion, so firm sample/authority matching fires reliably; re-tag the "other" bucket into real subtypes; and register the common types (light) so they appear in the Generate dropdown.

**Architecture:** New `motion_taxonomy.py` is the single source of truth (ordered id + keyword patterns + `normalize_motion_type`). `path_meta` delegates to it and subclassifies "Other" folders by filename. Oppose/Generate normalize their freeform/custom labels at match time. A one-off script re-tags existing `other` index rows in place. `config.BUILTIN_SEED` gains light generic-engine entries for the common types.

**Tech Stack:** Python, sqlite (in-place UPDATE), pytest. Tests: `C:\geminiterminal2\.venv\Scripts\python.exe`. **Work in worktree `C:\firm-briefs-p25` on branch `feat/firm-briefs-taxonomy`.**

**Environment for implementers:** shell cwd RESETS to `C:\geminiterminal2` between commands — begin every PowerShell command with `Set-Location 'C:\firm-briefs-p25';`. Tests: `Set-Location 'C:\firm-briefs-p25'; & 'C:\geminiterminal2\.venv\Scripts\python.exe' -m pytest <path> -v`. PowerShell only; no bash compound `cd && `. git via `git -C "C:/firm-briefs-p25" ...`.

**Canonical ids (must match existing index ids):** `msj, compel, demurrer, strike, in_limine, quash, sanctions, relieve_counsel, continue_trial, ex_parte, leave, dismiss, ime, gfs, consolidate, reconsider, protective_order, set_aside_default, other`. (Use `leave`/`dismiss`, NOT `leave_to_amend`/`motion_to_dismiss` — the index already uses the short ids.)

---

### Task 1: `motion_taxonomy.py` — canonical types + normalizer

**Files:**
- Create: `icharlotte_core/firm_briefs/motion_taxonomy.py`
- Test: `tests/test_firm_briefs/test_motion_taxonomy.py`

- [ ] **Step 1: Write the failing test**

```python
# tests/test_firm_briefs/test_motion_taxonomy.py
import pytest
from icharlotte_core.firm_briefs.motion_taxonomy import normalize_motion_type, display_name

@pytest.mark.parametrize("text,expected", [
    ("Defendant's Notice of Motion and Motion for Summary Judgment", "msj"),
    ("MSJ", "msj"),
    ("Motion for Summary Adjudication", "msj"),
    ("Motion to Compel Further Responses to RFP, Set One", "compel"),
    ("MTC", "compel"),
    ("Demurrer to First Amended Complaint", "demurrer"),
    ("Motion to Strike Punitive Damages", "strike"),
    ("Motion in Limine No. 3", "in_limine"),
    ("Motion to Quash Service of Summons", "quash"),
    ("Motion for Terminating Sanctions", "sanctions"),
    ("Motion to be Relieved as Counsel", "relieve_counsel"),
    ("Motion to Continue Trial and All Related Dates", "continue_trial"),
    ("Motion for Trial Preference", "continue_trial"),
    ("Ex Parte Application to Advance Hearing", "ex_parte"),
    ("Ex Parte Application to Continue Trial", "ex_parte"),   # ex parte wins over continue
    ("Motion for Leave to Amend Complaint", "leave"),
    ("Motion for Leave to File Cross-Complaint", "leave"),
    ("Motion for Leave to Conduct IME of Plaintiff", "ime"),  # IME wins over leave
    ("Notice of Motion for Independent Medical Examination", "ime"),
    ("Motion for Determination of Good Faith Settlement", "gfs"),
    ("Defendant's Motion GFS", "gfs"),
    ("Motion to Dismiss for Forum Non Conveniens", "dismiss"),
    ("Motion to Consolidate Related Cases", "consolidate"),
    ("Motion for Reconsideration", "reconsider"),
    ("Motion for Protective Order", "protective_order"),
    ("Motion to Set Aside Default", "set_aside_default"),
    ("Motion to Tax Costs", "other"),          # real but unmapped -> other
    ("", "other"),
    (None, "other"),
])
def test_normalize(text, expected):
    assert normalize_motion_type(text) == expected

def test_display_name():
    assert display_name("msj") == "Motion for Summary Judgment/Adjudication"
    assert display_name("ime") == "Motion for Leave to Conduct IME"
    assert display_name("nonexistent") == "nonexistent"
```

- [ ] **Step 2: Run test to verify it fails**

Run: `Set-Location 'C:\firm-briefs-p25'; & 'C:\geminiterminal2\.venv\Scripts\python.exe' -m pytest tests/test_firm_briefs/test_motion_taxonomy.py -v`
Expected: FAIL (`ModuleNotFoundError`)

- [ ] **Step 3: Implement**

```python
# icharlotte_core/firm_briefs/motion_taxonomy.py
"""Canonical motion-type vocabulary + a freeform->id normalizer.

Single source of truth shared by ingest (path_meta), Oppose-a-Motion (analyzer
output), and Generate-a-Motion (intake). Matching is keyed on these ids, so all
three must agree. Order is MOST-SPECIFIC FIRST: the first pattern that matches
the lowercased text wins. Unknown -> "other".
"""
from __future__ import annotations

import re
from typing import List, Tuple

# (id, display_name, [regex patterns]) — ORDER MATTERS (specific before generic).
CANONICAL_TYPES: List[Tuple[str, str, List[str]]] = [
    ("ime",              "Motion for Leave to Conduct IME",
        [r"\bime\b", r"independent medical exam", r"medical examination",
         r"physical examination", r"\bdme\b", r"conduct.{0,20}examination"]),
    ("gfs",              "Motion for Good Faith Settlement Determination",
        [r"good faith settlement", r"\bgfs\b"]),
    ("set_aside_default","Motion to Set Aside Default",
        [r"set[\s-]?aside", r"vacate.{0,15}default", r"relief from default"]),
    ("protective_order", "Motion for Protective Order",
        [r"protective order"]),
    ("in_limine",        "Motion in Limine",
        [r"in limine", r"\bmil\b"]),
    ("summary_judgment_alias", "", []),  # placeholder removed below (see note)
    ("msj",              "Motion for Summary Judgment/Adjudication",
        [r"summary judgment", r"summary adjudication", r"\bmsj\b", r"\bmsa\b"]),
    ("demurrer",         "Demurrer",
        [r"demurrer", r"\bdemur"]),
    ("strike",           "Motion to Strike",
        [r"motion to strike", r"\bmts\b", r"strike.{0,20}(punitive|portions|answer|complaint)"]),
    ("compel",           "Motion to Compel",
        [r"compel", r"\bmtc\b", r"\bmtca\b", r"\bmtcf\b"]),
    ("quash",            "Motion to Quash",
        [r"quash"]),
    ("sanctions",        "Motion for Sanctions",
        [r"sanction"]),
    ("relieve_counsel",  "Motion to be Relieved as Counsel",
        [r"relieved? as counsel", r"be relieved", r"withdraw as counsel", r"motion to withdraw"]),
    ("ex_parte",         "Ex Parte Application",
        [r"ex[\s-]?parte", r"\bepa\b"]),
    ("consolidate",      "Motion to Consolidate",
        [r"consolidat"]),
    ("reconsider",       "Motion for Reconsideration",
        [r"reconsider"]),
    ("dismiss",          "Motion to Dismiss",
        [r"dismiss"]),
    ("leave",            "Motion for Leave",
        [r"leave to amend", r"leave to file", r"motion for leave", r"\bleave\b"]),
    ("continue_trial",   "Motion to Continue Trial",
        [r"continue trial", r"cont(?:inuance)?.{0,12}trial", r"trial continuance",
         r"trial preference", r"preferential", r"\bpreference\b", r"specially set"]),
]

# Drop the placeholder row (kept above only to make the ordering intent explicit
# in review diffs); real lookups skip empty-pattern rows anyway.
CANONICAL_TYPES = [t for t in CANONICAL_TYPES if t[2]]

_DISPLAY = {tid: name for tid, name, _ in CANONICAL_TYPES}


def normalize_motion_type(text: str) -> str:
    """Return the canonical motion-type id for a freeform label, else 'other'."""
    low = (text or "").strip().lower()
    if not low:
        return "other"
    for tid, _name, patterns in CANONICAL_TYPES:
        for pat in patterns:
            if re.search(pat, low):
                return tid
    return "other"


def display_name(type_id: str) -> str:
    return _DISPLAY.get(type_id, type_id)
```

NOTE: the `summary_judgment_alias` placeholder row exists only so the
ordering reads clearly; the `CANONICAL_TYPES = [t for t in ... if t[2]]` line
removes empty-pattern rows. Keep both lines.

- [ ] **Step 4: Run test to verify it passes**

Run: same as Step 2. Expected: PASS (all parametrized cases + display). If any ordering case fails (e.g. "Ex Parte … Continue Trial" → continue_trial), the offending type must move earlier; do NOT weaken the assertions.

- [ ] **Step 5: Commit**

```bash
git -C "C:/firm-briefs-p25" add icharlotte_core/firm_briefs/motion_taxonomy.py tests/test_firm_briefs/test_motion_taxonomy.py
git -C "C:/firm-briefs-p25" commit -m "feat(firm_briefs): canonical motion-type taxonomy + normalize_motion_type"
```

---

### Task 2: Refactor `path_meta` to use the normalizer + subclassify "Other" by filename

**Files:**
- Modify: `icharlotte_core/firm_briefs/path_meta.py`
- Test: `tests/test_firm_briefs/test_path_meta.py` (extend existing)

- [ ] **Step 1: Add failing tests** (append to the existing `tests/test_firm_briefs/test_path_meta.py`)

```python
# --- Phase 2.5: filename subclassification of the "Other" buckets ---
def test_motions_other_folder_subclassifies_by_filename():
    p = ROOT + r"\Motions - Other\072 - X__Motion for Leave to Conduct IME of Plaintiff.pdf"
    assert meta_for_path(p, ROOT) == ("ime", "moving")

def test_motions_other_unmappable_stays_other():
    p = ROOT + r"\Motions - Other\072 - X__Motion to Tax Costs.pdf"
    assert meta_for_path(p, ROOT) == ("other", "moving")

def test_oppositions_other_subfolder_subclassifies():
    p = ROOT + r"\Oppositions\Other\011 - Y__Opposition to Motion to Consolidate.pdf"
    assert meta_for_path(p, ROOT) == ("consolidate", "opposition")

def test_existing_specific_folder_still_wins():
    p = ROOT + r"\Oppositions\Motion to Compel\008 - Z__opp.pdf"
    assert meta_for_path(p, ROOT) == ("compel", "opposition")
```

(`ROOT` is already defined at the top of the existing test file.)

- [ ] **Step 2: Run to verify the new tests fail**

Run: `Set-Location 'C:\firm-briefs-p25'; & 'C:\geminiterminal2\.venv\Scripts\python.exe' -m pytest tests/test_firm_briefs/test_path_meta.py -v`
Expected: the 3 subclassify tests FAIL (they currently return `other`); existing tests still pass.

- [ ] **Step 3: Implement** — replace `_TYPE_ALIASES` + `_canon_type` usage with the taxonomy normalizer, and subclassify the "Other" buckets by filename. Edit `path_meta.py`:

Replace the `_TYPE_ALIASES` dict and `_canon_type` function with:

```python
from .motion_taxonomy import normalize_motion_type


def _canon_type(label: str) -> str:
    return normalize_motion_type(label)


def _type_from_filename(path: str) -> str:
    """Subtype for the generic 'Other' buckets, read from the filename
    (strip the 'Matter__' prefix the library uses)."""
    name = os.path.basename(path)
    name = name.split("__", 1)[-1]
    return normalize_motion_type(os.path.splitext(name)[0])
```

Then update `meta_for_path` so the "Other" buckets subclassify by filename:

```python
    if low == "oppositions":
        t = _canon_type(sub)
        if t == "other":
            t = _type_from_filename(path)
        return (t, "opposition")
    if low == "replies":
        t = _canon_type(sub)
        if t == "other":
            t = _type_from_filename(path)
        return (t, "reply")
    if low == "ex parte applications":
        return ("ex_parte", "moving")
    if low.startswith("motion - "):
        return (_canon_type(top[len("motion - "):]), "moving")
    if low == "motions - other":
        return (_type_from_filename(path), "moving")
    if low.startswith("pleadings - "):
        return (_canon_type(top[len("pleadings - "):]), "pleading")
    return None
```

Leave the pleadings handling and the `_Support`/`_Other` exclusion unchanged. Note: pleadings folder labels ("answer", "complaint", …) now route through `normalize_motion_type`, which returns `other` for them — but pleadings are out of the motion-index scope (the build filters to motions/opp/replies), so this is harmless. If a `test_path_meta.py` pleadings test exists and now fails, update that test's expectation to the normalizer's output (do not special-case pleadings in code).

- [ ] **Step 4: Run tests to verify they pass**

Run: same as Step 2. Expected: all pass (new subclassify + existing). Fix any pleadings-expectation drift per the note above.

- [ ] **Step 5: Commit**

```bash
git -C "C:/firm-briefs-p25" add icharlotte_core/firm_briefs/path_meta.py tests/test_firm_briefs/test_path_meta.py
git -C "C:/firm-briefs-p25" commit -m "refactor(firm_briefs): path_meta uses taxonomy normalizer + subclassifies Other by filename"
```

---

### Task 3: `retag_firm_index.py` — re-tag existing 'other' rows in place

**Files:**
- Create: `retag_firm_index.py` (repo root, alongside the other build scripts)
- Test: `tests/test_firm_briefs/test_retag.py`

- [ ] **Step 1: Write the failing test**

```python
# tests/test_firm_briefs/test_retag.py
from icharlotte_core.firm_briefs.index import FirmBriefIndex
from icharlotte_core.firm_briefs.citation_harvest import HarvestedCite
from retag_firm_index import retag_other
import numpy as np


def _vec():
    v = np.ones(384, dtype=np.float32)
    return v / np.linalg.norm(v)


def _add(idx, path, mtype, side):
    idx.upsert_brief(path=path, content_hash="h", motion_type=mtype, side=side,
                     heading="", profile="p", profile_vec=_vec(), char_len=10,
                     ocr_ratio=0.0, cites=[HarvestedCite(reporter_citation="1 Cal.5th 1",
                                                         norm_cite="1cal.5th1", proposition="x")])


def test_retag_reclassifies_other(tmp_path):
    idx = FirmBriefIndex(db_path=str(tmp_path / "fb.db"), vectors_path=str(tmp_path / "v.f16"))
    idx.create_schema()
    _add(idx, r"C:\lib\Motions - Other\072 - X__Motion for Leave to Conduct IME.pdf", "other", "moving")
    _add(idx, r"C:\lib\Motions - Other\072 - X__Motion to Tax Costs.pdf", "other", "moving")
    _add(idx, r"C:\lib\Motion - Compel\008__opp.pdf", "compel", "moving")  # not 'other' -> untouched
    changed = retag_other(idx)
    assert changed == 1  # only the IME row reclassifies; Tax Costs stays other
    con = idx._conn()
    types = dict(con.execute("SELECT motion_type, COUNT(*) FROM briefs GROUP BY motion_type").fetchall())
    assert types.get("ime") == 1
    assert types.get("compel") == 1
    assert types.get("other") == 1  # Tax Costs remains


def test_retag_idempotent(tmp_path):
    idx = FirmBriefIndex(db_path=str(tmp_path / "fb.db"), vectors_path=str(tmp_path / "v.f16"))
    idx.create_schema()
    _add(idx, r"C:\lib\Motions - Other\x__Motion for Reconsideration.pdf", "other", "moving")
    assert retag_other(idx) == 1
    assert retag_other(idx) == 0  # nothing left to change
```

- [ ] **Step 2: Run test to verify it fails**

Run: `Set-Location 'C:\firm-briefs-p25'; & 'C:\geminiterminal2\.venv\Scripts\python.exe' -m pytest tests/test_firm_briefs/test_retag.py -v`
Expected: FAIL (`ModuleNotFoundError: retag_firm_index`)

- [ ] **Step 3: Implement**

```python
# retag_firm_index.py
"""Re-tag existing 'other' firm-index rows by normalizing their filename.

In-place UPDATE only: no re-extraction, no re-embedding (profile vectors and
citations are unchanged). Idempotent. Usage:
    python retag_firm_index.py
"""
import os
import sys

sys.path.insert(0, r"C:\geminiterminal2")

from icharlotte_core.firm_briefs.motion_taxonomy import normalize_motion_type


def _type_from_path(path: str) -> str:
    name = os.path.basename(path).split("__", 1)[-1]
    return normalize_motion_type(os.path.splitext(name)[0])


def retag_other(index) -> int:
    """Reclassify rows currently tagged 'other' using the filename. Returns count changed."""
    con = index._conn()
    rows = con.execute("SELECT id, path FROM briefs WHERE motion_type='other'").fetchall()
    changed = 0
    for r in rows:
        new_type = _type_from_path(r["path"])
        if new_type and new_type != "other":
            con.execute("UPDATE briefs SET motion_type=? WHERE id=?", (new_type, r["id"]))
            changed += 1
    con.commit()
    return changed


def main() -> int:
    from icharlotte_core.firm_briefs import factory
    from icharlotte_core.firm_briefs.index import FirmBriefIndex
    if not factory.index_available():
        print("No firm index built; nothing to re-tag.")
        return 1
    db, vec = factory.index_paths()
    idx = FirmBriefIndex(db_path=db, vectors_path=vec)
    idx.create_schema()
    before = dict(idx._conn().execute(
        "SELECT motion_type, COUNT(*) FROM briefs WHERE status='ok' GROUP BY motion_type").fetchall())
    n = retag_other(idx)
    after = dict(idx._conn().execute(
        "SELECT motion_type, COUNT(*) FROM briefs WHERE status='ok' GROUP BY motion_type").fetchall())
    print(f"Re-tagged {n} 'other' rows.")
    print("before:", before)
    print("after: ", after)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
```

- [ ] **Step 4: Run tests to verify they pass**

Run: same as Step 2. Expected: PASS (2 passed)

- [ ] **Step 5: Commit**

```bash
git -C "C:/firm-briefs-p25" add retag_firm_index.py tests/test_firm_briefs/test_retag.py
git -C "C:/firm-briefs-p25" commit -m "feat(firm_briefs): retag_firm_index in-place reclassifies 'other' by filename"
```

---

### Task 4: Register common types (light) in `config.BUILTIN_SEED`

**Files:**
- Modify: `icharlotte_core/motion_generation/config.py` (the `MOTION_TYPE_CONFIGS` dict)
- Test: `tests/test_firm_briefs/test_motion_types_registered.py`

- [ ] **Step 1: Write the failing test**

```python
# tests/test_firm_briefs/test_motion_types_registered.py
from icharlotte_core.motion_generation.config import list_motion_types, get_motion_config

EXPECTED = {"msj", "ex_parte", "ime", "gfs", "dismiss", "leave", "consolidate",
            "quash", "sanctions", "continue_trial", "protective_order",
            "compel", "demurrer", "strike"}


def test_common_types_registered():
    ids = {c.type_id for c in list_motion_types()}
    assert EXPECTED.issubset(ids)


def test_new_types_have_display_and_legal_standard():
    for tid in ["msj", "ex_parte", "ime", "gfs"]:
        cfg = get_motion_config(tid)
        assert cfg.type_id == tid
        assert cfg.display_name
        assert cfg.legal_standard_hint
        assert cfg.section_plan  # non-empty spine


def test_unknown_still_generic():
    assert get_motion_config("totally-unknown-xyz").type_id == "generic"
```

- [ ] **Step 2: Run test to verify it fails**

Run: `Set-Location 'C:\firm-briefs-p25'; & 'C:\geminiterminal2\.venv\Scripts\python.exe' -m pytest tests/test_firm_briefs/test_motion_types_registered.py -v`
Expected: FAIL (msj/ex_parte/etc. not in list_motion_types)

- [ ] **Step 3: Implement** — add these entries to the `MOTION_TYPE_CONFIGS` dict in `config.py` (after the existing `strike` entry, before `generic`). They reuse the generic engine (empty `analyzer_prompt`/`grounds_prompt`) with a per-type `legal_standard_hint` and the shared `_BASE_SECTIONS` spine:

```python
    "msj": MotionTypeConfig(
        type_id="msj", display_name="Motion for Summary Judgment/Adjudication",
        target_doc_guidance="Add the separate statement of undisputed material facts and the supporting evidence (declarations, discovery excerpts) the motion relies on.",
        legal_standard_hint="A motion for summary judgment/adjudication is governed by Code of Civil Procedure section 437c. Summary judgment is proper only where there is no triable issue of material fact and the moving party is entitled to judgment as a matter of law (CCP 437c(c)); a defendant meets its burden by showing an element cannot be established or a complete defense exists (CCP 437c(p)(2)).",
        section_plan=_BASE_SECTIONS),
    "ex_parte": MotionTypeConfig(
        type_id="ex_parte", display_name="Ex Parte Application",
        target_doc_guidance="Add the supporting declaration showing irreparable harm/urgency and the notice given to opposing counsel (Cal. Rules of Court, rule 3.1204).",
        legal_standard_hint="Ex parte relief is governed by California Rules of Court, rules 3.1200-3.1207. The applicant must make an affirmative factual showing of irreparable harm, immediate danger, or other statutory basis for ex parte relief, and must give notice by 10:00 a.m. the court day before.",
        section_plan=_BASE_SECTIONS),
    "ime": MotionTypeConfig(
        type_id="ime", display_name="Motion for Leave to Conduct IME",
        target_doc_guidance="Add the showing of good cause and the proposed examination's time, place, manner, conditions, scope, and examiner (CCP 2032.310, 2032.320).",
        legal_standard_hint="A motion for a physical or mental examination is governed by Code of Civil Procedure sections 2032.310 and 2032.320. The motion must specify the time, place, manner, conditions, scope, and nature of the examination and the examiner, and must be supported by a showing of good cause.",
        section_plan=_BASE_SECTIONS),
    "gfs": MotionTypeConfig(
        type_id="gfs", display_name="Motion for Determination of Good Faith Settlement",
        target_doc_guidance="Add the settlement terms and the facts bearing on the Tech-Bilt factors (settlor's proportionate liability, amount paid, allocation).",
        legal_standard_hint="A determination of good faith settlement is governed by Code of Civil Procedure section 877.6 and evaluated under the factors in Tech-Bilt, Inc. v. Woodward-Clyde & Associates (1985) 38 Cal.3d 488. A good faith determination bars other defendants' claims for equitable contribution or indemnity.",
        section_plan=_BASE_SECTIONS),
    "dismiss": MotionTypeConfig(
        type_id="dismiss", display_name="Motion to Dismiss",
        target_doc_guidance="Add the procedural basis for dismissal (e.g., failure to prosecute, forum non conveniens) and supporting facts.",
        legal_standard_hint="Grounds for dismissal include the discretionary and mandatory dismissal statutes (Code of Civil Procedure sections 583.130-583.430) and forum non conveniens (CCP 410.30).",
        section_plan=_BASE_SECTIONS),
    "leave": MotionTypeConfig(
        type_id="leave", display_name="Motion for Leave (Amend/File)",
        target_doc_guidance="Add the proposed amended pleading or cross-complaint and a declaration explaining the amendment and its timing (Cal. Rules of Court, rule 3.1324).",
        legal_standard_hint="Leave to amend is governed by Code of Civil Procedure sections 473(a) and 576 and is liberally granted; the motion must comply with California Rules of Court, rule 3.1324 (copy of the proposed amendment and a supporting declaration). Leave to file a cross-complaint is governed by CCP 426.50 / 428.50.",
        section_plan=_BASE_SECTIONS),
    "consolidate": MotionTypeConfig(
        type_id="consolidate", display_name="Motion to Consolidate",
        target_doc_guidance="Add the case captions/numbers to be consolidated and the common questions of law or fact.",
        legal_standard_hint="Consolidation is governed by Code of Civil Procedure section 1048(a): the court may consolidate actions involving a common question of law or fact to avoid unnecessary cost or delay.",
        section_plan=_BASE_SECTIONS),
    "quash": MotionTypeConfig(
        type_id="quash", display_name="Motion to Quash",
        target_doc_guidance="Add the summons/subpoena at issue and the facts showing lack of jurisdiction or an improper/overbroad subpoena.",
        legal_standard_hint="A motion to quash service of summons for lack of personal jurisdiction is governed by Code of Civil Procedure section 418.10. A motion to quash a subpoena is governed by CCP 1987.1.",
        section_plan=_BASE_SECTIONS),
    "sanctions": MotionTypeConfig(
        type_id="sanctions", display_name="Motion for Sanctions",
        target_doc_guidance="Add the conduct at issue and the prior orders/meet-and-confer establishing the basis for sanctions.",
        legal_standard_hint="Discovery sanctions are governed by Code of Civil Procedure sections 2023.010-2023.030; monetary, issue, evidence, and terminating sanctions escalate with the abuse. CCP 128.5/128.7 govern sanctions for bad-faith actions or tactics.",
        section_plan=_BASE_SECTIONS),
    "continue_trial": MotionTypeConfig(
        type_id="continue_trial", display_name="Motion to Continue Trial",
        target_doc_guidance="Add the declaration showing good cause for the continuance and the current trial/related dates (Cal. Rules of Court, rule 3.1332).",
        legal_standard_hint="Trial continuances are disfavored and granted only on an affirmative showing of good cause under California Rules of Court, rule 3.1332. Trial preference is governed by Code of Civil Procedure section 36.",
        section_plan=_BASE_SECTIONS),
    "protective_order": MotionTypeConfig(
        type_id="protective_order", display_name="Motion for Protective Order",
        target_doc_guidance="Add the discovery at issue and the facts showing annoyance, oppression, or undue burden justifying protection.",
        legal_standard_hint="Protective orders are governed by the method-specific statutes (e.g., Code of Civil Procedure sections 2025.420 for depositions, 2030.090 for interrogatories, 2031.060 for inspection demands); the court may make any order that justice requires to protect a party from unwarranted annoyance, oppression, or undue burden and expense.",
        section_plan=_BASE_SECTIONS),
```

- [ ] **Step 4: Run tests to verify they pass**

Run: same as Step 2. Expected: PASS (3 passed)

- [ ] **Step 5: Run the motion_generation regression**

Run: `Set-Location 'C:\firm-briefs-p25'; & 'C:\geminiterminal2\.venv\Scripts\python.exe' -m pytest tests -q -k "motion or generate or workbench"`
Expected: PASS (existing compel/demurrer/strike/generic behavior intact; `BUILTIN_SEED = dict(MOTION_TYPE_CONFIGS)` picks up the new entries automatically).

- [ ] **Step 6: Commit**

```bash
git -C "C:/firm-briefs-p25" add icharlotte_core/motion_generation/config.py tests/test_firm_briefs/test_motion_types_registered.py
git -C "C:/firm-briefs-p25" commit -m "feat(generate_motion): register common motion types (light, generic engine + legal-standard hints)"
```

---

### Task 5: Normalize the motion type at match time in Oppose-a-Motion

**Files:**
- Modify: `icharlotte_core/ui/wizard/pages/oppose_motion_page.py`
- Test: `tests/test_firm_briefs/test_oppose_normalize.py`

The analyzer's freeform `metadata.motion_type` stays shown in the editable field; only the value passed to the firm provider / style / research is normalized.

- [ ] **Step 1: Write the failing test**

```python
# tests/test_firm_briefs/test_oppose_normalize.py
from icharlotte_core.ui.wizard.pages import oppose_motion_page as omp


def test_module_exposes_normalizer():
    # The page must import the canonical normalizer for match-time use.
    assert omp.normalize_motion_type("Defendant's Motion for Summary Judgment") == "msj"
    assert omp.normalize_motion_type("Motion to Compel Further Responses") == "compel"
```

- [ ] **Step 2: Run test to verify it fails**

Run: `Set-Location 'C:\firm-briefs-p25'; & 'C:\geminiterminal2\.venv\Scripts\python.exe' -m pytest tests/test_firm_briefs/test_oppose_normalize.py -v`
Expected: FAIL (`AttributeError: module ... has no attribute 'normalize_motion_type'`)

- [ ] **Step 3: Implement** — in `oppose_motion_page.py`:

1. Add the import near the other firm_briefs imports:
```python
from icharlotte_core.firm_briefs.motion_taxonomy import normalize_motion_type
```
2. In the worker, where the firm provider / style / research currently receive `metadata.motion_type`, compute the canonical id ONCE and use it for all three. Find the block that builds `firm_provider`/`firm_style`/calls `research_arguments(... motion_type=metadata.motion_type ...)` and introduce:
```python
            firm_motion_type = normalize_motion_type(metadata.motion_type)
```
Then replace the three `metadata.motion_type` usages that feed matching:
- `_make_firm_provider(corpus)` is unaffected (it doesn't take type), but the `research_arguments(..., motion_type=metadata.motion_type, side="opposition")` call → `motion_type=firm_motion_type`.
- `_firm_style_exemplars(metadata.motion_type, "opposition", metadata)` → `_firm_style_exemplars(firm_motion_type, "opposition", metadata)`.
Leave the editable field (`self.motion_type_edit` / `metadata.motion_type`) and the manual StyleExampleRegistry lookup unchanged.

- [ ] **Step 4: Run test to verify it passes**

Run: same as Step 2. Expected: PASS

- [ ] **Step 5: Commit**

```bash
git -C "C:/firm-briefs-p25" add icharlotte_core/ui/wizard/pages/oppose_motion_page.py tests/test_firm_briefs/test_oppose_normalize.py
git -C "C:/firm-briefs-p25" commit -m "fix(oppose_motion): normalize analyzer motion type to canonical id for firm matching"
```

---

### Task 6: Normalize the "Other (specify…)" custom name in Generate-a-Motion

**Files:**
- Modify: `icharlotte_core/ui/wizard/pages/generate_motion_page.py`
- Test: `tests/test_firm_briefs/test_generate_normalize.py`

When the user picks a registered type, the id is already canonical. When they pick "Other (specify…)" and type a name (which today maps to `generic`), normalize that name so e.g. "MSJ" resolves to `msj` and matches the index. This only affects the id used for firm matching (provider/style/research), not the drafter engine selection (which still uses the chosen config).

- [ ] **Step 1: Write the failing test**

```python
# tests/test_firm_briefs/test_generate_normalize.py
from icharlotte_core.ui.wizard.pages import generate_motion_page as gmp


def test_module_exposes_normalizer():
    assert gmp.normalize_motion_type("MSJ") == "msj"
    assert gmp.normalize_motion_type("Motion for Good Faith Settlement") == "gfs"
```

- [ ] **Step 2: Run test to verify it fails**

Run: `Set-Location 'C:\firm-briefs-p25'; & 'C:\geminiterminal2\.venv\Scripts\python.exe' -m pytest tests/test_firm_briefs/test_generate_normalize.py -v`
Expected: FAIL (`AttributeError`)

- [ ] **Step 3: Implement** — in `generate_motion_page.py`:

1. Add the import:
```python
from icharlotte_core.firm_briefs.motion_taxonomy import normalize_motion_type
```
2. At the firm-matching site (where `_firm_style_exemplars(...)` is called and where `research_arguments(..., motion_type=...)` runs), compute a canonical id for matching. The chosen type id is `self.settings.get("motion_type_id")` (or the analyzer `metadata.motion_type`); when it is `generic` or empty, fall back to normalizing the custom name / metadata. Use:
```python
            raw_type = self.settings.get("motion_type_id") or getattr(metadata, "motion_type", "")
            firm_motion_type = raw_type if raw_type not in ("", "generic") else \
                normalize_motion_type(self.settings.get("custom_motion_name", "") or getattr(metadata, "motion_type", ""))
```
Then pass `firm_motion_type` to `_firm_style_exemplars(firm_motion_type, "moving", metadata)` and to `research_arguments(..., motion_type=firm_motion_type, ...)`. READ the surrounding code to find the exact setting key for the custom name (the intake stores it; e.g. `custom_motion_name` or similar — confirm and use the real key). If no separate custom-name key exists, normalize `getattr(metadata, "motion_type", "")`.

- [ ] **Step 4: Run test to verify it passes**

Run: same as Step 2. Expected: PASS

- [ ] **Step 5: Commit**

```bash
git -C "C:/firm-briefs-p25" add icharlotte_core/ui/wizard/pages/generate_motion_page.py tests/test_firm_briefs/test_generate_normalize.py
git -C "C:/firm-briefs-p25" commit -m "fix(generate_motion): normalize custom 'Other' motion name to canonical id for firm matching"
```

---

### Task 7: Full regression sweep + re-tag the real index

**Files:** none (verification + one-time data op)

- [ ] **Step 1: Run the suites**

Run:
```
Set-Location 'C:\firm-briefs-p25'; & 'C:\geminiterminal2\.venv\Scripts\python.exe' -m pytest tests/test_firm_briefs tests/test_opposition tests/test_wizard -q
```
Expected: all `tests/test_firm_briefs` pass; no NEW opposition/wizard failures (pre-existing Qt-fixture collection errors in unrelated wizard files are not regressions — classify mine vs pre-existing).

- [ ] **Step 2: Re-tag the real index** (the index lives at `C:\geminiterminal2\.gemini\firm_briefs`; point the worktree code at it via the env override):

Run:
```
Set-Location 'C:\firm-briefs-p25'; $env:FIRM_BRIEFS_DATA_DIR='C:\geminiterminal2\.gemini\firm_briefs'; & 'C:\geminiterminal2\.venv\Scripts\python.exe' 'C:\firm-briefs-p25\retag_firm_index.py'
```
Expected: prints "Re-tagged N 'other' rows" with a before/after distribution showing `other` shrinking and `ime`/`leave`/`dismiss`/`gfs`/`consolidate`/etc. appearing. Report the output.

- [ ] **Step 3: Commit** (only if any test-only fixups were needed; otherwise nothing to commit here)

---

## Self-review notes (for the implementer)
- **Canonical ids must match the existing index ids** (`leave`, `dismiss`, etc.) so re-tag + matching + config all align.
- **Ordering in `CANONICAL_TYPES` is the crux** — specific multiword patterns (IME, GFS, set-aside, in-limine, MSJ, ex-parte) before generic ones (leave, continue_trial, dismiss). The parametrized normalizer test locks the tricky cases ("Ex Parte … Continue Trial" → ex_parte; "Leave to Conduct IME" → ime).
- **Additive/guarded:** the editable oppose field still shows the freeform label; generic drafter behavior for unknown types is unchanged; absent index → still no-op.
- **Deferred:** Phase 3 panel UI; bespoke per-type prompts; content-based classification; 3800 build.
