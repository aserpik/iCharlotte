# Med-Cron Multi-Analysis Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Extend Med-Cron from a single rewrite task into a multi-analysis task: after the user selects a chronology, the wizard offers a curated catalog of analyses plus user-typed custom analyses; selected analyses run in parallel and each writes its own docx.

**Architecture:** Two-phase script following the Summarize-Depositions pattern. Phase 1 extracts text (narrative-only + full-with-tables), writes a session JSON listing the catalog, prints `AWAITING_INPUT:`. The wizard settings page shows the picker; on Proceed it writes `user_config` back. Phase 2 reads the session, fans out the selected analyses inside one `ThreadPoolExecutor`, writes one docx per analysis.

**Tech Stack:** Python 3.x, PySide6, pypdf, python-docx, `icharlotte_core.llm_config.LLMCaller`, `icharlotte_core.document_processor.extract_docx_text`, `icharlotte_core.agent_logger.AgentLogger`.

**Spec:** `docs/superpowers/specs/2026-05-17-medcron-analyses-design.md`

**Branch:** continue on `ui-redesign` (current branch).

---

## File Map

### New files

- `Scripts/MED_CHRON_ANALYSES/__init__.py` — empty package marker
- `Scripts/MED_CHRON_ANALYSES/catalog.py` — `AnalysisDef` dataclass + `CATALOG` list + `CATALOG_BY_ID` lookup + `load_prompt(name)` helper
- `Scripts/MED_CHRON_ANALYSES/prompts/rewrite_chronology.txt` — existing `MED_CHRON_PROMPT.txt` content
- `Scripts/MED_CHRON_ANALYSES/prompts/inconsistencies.txt` — placeholder (one-liner the user will fill in later)
- `Scripts/MED_CHRON_ANALYSES/prompts/treatment_gaps.txt` — placeholder
- `Scripts/MED_CHRON_ANALYSES/prompts/_custom_wrapper.txt` — wrapper template with `{user_instruction}` placeholder
- `icharlotte_core/med_chron/__init__.py` — empty package marker
- `icharlotte_core/med_chron/session_manager.py` — session JSON read/write, cache path computation
- `icharlotte_core/ui/wizard/pages/med_chron_settings_page.py` — `MedChronSettingsPage(SettingsPage)`, mirrors `DepositionSettingsPage`
- `icharlotte_core/ui/med_chron_config_form.py` — `MedChronConfigForm` + `CustomAnalysisRow` widget
- `tests/test_med_chron/__init__.py`
- `tests/test_med_chron/test_catalog.py`
- `tests/test_med_chron/test_session_manager.py`
- `tests/test_med_chron/test_phase1_prep.py`
- `tests/test_med_chron/test_phase2_runner.py`
- `tests/test_med_chron/test_legacy_cli_compatibility.py`
- `tests/test_wizard/test_med_chron_settings_page.py`
- `tests/test_wizard/test_med_chron_registry.py`

### Modified files

- `Scripts/med_chron.py` — rewrite into a CLI with three modes (legacy, prep, run). Reuse existing `extract_text`, `extract_provider_from_filename`, `sanitize_filename`, `filter_content`, `add_markdown_to_doc`. Replace `call_gemini` with `LLMCaller`. Refactor `save_to_docx` to accept an output path.
- `icharlotte_core/ui/wizard/registry.py` — add `med_chron_analysis` `TaskSpec` and `_med_chron_settings_page_cls()` factory; add `phase1_args` and `phase2_flag` fields to `TaskSpec` dataclass.
- `icharlotte_core/ui/wizard/runners/subprocess_worker.py` — accept `phase1_args` and `phase2_flag` constructor args; use them in `_start_file` and `resume_with_config`.
- `icharlotte_core/ui/wizard/runners/parallel_subprocess_worker.py` — add `"med_chron.py"` to `_TWO_PHASE_SCRIPTS`.
- `icharlotte_core/ui/wizard/task_tab.py` — pass `spec.phase1_args` and `spec.phase2_flag` to `SubprocessWorker` (both creation sites).

### Files intentionally left alone

- `Scripts/MED_CHRON_PROMPT.txt` — keep on disk as a backstop; the new file at `Scripts/MED_CHRON_ANALYSES/prompts/rewrite_chronology.txt` will contain the same content. Legacy CLI mode reads the new path.

---

## Task 1: Catalog module + prompt files

**Files:**
- Create: `Scripts/MED_CHRON_ANALYSES/__init__.py`
- Create: `Scripts/MED_CHRON_ANALYSES/catalog.py`
- Create: `Scripts/MED_CHRON_ANALYSES/prompts/rewrite_chronology.txt`
- Create: `Scripts/MED_CHRON_ANALYSES/prompts/inconsistencies.txt`
- Create: `Scripts/MED_CHRON_ANALYSES/prompts/treatment_gaps.txt`
- Create: `Scripts/MED_CHRON_ANALYSES/prompts/_custom_wrapper.txt`
- Test: `tests/test_med_chron/__init__.py` (empty)
- Test: `tests/test_med_chron/test_catalog.py`

- [ ] **Step 1: Write the failing tests**

Create `tests/test_med_chron/__init__.py` (empty file).

Create `tests/test_med_chron/test_catalog.py`:

```python
"""Tests for the Med-Cron analysis catalog."""

import sys
from pathlib import Path

import pytest

PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(PROJECT_ROOT / "Scripts"))

from MED_CHRON_ANALYSES import catalog  # noqa: E402


def test_catalog_is_non_empty():
    assert len(catalog.CATALOG) > 0


def test_catalog_ids_are_unique():
    ids = [d.id for d in catalog.CATALOG]
    assert len(ids) == len(set(ids)), f"duplicate catalog ids: {ids}"


def test_only_rewrite_is_default_selected():
    defaults = [d for d in catalog.CATALOG if d.default_selected]
    assert len(defaults) == 1
    assert defaults[0].id == "rewrite_chronology"


def test_rewrite_does_not_use_tables():
    rewrite = catalog.CATALOG_BY_ID["rewrite_chronology"]
    assert rewrite.uses_tables is False


def test_every_catalog_entry_has_prompt_file_on_disk():
    prompts_dir = PROJECT_ROOT / "Scripts" / "MED_CHRON_ANALYSES" / "prompts"
    for d in catalog.CATALOG:
        path = prompts_dir / d.prompt_file
        assert path.exists(), f"missing prompt file: {path}"


def test_load_prompt_returns_file_contents():
    text = catalog.load_prompt("rewrite_chronology.txt")
    assert text
    assert isinstance(text, str)


def test_load_prompt_rejects_path_traversal():
    with pytest.raises(ValueError):
        catalog.load_prompt("../../../etc/passwd")


def test_custom_wrapper_contains_placeholder():
    wrapper = catalog.load_prompt("_custom_wrapper.txt")
    assert "{user_instruction}" in wrapper
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `python -m pytest tests/test_med_chron/test_catalog.py -v`
Expected: ImportError / module not found for `MED_CHRON_ANALYSES`.

- [ ] **Step 3: Create the package marker**

Create `Scripts/MED_CHRON_ANALYSES/__init__.py` as an empty file.

- [ ] **Step 4: Create the catalog module**

Create `Scripts/MED_CHRON_ANALYSES/catalog.py`:

```python
"""Curated catalog of Med-Cron analyses.

Each entry is a curated analysis that can be selected by the user in the
wizard. The catalog also exposes a helper to load prompt files from the
``prompts/`` directory next to this module.

Custom user-typed analyses are NOT stored here — they live per-session in
the session JSON and use the ``_custom_wrapper.txt`` prompt.
"""

from dataclasses import dataclass
from pathlib import Path

_PROMPTS_DIR = Path(__file__).parent / "prompts"


@dataclass(frozen=True)
class AnalysisDef:
    """One curated analysis. Source of truth for prompt_file + uses_tables."""
    id: str            # stable slug for filenames + session JSON
    title: str         # shown in checkbox UI
    description: str   # short tooltip under the checkbox
    uses_tables: bool  # True -> full text; False -> narrative only
    prompt_file: str   # filename in prompts/ (no directory components)
    default_selected: bool = False


CATALOG: list[AnalysisDef] = [
    AnalysisDef(
        id="rewrite_chronology",
        title="Rewrite Chronology (readable narrative)",
        description="Reformats the pre/post-injury synopsis into a clean narrative.",
        uses_tables=False,
        prompt_file="rewrite_chronology.txt",
        default_selected=True,
    ),
    AnalysisDef(
        id="inconsistencies",
        title="Inconsistency Check",
        description="Flags contradictions between narrative and table entries.",
        uses_tables=True,
        prompt_file="inconsistencies.txt",
    ),
    AnalysisDef(
        id="treatment_gaps",
        title="Treatment Gap Detector",
        description="Identifies unexplained gaps in treatment dates.",
        uses_tables=True,
        prompt_file="treatment_gaps.txt",
    ),
]

CATALOG_BY_ID: dict[str, AnalysisDef] = {d.id: d for d in CATALOG}


def load_prompt(name: str) -> str:
    """Read a prompt file from the prompts/ directory.

    Rejects any name containing path separators or '..' to prevent
    traversal outside the prompts directory.
    """
    if "/" in name or "\\" in name or ".." in name:
        raise ValueError(f"invalid prompt name: {name!r}")
    return (_PROMPTS_DIR / name).read_text(encoding="utf-8")
```

- [ ] **Step 5: Create the prompt files**

Create `Scripts/MED_CHRON_ANALYSES/prompts/rewrite_chronology.txt` with the contents of the existing `Scripts/MED_CHRON_PROMPT.txt` (copy verbatim — read the existing file first, then write to the new path).

Create `Scripts/MED_CHRON_ANALYSES/prompts/inconsistencies.txt`:

```
You will be given the BRIEF SYNOPSIS sections of a medical chronology PLUS
the underlying tables of medical entries (visits, providers, dates,
findings, treatments).

Identify any contradictions or inconsistencies between what the narrative
synopsis claims and what the tables actually show. Examples:

- Narrative says "no prior back complaints" but a table row predates the
  alleged injury and references back pain.
- Narrative cites a specific date of injury but the earliest treatment
  entry post-dates that by months without explanation.
- Narrative attributes a finding to one provider but the tables show
  another provider made it.

For each inconsistency, cite the narrative claim, cite the contradicting
table entry (date + provider + entry), and briefly explain the conflict.

Use Markdown. One heading per inconsistency. Be specific and grounded —
do NOT speculate beyond what the documents support.
```

Create `Scripts/MED_CHRON_ANALYSES/prompts/treatment_gaps.txt`:

```
You will be given the BRIEF SYNOPSIS sections of a medical chronology PLUS
the underlying tables of medical entries.

Identify unexplained gaps in treatment. A "gap" is a period (typically
30+ days, but flag shorter ones if clinically odd) between entries where
ongoing treatment was expected but did not occur.

For each gap, report:

- Date range of the gap (from last entry → next entry).
- Provider type involved (PT, ortho, pain management, etc.).
- The clinical context around the gap (what was the last documented plan
  before the gap? what happened after?).
- A brief note on whether the gap appears explained by the documents
  (insurance lapse, treater discharge, patient nonresponse) or
  unexplained.

Use Markdown. One heading per gap.
```

Create `Scripts/MED_CHRON_ANALYSES/prompts/_custom_wrapper.txt`:

```
You will be given the BRIEF SYNOPSIS sections of a medical chronology PLUS
the underlying tables of medical entries.

The user has asked you to perform the following analysis on this
chronology:

{user_instruction}

Ground every finding in the document. Cite specific dates, providers, and
entries where applicable. Use Markdown. Be specific.
```

- [ ] **Step 6: Run tests to verify they pass**

Run: `python -m pytest tests/test_med_chron/test_catalog.py -v`
Expected: 8 passed.

- [ ] **Step 7: Commit**

```bash
git add Scripts/MED_CHRON_ANALYSES tests/test_med_chron/__init__.py tests/test_med_chron/test_catalog.py
git commit -m "feat(med-chron): add analysis catalog + prompt files"
```

---

## Task 2: Session manager

**Files:**
- Create: `icharlotte_core/med_chron/__init__.py`
- Create: `icharlotte_core/med_chron/session_manager.py`
- Test: `tests/test_med_chron/test_session_manager.py`

- [ ] **Step 1: Write the failing tests**

Create `tests/test_med_chron/test_session_manager.py`:

```python
"""Tests for the Med-Cron session manager."""

import json
from pathlib import Path

import pytest

from icharlotte_core.med_chron import session_manager


def _make_input_file(tmp_path: Path, name: str = "Acme PT Records.docx") -> Path:
    p = tmp_path / name
    p.write_bytes(b"\x00" * 1024)
    return p


def test_compute_session_paths_layout(tmp_path):
    inp = _make_input_file(tmp_path)
    out_dir = tmp_path / "NOTES" / "AI OUTPUT"
    out_dir.mkdir(parents=True)
    paths = session_manager.compute_session_paths(str(inp), str(out_dir))

    cache_root = out_dir / ".med_chron"
    assert paths.cache_dir.parent == cache_root
    assert paths.cache_dir.name and len(paths.cache_dir.name) == 12  # hash
    assert paths.session_path.name == "session.json"
    assert paths.narrative_text_path.name == "narrative.txt"
    assert paths.full_text_path.name == "full.txt"


def test_hash_changes_when_mtime_changes(tmp_path):
    inp = _make_input_file(tmp_path)
    out_dir = tmp_path / "out"
    out_dir.mkdir()

    p1 = session_manager.compute_session_paths(str(inp), str(out_dir))
    # Force a different mtime
    import os, time
    new_time = inp.stat().st_mtime_ns + 1_000_000_000
    os.utime(inp, ns=(new_time, new_time))
    p2 = session_manager.compute_session_paths(str(inp), str(out_dir))

    assert p1.cache_dir != p2.cache_dir


def test_write_then_read_session_round_trip(tmp_path):
    inp = _make_input_file(tmp_path)
    out_dir = tmp_path / "out"
    out_dir.mkdir()
    paths = session_manager.compute_session_paths(str(inp), str(out_dir))

    data = {"phase": "awaiting_input", "user_config": None, "k": 1}
    session_manager.write_session(paths.session_path, data)

    loaded = session_manager.read_session(paths.session_path)
    assert loaded == data


def test_update_user_config_flips_phase(tmp_path):
    inp = _make_input_file(tmp_path)
    out_dir = tmp_path / "out"
    out_dir.mkdir()
    paths = session_manager.compute_session_paths(str(inp), str(out_dir))

    session_manager.write_session(
        paths.session_path,
        {"phase": "awaiting_input", "user_config": None},
    )
    session_manager.update_user_config(
        paths.session_path,
        {"selected_catalog_ids": ["rewrite_chronology"], "custom_analyses": []},
    )

    loaded = session_manager.read_session(paths.session_path)
    assert loaded["phase"] == "ready_to_run"
    assert loaded["user_config"]["selected_catalog_ids"] == ["rewrite_chronology"]


def test_write_session_creates_parent_dirs(tmp_path):
    deeply_nested = tmp_path / "a" / "b" / "c" / "session.json"
    session_manager.write_session(deeply_nested, {"phase": "x"})
    assert deeply_nested.exists()
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `python -m pytest tests/test_med_chron/test_session_manager.py -v`
Expected: ImportError for `icharlotte_core.med_chron`.

- [ ] **Step 3: Create the package + module**

Create `icharlotte_core/med_chron/__init__.py` as an empty file.

Create `icharlotte_core/med_chron/session_manager.py`:

```python
"""Session JSON management for the two-phase Med-Cron agent.

Phase 1 (prep) writes the session and pauses. The wizard UI fills in
``user_config`` and flips ``phase`` to ``"ready_to_run"``. Phase 2 (run)
reads the session and produces output.

Cache layout, scoped to the case's output directory:

    <output_dir>/.med_chron/<file_hash>/
        narrative.txt
        full.txt
        session.json

``<file_hash>`` is sha1(abspath + mtime_ns), truncated to 12 hex chars,
so touching the source file invalidates the cache.
"""

import hashlib
import json
import os
from dataclasses import dataclass
from pathlib import Path


@dataclass(frozen=True)
class SessionPaths:
    cache_dir: Path
    session_path: Path
    narrative_text_path: Path
    full_text_path: Path


def _file_hash(input_path: str) -> str:
    abspath = os.path.abspath(input_path)
    try:
        mtime_ns = os.stat(input_path).st_mtime_ns
    except FileNotFoundError:
        mtime_ns = 0
    raw = f"{abspath}|{mtime_ns}".encode("utf-8")
    return hashlib.sha1(raw).hexdigest()[:12]


def compute_session_paths(input_path: str, output_dir: str) -> SessionPaths:
    """Return the cache + session paths for an input file under output_dir."""
    cache_dir = Path(output_dir) / ".med_chron" / _file_hash(input_path)
    return SessionPaths(
        cache_dir=cache_dir,
        session_path=cache_dir / "session.json",
        narrative_text_path=cache_dir / "narrative.txt",
        full_text_path=cache_dir / "full.txt",
    )


def write_session(session_path, data: dict) -> None:
    """Atomically write the session JSON via tmp file + os.replace."""
    session_path = Path(session_path)
    session_path.parent.mkdir(parents=True, exist_ok=True)
    tmp = session_path.with_suffix(session_path.suffix + ".tmp")
    tmp.write_text(json.dumps(data, indent=2), encoding="utf-8")
    os.replace(tmp, session_path)


def read_session(session_path) -> dict:
    return json.loads(Path(session_path).read_text(encoding="utf-8"))


def update_user_config(session_path, user_config: dict) -> None:
    """Load, set user_config, flip phase to 'ready_to_run', write atomically."""
    data = read_session(session_path)
    data["user_config"] = user_config
    data["phase"] = "ready_to_run"
    write_session(session_path, data)
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `python -m pytest tests/test_med_chron/test_session_manager.py -v`
Expected: 5 passed.

- [ ] **Step 5: Commit**

```bash
git add icharlotte_core/med_chron tests/test_med_chron/test_session_manager.py
git commit -m "feat(med-chron): add session manager for two-phase flow"
```

---

## Task 3: Phase 1 (prep) — text extraction + session JSON

**Files:**
- Modify: `Scripts/med_chron.py` — add `process_prep()` function and `--phase=prep` dispatch in `main()`. Do NOT touch the existing default-mode behavior yet (legacy mode preserved in Task 5).
- Test: `tests/test_med_chron/test_phase1_prep.py`

- [ ] **Step 1: Write the failing tests**

Create `tests/test_med_chron/test_phase1_prep.py`:

```python
"""Tests for Phase 1 (prep) of the Med-Cron agent."""

import json
import os
import sys
from pathlib import Path
from unittest.mock import patch

import pytest
from docx import Document

PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(PROJECT_ROOT / "Scripts"))

import med_chron  # noqa: E402


def _make_chrono_docx(path: Path, narrative_pre: str, narrative_post: str,
                     table_rows: list[list[str]]) -> Path:
    doc = Document()
    doc.add_paragraph(f"BRIEF SYNOPSIS OF PRE-INJURY MEDICAL RECORD: {narrative_pre}")
    doc.add_paragraph(f"BRIEF SYNOPSIS OF POST-INJURY MEDICAL RECORD: {narrative_post}")
    table = doc.add_table(rows=len(table_rows), cols=len(table_rows[0]))
    for r, row in enumerate(table_rows):
        for c, cell in enumerate(row):
            table.cell(r, c).text = cell
    doc.save(str(path))
    return path


def test_phase1_writes_session_with_both_text_caches(tmp_path):
    src = tmp_path / "1234-001_ ACME PT.docx"
    _make_chrono_docx(
        src,
        narrative_pre="No prior back issues.",
        narrative_post="Treated for lumbar strain following the accident.",
        table_rows=[["Date", "Provider"], ["2024-01-15", "Acme PT"]],
    )
    out_dir = tmp_path / "NOTES" / "AI OUTPUT"

    rc = med_chron.process_prep(str(src), str(out_dir))
    assert rc == 0

    from icharlotte_core.med_chron import session_manager
    paths = session_manager.compute_session_paths(str(src), str(out_dir))

    data = session_manager.read_session(paths.session_path)
    assert data["phase"] == "awaiting_input"
    assert data["user_config"] is None
    assert data["narrative_missing"] is False
    assert paths.narrative_text_path.exists()
    assert paths.full_text_path.exists()


def test_phase1_narrative_only_excludes_table_content(tmp_path):
    src = tmp_path / "rec.docx"
    _make_chrono_docx(
        src,
        narrative_pre="No prior back issues.",
        narrative_post="Treated for lumbar strain.",
        table_rows=[["Date", "Provider"], ["2024-01-15", "UNIQUE_TABLE_TOKEN"]],
    )
    out_dir = tmp_path / "out"

    rc = med_chron.process_prep(str(src), str(out_dir))
    assert rc == 0

    from icharlotte_core.med_chron import session_manager
    paths = session_manager.compute_session_paths(str(src), str(out_dir))
    narrative = paths.narrative_text_path.read_text(encoding="utf-8")
    assert "UNIQUE_TABLE_TOKEN" not in narrative
    assert "lumbar strain" in narrative


def test_phase1_full_text_includes_table_rows(tmp_path):
    src = tmp_path / "rec.docx"
    _make_chrono_docx(
        src,
        narrative_pre="No prior issues.",
        narrative_post="Treated.",
        table_rows=[["Date", "Provider"], ["2024-01-15", "UNIQUE_TABLE_TOKEN"]],
    )
    out_dir = tmp_path / "out"

    rc = med_chron.process_prep(str(src), str(out_dir))
    assert rc == 0

    from icharlotte_core.med_chron import session_manager
    paths = session_manager.compute_session_paths(str(src), str(out_dir))
    full = paths.full_text_path.read_text(encoding="utf-8")
    assert "UNIQUE_TABLE_TOKEN" in full


def test_phase1_missing_synopsis_marks_narrative_missing(tmp_path):
    src = tmp_path / "no_synopsis.docx"
    doc = Document()
    doc.add_paragraph("This document has no BRIEF SYNOPSIS sections.")
    doc.add_paragraph("Just regular medical content.")
    doc.save(str(src))
    out_dir = tmp_path / "out"

    rc = med_chron.process_prep(str(src), str(out_dir))
    assert rc == 0  # still succeeds

    from icharlotte_core.med_chron import session_manager
    paths = session_manager.compute_session_paths(str(src), str(out_dir))
    data = session_manager.read_session(paths.session_path)
    assert data["narrative_missing"] is True
    # narrative.txt is written empty
    assert paths.narrative_text_path.exists()
    assert paths.narrative_text_path.read_text(encoding="utf-8").strip() == ""


def test_phase1_prints_awaiting_input_token_on_success(tmp_path, capsys):
    src = tmp_path / "rec.docx"
    _make_chrono_docx(
        src,
        narrative_pre="Pre.",
        narrative_post="Post.",
        table_rows=[["A", "B"], ["1", "2"]],
    )
    out_dir = tmp_path / "out"

    med_chron.process_prep(str(src), str(out_dir))
    out = capsys.readouterr().out
    awaiting = [ln for ln in out.splitlines() if ln.startswith("AWAITING_INPUT:")]
    assert awaiting, f"AWAITING_INPUT token not printed; stdout: {out!r}"


def test_phase1_session_includes_catalog_snapshot(tmp_path):
    src = tmp_path / "rec.docx"
    _make_chrono_docx(
        src,
        narrative_pre="P.",
        narrative_post="Q.",
        table_rows=[["A", "B"], ["1", "2"]],
    )
    out_dir = tmp_path / "out"

    med_chron.process_prep(str(src), str(out_dir))
    from icharlotte_core.med_chron import session_manager
    paths = session_manager.compute_session_paths(str(src), str(out_dir))
    data = session_manager.read_session(paths.session_path)

    assert isinstance(data["catalog"], list)
    rewrite = next(e for e in data["catalog"] if e["id"] == "rewrite_chronology")
    assert rewrite["default_selected"] is True
    assert rewrite["uses_tables"] is False


def test_phase1_reuses_cache_on_unchanged_input(tmp_path):
    src = tmp_path / "rec.docx"
    _make_chrono_docx(
        src,
        narrative_pre="P.",
        narrative_post="Q.",
        table_rows=[["A", "B"], ["1", "2"]],
    )
    out_dir = tmp_path / "out"

    med_chron.process_prep(str(src), str(out_dir))
    from icharlotte_core.med_chron import session_manager
    paths = session_manager.compute_session_paths(str(src), str(out_dir))

    # Mutate the cached file to detect re-extraction.
    sentinel = "SENTINEL_TEXT_DO_NOT_OVERWRITE"
    paths.full_text_path.write_text(sentinel, encoding="utf-8")

    # Re-run with the source unchanged.
    rc = med_chron.process_prep(str(src), str(out_dir))
    assert rc == 0
    # Cache reuse means our sentinel survived.
    assert paths.full_text_path.read_text(encoding="utf-8") == sentinel
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `python -m pytest tests/test_med_chron/test_phase1_prep.py -v`
Expected: AttributeError — `med_chron.process_prep` does not exist.

- [ ] **Step 3: Add Phase 1 to med_chron.py**

Open `Scripts/med_chron.py`. At the top, after the existing imports, add:

```python
from concurrent.futures import ThreadPoolExecutor, as_completed
```

After the existing helper functions (`filter_content`, `save_to_docx`, etc.) and BEFORE the existing `main()`, add:

```python
def _extract_full_text(file_path: str) -> str:
    """Extract narrative + table text from a chronology file.

    .docx → ``icharlotte_core.document_processor.extract_docx_text``
            (canonical extractor that includes tables as pipe-separated rows).
    .pdf  → ``extract_text`` (same as narrative path; PDFs don't have a
            paragraphs-vs-tables split in extraction).
    .doc  → Word COM read-only, never calls word.Quit().
    """
    ext = os.path.splitext(file_path)[1].lower()
    if ext == ".docx":
        from icharlotte_core.document_processor import extract_docx_text
        return extract_docx_text(file_path)
    if ext == ".pdf":
        # PDFs flatten tables into narrative text already.
        return extract_text(file_path) or ""
    if ext == ".doc":
        return _extract_doc_via_word_com(file_path)
    # Fallback to existing extractor (txt etc.).
    return extract_text(file_path) or ""


def _extract_doc_via_word_com(file_path: str) -> str:
    """Read a legacy .doc by attaching to the user's running Word.

    Mirrors ChatTab._extract_doc_text: never set word.Visible and never
    call word.Quit() — only close the Document we opened. Open ReadOnly
    so the user's session is untouched.
    """
    try:
        import win32com.client  # type: ignore
    except ImportError:
        log_event("win32com not available; cannot extract .doc files", level="warning")
        return ""
    word = None
    doc = None
    try:
        word = win32com.client.Dispatch("Word.Application")
        doc = word.Documents.Open(
            FileName=os.path.abspath(file_path),
            ReadOnly=True,
            AddToRecentFiles=False,
            ConfirmConversions=False,
        )
        return doc.Content.Text or ""
    except Exception as e:
        log_event(f".doc extraction failed for {file_path}: {e}", level="warning")
        return ""
    finally:
        if doc is not None:
            try:
                doc.Close(SaveChanges=False)
            except Exception:
                pass


def _build_catalog_snapshot() -> list:
    """Serialise the curated catalog into a JSON-friendly list."""
    # Ensure Scripts/ is importable so MED_CHRON_ANALYSES is found.
    scripts_dir = os.path.dirname(os.path.abspath(__file__))
    if scripts_dir not in sys.path:
        sys.path.insert(0, scripts_dir)
    from MED_CHRON_ANALYSES.catalog import CATALOG
    return [
        {
            "id": d.id,
            "title": d.title,
            "description": d.description,
            "uses_tables": d.uses_tables,
            "default_selected": d.default_selected,
        }
        for d in CATALOG
    ]


def process_prep(input_path: str, output_dir: str) -> int:
    """Phase 1: extract text twice, write session JSON, print AWAITING_INPUT.

    Returns process-style exit code (0 success, non-zero failure). Does
    NOT call sys.exit — leaves that to main().
    """
    from icharlotte_core.med_chron import session_manager

    paths = session_manager.compute_session_paths(input_path, output_dir)

    # Cache reuse: if session.json already exists and both text files
    # are present, skip extraction. The hash incorporates mtime, so a
    # changed file routes to a different cache dir automatically.
    if (paths.session_path.exists()
            and paths.narrative_text_path.exists()
            and paths.full_text_path.exists()):
        log_event(f"Reusing cached prep at {paths.cache_dir}")
        print(f"AWAITING_INPUT:{paths.session_path}", flush=True)
        return 0

    paths.cache_dir.mkdir(parents=True, exist_ok=True)

    # --- Narrative-only text ---
    raw_text = extract_text(input_path)
    if not raw_text:
        log_event(f"Could not extract text from {input_path}", level="error")
        return 1
    narrative = filter_content(raw_text)
    narrative_missing = narrative is None
    paths.narrative_text_path.write_text(narrative or "", encoding="utf-8")

    # --- Full text (narrative + tables) ---
    full_text = _extract_full_text(input_path)
    if not full_text:
        log_event(f"Could not extract full text from {input_path}", level="error")
        return 1
    paths.full_text_path.write_text(full_text, encoding="utf-8")

    # --- Session JSON ---
    filename = os.path.basename(input_path)
    provider_name = extract_provider_from_filename(filename)
    file_num_match = re.search(r"(\d{4}\.\d{3})", input_path)
    file_number = file_num_match.group(1) if file_num_match else None

    session_data = {
        "version": 1,
        "phase": "awaiting_input",
        "input_path": str(input_path),
        "narrative_text_path": str(paths.narrative_text_path),
        "full_text_path": str(paths.full_text_path),
        "narrative_missing": narrative_missing,
        "provider_name": provider_name,
        "file_number": file_number,
        "catalog": _build_catalog_snapshot(),
        "user_config": None,
    }
    session_manager.write_session(paths.session_path, session_data)

    log_event(
        f"Phase 1 complete: cached {len(narrative or '')} narrative chars "
        f"+ {len(full_text)} full chars; session at {paths.session_path}"
    )
    print(f"AWAITING_INPUT:{paths.session_path}", flush=True)
    return 0
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `python -m pytest tests/test_med_chron/test_phase1_prep.py -v`
Expected: 7 passed.

- [ ] **Step 5: Commit**

```bash
git add Scripts/med_chron.py tests/test_med_chron/test_phase1_prep.py
git commit -m "feat(med-chron): add Phase 1 (prep) text extraction + session"
```

---

## Task 4: Phase 2 (run) — parallel analysis runner

**Files:**
- Modify: `Scripts/med_chron.py` — add `process_run()`, `_run_one_analysis()`, helper `_slug()`, refactor `save_to_docx` to accept output path.
- Test: `tests/test_med_chron/test_phase2_runner.py`

- [ ] **Step 1: Write the failing tests**

Create `tests/test_med_chron/test_phase2_runner.py`:

```python
"""Tests for Phase 2 (run) of the Med-Cron agent."""

import json
import os
import sys
from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest

PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(PROJECT_ROOT / "Scripts"))

import med_chron  # noqa: E402
from icharlotte_core.med_chron import session_manager  # noqa: E402


def _prep_session(tmp_path: Path, *, narrative: str = "narr",
                  full: str = "full text", selected: list[str] = None,
                  custom: list[dict] = None) -> Path:
    """Hand-build a ready_to_run session for direct Phase 2 tests."""
    cache = tmp_path / ".med_chron" / "abc123"
    cache.mkdir(parents=True)
    narrative_path = cache / "narrative.txt"
    narrative_path.write_text(narrative, encoding="utf-8")
    full_path = cache / "full.txt"
    full_path.write_text(full, encoding="utf-8")
    session_path = cache / "session.json"

    data = {
        "version": 1,
        "phase": "ready_to_run",
        "input_path": str(tmp_path / "rec.docx"),
        "narrative_text_path": str(narrative_path),
        "full_text_path": str(full_path),
        "narrative_missing": narrative == "",
        "provider_name": "Acme PT",
        "file_number": "1234.567",
        "catalog": [],
        "user_config": {
            "selected_catalog_ids": selected if selected is not None else ["rewrite_chronology"],
            "custom_analyses": custom or [],
        },
    }
    session_manager.write_session(session_path, data)
    return session_path


def test_run_one_uses_narrative_when_uses_tables_false(tmp_path):
    session_path = _prep_session(
        tmp_path,
        narrative="NARR_ONLY",
        full="NARR_ONLY plus TABLE_TOKEN",
        selected=["rewrite_chronology"],
    )

    captured = {}

    def fake_call(prompt, text, **kw):
        captured["text"] = text
        return "# Result\nbody"

    with patch.object(med_chron.LLMCaller, "call", side_effect=fake_call):
        rc = med_chron.process_run(str(session_path), str(tmp_path / "out"))

    assert rc == 0
    assert captured["text"] == "NARR_ONLY"
    assert "TABLE_TOKEN" not in captured["text"]


def test_run_one_uses_full_text_when_uses_tables_true(tmp_path):
    session_path = _prep_session(
        tmp_path,
        narrative="NARR_ONLY",
        full="NARR_ONLY plus TABLE_TOKEN",
        selected=["inconsistencies"],
    )

    captured = {}
    with patch.object(med_chron.LLMCaller, "call",
                      side_effect=lambda p, t, **kw: captured.setdefault("text", t) or "# X"):
        med_chron.process_run(str(session_path), str(tmp_path / "out"))

    assert "TABLE_TOKEN" in captured["text"]


def test_per_run_failure_does_not_abort_siblings(tmp_path):
    session_path = _prep_session(
        tmp_path,
        narrative="narr",
        full="full",
        selected=["rewrite_chronology", "inconsistencies"],
    )

    call_count = {"n": 0}

    def flaky(prompt, text, **kw):
        call_count["n"] += 1
        if call_count["n"] == 1:
            raise RuntimeError("simulated LLM error")
        return "# Survivor"

    with patch.object(med_chron.LLMCaller, "call", side_effect=flaky):
        rc = med_chron.process_run(str(session_path), str(tmp_path / "out"))

    # At least one succeeded → exit code 0.
    assert rc == 0
    out_files = list((tmp_path / "out").rglob("*.docx"))
    assert len(out_files) == 1  # only the survivor wrote a docx


def test_all_runs_failed_exits_nonzero(tmp_path):
    session_path = _prep_session(
        tmp_path,
        selected=["rewrite_chronology", "inconsistencies"],
    )
    with patch.object(med_chron.LLMCaller, "call",
                      side_effect=RuntimeError("boom")):
        rc = med_chron.process_run(str(session_path), str(tmp_path / "out"))
    assert rc == 1


def test_skips_rewrite_when_narrative_missing(tmp_path):
    session_path = _prep_session(
        tmp_path,
        narrative="",  # empty narrative
        full="full body",
        selected=["rewrite_chronology", "inconsistencies"],
    )

    calls = []
    def stub_call(prompt, text, **kw):
        calls.append((prompt[:30], text[:30]))
        return "# OK"

    with patch.object(med_chron.LLMCaller, "call", side_effect=stub_call):
        rc = med_chron.process_run(str(session_path), str(tmp_path / "out"))

    assert rc == 0
    # Only one analysis (inconsistencies) ran — rewrite was skipped.
    assert len(calls) == 1


def test_custom_analysis_wraps_user_instruction_in_template(tmp_path):
    session_path = _prep_session(
        tmp_path,
        selected=[],  # only custom
        custom=[{"label": "Left-knee mentions",
                  "instruction": "Find every entry mentioning the left knee."}],
    )

    captured_prompts = []
    with patch.object(med_chron.LLMCaller, "call",
                      side_effect=lambda p, t, **kw: captured_prompts.append(p) or "# X"):
        med_chron.process_run(str(session_path), str(tmp_path / "out"))

    assert len(captured_prompts) == 1
    assert "Find every entry mentioning the left knee." in captured_prompts[0]
    # The wrapper template's placeholder must have been substituted.
    assert "{user_instruction}" not in captured_prompts[0]


def test_output_filenames_use_analysis_id(tmp_path):
    session_path = _prep_session(
        tmp_path,
        selected=["rewrite_chronology", "inconsistencies"],
    )
    with patch.object(med_chron.LLMCaller, "call", return_value="# Result"):
        med_chron.process_run(str(session_path), str(tmp_path / "out"))

    out_files = sorted(p.name for p in (tmp_path / "out").rglob("*.docx"))
    # input basename is "rec" (from rec.docx)
    assert any("med_chron_rewrite_chronology_rec.docx" == n for n in out_files)
    assert any("med_chron_inconsistencies_rec.docx" == n for n in out_files)


def test_custom_output_includes_index_and_slug(tmp_path):
    session_path = _prep_session(
        tmp_path,
        selected=[],
        custom=[
            {"label": "Left knee", "instruction": "..."},
            {"label": "Left knee", "instruction": "..."},  # same label
        ],
    )
    with patch.object(med_chron.LLMCaller, "call", return_value="# Result"):
        med_chron.process_run(str(session_path), str(tmp_path / "out"))

    out_files = sorted(p.name for p in (tmp_path / "out").rglob("*.docx"))
    # Both must exist (no collision) and include index 1/2.
    assert any("custom_1_left_knee_rec.docx" in n for n in out_files)
    assert any("custom_2_left_knee_rec.docx" in n for n in out_files)


def test_bails_if_phase_not_ready_to_run(tmp_path):
    session_path = _prep_session(tmp_path)
    data = session_manager.read_session(session_path)
    data["phase"] = "awaiting_input"
    session_manager.write_session(session_path, data)

    rc = med_chron.process_run(str(session_path), str(tmp_path / "out"))
    assert rc == 1


def test_slug_helper_lowercases_and_replaces_special_chars():
    assert med_chron._slug("Left-knee mentions") == "left-knee_mentions"
    assert med_chron._slug("  whitespace  ") == "whitespace"
    assert med_chron._slug("Has !! punct?") == "has_punct"


def test_max_workers_capped_at_4(tmp_path):
    """Even with many analyses queued, ThreadPoolExecutor uses at most 4 workers."""
    custom = [{"label": f"c{i}", "instruction": "do x"} for i in range(10)]
    session_path = _prep_session(tmp_path, selected=[], custom=custom)

    captured = {}
    real_tpe = med_chron.ThreadPoolExecutor
    class SpyTPE(real_tpe):
        def __init__(self, max_workers=None, *a, **kw):
            captured["max_workers"] = max_workers
            super().__init__(max_workers=max_workers, *a, **kw)

    with patch.object(med_chron, "ThreadPoolExecutor", SpyTPE):
        with patch.object(med_chron.LLMCaller, "call", return_value="# X"):
            med_chron.process_run(str(session_path), str(tmp_path / "out"))

    assert captured["max_workers"] == 4  # min(10, 4) → 4
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `python -m pytest tests/test_med_chron/test_phase2_runner.py -v`
Expected: AttributeError — `med_chron.process_run` does not exist.

- [ ] **Step 3: Add the new helpers + Phase 2 to med_chron.py**

Append to `Scripts/med_chron.py` (after `process_prep`):

```python
# =============================================================================
# Phase 2: Run selected analyses
# =============================================================================

from dataclasses import dataclass
from icharlotte_core.llm_config import LLMCaller


def _slug(value: str) -> str:
    """Lowercase + sanitize for use in filenames/run ids."""
    cleaned = re.sub(r"[^a-zA-Z0-9_\-]+", "_", value or "")
    return cleaned.strip("_").lower()


def _safe_basename(input_path: str) -> str:
    base = os.path.splitext(os.path.basename(input_path))[0]
    return _slug(base)


@dataclass
class RunSpec:
    id: str
    title: str
    prompt_text: str
    input_text: str
    output_path: str


@dataclass
class RunResult:
    spec: RunSpec
    success: bool
    error: str = ""
    output_chars: int = 0


def _build_run_list(session: dict, narrative: str, full: str,
                    safe_basename: str, output_dir: str) -> list[RunSpec]:
    """Translate user_config + catalog into concrete RunSpec instances."""
    # Ensure Scripts/ is importable so MED_CHRON_ANALYSES is found.
    scripts_dir = os.path.dirname(os.path.abspath(__file__))
    if scripts_dir not in sys.path:
        sys.path.insert(0, scripts_dir)
    from MED_CHRON_ANALYSES.catalog import CATALOG_BY_ID, load_prompt

    cfg = session["user_config"]
    runs: list[RunSpec] = []

    for cat_id in cfg.get("selected_catalog_ids", []):
        if cat_id not in CATALOG_BY_ID:
            log_event(f"Skipping unknown catalog id: {cat_id}", level="warning")
            continue
        d = CATALOG_BY_ID[cat_id]
        runs.append(RunSpec(
            id=cat_id,
            title=d.title,
            prompt_text=load_prompt(d.prompt_file),
            input_text=narrative if not d.uses_tables else full,
            output_path=os.path.join(
                output_dir, f"med_chron_{cat_id}_{safe_basename}.docx"
            ),
        ))

    wrapper = None
    for i, c in enumerate(cfg.get("custom_analyses", []), 1):
        if wrapper is None:
            wrapper = load_prompt("_custom_wrapper.txt")
        label_slug = _slug(c["label"])
        runs.append(RunSpec(
            id=f"custom_{i}_{label_slug}",
            title=c["label"],
            prompt_text=wrapper.replace("{user_instruction}", c["instruction"]),
            input_text=full,
            output_path=os.path.join(
                output_dir, f"med_chron_custom_{i}_{label_slug}_{safe_basename}.docx"
            ),
        ))

    return runs


def _drop_rewrite_if_narrative_missing(runs: list[RunSpec], narrative: str) -> list[RunSpec]:
    if narrative.strip():
        return runs
    kept = []
    for r in runs:
        if r.id == "rewrite_chronology":
            log_event(
                "Skipping Rewrite Chronology — no pre/post-injury synopsis "
                "headings found in this document.",
                level="warning",
            )
            continue
        kept.append(r)
    return kept


def _run_one_analysis(spec: RunSpec, llm_caller: LLMCaller,
                       provider_name: str) -> RunResult:
    """Execute a single analysis. Caller MUST NOT let exceptions escape."""
    try:
        log_event(f"[{spec.id}] starting LLM call ({len(spec.input_text)} chars)")
        result = llm_caller.call(
            prompt=spec.prompt_text,
            text=spec.input_text,
            task_type="summary",
        )
        if not result:
            return RunResult(spec=spec, success=False, error="LLM returned empty result")

        os.makedirs(os.path.dirname(spec.output_path), exist_ok=True)
        save_to_docx_at_path(result, spec.output_path, provider_name, spec.title)
        log_event(f"[{spec.id}] done: {len(result)} chars → {spec.output_path}")
        return RunResult(spec=spec, success=True, output_chars=len(result))
    except Exception as e:
        log_event(f"[{spec.id}] failed: {e}", level="error")
        return RunResult(spec=spec, success=False, error=str(e))


def save_to_docx_at_path(content: str, output_path: str,
                         provider_name: str, analysis_title: str) -> None:
    """Write content to output_path with the existing Med-Cron styling.

    If the destination is locked (e.g., open in Word), auto-version up to
    10 attempts: ``out.docx`` -> ``out v.2.docx`` -> ``out v.3.docx``.
    """
    from docx import Document
    from docx.shared import Pt

    base, ext = os.path.splitext(output_path)
    attempt = 1
    last_err = None
    while attempt <= 10:
        candidate = output_path if attempt == 1 else f"{base} v.{attempt}{ext}"
        try:
            doc = Document()
            style = doc.styles['Normal']
            style.font.name = 'Times New Roman'
            style.font.size = Pt(12)
            style.paragraph_format.line_spacing = 1.0

            p = doc.add_paragraph()
            run = p.add_run(f"{analysis_title} — {provider_name}")
            run.bold = True
            run.underline = True
            run.font.name = 'Times New Roman'
            run.font.size = Pt(12)

            doc.add_paragraph(f"Generated on: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
            add_markdown_to_doc(doc, content)
            doc.save(candidate)
            return
        except (PermissionError, IOError) as e:
            last_err = e
            attempt += 1
    raise RuntimeError(f"Could not save after {attempt - 1} attempts: {last_err}")


def process_run(session_path: str, output_dir: str) -> int:
    """Phase 2: load session, fan out analyses in parallel, write docx each.

    Returns 0 if at least one analysis succeeded, 1 if all failed or the
    session is malformed.
    """
    from icharlotte_core.med_chron import session_manager

    try:
        session = session_manager.read_session(session_path)
    except Exception as e:
        log_event(f"Could not load session at {session_path}: {e}", level="error")
        return 1

    if session.get("phase") != "ready_to_run":
        log_event(
            f"Session phase is {session.get('phase')!r}; expected ready_to_run",
            level="error",
        )
        return 1

    narrative = Path(session["narrative_text_path"]).read_text(encoding="utf-8")
    full = Path(session["full_text_path"]).read_text(encoding="utf-8")
    safe_basename = _safe_basename(session["input_path"])

    runs = _build_run_list(session, narrative, full, safe_basename, output_dir)
    runs = _drop_rewrite_if_narrative_missing(runs, narrative)

    if not runs:
        log_event("No runs scheduled (after skip rules). Nothing to do.", level="warning")
        return 1

    llm_caller = LLMCaller()
    provider_name = session.get("provider_name") or "Unknown Provider"

    successes = 0
    failures = 0
    total = len(runs)

    log_event(f"Starting {total} analyses (max 4 concurrent)")
    with ThreadPoolExecutor(max_workers=min(total, 4)) as ex:
        futures = {
            ex.submit(_run_one_analysis, r, llm_caller, provider_name): r
            for r in runs
        }
        done = 0
        for f in as_completed(futures):
            result = f.result()
            done += 1
            if result.success:
                successes += 1
            else:
                failures += 1
            pct = int(20 + (done * 70 / total))  # leave 20% for prep, 10% for final
            print(f"PROGRESS:{pct}:{done}/{total} done ({failures} failed)", flush=True)

    log_event(f"Phase 2 complete: {successes}/{total} succeeded, {failures} failed")
    print(f"PROGRESS:100:{successes}/{total} analyses complete ({failures} failed)", flush=True)
    return 0 if successes > 0 else 1
```

Note: `Path` is already imported (from `pathlib`? — check the existing imports at the top of `med_chron.py`). If not, add `from pathlib import Path` near the other imports.

- [ ] **Step 4: Run tests to verify they pass**

Run: `python -m pytest tests/test_med_chron/test_phase2_runner.py -v`
Expected: 10 passed.

- [ ] **Step 5: Commit**

```bash
git add Scripts/med_chron.py tests/test_med_chron/test_phase2_runner.py
git commit -m "feat(med-chron): add Phase 2 parallel analysis runner"
```

---

## Task 5: CLI dispatcher + legacy compatibility

**Files:**
- Modify: `Scripts/med_chron.py` — rewrite `main()` to dispatch on `--phase` flag; legacy mode (no flag) runs only the rewrite analysis.
- Test: `tests/test_med_chron/test_legacy_cli_compatibility.py`

- [ ] **Step 1: Write the failing tests**

Create `tests/test_med_chron/test_legacy_cli_compatibility.py`:

```python
"""Tests that ``python med_chron.py <file>`` (no --phase) still works.

This is the IndexTab agent-runner path. It must produce
``med_chron_<filename>.docx`` exactly like the pre-refactor agent.
"""

import os
import sys
from pathlib import Path
from unittest.mock import patch

import pytest
from docx import Document

PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(PROJECT_ROOT / "Scripts"))

import med_chron  # noqa: E402


def _make_chrono_docx(path: Path, *, pre="Pre.", post="Post.") -> Path:
    doc = Document()
    doc.add_paragraph(f"BRIEF SYNOPSIS OF PRE-INJURY MEDICAL RECORD: {pre}")
    doc.add_paragraph(f"BRIEF SYNOPSIS OF POST-INJURY MEDICAL RECORD: {post}")
    doc.save(str(path))
    return path


def test_legacy_dispatcher_dispatches_to_legacy_when_no_phase(tmp_path, monkeypatch):
    src = _make_chrono_docx(tmp_path / "1234-001_ ACME.docx")

    called = {"legacy": False, "prep": False, "run": False}

    monkeypatch.setattr(med_chron, "process_legacy",
                        lambda p, **kw: called.__setitem__("legacy", True) or 0)
    monkeypatch.setattr(med_chron, "process_prep",
                        lambda p, o, **kw: called.__setitem__("prep", True) or 0)
    monkeypatch.setattr(med_chron, "process_run",
                        lambda p, o, **kw: called.__setitem__("run", True) or 0)

    monkeypatch.setattr(sys, "argv", ["med_chron.py", str(src)])
    with pytest.raises(SystemExit) as exc:
        med_chron.main()
    assert exc.value.code == 0
    assert called == {"legacy": True, "prep": False, "run": False}


def test_dispatcher_routes_prep_phase(tmp_path, monkeypatch):
    src = _make_chrono_docx(tmp_path / "rec.docx")
    called = []
    monkeypatch.setattr(med_chron, "process_prep",
                        lambda p, o, **kw: called.append(("prep", p, o)) or 0)
    monkeypatch.setattr(sys, "argv",
                        ["med_chron.py", "--phase=prep", str(src)])
    with pytest.raises(SystemExit) as exc:
        med_chron.main()
    assert exc.value.code == 0
    assert called and called[0][0] == "prep"


def test_dispatcher_routes_run_phase(tmp_path, monkeypatch):
    fake_session = tmp_path / "s.json"
    fake_session.write_text("{}", encoding="utf-8")
    called = []
    monkeypatch.setattr(med_chron, "process_run",
                        lambda p, o, **kw: called.append(("run", p, o)) or 0)
    monkeypatch.setattr(sys, "argv",
                        ["med_chron.py", "--phase=run", str(fake_session)])
    with pytest.raises(SystemExit) as exc:
        med_chron.main()
    assert exc.value.code == 0
    assert called and called[0][0] == "run"


def test_legacy_mode_writes_existing_filename_pattern(tmp_path):
    src = _make_chrono_docx(tmp_path / "1234-001_ ACME PT.docx")
    out_dir = tmp_path / "NOTES" / "AI OUTPUT"
    out_dir.mkdir(parents=True)

    with patch.object(med_chron.LLMCaller, "call", return_value="# Rewrite\nbody"):
        rc = med_chron.process_legacy(str(src), output_dir_override=str(out_dir))

    assert rc == 0
    # The existing sanitizer keeps dashes (regex is [^a-zA-Z0-9_\-]).
    # Input "1234-001_ ACME PT.docx" → safe_name "1234-001__ACME_PT".
    expected = out_dir / "med_chron_1234-001__ACME_PT.docx"
    assert expected.exists(), f"missing expected output: {expected}; have: {list(out_dir.iterdir())}"
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `python -m pytest tests/test_med_chron/test_legacy_cli_compatibility.py -v`
Expected: AttributeError — `med_chron.process_legacy` or new main() not yet wired.

- [ ] **Step 3: Refactor `main()` and add `process_legacy()`**

In `Scripts/med_chron.py`, REPLACE the entire existing `def main():` function with the dispatcher below. Also add the new `process_legacy` function above it.

```python
def _resolve_output_dir(input_path: str) -> str:
    """Compute the case AI-OUTPUT directory using the existing rules.

    Lifted from the old main() so all three phases share one implementation.
    """
    parts = input_path.split(os.sep)
    output_dir = None
    case_root_parts = None

    # Priority 1: folder starting with exactly 3 digits.
    for i in range(len(parts) - 1, -1, -1):
        if re.match(r'^\d{3}(\D|$)', parts[i]):
            case_root_parts = parts[:i + 1]
            break

    # Priority 2: "Current Clients" / Client / Matter pattern.
    if not case_root_parts:
        for i, part in enumerate(parts):
            if part.lower() == "current clients":
                if i + 2 < len(parts):
                    case_root_parts = parts[:i + 3]
                break

    if case_root_parts:
        output_dir = os.sep.join(case_root_parts + ["NOTES", "AI OUTPUT"])

    if not output_dir:
        for i in range(len(parts) - 1, -1, -1):
            if parts[i].upper() == "NOTES":
                output_dir = os.path.join(os.sep.join(parts[:i + 1]), "AI OUTPUT")
                break

    if not output_dir:
        input_dir = os.path.dirname(input_path)
        parent_dir = os.path.dirname(input_dir)
        output_dir = os.path.join(parent_dir, "NOTES", "AI OUTPUT")

    os.makedirs(output_dir, exist_ok=True)
    return output_dir


def process_legacy(input_path: str, *, output_dir_override: str | None = None) -> int:
    """Legacy single-rewrite mode: ``python med_chron.py <file>`` with no --phase.

    Used by the older IndexTab agent runner. Runs only the Rewrite analysis
    on the narrative-only text, writing to the existing filename pattern
    ``med_chron_<safe_filename>.docx`` so external callers keep working.
    """
    if os.path.isdir(input_path):
        # Existing behavior: recurse on directory inputs.
        for root, _, files in os.walk(input_path):
            for file in files:
                if file.lower().endswith(('.pdf', '.docx')) and "med_chron" not in file.lower():
                    try:
                        subprocess.run([sys.executable, sys.argv[0], os.path.join(root, file)],
                                       check=True)
                    except subprocess.CalledProcessError as e:
                        log_event(f"Subprocess failed for {file}: {e}", level="error")
        return 0

    raw_text = extract_text(input_path)
    if not raw_text:
        log_event(f"Could not extract text from {input_path}", level="error")
        return 1
    narrative = filter_content(raw_text)
    if not narrative:
        log_event("No valid content under PRE/POST-INJURY headings.", level="warning")
        return 0

    scripts_dir = os.path.dirname(os.path.abspath(__file__))
    if scripts_dir not in sys.path:
        sys.path.insert(0, scripts_dir)
    from MED_CHRON_ANALYSES.catalog import load_prompt

    prompt = load_prompt("rewrite_chronology.txt")
    llm = LLMCaller()
    content = llm.call(prompt=prompt, text=narrative, task_type="summary")
    if not content:
        return 1

    filename = os.path.basename(input_path)
    provider_name = extract_provider_from_filename(filename)
    output_dir = output_dir_override or _resolve_output_dir(input_path)
    safe_name = re.sub(r"[^a-zA-Z0-9_\-]", "_", os.path.splitext(filename)[0])
    output_path = os.path.join(output_dir, f"med_chron_{safe_name}.docx")
    save_to_docx_at_path(content, output_path, provider_name,
                          "Medical Record Chronology")

    # Existing CaseDataManager wiring (best-effort).
    try:
        data_manager = CaseDataManager()
        file_num_match = re.search(r"(\d{4}\.\d{3})", input_path)
        if file_num_match:
            safe_provider = re.sub(r"[^a-zA-Z0-9_]", "_", provider_name.lower())
            data_manager.save_variable(
                file_num_match.group(1),
                f"med_chron_{safe_provider}",
                content,
                source="med_chron_agent",
                extra_tags=["Evidence", "Medical Records", "Chronology"],
            )
    except Exception as e:
        log_event(f"Could not save to case data: {e}", level="warning")

    log_event(f"Legacy rewrite done → {output_path}")
    return 0


def main():
    """CLI dispatcher.

    Modes:
      med_chron.py <file>                       → legacy single-rewrite
      med_chron.py --phase=prep <file>          → Phase 1 (prep)
      med_chron.py --phase=run  <session.json>  → Phase 2 (run)
    """
    args = sys.argv[1:]
    if not args:
        log_event("Error: No file path provided.", level="error")
        sys.exit(1)

    phase = None
    positional = []
    output_dir_override = None
    i = 0
    while i < len(args):
        a = args[i]
        if a.startswith("--phase="):
            phase = a.split("=", 1)[1].strip().lower()
            i += 1
        elif a == "--output_path" and i + 1 < len(args):
            output_dir_override = args[i + 1]
            i += 2
        else:
            positional.append(a)
            i += 1

    combined = " ".join(positional).strip().strip('"').strip("'")
    if positional and os.path.exists(combined):
        target = combined
    elif positional and os.path.exists(positional[0]):
        target = positional[0]
    else:
        log_event(f"Error: path not found: {combined or '(empty)'}", level="error")
        sys.exit(1)
    target = os.path.abspath(target)

    if phase == "prep":
        out_dir = output_dir_override or _resolve_output_dir(target)
        rc = process_prep(target, out_dir)
        sys.exit(rc)

    if phase == "run":
        # target is a session.json path. Output dir is the cache dir's
        # great-grandparent — i.e., the original NOTES/AI OUTPUT folder.
        out_dir = output_dir_override or str(Path(target).parent.parent.parent)
        rc = process_run(target, out_dir)
        sys.exit(rc)

    # No --phase flag: legacy single-rewrite mode.
    rc = process_legacy(target, output_dir_override=output_dir_override)
    sys.exit(rc)
```

DELETE the existing `call_gemini` function (no longer used — `LLMCaller` replaces it) and the old `save_to_docx` (replaced by `save_to_docx_at_path`).

If `Path` is not yet imported at the top of the file, add `from pathlib import Path`.

If `from icharlotte_core.llm_config import LLMCaller` isn't already added (Task 4), it should be present now.

- [ ] **Step 4: Run all med_chron tests to verify everything passes**

Run: `python -m pytest tests/test_med_chron/ -v`
Expected: All tests pass (catalog + session_manager + phase1 + phase2 + legacy CLI).

- [ ] **Step 5: Commit**

```bash
git add Scripts/med_chron.py tests/test_med_chron/test_legacy_cli_compatibility.py
git commit -m "feat(med-chron): CLI dispatcher with legacy + prep + run modes"
```

---

## Task 6: Parameterize SubprocessWorker with phase flags

**Files:**
- Modify: `icharlotte_core/ui/wizard/registry.py` — add `phase1_args: list[str] = []` and `phase2_flag: str = "--phase=summary"` to `TaskSpec`.
- Modify: `icharlotte_core/ui/wizard/runners/subprocess_worker.py` — accept and use the new flags.
- Modify: `icharlotte_core/ui/wizard/task_tab.py` — pass them through.
- Modify: `icharlotte_core/ui/wizard/runners/parallel_subprocess_worker.py` — add `"med_chron.py"` to `_TWO_PHASE_SCRIPTS`.
- Test: `tests/test_wizard/test_subprocess_worker.py` (extend existing file).

- [ ] **Step 1: Write the failing tests**

Open `tests/test_wizard/test_subprocess_worker.py` (existing file — already there per repo listing). Append these tests:

```python
def test_phase1_args_are_prepended_to_invocation(qtbot, tmp_path):
    """SubprocessWorker with phase1_args=['--phase=prep'] invokes the script
    with that flag before the file path."""
    from icharlotte_core.ui.wizard.runners.subprocess_worker import SubprocessWorker

    file_path = str(tmp_path / "in.docx")
    open(file_path, "w").close()

    captured = {}
    def fake_launch(self, extra_argv):
        captured["argv"] = list(extra_argv)
    SubprocessWorker._launch_process = fake_launch  # type: ignore

    w = SubprocessWorker(
        script_name="med_chron.py",
        case_path=str(tmp_path),
        file_number="",
        files=[file_path],
        settings={},
        phase1_args=["--phase=prep"],
        phase2_flag="--phase=run",
    )
    w.start()

    # Expect: [script_path, "--phase=prep", file_path]
    assert captured["argv"][1] == "--phase=prep"
    assert captured["argv"][2] == file_path


def test_phase2_flag_is_used_in_resume_with_config(qtbot, tmp_path):
    from icharlotte_core.ui.wizard.runners.subprocess_worker import SubprocessWorker

    captured = {}
    def fake_launch(self, extra_argv):
        captured["argv"] = list(extra_argv)
    SubprocessWorker._launch_process = fake_launch  # type: ignore

    w = SubprocessWorker(
        script_name="med_chron.py",
        case_path=str(tmp_path),
        file_number="",
        files=[str(tmp_path / "in.docx")],
        settings={},
        phase1_args=["--phase=prep"],
        phase2_flag="--phase=run",
    )
    w.resume_with_config(str(tmp_path / "session.json"))

    assert "--phase=run" in captured["argv"]
    assert str(tmp_path / "session.json") in captured["argv"]


def test_default_phase_flags_preserve_deposition_behavior(qtbot, tmp_path):
    """Default constructor args must still produce --phase=summary so the
    existing deposition flow keeps working."""
    from icharlotte_core.ui.wizard.runners.subprocess_worker import SubprocessWorker

    captured = {}
    def fake_launch(self, extra_argv):
        captured["argv"] = list(extra_argv)
    SubprocessWorker._launch_process = fake_launch  # type: ignore

    w = SubprocessWorker(
        script_name="summarize_deposition.py",
        case_path=str(tmp_path),
        file_number="",
        files=[str(tmp_path / "depo.pdf")],
        settings={},
    )
    w.resume_with_config(str(tmp_path / "session.json"))
    assert "--phase=summary" in captured["argv"]
```

At the top of the same file, ensure `import pytest` and `pytest.importorskip("pytestqt")` (no underscore) are present.

- [ ] **Step 2: Run tests to verify they fail**

Run: `python -m pytest tests/test_wizard/test_subprocess_worker.py -v`
Expected: TypeError on `SubprocessWorker(... phase1_args=..., phase2_flag=...)`.

- [ ] **Step 3: Update SubprocessWorker**

In `icharlotte_core/ui/wizard/runners/subprocess_worker.py`:

In `__init__`, after the existing parameters, add `phase1_args: List[str] | None = None` and `phase2_flag: str = "--phase=summary"`. Store them as `self._phase1_args = list(phase1_args or [])` and `self._phase2_flag = phase2_flag`.

Updated constructor signature:

```python
def __init__(
    self,
    script_name: str,
    case_path: str,
    file_number: str,
    files: List[str],
    settings: dict,
    phase1_args: List[str] | None = None,
    phase2_flag: str = "--phase=summary",
    parent=None,
):
    super().__init__(case_path, file_number, files, settings, parent)
    self._script_name = script_name
    self._phase1_args = list(phase1_args or [])
    self._phase2_flag = phase2_flag
    # ... rest unchanged
```

In `_start_file`, change the launch line:

```python
def _start_file(self, file_path: str) -> None:
    """Snapshot outputs then launch Phase 1 for file_path."""
    self._pre_existing_outputs = self._scan_outputs()
    self._awaiting_session_path = None
    self._stdout_buf = bytearray()
    argv = [self._script_path()] + self._phase1_args + [file_path]
    self._launch_process(argv)
```

In `resume_with_config`, change the launch line:

```python
def resume_with_config(self, session_path: str) -> None:
    """Start Phase 2 with the configured phase2_flag."""
    self._awaiting_session_path = None
    self._stdout_buf = bytearray()
    self._pre_existing_outputs = self._scan_outputs()
    self._launch_process([self._script_path(), self._phase2_flag, session_path])
```

- [ ] **Step 4: Update TaskSpec**

In `icharlotte_core/ui/wizard/registry.py`, extend the `TaskSpec` dataclass to include the new fields. Update the existing block:

```python
@dataclass(frozen=True)
class TaskSpec:
    task_id: str
    title: str
    description: str
    icon_glyph: str
    script_name: str
    default_folders: List[str] = field(default_factory=list)
    phase1_args: List[str] = field(default_factory=list)
    phase2_flag: str = "--phase=summary"
    _settings_page_cls_factory: Optional[object] = field(default=None, repr=False, compare=False)

    @property
    def settings_page_cls(self) -> type:
        if self._settings_page_cls_factory is not None:
            return self._settings_page_cls_factory()
        return _default_settings_page_cls()
```

The existing entries (`summarize_documents`, etc.) all use defaults so they keep working unchanged.

- [ ] **Step 5: Update task_tab.py**

In `icharlotte_core/ui/wizard/task_tab.py`, modify BOTH `SubprocessWorker(...)` calls to thread the flags through.

For `start_speculative_run()` around line 105, update to:

```python
worker = SubprocessWorker(
    script_name=self._spec.script_name,
    case_path=self._case_path,
    file_number=self._file_number,
    files=self._files,
    settings={},
    phase1_args=list(self._spec.phase1_args),
    phase2_flag=self._spec.phase2_flag,
    parent=self,
)
```

For `_start_run()` around line 175, update to:

```python
self._worker = SubprocessWorker(
    script_name=self._spec.script_name,
    case_path=self._case_path,
    file_number=self._file_number,
    files=self._files,
    settings=settings_dict,
    phase1_args=list(self._spec.phase1_args),
    phase2_flag=self._spec.phase2_flag,
    parent=self,
)
```

- [ ] **Step 6: Add med_chron.py to _TWO_PHASE_SCRIPTS**

In `icharlotte_core/ui/wizard/runners/parallel_subprocess_worker.py`, update line ~55:

```python
_TWO_PHASE_SCRIPTS = {"summarize_deposition.py", "med_chron.py"}
```

- [ ] **Step 7: Run tests to verify they pass**

Run: `python -m pytest tests/test_wizard/test_subprocess_worker.py -v`
Expected: New tests pass; existing tests still pass.

Run the broader wizard test suite:
`python -m pytest tests/test_wizard/ -v`
Expected: all existing wizard tests still pass.

- [ ] **Step 8: Commit**

```bash
git add icharlotte_core/ui/wizard/registry.py \
        icharlotte_core/ui/wizard/runners/subprocess_worker.py \
        icharlotte_core/ui/wizard/runners/parallel_subprocess_worker.py \
        icharlotte_core/ui/wizard/task_tab.py \
        tests/test_wizard/test_subprocess_worker.py
git commit -m "feat(wizard): parameterize SubprocessWorker phase flags via TaskSpec"
```

---

## Task 7: MedChronConfigForm widget

**Files:**
- Create: `icharlotte_core/ui/med_chron_config_form.py`
- Test: `tests/test_wizard/test_med_chron_settings_page.py` (this file will host both form + page tests; we add the form tests now and the page tests in Task 8)

- [ ] **Step 1: Write the failing tests**

Create `tests/test_wizard/test_med_chron_settings_page.py`:

```python
"""Tests for the MedChronConfigForm + MedChronSettingsPage UI."""

import json
import sys
from pathlib import Path

import pytest

pytest.importorskip("pytestqt")  # NOTE: no underscore — pytest_qt silently skips

PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))


def _write_session(tmp_path: Path, *, narrative_missing: bool = False) -> Path:
    cache = tmp_path / ".med_chron" / "abc123"
    cache.mkdir(parents=True)
    session_path = cache / "session.json"
    session_path.write_text(json.dumps({
        "version": 1,
        "phase": "awaiting_input",
        "input_path": str(tmp_path / "rec.docx"),
        "narrative_text_path": str(cache / "narrative.txt"),
        "full_text_path": str(cache / "full.txt"),
        "narrative_missing": narrative_missing,
        "provider_name": "Acme PT",
        "file_number": "1234.567",
        "catalog": [
            {"id": "rewrite_chronology", "title": "Rewrite Chronology",
             "description": "Reformats narrative.", "uses_tables": False,
             "default_selected": True},
            {"id": "inconsistencies", "title": "Inconsistency Check",
             "description": "Find contradictions.", "uses_tables": True,
             "default_selected": False},
        ],
        "user_config": None,
    }, indent=2), encoding="utf-8")
    return session_path


def test_form_renders_one_checkbox_per_catalog_entry(qtbot, tmp_path):
    from icharlotte_core.ui.med_chron_config_form import MedChronConfigForm
    session_path = _write_session(tmp_path)
    form = MedChronConfigForm(session_path)
    qtbot.addWidget(form)
    assert len(form.catalog_checkboxes) == 2


def test_default_selected_checkbox_starts_checked(qtbot, tmp_path):
    from icharlotte_core.ui.med_chron_config_form import MedChronConfigForm
    session_path = _write_session(tmp_path)
    form = MedChronConfigForm(session_path)
    qtbot.addWidget(form)
    rewrite_cb = form.catalog_checkboxes["rewrite_chronology"]
    other_cb = form.catalog_checkboxes["inconsistencies"]
    assert rewrite_cb.isChecked() is True
    assert other_cb.isChecked() is False


def test_narrative_missing_banner_shown(qtbot, tmp_path):
    from icharlotte_core.ui.med_chron_config_form import MedChronConfigForm
    session_path = _write_session(tmp_path, narrative_missing=True)
    form = MedChronConfigForm(session_path)
    qtbot.addWidget(form)
    assert form.narrative_missing_banner.isVisible()


def test_narrative_missing_banner_hidden_by_default(qtbot, tmp_path):
    from icharlotte_core.ui.med_chron_config_form import MedChronConfigForm
    session_path = _write_session(tmp_path, narrative_missing=False)
    form = MedChronConfigForm(session_path)
    qtbot.addWidget(form)
    assert form.narrative_missing_banner.isVisible() is False


def test_proceed_requires_at_least_one_selection(qtbot, tmp_path):
    from icharlotte_core.ui.med_chron_config_form import MedChronConfigForm
    session_path = _write_session(tmp_path)
    form = MedChronConfigForm(session_path)
    qtbot.addWidget(form)
    # Uncheck the default-selected box.
    form.catalog_checkboxes["rewrite_chronology"].setChecked(False)
    # No custom rows. commit should fail.
    assert form.commit_user_config() is False


def test_custom_row_requires_label_and_instruction(qtbot, tmp_path):
    from icharlotte_core.ui.med_chron_config_form import MedChronConfigForm
    session_path = _write_session(tmp_path)
    form = MedChronConfigForm(session_path)
    qtbot.addWidget(form)
    form.catalog_checkboxes["rewrite_chronology"].setChecked(False)
    # Add a row with only a label, no instruction.
    form.add_custom_row()
    row = form.custom_rows[0]
    row.label_edit.setText("Some label")
    # instruction is empty — should fail
    assert form.commit_user_config() is False


def test_empty_custom_rows_silently_dropped(qtbot, tmp_path):
    from icharlotte_core.ui.med_chron_config_form import MedChronConfigForm
    from icharlotte_core.med_chron import session_manager
    session_path = _write_session(tmp_path)
    form = MedChronConfigForm(session_path)
    qtbot.addWidget(form)
    # Add three rows: one valid, two empty.
    form.add_custom_row()
    form.add_custom_row()
    form.add_custom_row()
    form.custom_rows[1].label_edit.setText("Real one")
    form.custom_rows[1].instruction_edit.setPlainText("Do this thing.")
    # rows 0 and 2 are blank — should be dropped.

    assert form.commit_user_config() is True
    data = session_manager.read_session(session_path)
    cfg = data["user_config"]
    assert len(cfg["custom_analyses"]) == 1
    assert cfg["custom_analyses"][0]["label"] == "Real one"


def test_commit_flips_phase_to_ready_to_run(qtbot, tmp_path):
    from icharlotte_core.ui.med_chron_config_form import MedChronConfigForm
    from icharlotte_core.med_chron import session_manager
    session_path = _write_session(tmp_path)
    form = MedChronConfigForm(session_path)
    qtbot.addWidget(form)
    assert form.commit_user_config() is True
    data = session_manager.read_session(session_path)
    assert data["phase"] == "ready_to_run"
    assert data["user_config"]["selected_catalog_ids"] == ["rewrite_chronology"]
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `python -m pytest tests/test_wizard/test_med_chron_settings_page.py -v`
Expected: ImportError — `MedChronConfigForm` does not exist.

- [ ] **Step 3: Create the form widget**

Create `icharlotte_core/ui/med_chron_config_form.py`:

```python
"""Configuration form for the multi-analysis Med-Cron task.

Reads the Phase 1 session JSON, presents:
- a checkbox per curated analysis (Rewrite Chronology pre-checked)
- a "Custom analyses" panel with add/remove rows
- a "narrative missing" warning banner when applicable

On commit_user_config(), validates selection and writes user_config back
to the session, flipping phase to ``ready_to_run``.
"""

from pathlib import Path
from typing import Optional

from PySide6.QtWidgets import (
    QCheckBox, QHBoxLayout, QLabel, QLineEdit, QMessageBox, QPlainTextEdit,
    QPushButton, QScrollArea, QVBoxLayout, QWidget,
)

from icharlotte_core.med_chron import session_manager


class CustomAnalysisRow(QWidget):
    """One row in the custom-analyses list: label + instruction + remove btn."""

    def __init__(self, parent: QWidget, on_remove):
        super().__init__(parent)
        self._on_remove = on_remove

        layout = QVBoxLayout(self)
        layout.setContentsMargins(4, 4, 4, 4)
        layout.setSpacing(4)

        top = QHBoxLayout()
        top.addWidget(QLabel("Label:"))
        self.label_edit = QLineEdit()
        self.label_edit.setPlaceholderText("Short name (e.g. 'Left-knee mentions')")
        top.addWidget(self.label_edit, 1)
        self.remove_btn = QPushButton("−")
        self.remove_btn.setFixedSize(24, 24)
        self.remove_btn.setStyleSheet("QPushButton { color: #c62828; font-weight: bold; }")
        self.remove_btn.clicked.connect(self._handle_remove)
        top.addWidget(self.remove_btn)
        layout.addLayout(top)

        layout.addWidget(QLabel("Request:"))
        self.instruction_edit = QPlainTextEdit()
        self.instruction_edit.setPlaceholderText("Describe the analysis…")
        self.instruction_edit.setFixedHeight(60)
        layout.addWidget(self.instruction_edit)

        self.setStyleSheet(
            "CustomAnalysisRow { border: 1px solid #ddd; border-radius: 4px; }"
        )

    def _handle_remove(self):
        self._on_remove(self)

    def label(self) -> str:
        return self.label_edit.text().strip()

    def instruction(self) -> str:
        return self.instruction_edit.toPlainText().strip()

    def is_empty(self) -> bool:
        return not self.label() and not self.instruction()


class MedChronConfigForm(QWidget):
    """Pickable list of catalog analyses + custom analysis rows."""

    def __init__(self, session_path, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.session_path = Path(session_path)
        self._session = session_manager.read_session(self.session_path)

        root = QVBoxLayout(self)
        root.setContentsMargins(8, 8, 8, 8)
        root.setSpacing(8)

        header = QLabel(
            f"<b>Analyses to run on {self._session.get('provider_name', '')}</b>"
        )
        root.addWidget(header)

        # Narrative-missing warning banner
        self.narrative_missing_banner = QLabel(
            "⚠ Narrative text not found in this document — "
            "Rewrite Chronology will be skipped."
        )
        self.narrative_missing_banner.setStyleSheet(
            "background-color: #FFF3CD; color: #856404; "
            "padding: 6px; border-radius: 4px;"
        )
        self.narrative_missing_banner.setVisible(
            bool(self._session.get("narrative_missing"))
        )
        root.addWidget(self.narrative_missing_banner)

        # Curated catalog section
        cat_label = QLabel("<b>Curated analyses:</b>")
        root.addWidget(cat_label)

        self.catalog_checkboxes: dict[str, QCheckBox] = {}
        for entry in self._session.get("catalog", []):
            row = QWidget()
            row_layout = QVBoxLayout(row)
            row_layout.setContentsMargins(16, 2, 0, 2)
            row_layout.setSpacing(0)
            cb = QCheckBox(entry.get("title", entry.get("id", "")))
            cb.setChecked(bool(entry.get("default_selected")))
            self.catalog_checkboxes[entry["id"]] = cb
            row_layout.addWidget(cb)
            desc = QLabel(entry.get("description", ""))
            desc.setStyleSheet("color: #666; font-size: 11px; padding-left: 22px;")
            desc.setWordWrap(True)
            row_layout.addWidget(desc)
            root.addWidget(row)

        # Custom analyses section
        custom_header = QHBoxLayout()
        custom_header.addWidget(QLabel("<b>Custom analyses:</b>"))
        custom_header.addStretch()
        root.addLayout(custom_header)

        # Scrollable container for custom rows.
        self._custom_container = QWidget()
        self._custom_container_layout = QVBoxLayout(self._custom_container)
        self._custom_container_layout.setContentsMargins(0, 0, 0, 0)
        self._custom_container_layout.setSpacing(4)
        self._custom_container_layout.addStretch()  # pin rows to top

        scroll = QScrollArea()
        scroll.setWidget(self._custom_container)
        scroll.setWidgetResizable(True)
        scroll.setMinimumHeight(80)
        scroll.setMaximumHeight(220)
        root.addWidget(scroll)

        self.custom_rows: list[CustomAnalysisRow] = []

        add_btn = QPushButton("+ Add custom analysis")
        add_btn.clicked.connect(self.add_custom_row)
        root.addWidget(add_btn)

        # Inline validation error label.
        self._error_label = QLabel("")
        self._error_label.setStyleSheet("color: #c62828; font-style: italic;")
        self._error_label.setVisible(False)
        root.addWidget(self._error_label)

    def add_custom_row(self) -> CustomAnalysisRow:
        row = CustomAnalysisRow(self._custom_container, self._remove_custom_row)
        # Insert above the stretch at the end.
        idx = self._custom_container_layout.count() - 1
        self._custom_container_layout.insertWidget(idx, row)
        self.custom_rows.append(row)
        return row

    def _remove_custom_row(self, row: CustomAnalysisRow) -> None:
        if row in self.custom_rows:
            self.custom_rows.remove(row)
        self._custom_container_layout.removeWidget(row)
        row.deleteLater()

    def _selected_catalog_ids(self) -> list[str]:
        return [cid for cid, cb in self.catalog_checkboxes.items() if cb.isChecked()]

    def _validated_custom_rows(self) -> tuple[list[dict], str]:
        """Return (clean_rows, error_msg). Empty rows are silently dropped.

        Partially-filled rows (one of label/instruction missing) are an error.
        """
        clean = []
        for r in self.custom_rows:
            if r.is_empty():
                continue
            lbl, instr = r.label(), r.instruction()
            if not lbl or not instr:
                return [], (
                    "Custom analyses need both a label and an instruction. "
                    "Fill in (or remove) the partially-completed row."
                )
            clean.append({"label": lbl, "instruction": instr})
        return clean, ""

    def commit_user_config(self) -> bool:
        """Validate and write user_config. Returns True on success."""
        self._error_label.setVisible(False)

        selected = self._selected_catalog_ids()
        clean_custom, err = self._validated_custom_rows()
        if err:
            self._error_label.setText(err)
            self._error_label.setVisible(True)
            return False
        if not selected and not clean_custom:
            self._error_label.setText(
                "Select at least one analysis, or add a custom analysis."
            )
            self._error_label.setVisible(True)
            return False

        session_manager.update_user_config(
            self.session_path,
            {
                "selected_catalog_ids": selected,
                "custom_analyses": clean_custom,
            },
        )
        return True
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `python -m pytest tests/test_wizard/test_med_chron_settings_page.py -v`
Expected: 8 passed.

- [ ] **Step 5: Commit**

```bash
git add icharlotte_core/ui/med_chron_config_form.py \
        tests/test_wizard/test_med_chron_settings_page.py
git commit -m "feat(med-chron): MedChronConfigForm with catalog + custom rows"
```

---

## Task 8: MedChronSettingsPage

**Files:**
- Create: `icharlotte_core/ui/wizard/pages/med_chron_settings_page.py`
- Test: `tests/test_wizard/test_med_chron_settings_page.py` (append page-level tests)

- [ ] **Step 1: Write the failing tests**

Append to `tests/test_wizard/test_med_chron_settings_page.py`:

```python
def test_settings_page_proceed_disabled_until_phase1_completes(qtbot, tmp_path):
    from icharlotte_core.ui.wizard.pages.med_chron_settings_page import MedChronSettingsPage
    from icharlotte_core.ui.wizard.registry import TaskSpec

    spec = TaskSpec(
        task_id="med_chron_analysis",
        title="Med Chron Analysis",
        description="…",
        icon_glyph="🩺",
        script_name="med_chron.py",
    )
    page = MedChronSettingsPage(spec, files=[str(tmp_path / "rec.docx")])
    qtbot.addWidget(page)
    assert page.proceed_btn.isEnabled() is False


def test_settings_page_swaps_to_form_on_awaiting_input(qtbot, tmp_path):
    from icharlotte_core.ui.wizard.pages.med_chron_settings_page import MedChronSettingsPage
    from icharlotte_core.ui.wizard.registry import TaskSpec

    spec = TaskSpec(
        task_id="med_chron_analysis",
        title="Med Chron Analysis",
        description="…",
        icon_glyph="🩺",
        script_name="med_chron.py",
    )
    page = MedChronSettingsPage(spec, files=[str(tmp_path / "rec.docx")])
    qtbot.addWidget(page)

    # Build a valid session for the form to read.
    session_path = _write_session(tmp_path)

    # Simulate Phase 1 completion.
    page._on_phase1_complete(str(session_path))

    assert page._stack.currentIndex() == 1   # form page
    assert page.proceed_btn.isEnabled() is True
    assert page._form is not None
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `python -m pytest tests/test_wizard/test_med_chron_settings_page.py -v`
Expected: ImportError — `MedChronSettingsPage` does not exist.

- [ ] **Step 3: Create the settings page**

Create `icharlotte_core/ui/wizard/pages/med_chron_settings_page.py`:

```python
"""MedChronSettingsPage — inline settings for the Med Chron Analysis task.

Mirrors DepositionSettingsPage:
1. Speculative Phase 1 (prep) launched as soon as the tab opens.
2. While extraction runs, a "Preparing chronology…" spinner is shown.
3. On AWAITING_INPUT, the MedChronConfigForm is built and swapped in.
4. User picks analyses + custom rows, clicks Proceed.
5. _on_proceed() validates via form.commit_user_config() then emits
   phase2_requested(session_path).
"""

from typing import Optional

from PySide6.QtCore import Signal
from PySide6.QtWidgets import (
    QLabel, QProgressBar, QSizePolicy, QStackedWidget, QVBoxLayout, QWidget,
)

from ..registry import TaskSpec
from .settings_page import SettingsPage


class MedChronSettingsPage(SettingsPage):
    """SettingsPage subclass that embeds MedChronConfigForm inline."""

    phase2_requested = Signal(str)  # carries session_path

    def __init__(
        self,
        spec: TaskSpec,
        files,
        case_root: str | None = None,
        parent: QWidget | None = None,
    ):
        super().__init__(spec, files=files, case_root=case_root, parent=parent)

        self._session_path: Optional[str] = None
        self._form = None

        # Page 0: "Preparing chronology…"
        prep_widget = QWidget()
        prep_layout = QVBoxLayout(prep_widget)
        prep_layout.setContentsMargins(8, 16, 8, 8)
        prep_layout.setSpacing(8)

        title = QLabel("Preparing chronology…")
        title.setStyleSheet("font-size: 13px; font-weight: 600; color: #1976D2;")
        prep_layout.addWidget(title)

        self._progress = QProgressBar()
        self._progress.setRange(0, 100)
        self._progress.setValue(0)
        prep_layout.addWidget(self._progress)

        self._small_status = QLabel("")
        self._small_status.setStyleSheet("color: #555; font-size: 11px;")
        self._small_status.setWordWrap(True)
        prep_layout.addWidget(self._small_status)
        prep_layout.addStretch()

        # Page 1: form placeholder (populated on phase 1 complete)
        self._form_placeholder = QWidget()
        self._form_placeholder_layout = QVBoxLayout(self._form_placeholder)
        self._form_placeholder_layout.setContentsMargins(0, 0, 0, 0)

        # Stack
        self._stack = QStackedWidget()
        self._stack.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding
        )
        self._stack.addWidget(prep_widget)            # index 0
        self._stack.addWidget(self._form_placeholder) # index 1
        self._stack.setCurrentIndex(0)

        outer = self.layout()
        # Remove the "Settings for … to be defined." placeholder body (index 3)
        item = outer.itemAt(3)
        if item is not None:
            w = item.widget()
            if w is not None:
                outer.removeWidget(w)
                w.deleteLater()
        outer.insertWidget(3, self._stack, 1)

        self.proceed_btn.setEnabled(False)

    def attach_worker(self, worker) -> bool:
        worker.status.connect(self._small_status.setText)
        worker.progress.connect(self._progress.setValue)
        worker.awaiting_input.connect(self._on_phase1_complete)
        worker.failed.connect(self._on_phase1_failed)
        return True

    def _on_phase1_complete(self, session_path: str) -> None:
        from icharlotte_core.ui.med_chron_config_form import MedChronConfigForm

        self._session_path = session_path
        try:
            self._form = MedChronConfigForm(session_path, parent=self._form_placeholder)
        except Exception as e:
            self._on_phase1_failed(f"Could not load analysis picker: {e}")
            return
        self._form_placeholder_layout.addWidget(self._form)
        self._stack.setCurrentIndex(1)
        self.proceed_btn.setEnabled(True)

    def _on_phase1_failed(self, err: str) -> None:
        self._small_status.setStyleSheet("color: #c62828; font-size: 11px;")
        self._small_status.setText(f"Preparation failed: {err}")

    def _on_proceed(self) -> None:
        if self._form is None:
            return
        if not self._form.commit_user_config():
            return  # validation failed; form showed error
        self.phase2_requested.emit(self._session_path)
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `python -m pytest tests/test_wizard/test_med_chron_settings_page.py -v`
Expected: all tests pass (form + page).

- [ ] **Step 5: Commit**

```bash
git add icharlotte_core/ui/wizard/pages/med_chron_settings_page.py \
        tests/test_wizard/test_med_chron_settings_page.py
git commit -m "feat(med-chron): MedChronSettingsPage with speculative Phase 1"
```

---

## Task 9: Register the task in the wizard

**Files:**
- Modify: `icharlotte_core/ui/wizard/registry.py` — add `_med_chron_settings_page_cls` factory + `med_chron_analysis` TaskSpec entry.
- Test: `tests/test_wizard/test_med_chron_registry.py`

- [ ] **Step 1: Write the failing tests**

Create `tests/test_wizard/test_med_chron_registry.py`:

```python
"""Tests for the med_chron_analysis wizard task registration."""

import sys
from pathlib import Path

import pytest

pytest.importorskip("pytestqt")

PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))


def test_med_chron_analysis_task_is_registered():
    from icharlotte_core.ui.wizard.registry import TASK_REGISTRY
    assert "med_chron_analysis" in TASK_REGISTRY


def test_med_chron_task_uses_med_chron_settings_page():
    from icharlotte_core.ui.wizard.registry import get_task
    from icharlotte_core.ui.wizard.pages.med_chron_settings_page import MedChronSettingsPage
    spec = get_task("med_chron_analysis")
    assert spec.settings_page_cls is MedChronSettingsPage


def test_med_chron_task_phase_flags_set():
    from icharlotte_core.ui.wizard.registry import get_task
    spec = get_task("med_chron_analysis")
    assert "--phase=prep" in list(spec.phase1_args)
    assert spec.phase2_flag == "--phase=run"


def test_med_chron_task_script_name():
    from icharlotte_core.ui.wizard.registry import get_task
    spec = get_task("med_chron_analysis")
    assert spec.script_name == "med_chron.py"


def test_existing_deposition_task_phase_flags_unchanged():
    """Defaults must keep existing deposition behavior intact."""
    from icharlotte_core.ui.wizard.registry import get_task
    spec = get_task("summarize_depositions")
    assert spec.phase2_flag == "--phase=summary"
    assert list(spec.phase1_args) == []  # no extra flag for phase 1
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `python -m pytest tests/test_wizard/test_med_chron_registry.py -v`
Expected: KeyError on `"med_chron_analysis"` — not yet registered.

- [ ] **Step 3: Register the task**

In `icharlotte_core/ui/wizard/registry.py`, ADD the factory function near the other ones:

```python
def _med_chron_settings_page_cls():
    from .pages.med_chron_settings_page import MedChronSettingsPage
    return MedChronSettingsPage
```

ADD the new entry to `TASK_REGISTRY` (after the existing `medical_records` entry):

```python
"med_chron_analysis": TaskSpec(
    task_id="med_chron_analysis",
    title="Med Chron Analysis",
    description="Run selectable analyses on a medical chronology.",
    icon_glyph="\U0001FA7A",  # 🩺
    script_name="med_chron.py",
    default_folders=["NOTES/AI OUTPUT", "RECORDS"],
    phase1_args=["--phase=prep"],
    phase2_flag="--phase=run",
    _settings_page_cls_factory=_med_chron_settings_page_cls,
),
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `python -m pytest tests/test_wizard/test_med_chron_registry.py -v`
Expected: 5 passed.

Run the wider wizard suite for regressions:
`python -m pytest tests/test_wizard/ -v`
Expected: all green.

- [ ] **Step 5: Commit**

```bash
git add icharlotte_core/ui/wizard/registry.py \
        tests/test_wizard/test_med_chron_registry.py
git commit -m "feat(wizard): register med_chron_analysis task"
```

---

## Task 10: End-to-end smoke test + manual verification

**Files:**
- Modify: `tests/test_med_chron/test_phase2_runner.py` — add one integration test that runs prep → user_config write → run, end-to-end, with a real (stubbed-LLM) call chain.

- [ ] **Step 1: Write the end-to-end integration test**

Append to `tests/test_med_chron/test_phase2_runner.py`:

```python
def test_full_prep_to_run_pipeline(tmp_path):
    """End-to-end: prep produces session, simulate user_config, run produces docx."""
    from icharlotte_core.med_chron import session_manager
    from docx import Document

    # Build a chronology with both narrative and table content.
    src = tmp_path / "1234-001_ E2E PT.docx"
    doc = Document()
    doc.add_paragraph(
        "BRIEF SYNOPSIS OF PRE-INJURY MEDICAL RECORD: "
        "No prior complaints."
    )
    doc.add_paragraph(
        "BRIEF SYNOPSIS OF POST-INJURY MEDICAL RECORD: "
        "Lumbar strain treated by PT."
    )
    table = doc.add_table(rows=2, cols=2)
    table.cell(0, 0).text = "Date"
    table.cell(0, 1).text = "Provider"
    table.cell(1, 0).text = "2024-02-01"
    table.cell(1, 1).text = "E2E_TABLE_TOKEN"
    doc.save(str(src))

    out_dir = tmp_path / "NOTES" / "AI OUTPUT"

    # Phase 1
    assert med_chron.process_prep(str(src), str(out_dir)) == 0

    paths = session_manager.compute_session_paths(str(src), str(out_dir))

    # Simulate the UI writing user_config.
    session_manager.update_user_config(
        paths.session_path,
        {
            "selected_catalog_ids": ["rewrite_chronology", "inconsistencies"],
            "custom_analyses": [
                {"label": "E2E custom", "instruction": "Find E2E_TABLE_TOKEN."},
            ],
        },
    )

    # Phase 2 with stubbed LLM
    seen_texts = []
    def stub(prompt, text, **kw):
        seen_texts.append(text)
        return "# stub\nbody"

    with patch.object(med_chron.LLMCaller, "call", side_effect=stub):
        assert med_chron.process_run(str(paths.session_path), str(out_dir)) == 0

    out_docs = sorted(p.name for p in out_dir.glob("*.docx"))
    # 3 analyses → 3 docx files at out_dir root
    assert len(out_docs) == 3
    assert any("rewrite_chronology" in n for n in out_docs)
    assert any("inconsistencies" in n for n in out_docs)
    assert any("custom_1_e2e_custom" in n for n in out_docs)

    # Confirm narrative-only vs full-text routing.
    # The rewrite call sees no table token; the others do.
    rewrite_text = [t for t in seen_texts if "E2E_TABLE_TOKEN" not in t]
    full_text = [t for t in seen_texts if "E2E_TABLE_TOKEN" in t]
    assert len(rewrite_text) == 1  # rewrite only
    assert len(full_text) == 2     # inconsistencies + custom
```

- [ ] **Step 2: Run all med-chron tests**

Run: `python -m pytest tests/test_med_chron/ tests/test_wizard/test_med_chron_settings_page.py tests/test_wizard/test_med_chron_registry.py -v`
Expected: all tests pass.

- [ ] **Step 3: Manual UI verification**

Launch the app: `python iCharlotte.py`

In Wizard Mode:
1. Pick a real medical chronology .docx that has the BRIEF SYNOPSIS PRE/POST-INJURY headings.
2. Click the "Med Chron Analysis" task card.
3. Verify the file appears in the file list.
4. Verify "Preparing chronology…" appears, then swaps to the picker.
5. Verify "Rewrite Chronology" is pre-checked, others are not.
6. Check "Inconsistency Check". Click "+ Add custom analysis", give it a label and instruction.
7. Click Proceed. Status page shows progress per analysis.
8. Output page shows the produced docx files. Open each to confirm content.

Repeat with a chronology that LACKS the BRIEF SYNOPSIS headings:
- Verify the yellow "narrative missing" banner appears in the picker.
- Verify Rewrite Chronology is skipped at run-time (only the other analyses produce docx files).

Verify legacy IndexTab path still works (or at minimum that `python Scripts/med_chron.py <file>` from the command line produces the original `med_chron_<file>.docx`).

- [ ] **Step 4: Commit if anything changed during manual verification**

If you spotted and fixed an issue during step 3:

```bash
git add <changed files>
git commit -m "fix(med-chron): <one-line fix description>"
```

Otherwise no commit needed — the test commit from step 1 stands on its own.

---

## Final checklist

- [ ] All new tests pass: `python -m pytest tests/test_med_chron/ tests/test_wizard/test_med_chron_settings_page.py tests/test_wizard/test_med_chron_registry.py -v`
- [ ] Wizard regression suite passes: `python -m pytest tests/test_wizard/ -v`
- [ ] Legacy `python Scripts/med_chron.py <file>` still produces `med_chron_<filename>.docx`
- [ ] Manual UI verification of the full flow (curated + custom analyses)
- [ ] Manual UI verification of the narrative-missing path
- [ ] No new `# TODO` markers left in the diff

---

## Risks & Things to Watch

1. **`Path` import in `med_chron.py`**: the existing file may not import `Path` from `pathlib`. Task 4 assumes it does; if not, add `from pathlib import Path` to the imports up top.

2. **`extract_docx_text` for older `.docx`**: occasionally a malformed docx makes the document_processor raise. Phase 1 catches via the `if not full_text` check and exits 1. The wizard surfaces this via the `failed` signal; the user sees the error in the small-status label.

3. **LLMCaller rate limits**: 4 concurrent threads × multiple chronologies in parallel task tabs is a lot. If the team hits 429s in practice, lower `max_workers` from 4 to 2 in `process_run`.

4. **`process_legacy` filename collision risk**: legacy mode writes to `med_chron_<file>.docx`, but Phase 2 writes to `med_chron_<analysis_id>_<file>.docx`. Different filename patterns — no collision.

5. **Existing `MED_CHRON_PROMPT.txt`** at `Scripts/MED_CHRON_PROMPT.txt` is now redundant. The plan does NOT delete it (we keep it as a backstop in case anything outside this codebase reads it). If desired, a follow-up cleanup task can remove it once we're confident nothing references it externally.
