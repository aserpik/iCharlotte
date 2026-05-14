# Interactive Deposition Summary Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Split `summarize_deposition` into a two-phase interactive workflow — phase 1 proposes ranked testimony topics and pauses for user input via a popup, phase 2 generates a topic-locked summary from the user's choices.

**Architecture:** Two subprocess invocations of `Scripts/summarize_deposition.py` coordinated through a sidecar JSON session file written to `logs/depo_sessions/`. A new `AWAITING_INPUT:` stdout token tells `AgentRunner` to surface a READY button on the status row; the popup writes a `user_config` block back into the session before the second subprocess is launched.

**Tech Stack:** Python 3, PyQt6/PySide6 (`QDialog`, `QProcess`, `QObject` signals), pytest + pytest-qt, existing `LLMCaller` / `AgentLogger` / `DocumentProcessor` infrastructure.

**Spec:** `docs/superpowers/specs/2026-05-14-interactive-depo-summary-design.md`

---

## File map

**New files:**
- `icharlotte_core/deposition/session_manager.py` — read/write the sidecar session JSON; atomic update; cleanup
- `icharlotte_core/ui/depo_summary_config_dialog.py` — popup `QDialog`
- `Scripts/DEPOSITION_TOPIC_DISCOVERY_PROMPT.txt` — new LLM prompt
- `tests/test_deposition/__init__.py`
- `tests/test_deposition/test_session_manager.py`
- `tests/test_deposition/test_summarize_deposition_phases.py`
- `tests/test_deposition/test_agent_runner_awaiting_input.py`
- `tests/test_deposition/test_depo_summary_config_dialog.py`
- `tests/test_deposition/test_full_flow_smoke.py`

**Modified files:**
- `Scripts/summarize_deposition.py` — split `process_document` into `process_topics` and `process_summary`; add `--phase` dispatcher in `main()`; delete `ExhibitExtractor`, `ImpeachmentDetector`, parallel extraction pass, `build_narrative_prompt` topic-count logic
- `Scripts/SUMMARIZE_DEPOSITION_PROMPT.txt` — rewrite as topic-locked template with `{deponent_label}`, `{bullets_per_topic}`, `{topic_list}`, `{custom_rules}` placeholders
- `Scripts/DEPOSITION_CROSS_CHECK_PROMPT.txt` — drop `{extraction}` placeholder and all impeachment / exhibits language
- `icharlotte_core/ui/widgets.py` — `AgentRunner.awaiting_input` signal, `AWAITING_INPUT:` parse branch, paused-exit handling in `handle_finished`, `resume_with_config` method; `StatusWidget.on_awaiting_input` slot, READY button, `ready_clicked` signal
- `icharlotte_core/ui/tabs.py` — wire `AgentRunner.awaiting_input` → open dialog → `resume_with_config` for the deposition agent

---

## Task 1: Session manager module

Creates the shared module both phases of the agent (and the popup) use to read/write the sidecar JSON. Pure logic — easy to TDD.

**Files:**
- Create: `icharlotte_core/deposition/session_manager.py`
- Create: `tests/test_deposition/__init__.py`
- Create: `tests/test_deposition/test_session_manager.py`

- [ ] **Step 1: Write the failing tests**

Create `tests/test_deposition/__init__.py` as an empty file.

Create `tests/test_deposition/test_session_manager.py`:

```python
"""Tests for icharlotte_core.deposition.session_manager."""

import json
import os
from pathlib import Path

import pytest

from icharlotte_core.deposition import session_manager


def test_compute_session_paths_are_deterministic_per_input(tmp_path, monkeypatch):
    monkeypatch.setattr(session_manager, "SESSION_DIR", tmp_path)
    # Same input twice in same second yields different timestamps but same hash prefix
    p1 = session_manager.compute_session_paths(r"Z:\foo\Smith Depo.pdf")
    p2 = session_manager.compute_session_paths(r"Z:\foo\Smith Depo.pdf")
    assert p1.session_path.name.split("_", 1)[0] == p2.session_path.name.split("_", 1)[0]
    assert p1.cached_text_path.suffix == ".txt"
    assert p1.session_path.suffix == ".json"
    assert p1.session_path.parent == tmp_path


def test_write_and_read_session_round_trip(tmp_path, monkeypatch):
    monkeypatch.setattr(session_manager, "SESSION_DIR", tmp_path)
    paths = session_manager.compute_session_paths(r"Z:\foo\X.pdf")
    data = {
        "version": 1,
        "phase": "awaiting_input",
        "input_path": r"Z:\foo\X.pdf",
        "cached_text_path": str(paths.cached_text_path),
        "deponent_name": "John Smith",
        "deposition_date": "January 15, 2024",
        "deponent_type": "Plaintiff",
        "file_number": "3850.084",
        "topics": [{"id": 1, "title": "Topic A", "rank": 1, "discussion_density": "high"}],
        "user_config": None,
    }
    session_manager.write_session(paths.session_path, data)
    loaded = session_manager.read_session(paths.session_path)
    assert loaded == data


def test_write_session_is_atomic(tmp_path, monkeypatch):
    """If os.replace fails, the original file is not corrupted."""
    monkeypatch.setattr(session_manager, "SESSION_DIR", tmp_path)
    session_path = tmp_path / "s.json"
    session_path.write_text('{"phase": "awaiting_input", "version": 1}', encoding="utf-8")

    def boom(*a, **kw):
        raise OSError("simulated replace failure")

    monkeypatch.setattr(os, "replace", boom)
    with pytest.raises(OSError):
        session_manager.write_session(session_path, {"phase": "ready_for_summary", "version": 1})

    # Original file untouched
    assert json.loads(session_path.read_text(encoding="utf-8"))["phase"] == "awaiting_input"


def test_update_user_config_flips_phase_and_writes_config(tmp_path, monkeypatch):
    monkeypatch.setattr(session_manager, "SESSION_DIR", tmp_path)
    session_path = tmp_path / "s.json"
    session_manager.write_session(session_path, {
        "version": 1,
        "phase": "awaiting_input",
        "user_config": None,
        "topics": [],
    })

    config = {
        "selected_topics": ["A"],
        "added_topics": ["B"],
        "bullets_per_topic": 5,
        "deponent_label": "Plaintiff",
        "custom_rules": "Use past tense.",
        "cross_check_enabled": True,
    }
    session_manager.update_user_config(session_path, config)

    loaded = session_manager.read_session(session_path)
    assert loaded["phase"] == "ready_for_summary"
    assert loaded["user_config"] == config


def test_cleanup_session_removes_both_files(tmp_path, monkeypatch):
    monkeypatch.setattr(session_manager, "SESSION_DIR", tmp_path)
    session_path = tmp_path / "s.json"
    cached = tmp_path / "s.txt"
    session_path.write_text('{"cached_text_path": "' + str(cached).replace("\\", "\\\\") + '"}', encoding="utf-8")
    cached.write_text("transcript text", encoding="utf-8")

    session_manager.cleanup_session(session_path)
    assert not session_path.exists()
    assert not cached.exists()


def test_cleanup_session_tolerates_missing_files(tmp_path, monkeypatch):
    monkeypatch.setattr(session_manager, "SESSION_DIR", tmp_path)
    session_path = tmp_path / "missing.json"
    # Should not raise even though neither file exists.
    session_manager.cleanup_session(session_path)
```

- [ ] **Step 2: Run tests to verify they fail**

```
python -m pytest tests/test_deposition/test_session_manager.py -v
```

Expected: ImportError on `icharlotte_core.deposition.session_manager`.

- [ ] **Step 3: Implement the session manager**

Create `icharlotte_core/deposition/session_manager.py`:

```python
"""Sidecar JSON session management for the interactive deposition summary flow.

Phase 1 of the deposition agent writes a session file describing the
proposed topics and pauses. The UI loads it, lets the user edit, writes
back a user_config block, and launches phase 2. Phase 2 reads the
session, generates the summary, and cleans up.
"""

import hashlib
import json
import os
import time
from dataclasses import dataclass
from pathlib import Path

# Resolved relative to the iCharlotte project root (icharlotte_core/.. == project)
_PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent
SESSION_DIR = _PROJECT_ROOT / "logs" / "depo_sessions"


@dataclass(frozen=True)
class SessionPaths:
    session_path: Path
    cached_text_path: Path


def compute_session_paths(input_path: str) -> SessionPaths:
    """Build a unique (session_json, cached_text) path pair for an input file."""
    digest = hashlib.sha1(os.fspath(input_path).encode("utf-8")).hexdigest()[:12]
    ts = time.strftime("%Y%m%d_%H%M%S")
    base = f"{digest}_{ts}"
    SESSION_DIR.mkdir(parents=True, exist_ok=True)
    return SessionPaths(
        session_path=SESSION_DIR / f"{base}.json",
        cached_text_path=SESSION_DIR / f"{base}.txt",
    )


def write_session(session_path: Path, data: dict) -> None:
    """Atomically write the session JSON via tmp file + os.replace."""
    session_path = Path(session_path)
    session_path.parent.mkdir(parents=True, exist_ok=True)
    tmp = session_path.with_suffix(session_path.suffix + ".tmp")
    tmp.write_text(json.dumps(data, indent=2), encoding="utf-8")
    os.replace(tmp, session_path)


def read_session(session_path: Path) -> dict:
    return json.loads(Path(session_path).read_text(encoding="utf-8"))


def update_user_config(session_path: Path, user_config: dict) -> None:
    """Load, set user_config, flip phase to 'ready_for_summary', write atomically."""
    data = read_session(session_path)
    data["user_config"] = user_config
    data["phase"] = "ready_for_summary"
    write_session(session_path, data)


def cleanup_session(session_path: Path) -> None:
    """Delete the session JSON and its cached transcript. Tolerant of missing files."""
    session_path = Path(session_path)
    try:
        data = read_session(session_path)
        cached = data.get("cached_text_path")
        if cached:
            cached_path = Path(cached)
            if cached_path.exists():
                cached_path.unlink()
    except (FileNotFoundError, json.JSONDecodeError):
        pass
    if session_path.exists():
        session_path.unlink()
```

- [ ] **Step 4: Run tests to verify they pass**

```
python -m pytest tests/test_deposition/test_session_manager.py -v
```

Expected: all 6 tests PASS.

- [ ] **Step 5: Commit**

```bash
git add icharlotte_core/deposition/session_manager.py tests/test_deposition/__init__.py tests/test_deposition/test_session_manager.py
git commit -m "feat(deposition): session manager for interactive summary flow"
```

---

## Task 2: Topic-discovery prompt file

The new LLM prompt that ranks testimony topics. Plain content file — no test needed; covered by the phase-1 LLM mock tests in Task 5.

**Files:**
- Create: `Scripts/DEPOSITION_TOPIC_DISCOVERY_PROMPT.txt`

- [ ] **Step 1: Write the prompt**

Create `Scripts/DEPOSITION_TOPIC_DISCOVERY_PROMPT.txt`:

```
You are reviewing a deposition transcript to identify the testimony topics that a litigator would want to highlight in a summary. Your job is NOT to summarize — only to produce a ranked list of topic titles.

INSTRUCTIONS:
1. Read the entire transcript carefully.
2. Identify the discrete testimony topics the deponent was questioned about.
3. Rank them from most important / most discussed at the top to least important / least discussed at the bottom. Importance considers both how much time was spent on the topic and how legally significant the testimony is (admissions, mechanism of injury, damages, credibility issues, etc.).
4. Return your answer as a JSON array. No prose, no markdown fences, no commentary. JSON only.

JSON SCHEMA:
[
  {"title": "Short topic title in title case", "rank": 1, "discussion_density": "high|medium|low"},
  {"title": "Next topic", "rank": 2, "discussion_density": "high"},
  ...
]

RULES:
- "title" is 3–8 words, capitalized title case, no trailing punctuation.
- "rank" is a 1-based integer matching the array position.
- "discussion_density" is one of: "high" (extensive testimony, many pages), "medium" (multiple exchanges), "low" (touched on briefly).
- Produce between 8 and 25 topics. Prefer specific topics over broad ones (e.g., "Pre-Accident Lower Back Treatment" beats "Medical History").
- Do not include topics about procedural matters (objections, breaks, exhibit-marking discussions).
- Return JSON only. The first character of your response must be "[" and the last must be "]".
```

- [ ] **Step 2: Commit**

```bash
git add Scripts/DEPOSITION_TOPIC_DISCOVERY_PROMPT.txt
git commit -m "feat(deposition): topic-discovery LLM prompt"
```

---

## Task 3: Rewrite summary prompt as topic-locked template

The current `SUMMARIZE_DEPOSITION_PROMPT.txt` asks the LLM to pick topics. Phase 2 now provides the topics, bullet count, deponent label, and custom rules — the LLM just fills in bullets under the supplied headings.

**Files:**
- Modify: `Scripts/SUMMARIZE_DEPOSITION_PROMPT.txt` (full rewrite)

- [ ] **Step 1: Rewrite the file**

Replace the entire contents of `Scripts/SUMMARIZE_DEPOSITION_PROMPT.txt` with:

```
You are summarizing a deposition transcript using a fixed set of topic headings supplied by the attorney. You do not choose the topics. You do not add or omit topics. You write bullet points under the headings provided.

DEPONENT LABEL: {deponent_label}
BULLETS PER TOPIC: {bullets_per_topic}

TOPIC HEADINGS (use exactly these, in this order):
{topic_list}

OUTPUT RULES:
1. Start the summary with this sentence: "The parties took the deposition of [name of deponent] on [date of deposition]. Below are the most salient portions of [name of deponent]'s testimony:"
2. Under each topic heading, write exactly {bullets_per_topic} bullet points summarizing what the deponent testified regarding that topic. If the transcript contains less testimony on a topic than requested, write fewer bullets — never invent testimony.
3. Each topic heading appears on its own line, in bold (use **Heading** markdown), title-cased, with no numbering and no trailing punctuation.
4. Each bullet is at least two complete sentences. Use markdown dashes ("- ") at the start of each bullet.
5. Refer to the deponent as "{deponent_label}" throughout. Do not use the deponent's full name except in the opening sentence.
6. Summarize testimony directly. Do not use introductory clauses like "Regarding," "Concerning," "With respect to," or "When asked about." Write "Plaintiff testified he broke his leg," not "Regarding his injuries, Plaintiff testified he broke his leg."
7. Avoid repeating phrases that flag content as testimony. Write "Joe texted Plaintiff photos," not "Plaintiff testified that Joe texted her photos."
8. Include testimony favorable to the defense where present.
9. Do not add introductory or concluding paragraphs beyond rule 1's opening sentence. Do not add an exhibits section, impeachment section, or any section not in the topic list.

CUSTOM RULES FROM THE ATTORNEY (apply in addition to the rules above):
{custom_rules}
```

- [ ] **Step 2: Commit**

```bash
git add Scripts/SUMMARIZE_DEPOSITION_PROMPT.txt
git commit -m "feat(deposition): rewrite summary prompt as topic-locked template"
```

---

## Task 4: Strip impeachment/exhibits from cross-check prompt

Cross-check now compares summary against the original transcript only — no structured extraction input, no impeachment alerts, no exhibit-list enhancement.

**Files:**
- Modify: `Scripts/DEPOSITION_CROSS_CHECK_PROMPT.txt` (full rewrite)

- [ ] **Step 1: Rewrite the file**

Replace the entire contents of `Scripts/DEPOSITION_CROSS_CHECK_PROMPT.txt` with:

```
You are a quality assurance reviewer for legal deposition summaries. Compare the NARRATIVE SUMMARY against the ORIGINAL TRANSCRIPT and return a corrected version of the summary.

WHAT TO CHECK:
- Every factual claim in the summary is supported by the transcript.
- Deponent name and deposition date are accurate.
- Attribution is correct (testimony is attributed to the right speaker).
- Wording matches what the deponent actually said in substance.

WHAT TO PRESERVE:
- The set of topic headings and their order — do not add, remove, or reorder headings.
- The bullet count under each heading — keep the same number of bullets unless a bullet must be removed because it is unsupported by the transcript.
- Markdown formatting: bold headings (**Heading**), bullets starting with "- ".
- The opening sentence ("The parties took the deposition of...").

WHAT TO RETURN:
- The corrected summary, ready for direct use. No meta-commentary about your review process.
- No new sections (no exhibits list, no impeachment alerts, no closing remarks).

=== NARRATIVE SUMMARY ===
{summary}

=== ORIGINAL TRANSCRIPT (first 75000 characters) ===
{original}
```

- [ ] **Step 2: Commit**

```bash
git add Scripts/DEPOSITION_CROSS_CHECK_PROMPT.txt
git commit -m "feat(deposition): simplify cross-check prompt for topic-locked flow"
```

---

## Task 5: Phase 1 — `process_topics` function

Phase 1 extracts text, caches it, runs the topic-discovery LLM call, writes the session JSON, prints `AWAITING_INPUT:` and exits.

**Files:**
- Modify: `Scripts/summarize_deposition.py` (add `process_topics`, helpers, prompt path constant)
- Create: `tests/test_deposition/test_summarize_deposition_phases.py` (initial tests for phase 1)

- [ ] **Step 1: Write the failing tests**

Create `tests/test_deposition/test_summarize_deposition_phases.py`:

```python
"""Tests for the two-phase summarize_deposition agent."""

import json
import os
import sys
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

import pytest

# Make Scripts/ importable
PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(PROJECT_ROOT / "Scripts"))

import summarize_deposition  # noqa: E402
from icharlotte_core.deposition import session_manager  # noqa: E402


FAKE_TRANSCRIPT = (
    "DEPOSITION OF JOHN SMITH\n"
    "Taken on January 15, 2024\n\n"
    "Q. Please state your name.\n"
    "A. John Smith.\n"
    "Q. Where do you live?\n"
    "A. Riverside, California.\n"
) * 50


def _stub_extract_with_dynamic_ocr(self, path):
    return SimpleNamespace(
        success=True,
        text=FAKE_TRANSCRIPT,
        char_count=len(FAKE_TRANSCRIPT),
        page_count=20,
        ocr_pages=[],
        ocr_percentage=0.0,
        error=None,
    )


@pytest.fixture
def isolated_sessions(tmp_path, monkeypatch):
    monkeypatch.setattr(session_manager, "SESSION_DIR", tmp_path / "sessions")
    return tmp_path


def test_phase1_writes_session_json_and_caches_text(isolated_sessions, capsys, monkeypatch):
    canned_topics = json.dumps([
        {"title": "Pre-Accident Medical History", "rank": 1, "discussion_density": "high"},
        {"title": "Mechanism Of Injury", "rank": 2, "discussion_density": "high"},
    ])

    monkeypatch.setattr(
        summarize_deposition.DocumentProcessor,
        "extract_with_dynamic_ocr",
        _stub_extract_with_dynamic_ocr,
    )

    with patch.object(summarize_deposition.LLMCaller, "call", return_value=canned_topics):
        input_path = str(isolated_sessions / "Smith Depo.pdf")
        Path(input_path).write_bytes(b"%PDF-1.4\n")  # presence only; extractor is stubbed
        logger = summarize_deposition.AgentLogger("DepositionTest", log_to_file=False)
        result = summarize_deposition.process_topics(input_path, logger)

    assert result is True
    out = capsys.readouterr().out
    awaiting_lines = [ln for ln in out.splitlines() if ln.startswith("AWAITING_INPUT:")]
    assert awaiting_lines, "AWAITING_INPUT token not printed"
    session_path = Path(awaiting_lines[-1][len("AWAITING_INPUT:"):])
    assert session_path.exists()

    data = json.loads(session_path.read_text(encoding="utf-8"))
    assert data["phase"] == "awaiting_input"
    assert data["user_config"] is None
    assert len(data["topics"]) == 2
    assert data["topics"][0]["title"] == "Pre-Accident Medical History"
    assert Path(data["cached_text_path"]).exists()
    assert Path(data["cached_text_path"]).read_text(encoding="utf-8") == FAKE_TRANSCRIPT


def test_phase1_handles_malformed_llm_json(isolated_sessions, capsys, monkeypatch):
    # Bulleted list instead of JSON
    malformed = "- Topic A\n- Topic B\n- Topic C\n"

    monkeypatch.setattr(
        summarize_deposition.DocumentProcessor,
        "extract_with_dynamic_ocr",
        _stub_extract_with_dynamic_ocr,
    )

    with patch.object(summarize_deposition.LLMCaller, "call", return_value=malformed):
        input_path = str(isolated_sessions / "X.pdf")
        Path(input_path).write_bytes(b"%PDF-1.4\n")
        logger = summarize_deposition.AgentLogger("DepositionTest", log_to_file=False)
        result = summarize_deposition.process_topics(input_path, logger)

    assert result is True
    out = capsys.readouterr().out
    awaiting_lines = [ln for ln in out.splitlines() if ln.startswith("AWAITING_INPUT:")]
    session_path = Path(awaiting_lines[-1][len("AWAITING_INPUT:"):])
    data = json.loads(session_path.read_text(encoding="utf-8"))
    titles = [t["title"] for t in data["topics"]]
    assert "Topic A" in titles
    assert "Topic B" in titles
    assert "Topic C" in titles


def test_phase1_caps_topic_count_at_25(isolated_sessions, capsys, monkeypatch):
    many = json.dumps([
        {"title": f"Topic {i}", "rank": i, "discussion_density": "medium"}
        for i in range(1, 51)
    ])

    monkeypatch.setattr(
        summarize_deposition.DocumentProcessor,
        "extract_with_dynamic_ocr",
        _stub_extract_with_dynamic_ocr,
    )

    with patch.object(summarize_deposition.LLMCaller, "call", return_value=many):
        input_path = str(isolated_sessions / "Y.pdf")
        Path(input_path).write_bytes(b"%PDF-1.4\n")
        logger = summarize_deposition.AgentLogger("DepositionTest", log_to_file=False)
        summarize_deposition.process_topics(input_path, logger)

    out = capsys.readouterr().out
    session_path = Path([ln for ln in out.splitlines() if ln.startswith("AWAITING_INPUT:")][-1][len("AWAITING_INPUT:"):])
    data = json.loads(session_path.read_text(encoding="utf-8"))
    assert len(data["topics"]) == 25
    # Truncation keeps the top-ranked topics
    assert data["topics"][0]["title"] == "Topic 1"
    assert data["topics"][-1]["title"] == "Topic 25"
```

- [ ] **Step 2: Run tests to verify they fail**

```
python -m pytest tests/test_deposition/test_summarize_deposition_phases.py -v
```

Expected: AttributeError or similar — `process_topics` does not exist yet on `summarize_deposition`.

- [ ] **Step 3: Add the topic-discovery prompt path constant and `process_topics`**

In `Scripts/summarize_deposition.py`:

a) Add a new constant alongside the existing prompt path constants (around line 59–62):

```python
TOPIC_DISCOVERY_PROMPT_FILE = os.path.join(SCRIPTS_DIR, "DEPOSITION_TOPIC_DISCOVERY_PROMPT.txt")
```

b) Add this `import` near the top of the file (after the other `icharlotte_core` imports):

```python
from icharlotte_core.deposition import session_manager
```

c) Add this helper function near the top of the file (after the `DeponentExtractor` class, before the existing `add_markdown_to_doc`):

```python
import json as _json  # local alias to avoid shadowing in helpers


def _parse_topic_response(response: str) -> list:
    """Parse the LLM topic-discovery response into a list of topic dicts.

    Falls back to bullet-list parsing if the response is not valid JSON.
    Always returns at most 25 topics, sorted by rank (ascending).
    """
    response = (response or "").strip()
    topics: list = []

    # Strip code fences if present
    if response.startswith("```"):
        lines = response.splitlines()
        # Drop the opening ``` line (and ```json variants) and the closing ``` if any.
        if lines and lines[0].startswith("```"):
            lines = lines[1:]
        if lines and lines[-1].startswith("```"):
            lines = lines[:-1]
        response = "\n".join(lines).strip()

    # Try JSON first
    try:
        parsed = _json.loads(response)
        if isinstance(parsed, list):
            for i, entry in enumerate(parsed, 1):
                if not isinstance(entry, dict):
                    continue
                title = str(entry.get("title", "")).strip()
                if not title:
                    continue
                rank = entry.get("rank", i)
                try:
                    rank = int(rank)
                except (TypeError, ValueError):
                    rank = i
                density = str(entry.get("discussion_density", "medium")).strip().lower()
                if density not in ("high", "medium", "low"):
                    density = "medium"
                topics.append({
                    "id": len(topics) + 1,
                    "title": title,
                    "rank": rank,
                    "discussion_density": density,
                })
    except _json.JSONDecodeError:
        pass

    # Fallback: bullet-list parsing
    if not topics:
        for line in response.splitlines():
            stripped = line.strip()
            if not stripped:
                continue
            for prefix in ("- ", "* ", "• "):
                if stripped.startswith(prefix):
                    stripped = stripped[len(prefix):].strip()
                    break
            # Drop leading numbering like "1." or "1)"
            stripped = re.sub(r"^\d+[\.\)]\s*", "", stripped)
            if not stripped:
                continue
            topics.append({
                "id": len(topics) + 1,
                "title": stripped,
                "rank": len(topics) + 1,
                "discussion_density": "medium",
            })

    # Sort by rank and truncate
    topics.sort(key=lambda t: t["rank"])
    topics = topics[:25]
    # Re-stamp ids and ranks to match final order
    for i, t in enumerate(topics, 1):
        t["id"] = i
        t["rank"] = i
    return topics


def process_topics(input_path: str, logger) -> bool:
    """Phase 1: extract text, discover topics, write session JSON, await input."""
    memory_monitor = MemoryMonitor(warn_threshold_mb=1500, abort_threshold_mb=2000, logger=logger.info)
    llm_caller = LLMCaller(logger=logger)

    logger.progress(2, "Initializing topic discovery...")

    if not os.path.exists(TOPIC_DISCOVERY_PROMPT_FILE):
        logger.error(f"Missing prompt file: {TOPIC_DISCOVERY_PROMPT_FILE}")
        return False
    with open(TOPIC_DISCOVERY_PROMPT_FILE, "r", encoding="utf-8") as f:
        topic_prompt = f.read()

    logger.progress(5, "Extracting transcript text...")
    try:
        with memory_monitor.track_operation("Text Extraction"):
            processor = DocumentProcessor(ocr_config=OCRConfig(adaptive=True), logger=logger)
            result = processor.extract_with_dynamic_ocr(input_path)
            if not result.success:
                raise ExtractionError(f"Failed to extract text: {result.error}", file_path=input_path)
            text = result.text
    except Exception as e:
        logger.pass_failed("Text Extraction", str(e), recoverable=False)
        return False

    logger.progress(20, f"Extracted {len(text)} chars; caching transcript")

    paths = session_manager.compute_session_paths(input_path)
    paths.cached_text_path.write_text(text, encoding="utf-8")

    deponent_name = DeponentExtractor.extract_deponent_name(text) or os.path.splitext(os.path.basename(input_path))[0]
    deposition_date = DeponentExtractor.extract_deposition_date(text)
    deponent_type = DeponentExtractor.detect_deponent_type(text, deponent_name)
    file_number = extract_file_number(input_path)

    logger.progress(40, "Calling LLM for topic discovery...")
    try:
        response = llm_caller.call(topic_prompt, text, task_type="summary")
    except Exception as e:
        logger.pass_failed("Topic Discovery", str(e), recoverable=False)
        return False

    topics = _parse_topic_response(response)
    if not topics:
        logger.error("Topic discovery returned no usable topics")
        return False

    session_data = {
        "version": 1,
        "phase": "awaiting_input",
        "input_path": str(input_path),
        "cached_text_path": str(paths.cached_text_path),
        "deponent_name": deponent_name,
        "deposition_date": deposition_date,
        "deponent_type": deponent_type,
        "file_number": file_number,
        "topics": topics,
        "user_config": None,
    }
    session_manager.write_session(paths.session_path, session_data)

    logger.progress(95, f"Discovered {len(topics)} topics; awaiting user input")
    print(f"AWAITING_INPUT:{paths.session_path}", flush=True)
    return True
```

- [ ] **Step 4: Run tests to verify they pass**

```
python -m pytest tests/test_deposition/test_summarize_deposition_phases.py -v
```

Expected: all 3 tests PASS.

- [ ] **Step 5: Commit**

```bash
git add Scripts/summarize_deposition.py tests/test_deposition/test_summarize_deposition_phases.py
git commit -m "feat(deposition): phase 1 process_topics with session sidecar"
```

---

## Task 6: Phase 2 — `process_summary` function

Phase 2 loads the session JSON, builds a topic-locked prompt with the user's choices, runs the LLM, optionally runs cross-check, saves the docx, and cleans up.

**Files:**
- Modify: `Scripts/summarize_deposition.py` (add `process_summary` and `_build_topic_locked_prompt`)
- Modify: `tests/test_deposition/test_summarize_deposition_phases.py` (append phase-2 tests)

- [ ] **Step 1: Add the failing phase-2 tests**

Append to `tests/test_deposition/test_summarize_deposition_phases.py`:

```python
# ---------------------------------------------------------------------------
# Phase 2
# ---------------------------------------------------------------------------

def _write_ready_session(tmp_path, *, cross_check, selected, added, bullets=5, label="Plaintiff", rules=""):
    session_path = tmp_path / "session.json"
    cached_path = tmp_path / "session.txt"
    cached_path.write_text(FAKE_TRANSCRIPT, encoding="utf-8")
    session_manager.write_session(session_path, {
        "version": 1,
        "phase": "ready_for_summary",
        "input_path": str(tmp_path / "Smith Depo.pdf"),
        "cached_text_path": str(cached_path),
        "deponent_name": "John Smith",
        "deposition_date": "January 15, 2024",
        "deponent_type": "Plaintiff",
        "file_number": "3850.084",
        "topics": [{"id": 1, "title": "Topic A", "rank": 1, "discussion_density": "high"}],
        "user_config": {
            "selected_topics": selected,
            "added_topics": added,
            "bullets_per_topic": bullets,
            "deponent_label": label,
            "custom_rules": rules,
            "cross_check_enabled": cross_check,
        },
    })
    return session_path


def test_phase2_reads_session_and_generates_summary(tmp_path, monkeypatch):
    session_path = _write_ready_session(
        tmp_path, cross_check=False,
        selected=["Topic A"], added=["Topic B"],
    )

    canned_summary = "**Topic A**\n- Bullet about A.\n\n**Topic B**\n- Bullet about B."
    called_with = {}

    def fake_save(content, output_path, deponent, date, logger):
        called_with["content"] = content
        called_with["output_path"] = output_path
        return True

    monkeypatch.setattr(summarize_deposition, "save_to_docx", fake_save)
    # Stub out registry / case-data writes — they're orthogonal to this test.
    monkeypatch.setattr(summarize_deposition, "_register_outputs", lambda *a, **kw: None, raising=False)

    with patch.object(summarize_deposition.LLMCaller, "call", return_value=canned_summary) as mock_call:
        logger = summarize_deposition.AgentLogger("DepositionTest", log_to_file=False)
        ok = summarize_deposition.process_summary(str(session_path), logger)

    assert ok is True
    assert mock_call.call_count == 1  # cross-check disabled
    assert called_with["content"] == canned_summary
    # Session + cached text cleaned up on success
    assert not session_path.exists()


def test_phase2_cross_check_runs_only_when_enabled(tmp_path, monkeypatch):
    session_path = _write_ready_session(
        tmp_path, cross_check=True,
        selected=["Topic A"], added=[],
    )

    monkeypatch.setattr(summarize_deposition, "save_to_docx", lambda *a, **kw: True)
    monkeypatch.setattr(summarize_deposition, "_register_outputs", lambda *a, **kw: None, raising=False)

    long_enough = "**Topic A**\n" + ("- Filler bullet.\n" * 20)

    with patch.object(summarize_deposition.LLMCaller, "call", return_value=long_enough) as mock_call:
        logger = summarize_deposition.AgentLogger("DepositionTest", log_to_file=False)
        summarize_deposition.process_summary(str(session_path), logger)

    assert mock_call.call_count == 2  # summary + cross-check


def test_phase2_prompt_includes_all_selected_plus_added_topics_in_order(tmp_path, monkeypatch):
    session_path = _write_ready_session(
        tmp_path, cross_check=False,
        selected=["Pre-Accident History", "Mechanism Of Injury"],
        added=["Communications With Providers"],
        bullets=7,
        label="Mr. Smith",
        rules="Use past tense.",
    )

    monkeypatch.setattr(summarize_deposition, "save_to_docx", lambda *a, **kw: True)
    monkeypatch.setattr(summarize_deposition, "_register_outputs", lambda *a, **kw: None, raising=False)

    captured = {}

    def fake_call(self, prompt, text, task_type=None, **kw):
        captured["prompt"] = prompt
        captured["text"] = text
        return "**Pre-Accident History**\n- B.\n\n**Mechanism Of Injury**\n- B.\n\n**Communications With Providers**\n- B."

    monkeypatch.setattr(summarize_deposition.LLMCaller, "call", fake_call)

    logger = summarize_deposition.AgentLogger("DepositionTest", log_to_file=False)
    summarize_deposition.process_summary(str(session_path), logger)

    prompt = captured["prompt"]
    # All three topics appear in order
    a = prompt.find("Pre-Accident History")
    b = prompt.find("Mechanism Of Injury")
    c = prompt.find("Communications With Providers")
    assert 0 < a < b < c
    assert "Mr. Smith" in prompt
    assert "7" in prompt  # bullets per topic
    assert "Use past tense." in prompt
```

- [ ] **Step 2: Run tests to verify they fail**

```
python -m pytest tests/test_deposition/test_summarize_deposition_phases.py -v -k phase2
```

Expected: AttributeError / missing `process_summary`.

- [ ] **Step 3: Add `process_summary` and helpers**

In `Scripts/summarize_deposition.py`, add after `process_topics`:

```python
def _build_topic_locked_prompt(base_prompt: str, *, topic_list: list, bullets_per_topic: int,
                                deponent_label: str, custom_rules: str) -> str:
    """Render the topic-locked summary prompt with user-supplied substitutions."""
    rendered_topics = "\n".join(f"- {t}" for t in topic_list)
    return (base_prompt
            .replace("{deponent_label}", deponent_label)
            .replace("{bullets_per_topic}", str(bullets_per_topic))
            .replace("{topic_list}", rendered_topics)
            .replace("{custom_rules}", custom_rules or "(none)"))


def _register_outputs(input_path, summary, deponent_name, deponent_type, output_file, logger):
    """Wrapper around CaseDataManager + DocumentRegistry writes. Best-effort; logs but does not fail the run."""
    try:
        data_manager = CaseDataManager()
        file_num = extract_file_number(input_path)
        if file_num:
            clean_name = re.sub(r"[^a-zA-Z0-9_]", "_", (deponent_name or "unknown").lower())
            var_key = f"depo_summary_{clean_name}"
            data_manager.save_variable(
                file_num, var_key, summary,
                source="deposition_agent",
                extra_tags=["Deposition", deponent_type] if deponent_type else ["Deposition"],
            )
    except Exception as e:
        logger.warning(f"Could not save to case data: {e}")

    try:
        file_num = extract_file_number(input_path)
        if not file_num:
            return
        depo_type_map = {
            "Plaintiff": "Deposition - Plaintiff",
            "Defendant": "Deposition - Defendant",
            "Witness": "Deposition - Witness",
            "Expert Witness": "Deposition - Expert",
            "Expert/Physician": "Deposition - Expert",
            "Treating Physician": "Deposition - Expert",
            "Corporate Representative": "Deposition - Corporate Representative",
        }
        registry_doc_type = depo_type_map.get(deponent_type, "Deposition - Witness")
        from document_registry import DocumentClassifier
        classifier = DocumentClassifier(logger=logger)
        base_name = os.path.splitext(os.path.basename(input_path))[0]
        standardized_name = classifier.generate_name(summary, registry_doc_type, fallback_name=base_name)
        registry = DocumentRegistry()
        registry.register_document(
            file_number=file_num,
            name=standardized_name,
            document_type=registry_doc_type,
            source_path=input_path,
            summary_location=output_file,
            agent="summarize_deposition",
            char_count=len(summary),
        )
    except Exception as e:
        logger.warning(f"Could not register document: {e}")


def process_summary(session_path: str, logger) -> bool:
    """Phase 2: read session, generate topic-locked summary, save docx, cleanup."""
    from pathlib import Path
    memory_monitor = MemoryMonitor(warn_threshold_mb=1500, abort_threshold_mb=2000, logger=logger.info)
    llm_caller = LLMCaller(logger=logger)

    session_path = Path(session_path)
    logger.progress(5, "Loading session...")
    try:
        session = session_manager.read_session(session_path)
    except Exception as e:
        logger.error(f"Failed to load session: {e}")
        return False

    if session.get("phase") != "ready_for_summary" or not session.get("user_config"):
        logger.pass_failed("Phase 2 Init", "Session is not ready_for_summary", recoverable=False)
        return False

    cached_path = Path(session["cached_text_path"])
    if not cached_path.exists():
        logger.pass_failed("Phase 2 Init", f"Cached transcript missing: {cached_path}", recoverable=False)
        return False

    logger.progress(10, "Reading cached transcript...")
    text = cached_path.read_text(encoding="utf-8")

    cfg = session["user_config"]
    final_topics = [t for t in cfg.get("selected_topics", [])] + [t for t in cfg.get("added_topics", [])]
    if not final_topics:
        logger.pass_failed("Phase 2 Init", "No topics selected", recoverable=False)
        return False

    with open(NARRATIVE_PROMPT_FILE, "r", encoding="utf-8") as f:
        base_prompt = f.read()

    prompt = _build_topic_locked_prompt(
        base_prompt,
        topic_list=final_topics,
        bullets_per_topic=cfg.get("bullets_per_topic", 5),
        deponent_label=cfg.get("deponent_label") or session.get("deponent_type", "Deponent"),
        custom_rules=cfg.get("custom_rules", ""),
    )

    logger.progress(30, "Generating summary...")
    try:
        with memory_monitor.track_operation("Narrative Summary"):
            summary = llm_caller.call(prompt, text, task_type="summary")
        if not summary:
            raise SummaryPassError("LLM returned empty summary")
    except Exception as e:
        logger.pass_failed("Narrative Summary", str(e), recoverable=False)
        return False
    logger.progress(70, f"Summary generated: {len(summary)} chars")

    if cfg.get("cross_check_enabled"):
        try:
            with open(CROSS_CHECK_PROMPT_FILE, "r", encoding="utf-8") as f:
                cross_prompt = f.read()
            cross_prompt = cross_prompt.replace("{summary}", summary).replace("{original}", text[:75000])
            logger.progress(75, "Running cross-check...")
            verified = llm_caller.call(cross_prompt, "", task_type="cross_check")
            if verified and len(verified) > len(summary) * 0.8:
                summary = verified
        except Exception as e:
            logger.warning(f"Cross-check failed; using original summary: {e}")
    logger.progress(85, "Cross-check complete")

    output_dir = get_output_directory(session["input_path"])
    os.makedirs(output_dir, exist_ok=True)
    output_file = os.path.join(output_dir, "Deposition_summaries.docx")
    logger.progress(90, f"Saving to {os.path.basename(output_file)}...")
    if not save_to_docx(summary, output_file, session["deponent_name"], session["deposition_date"], logger):
        return False
    logger.progress(95, "Document saved")

    _register_outputs(session["input_path"], summary, session["deponent_name"],
                       session.get("deponent_type"), output_file, logger)

    session_manager.cleanup_session(session_path)
    logger.progress(100, "Deposition summarization complete")
    return True
```

- [ ] **Step 4: Run tests to verify they pass**

```
python -m pytest tests/test_deposition/test_summarize_deposition_phases.py -v
```

Expected: all 6 tests PASS (3 phase-1, 3 phase-2).

- [ ] **Step 5: Commit**

```bash
git add Scripts/summarize_deposition.py tests/test_deposition/test_summarize_deposition_phases.py
git commit -m "feat(deposition): phase 2 process_summary with topic-locked prompt"
```

---

## Task 7: `--phase` dispatcher in `main()`

`main()` currently calls `process_document` directly. Replace its dispatch with a `--phase=topics` / `--phase=summary` switch and keep backward compatibility for callers that don't pass `--phase` (default to `topics`, since that's what the UI will trigger for a file selection).

**Files:**
- Modify: `Scripts/summarize_deposition.py` (rewrite `main()`)

- [ ] **Step 1: Replace `main()`**

Replace the existing `main()` function (~line 1135) with:

```python
def main():
    """Main entry point. Accepts --phase=topics <input> or --phase=summary <session_path>."""
    if len(sys.argv) < 2:
        print("Error: No file path or session path provided.", flush=True)
        sys.exit(1)

    args = sys.argv[1:]

    # Parse --phase argument (default: topics)
    phase = "topics"
    positional = []
    for a in args:
        if a.startswith("--phase="):
            phase = a.split("=", 1)[1].strip().lower()
        else:
            positional.append(a)

    if not positional:
        print("Error: missing input path.", flush=True)
        sys.exit(1)

    # Handle quoted paths with spaces (preserve existing behavior)
    combined = " ".join(positional).strip().strip('"').strip("'")
    if os.path.exists(combined):
        target = combined
    else:
        target = positional[0]
    target = os.path.abspath(target.strip().strip('"').strip("'"))

    if not os.path.exists(target):
        print(f"Error: path not found: {target}", flush=True)
        sys.exit(1)

    file_number = extract_file_number(target) if phase == "topics" else None
    logger = AgentLogger("Deposition", file_number=file_number)

    if phase == "topics":
        # Directory dispatcher mode is only meaningful for phase 1.
        if os.path.isdir(target):
            logger.info("Input is a directory. Spawning per-file phase-1 agents...")
            for root, _, files in os.walk(target):
                for f in files:
                    if not f.lower().endswith((".pdf", ".docx")):
                        continue
                    if "Deposition_summaries" in f:
                        continue
                    p = os.path.join(root, f)
                    creationflags = 0x08000000 if os.name == "nt" else 0
                    subprocess.Popen([sys.executable, sys.argv[0], "--phase=topics", p],
                                     creationflags=creationflags if os.name == "nt" else 0)
            sys.exit(0)

        ok = process_topics(target, logger)
        sys.exit(0 if ok else 1)

    if phase == "summary":
        ok = process_summary(target, logger)
        sys.exit(0 if ok else 1)

    print(f"Error: unknown --phase value: {phase}", flush=True)
    sys.exit(2)
```

- [ ] **Step 2: Manual smoke test**

Run a quick sanity check (no LLM call — should fail fast on a missing prompt or session, but should not blow up at argument parsing):

```
python Scripts/summarize_deposition.py --phase=summary C:\tmp\nonexistent.json
```

Expected: non-zero exit with a clear "Failed to load session" log line, no traceback in stdout.

- [ ] **Step 3: Commit**

```bash
git add Scripts/summarize_deposition.py
git commit -m "feat(deposition): --phase dispatcher in summarize_deposition main()"
```

---

## Task 8: Remove dead code

Delete `ExhibitExtractor`, `ImpeachmentDetector`, the parallel extraction-and-summary pass, `build_narrative_prompt`'s topic-count substitution logic, the legacy `process_document` function, and unused prompt path constants.

**Files:**
- Modify: `Scripts/summarize_deposition.py`

- [ ] **Step 1: Delete the unused classes and helpers**

Remove from `Scripts/summarize_deposition.py`:

a) The entire `class ExhibitExtractor:` block (currently at lines ~281–348).

b) The entire `class ImpeachmentDetector:` block (currently at lines ~354–461).

c) The entire `build_narrative_prompt` function (currently at lines ~681–717).

d) The entire `process_document` function (currently at lines ~720–1104). It is replaced by `process_topics` + `process_summary`.

e) The `process_directory` function (currently at lines ~1107–1132). Folder mode is now handled inside `main()` for `--phase=topics`.

f) Remove the `EXTRACTION_PROMPT_FILE` constant — phase 2 does not use it.

g) Remove the unused imports: `from concurrent.futures import ThreadPoolExecutor, as_completed` (no parallelism in the new flow), and the `from icharlotte_core.exceptions import` line keeps only `LLMError`, `ExtractionError`, `SummaryPassError`, `MemoryLimitError` (drop `PassFailedError`, `CrossCheckPassError` if no longer referenced — verify with grep).

- [ ] **Step 2: Verify the script still imports cleanly**

```
python -c "import sys; sys.path.insert(0, 'Scripts'); import summarize_deposition; print('OK')"
```

Expected: `OK`.

- [ ] **Step 3: Run the full test file again**

```
python -m pytest tests/test_deposition/test_summarize_deposition_phases.py -v
```

Expected: all 6 tests still PASS (deletions should not regress anything).

- [ ] **Step 4: Commit**

```bash
git add Scripts/summarize_deposition.py
git commit -m "refactor(deposition): remove exhibits, impeachment, parallel extraction pass"
```

---

## Task 9: AgentRunner — `AWAITING_INPUT:` parsing + paused-exit handling

`AgentRunner` learns to recognize the new stdout token, expose an `awaiting_input` signal, and treat the subsequent exit-0 as "paused" (do not emit `finished`).

**Files:**
- Modify: `icharlotte_core/ui/widgets.py` (extend `AgentRunner`)
- Create: `tests/test_deposition/test_agent_runner_awaiting_input.py`

- [ ] **Step 1: Write the failing tests**

Create `tests/test_deposition/test_agent_runner_awaiting_input.py`:

```python
"""Tests for AgentRunner's AWAITING_INPUT handling."""

from unittest.mock import MagicMock

import pytest

pytest.importorskip("pytest_qt")  # PySide6 + pytest-qt required

from icharlotte_core.ui.widgets import AgentRunner


def test_agent_runner_emits_awaiting_input_signal_on_token(qtbot):
    runner = AgentRunner("python", ["--phase=topics", "X.pdf"])
    received = []
    runner.awaiting_input.connect(received.append)

    runner.parse_progress("AWAITING_INPUT:C:\\tmp\\session.json\n")

    assert received == ["C:\\tmp\\session.json"]
    assert runner.session_path == "C:\\tmp\\session.json"


def test_agent_runner_does_not_emit_finished_when_paused(qtbot):
    runner = AgentRunner("python", ["--phase=topics", "X.pdf"])
    finished_calls = []
    awaiting_calls = []
    runner.finished.connect(finished_calls.append)
    runner.awaiting_input.connect(awaiting_calls.append)

    # Simulate phase-1 token then a normal exit(0).
    runner.parse_progress("AWAITING_INPUT:C:\\tmp\\session.json\n")
    runner.process = MagicMock()
    runner.process.deleteLater = MagicMock()
    from PySide6.QtCore import QProcess
    runner.handle_finished(0, QProcess.ExitStatus.NormalExit)

    assert awaiting_calls == ["C:\\tmp\\session.json"]
    assert finished_calls == []
    assert runner.success is None  # still running from the UI's perspective


def test_agent_runner_emits_finished_when_no_pause(qtbot):
    """Sanity check: without an AWAITING_INPUT token, exit 0 still emits finished(True)."""
    runner = AgentRunner("python", ["X.pdf"])
    finished_calls = []
    runner.finished.connect(finished_calls.append)

    runner.process = MagicMock()
    runner.process.deleteLater = MagicMock()
    from PySide6.QtCore import QProcess
    runner.handle_finished(0, QProcess.ExitStatus.NormalExit)

    assert finished_calls == [True]
```

- [ ] **Step 2: Run tests to verify they fail**

```
python -m pytest tests/test_deposition/test_agent_runner_awaiting_input.py -v
```

Expected: AttributeError on `awaiting_input` / `session_path`.

- [ ] **Step 3: Extend `AgentRunner`**

In `icharlotte_core/ui/widgets.py`:

a) Add `awaiting_input` to the `Signal` declarations at the top of `AgentRunner`:

```python
class AgentRunner(QObject):
    # Signals to update UI safely
    progress_update = Signal(int, str)
    log_update = Signal(str)
    output_file_found = Signal(str)
    finished = Signal(bool)
    awaiting_input = Signal(str)  # session_path — phase-1 paused, ready for user input
    ...
```

b) Add initialization in `__init__` (after the existing `self.output_file = None`):

```python
        self.session_path = None
```

c) Add a parsing branch in `parse_progress`, placed immediately after the existing `OUTPUT_FILE:` branch (around line 829):

```python
            # Format: AWAITING_INPUT:<session_path>
            if line.startswith("AWAITING_INPUT:"):
                path = line[len("AWAITING_INPUT:"):].strip()
                if path:
                    self.session_path = path
                    self.awaiting_input.emit(path)
                continue
```

d) Modify `handle_finished` to treat exit-0 as "paused" when `session_path` is set and `success` is still `None`. Find the `success = (exit_code == 0 and exit_status == QProcess.ExitStatus.NormalExit)` line and replace the surrounding block (down through `self.finished.emit(success)`) with:

```python
            success = (exit_code == 0 and exit_status == QProcess.ExitStatus.NormalExit)

            # Phase-1 paused-exit: exit 0 after AWAITING_INPUT means "paused for user input",
            # not "done". Keep the widget alive; do NOT emit finished.
            if success and self.session_path is not None and self.success is None:
                self.watchdog_timer.stop()
                self.process.deleteLater()
                # success stays None — UI treats this as still-running until phase 2 resolves.
                return

            self.success = success

            # ...existing logging block stays as-is...
            self.finished.emit(success)
            self.process.deleteLater()
```

(Preserve the existing logging block between `self.success = success` and `self.finished.emit(success)` — only the early-return for the paused case is new.)

- [ ] **Step 4: Run tests to verify they pass**

```
python -m pytest tests/test_deposition/test_agent_runner_awaiting_input.py -v
```

Expected: all 3 tests PASS. If `pytest_qt` is not installed:

```
pip install pytest-qt
```

then re-run.

- [ ] **Step 5: Commit**

```bash
git add icharlotte_core/ui/widgets.py tests/test_deposition/test_agent_runner_awaiting_input.py
git commit -m "feat(ui): AgentRunner AWAITING_INPUT signal and paused-exit handling"
```

---

## Task 10: AgentRunner — `resume_with_config` method

After the user submits the popup, the parent tab calls `agent_runner.resume_with_config(session_path)`. The runner instantiates a fresh `QProcess` against the same widget and runs phase 2.

**Files:**
- Modify: `icharlotte_core/ui/widgets.py` (add method on `AgentRunner`)
- Modify: `tests/test_deposition/test_agent_runner_awaiting_input.py` (add resume test)

- [ ] **Step 1: Add the failing test**

Append to `tests/test_deposition/test_agent_runner_awaiting_input.py`:

```python
def test_resume_with_config_starts_phase_two_process(qtbot, monkeypatch):
    runner = AgentRunner("python", ["--phase=topics", "X.pdf"])
    runner.session_path = r"C:\tmp\session.json"

    started_with = {}

    class FakeProcess:
        def __init__(self):
            self.started = False
        def start(self, cmd, args):
            started_with["cmd"] = cmd
            started_with["args"] = args
            self.started = True
        def readyReadStandardOutput(self):  # signal stub
            pass
        def readyReadStandardError(self):
            pass
        def finished(self):
            pass
        def state(self):
            return 0
        def kill(self):
            pass
        def deleteLater(self):
            pass
        # Allow .connect on the signal stubs
        def __getattr__(self, name):
            return MagicMock()

    monkeypatch.setattr("icharlotte_core.ui.widgets.QProcess", FakeProcess)

    runner.resume_with_config(r"C:\tmp\session.json")

    assert started_with["cmd"] == "python"
    assert "--phase=summary" in started_with["args"]
    assert r"C:\tmp\session.json" in started_with["args"]
```

- [ ] **Step 2: Run to verify it fails**

```
python -m pytest tests/test_deposition/test_agent_runner_awaiting_input.py::test_resume_with_config_starts_phase_two_process -v
```

Expected: AttributeError — `resume_with_config` does not exist.

- [ ] **Step 3: Implement `resume_with_config`**

In `icharlotte_core/ui/widgets.py`, inside `AgentRunner`, add this method (place it just after `retry_pass`, before `handle_stdout`):

```python
    def resume_with_config(self, session_path: str):
        """Launch phase 2 against the same widget. Phase 1's QProcess has already exited."""
        import time
        # Build phase-2 args: keep the script path (args[0]) and any other non-input args,
        # then append --phase=summary <session_path>.
        script = self.args[0] if self.args else ""
        phase2_args = [script, "--phase=summary", session_path]

        # Fresh QProcess instance for phase 2
        self.process = QProcess()
        self.process.readyReadStandardOutput.connect(self.handle_stdout)
        self.process.readyReadStandardError.connect(self.handle_stderr)
        self.process.finished.connect(self.handle_finished)

        # Allow handle_finished to emit finished(True) on phase-2 completion
        self.session_path = None
        self.last_output_time = time.time()

        log_info(f"AgentRunner resuming phase 2: {self.command} {' '.join(phase2_args)}")
        self.process.start(self.command, phase2_args)
        self.watchdog_timer.start()
```

- [ ] **Step 4: Run all the AgentRunner tests**

```
python -m pytest tests/test_deposition/test_agent_runner_awaiting_input.py -v
```

Expected: 4 tests PASS.

- [ ] **Step 5: Commit**

```bash
git add icharlotte_core/ui/widgets.py tests/test_deposition/test_agent_runner_awaiting_input.py
git commit -m "feat(ui): AgentRunner resume_with_config to launch phase 2"
```

---

## Task 11: Status widget — READY button + `ready_clicked` signal

The status widget surfaces the READY button when phase 1 pauses. Clicking emits `ready_clicked(session_path)` for the parent tab to consume.

**Files:**
- Modify: `icharlotte_core/ui/widgets.py` — the existing `StatusWidget` class at line 213. (It owns `update_progress`, `append_log`, `set_finished`, etc. and emits `cancel_requested`.)

- [ ] **Step 1: Add the READY button signal, button, and slot to `StatusWidget`**

a) Add a new signal in the `StatusWidget` class header, alongside `cancel_requested` (line 214):

```python
class StatusWidget(QFrame):
    cancel_requested = Signal()
    retry_requested = Signal()
    retry_pass_requested = Signal(str)
    ready_clicked = Signal(str)  # session_path
```

b) In `StatusWidget.__init__`, immediately after the existing `self.cancel_btn` block (which ends with `header_layout.addWidget(self.cancel_btn)` around line 264), insert:

```python
        self.ready_btn = QPushButton("READY")
        self.ready_btn.setFixedSize(80, 25)
        self.ready_btn.setStyleSheet(
            "background-color: #2e7d32; color: white; font-weight: bold;"
        )
        self.ready_btn.setToolTip("Click to choose topics and configure summary.")
        self.ready_btn.setVisible(False)
        self.ready_btn.clicked.connect(self._on_ready_button_clicked)
        header_layout.addWidget(self.ready_btn)

        self._pending_session_path = None
```

c) Add the slot and click handler as methods of `StatusWidget` (place them after `update_progress`):

```python
    def on_awaiting_input(self, session_path: str):
        self._pending_session_path = session_path
        self.ready_btn.setVisible(True)

    def _on_ready_button_clicked(self):
        if self._pending_session_path:
            self.ready_clicked.emit(self._pending_session_path)

    def clear_ready_state(self):
        """Called when phase 2 starts (or the run is cancelled) — hide the button."""
        self._pending_session_path = None
        self.ready_btn.setVisible(False)
```

- [ ] **Step 2: Wire `AgentRunner.awaiting_input` → widget slot in `connect_widget`**

In `AgentRunner.connect_widget` (around line 506), add the connection alongside the existing ones:

```python
        self.awaiting_input.connect(self.status_widget.on_awaiting_input)
```

And in `disconnect_widget`, add the corresponding `safe_disconnect`:

```python
        safe_disconnect(self.awaiting_input, self.status_widget.on_awaiting_input)
```

- [ ] **Step 3: Replay logic in `reconnect_widget`**

In `AgentRunner.reconnect_widget` (after the existing replay block), add:

```python
        if self.session_path is not None:
            widget.on_awaiting_input(self.session_path)
```

- [ ] **Step 4: Verify the existing test suite still passes**

```
python -m pytest tests/test_deposition/ -v
```

Expected: all earlier tests still PASS. The new widget behavior is exercised by the manual flow in Task 14 and the smoke test in Task 14b.

- [ ] **Step 5: Commit**

```bash
git add icharlotte_core/ui/widgets.py
git commit -m "feat(ui): status widget READY button and awaiting-input slot"
```

---

## Task 12: Popup dialog — `DepoSummaryConfigDialog`

The modal popup that lets the user pick topics, set bullet count, set deponent label, set custom rules, and toggle cross-check.

**Files:**
- Create: `icharlotte_core/ui/depo_summary_config_dialog.py`
- Create: `tests/test_deposition/test_depo_summary_config_dialog.py`

- [ ] **Step 1: Write the failing tests**

Create `tests/test_deposition/test_depo_summary_config_dialog.py`:

```python
"""Tests for the DepoSummaryConfigDialog popup."""

import json
import os
from pathlib import Path
from unittest.mock import patch

import pytest

pytest.importorskip("pytest_qt")

from icharlotte_core.deposition import session_manager
from icharlotte_core.ui.depo_summary_config_dialog import DepoSummaryConfigDialog


def _make_session(tmp_path) -> Path:
    session_path = tmp_path / "session.json"
    cached = tmp_path / "session.txt"
    cached.write_text("fake transcript", encoding="utf-8")
    session_manager.write_session(session_path, {
        "version": 1,
        "phase": "awaiting_input",
        "input_path": str(tmp_path / "Smith.pdf"),
        "cached_text_path": str(cached),
        "deponent_name": "John Smith",
        "deposition_date": "January 15, 2024",
        "deponent_type": "Plaintiff",
        "file_number": "3850.084",
        "topics": [
            {"id": 1, "title": "Pre-Accident History", "rank": 1, "discussion_density": "high"},
            {"id": 2, "title": "Mechanism Of Injury", "rank": 2, "discussion_density": "high"},
            {"id": 3, "title": "Damages", "rank": 3, "discussion_density": "medium"},
        ],
        "user_config": None,
    })
    return session_path


def test_dialog_loads_session_and_populates_topics(qtbot, tmp_path):
    session_path = _make_session(tmp_path)
    dlg = DepoSummaryConfigDialog(session_path)
    qtbot.addWidget(dlg)

    titles = [row.title_edit.text() for row in dlg.topic_rows]
    assert titles == ["Pre-Accident History", "Mechanism Of Injury", "Damages"]
    assert all(row.checkbox.isChecked() for row in dlg.topic_rows)
    assert dlg.deponent_label_edit.text() == "Plaintiff"
    assert dlg.bullets_spinbox.value() == 5
    assert dlg.cross_check_checkbox.isChecked()


def test_dialog_accept_writes_user_config_back_to_session(qtbot, tmp_path):
    session_path = _make_session(tmp_path)
    dlg = DepoSummaryConfigDialog(session_path)
    qtbot.addWidget(dlg)

    # Uncheck topic 2, rename topic 1, add a custom topic, change settings.
    dlg.topic_rows[1].checkbox.setChecked(False)
    dlg.topic_rows[0].title_edit.setText("Pre-Accident Lower Back Treatment")
    dlg.added_topics_edit.setPlainText("Communications With Treating Providers\n")
    dlg.bullets_spinbox.setValue(7)
    dlg.deponent_label_edit.setText("Mr. Smith")
    dlg.custom_rules_edit.setPlainText("Use past tense.")
    dlg.cross_check_checkbox.setChecked(False)

    dlg.accept()

    loaded = session_manager.read_session(session_path)
    assert loaded["phase"] == "ready_for_summary"
    cfg = loaded["user_config"]
    assert cfg["selected_topics"] == ["Pre-Accident Lower Back Treatment", "Damages"]
    assert cfg["added_topics"] == ["Communications With Treating Providers"]
    assert cfg["bullets_per_topic"] == 7
    assert cfg["deponent_label"] == "Mr. Smith"
    assert cfg["custom_rules"] == "Use past tense."
    assert cfg["cross_check_enabled"] is False


def test_dialog_cancel_does_not_modify_session(qtbot, tmp_path):
    session_path = _make_session(tmp_path)
    before = session_path.read_text(encoding="utf-8")

    dlg = DepoSummaryConfigDialog(session_path)
    qtbot.addWidget(dlg)
    dlg.topic_rows[0].checkbox.setChecked(False)
    dlg.bullets_spinbox.setValue(99)
    dlg.reject()

    after = session_path.read_text(encoding="utf-8")
    assert before == after


def test_dialog_atomic_write_preserves_original_on_failure(qtbot, tmp_path, monkeypatch):
    session_path = _make_session(tmp_path)
    before = session_path.read_text(encoding="utf-8")

    dlg = DepoSummaryConfigDialog(session_path)
    qtbot.addWidget(dlg)
    dlg.bullets_spinbox.setValue(9)

    def boom(*a, **kw):
        raise OSError("simulated")

    monkeypatch.setattr(os, "replace", boom)
    with pytest.raises(OSError):
        dlg.accept()

    # Session file untouched
    after = session_path.read_text(encoding="utf-8")
    assert before == after
```

- [ ] **Step 2: Run to verify they fail**

```
python -m pytest tests/test_deposition/test_depo_summary_config_dialog.py -v
```

Expected: ImportError on `icharlotte_core.ui.depo_summary_config_dialog`.

- [ ] **Step 3: Implement the dialog**

Create `icharlotte_core/ui/depo_summary_config_dialog.py`:

```python
"""Modal dialog the user fills out after phase 1 of the deposition agent."""

from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QCheckBox, QDialog, QDialogButtonBox, QHBoxLayout, QLabel, QLineEdit,
    QPlainTextEdit, QScrollArea, QSpinBox, QVBoxLayout, QWidget,
)

from icharlotte_core.deposition import session_manager


class _TopicRow(QWidget):
    def __init__(self, title: str, parent=None):
        super().__init__(parent)
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        self.checkbox = QCheckBox()
        self.checkbox.setChecked(True)
        self.title_edit = QLineEdit(title)
        layout.addWidget(self.checkbox)
        layout.addWidget(self.title_edit, 1)


class DepoSummaryConfigDialog(QDialog):
    def __init__(self, session_path, parent=None):
        super().__init__(parent)
        self.session_path = Path(session_path)
        self._session = session_manager.read_session(self.session_path)

        self.setWindowTitle("Configure Deposition Summary")
        self.setModal(True)
        self.setWindowModality(Qt.ApplicationModal)
        self.resize(700, 600)

        root = QVBoxLayout(self)

        header_text = (
            f"Configure summary for <b>{self._session.get('deponent_name', '')}</b> "
            f"({self._session.get('deponent_type', '')}, "
            f"{self._session.get('deposition_date', 'date unknown')})"
        )
        root.addWidget(QLabel(header_text))

        # Topic rows in a scroll area
        root.addWidget(QLabel("Topics (uncheck to omit, edit text to rename):"))
        topics_container = QWidget()
        topics_layout = QVBoxLayout(topics_container)
        topics_layout.setContentsMargins(4, 4, 4, 4)
        self.topic_rows = []
        for t in self._session.get("topics", []):
            row = _TopicRow(t.get("title", ""))
            self.topic_rows.append(row)
            topics_layout.addWidget(row)
        topics_layout.addStretch(1)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setWidget(topics_container)
        root.addWidget(scroll, 1)

        # Additional topics
        root.addWidget(QLabel("Additional topics (one per line):"))
        self.added_topics_edit = QPlainTextEdit()
        self.added_topics_edit.setPlaceholderText(
            "One topic per line. These are appended after the checked topics above, in order."
        )
        self.added_topics_edit.setFixedHeight(70)
        root.addWidget(self.added_topics_edit)

        # Settings row
        settings_row = QHBoxLayout()
        settings_row.addWidget(QLabel("Bullets per topic:"))
        self.bullets_spinbox = QSpinBox()
        self.bullets_spinbox.setRange(1, 15)
        self.bullets_spinbox.setValue(5)
        settings_row.addWidget(self.bullets_spinbox)

        settings_row.addSpacing(20)
        settings_row.addWidget(QLabel("Deponent label:"))
        self.deponent_label_edit = QLineEdit(self._session.get("deponent_type", ""))
        settings_row.addWidget(self.deponent_label_edit, 1)

        settings_row.addSpacing(20)
        self.cross_check_checkbox = QCheckBox("Run cross-check pass")
        self.cross_check_checkbox.setChecked(True)
        settings_row.addWidget(self.cross_check_checkbox)
        root.addLayout(settings_row)

        # Custom rules
        root.addWidget(QLabel("Custom rules:"))
        self.custom_rules_edit = QPlainTextEdit()
        self.custom_rules_edit.setPlaceholderText(
            "Any extra instructions for the summary (tense, citation style, things to avoid, etc.)."
        )
        self.custom_rules_edit.setFixedHeight(90)
        root.addWidget(self.custom_rules_edit)

        # Buttons
        buttons = QDialogButtonBox()
        buttons.addButton("Cancel", QDialogButtonBox.RejectRole)
        generate_btn = buttons.addButton("Generate Summary", QDialogButtonBox.AcceptRole)
        generate_btn.setDefault(True)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        root.addWidget(buttons)

    def accept(self):
        selected_topics = [
            row.title_edit.text().strip()
            for row in self.topic_rows
            if row.checkbox.isChecked() and row.title_edit.text().strip()
        ]
        added_topics = [
            line.strip()
            for line in self.added_topics_edit.toPlainText().splitlines()
            if line.strip()
        ]
        cfg = {
            "selected_topics": selected_topics,
            "added_topics": added_topics,
            "bullets_per_topic": self.bullets_spinbox.value(),
            "deponent_label": self.deponent_label_edit.text().strip() or "Deponent",
            "custom_rules": self.custom_rules_edit.toPlainText().strip(),
            "cross_check_enabled": self.cross_check_checkbox.isChecked(),
        }
        session_manager.update_user_config(self.session_path, cfg)
        super().accept()
```

- [ ] **Step 4: Run tests to verify they pass**

```
python -m pytest tests/test_deposition/test_depo_summary_config_dialog.py -v
```

Expected: all 4 tests PASS.

- [ ] **Step 5: Commit**

```bash
git add icharlotte_core/ui/depo_summary_config_dialog.py tests/test_deposition/test_depo_summary_config_dialog.py
git commit -m "feat(ui): DepoSummaryConfigDialog popup for interactive depo flow"
```

---

## Task 13: IndexTab wiring — READY click → dialog → resume

The parent tab (`IndexTab` in `icharlotte_core/ui/tabs.py`) is where `AgentRunner` instances are constructed for the file/folder agent runs. Connect the status widget's `ready_clicked` signal to a handler that opens the dialog and, on acceptance, calls `agent_runner.resume_with_config`.

**Files:**
- Modify: `icharlotte_core/ui/tabs.py` (find where the deposition agent is spawned; wire the `ready_clicked` → handler)

- [ ] **Step 1: Locate the deposition agent spawn site**

Run:

```
python -c "import re,sys; t=open(r'C:\geminiterminal2\icharlotte_core\ui\tabs.py',encoding='utf-8').read(); [print(i+1,':',l) for i,l in enumerate(t.splitlines()) if 'summarize_deposition' in l.lower() or 'AgentRunner(' in l]"
```

You should see where an `AgentRunner` is constructed for the deposition agent, and where its status widget is added to the tab. (If there is no special-cased deposition spawn — i.e., all agents go through the same code path — then the wiring goes in that shared path with a guard on script name.)

- [ ] **Step 2: Add the wiring**

In `icharlotte_core/ui/tabs.py`, in the function/method that constructs the `AgentRunner` for the deposition agent (or in the shared agent-spawn path, guarded by `if script_name.endswith("summarize_deposition.py")`), after the status widget is connected to the runner, add:

```python
        # Wire the READY button (only meaningful for the deposition agent).
        status_widget.ready_clicked.connect(
            lambda session_path, runner=agent_runner, widget=status_widget:
                self._open_depo_summary_dialog(session_path, runner, widget)
        )
```

Then add a new method on the tab class:

```python
    def _open_depo_summary_dialog(self, session_path: str, agent_runner, status_widget):
        from icharlotte_core.ui.depo_summary_config_dialog import DepoSummaryConfigDialog
        dlg = DepoSummaryConfigDialog(session_path, parent=self)
        if dlg.exec() == dlg.Accepted:
            status_widget.clear_ready_state()
            agent_runner.resume_with_config(session_path)
        # On Cancel: leave READY button visible so user can re-open.
```

- [ ] **Step 3: Manual UI smoke test**

```
python iCharlotte.py
```

(or however the app is launched in this environment).

Steps:
1. Open a small deposition PDF from a test case folder.
2. Run the Summarize Deposition agent.
3. Wait for phase 1 to finish — confirm the READY button appears on the status row.
4. Click READY — confirm the popup opens with topics, deponent label pre-filled, default bullets=5.
5. Uncheck one topic, add a custom topic, click Generate Summary.
6. Confirm phase 2 progress bar moves and an output `.docx` is saved.
7. Open the `.docx` and confirm it contains only the topics you selected/added, with the bullet count you set.

If any step fails, debug and fix before committing.

- [ ] **Step 4: Commit**

```bash
git add icharlotte_core/ui/tabs.py
git commit -m "feat(ui): wire READY button to summary config dialog and phase 2 resume"
```

---

## Task 14: Integration smoke test

End-to-end test that spawns phase 1 and phase 2 as real subprocesses with a mocked LLM. Catches the stdout-token/JSON round-trip that unit tests miss.

**Files:**
- Create: `tests/test_deposition/test_full_flow_smoke.py`

- [ ] **Step 1: Write the test**

Create `tests/test_deposition/test_full_flow_smoke.py`:

```python
"""End-to-end smoke test: phase 1 → user config → phase 2 → output docx."""

import json
import os
import subprocess
import sys
from pathlib import Path
from unittest.mock import patch

import pytest

from icharlotte_core.deposition import session_manager

PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent
AGENT_SCRIPT = PROJECT_ROOT / "Scripts" / "summarize_deposition.py"


@pytest.fixture
def mock_llm_env(tmp_path, monkeypatch):
    """Stand up a fake LLM via an environment variable that the agent honors in test mode.

    For this smoke test we rely on monkey-patching at the Python level rather than spawning
    a real subprocess — see the inline note. If you want a true subprocess test, wrap LLMCaller
    behavior behind an env var so the subprocess can pick it up.
    """
    return tmp_path


def test_full_flow_in_process(tmp_path, monkeypatch, capsys):
    """In-process smoke test of the phase 1 → phase 2 handoff.

    True subprocess test requires building an env-var-driven stub of LLMCaller; this
    test instead exercises main() in-process which still covers the session-JSON round-trip
    and the AWAITING_INPUT contract.
    """
    monkeypatch.setattr(session_manager, "SESSION_DIR", tmp_path / "sessions")

    sys.path.insert(0, str(PROJECT_ROOT / "Scripts"))
    import summarize_deposition

    # Build a fake input pdf
    input_path = tmp_path / "Smith Depo.pdf"
    input_path.write_bytes(b"%PDF-1.4\n")

    # Stub extraction
    from types import SimpleNamespace
    fake_text = ("DEPOSITION OF JOHN SMITH\nTaken on January 15, 2024\n\n"
                 "Q. State your name.\nA. John Smith.\n") * 30

    def fake_extract(self, p):
        return SimpleNamespace(success=True, text=fake_text, char_count=len(fake_text),
                               page_count=10, ocr_pages=[], ocr_percentage=0.0, error=None)

    monkeypatch.setattr(summarize_deposition.DocumentProcessor, "extract_with_dynamic_ocr", fake_extract)
    monkeypatch.setattr(summarize_deposition, "save_to_docx",
                         lambda content, output_path, deponent, date, logger:
                             Path(output_path).parent.mkdir(parents=True, exist_ok=True) or
                             Path(output_path).write_text(content, encoding="utf-8") or True)
    monkeypatch.setattr(summarize_deposition, "_register_outputs", lambda *a, **kw: None, raising=False)

    canned_topics = json.dumps([
        {"title": "Pre-Accident History", "rank": 1, "discussion_density": "high"},
        {"title": "Mechanism Of Injury", "rank": 2, "discussion_density": "high"},
    ])
    canned_summary = "**Pre-Accident History**\n- Bullet.\n\n**Mechanism Of Injury**\n- Bullet."

    call_responses = iter([canned_topics, canned_summary])

    def fake_call(self, prompt, text, task_type=None, **kw):
        return next(call_responses)

    monkeypatch.setattr(summarize_deposition.LLMCaller, "call", fake_call)

    # Phase 1
    logger = summarize_deposition.AgentLogger("DepoSmoke", log_to_file=False)
    assert summarize_deposition.process_topics(str(input_path), logger) is True

    out = capsys.readouterr().out
    awaiting = [ln for ln in out.splitlines() if ln.startswith("AWAITING_INPUT:")]
    assert awaiting, "phase 1 did not emit AWAITING_INPUT"
    session_path = Path(awaiting[-1][len("AWAITING_INPUT:"):])
    assert session_path.exists()

    # Simulate the dialog writing user_config
    session_manager.update_user_config(session_path, {
        "selected_topics": ["Pre-Accident History", "Mechanism Of Injury"],
        "added_topics": [],
        "bullets_per_topic": 5,
        "deponent_label": "Plaintiff",
        "custom_rules": "",
        "cross_check_enabled": False,
    })

    # Phase 2
    assert summarize_deposition.process_summary(str(session_path), logger) is True

    # Session cleaned up
    assert not session_path.exists()
```

- [ ] **Step 2: Run the test**

```
python -m pytest tests/test_deposition/test_full_flow_smoke.py -v
```

Expected: PASS.

- [ ] **Step 3: Run the full deposition test suite**

```
python -m pytest tests/test_deposition/ -v
```

Expected: every test in this plan passes.

- [ ] **Step 4: Commit**

```bash
git add tests/test_deposition/test_full_flow_smoke.py
git commit -m "test(deposition): full-flow smoke test for two-phase agent"
```

---

## Task 15: Manual verification

Per CLAUDE.md (mandatory: test the feature after building it). Run these checks on the live app.

**Files:** (none — manual)

- [ ] **Step 1: Single-file run**

1. Launch the app: `python iCharlotte.py`.
2. Select a real (short) deposition PDF from a test case folder.
3. Trigger the Summarize Deposition agent.
4. Confirm phase 1 progresses to ~95%, then the **READY** button appears on the status row.
5. Click READY. Confirm the popup loads with the agent's topics pre-filled and checked, deponent label pre-filled.
6. Uncheck one topic, rename another, add a custom topic, set bullets=4, change the label, leave cross-check checked, type a custom rule.
7. Click Generate Summary. Confirm phase 2 progress moves and the output `.docx` is written.
8. Open the `.docx`. Verify it contains exactly the topics you ended up with (selected after rename + added), each with ~4 bullets, with the deponent referred to using your label, and reflecting your custom rule.

- [ ] **Step 2: Folder run (concurrent READY buttons)**

1. Select a folder containing 2–3 deposition files.
2. Confirm each gets its own status row, each progresses through phase 1 independently, and each surfaces its own READY button.
3. Open them one at a time, fill out the popup, run phase 2, confirm each `.docx` is generated.

- [ ] **Step 3: Cancel and re-open the popup**

1. After phase 1 surfaces READY, click READY, then click Cancel in the dialog.
2. Confirm the READY button is still visible. Click it again — popup re-opens cleanly with default values (no persistence between opens).

- [ ] **Step 4: Cancel while paused**

1. After phase 1 surfaces READY, click the status row's existing Cancel button.
2. Confirm the agent is marked cancelled and the session JSON + cached `.txt` are removed from `logs/depo_sessions/` (or noted to remain for cleanup later — out of scope here, but check the behavior).

- [ ] **Step 5: Cross-check disabled**

1. Run a deposition, in the popup uncheck "Run cross-check pass", click Generate.
2. Open the agent log and confirm only one `task_type="summary"` LLM call occurred for phase 2 (no `task_type="cross_check"` call).

- [ ] **Step 6: Save a memory note**

If any quirks or behaviors surprised you during manual testing, add a topic file to `C:\Users\ASerpik.DESKTOP-MRIMK0D\.claude\projects\C--geminiterminal2\memory\` (e.g., `interactive_depo_summary.md`) and link it from `MEMORY.md`.

- [ ] **Step 7: Final commit (only if any cleanup or fix needed)**

If manual testing surfaces a bug that needs a quick fix, commit it:

```bash
git add <changed files>
git commit -m "fix(deposition): <one-line description of the issue>"
```

If everything passes cleanly, no extra commit needed.

---

## Done

All tasks complete:
- ✅ Session manager module
- ✅ Three prompt files (new + 2 rewritten)
- ✅ `Scripts/summarize_deposition.py` split into `process_topics` and `process_summary`, with `--phase` dispatcher
- ✅ Dead code removed (exhibits, impeachment, parallel extraction)
- ✅ `AgentRunner` extended with `awaiting_input` signal, paused-exit, `resume_with_config`
- ✅ Status widget READY button and `ready_clicked` signal
- ✅ `DepoSummaryConfigDialog` popup
- ✅ `IndexTab` wires it together
- ✅ Unit + integration tests
- ✅ Manual verification on a real deposition
