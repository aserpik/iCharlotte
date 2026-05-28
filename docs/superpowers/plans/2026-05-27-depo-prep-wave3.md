# Depo Prep — Wave 3: Phase 2 + UI + Integration

> Sub-plan of [2026-05-27-depo-prep.md](2026-05-27-depo-prep.md). Waves 1 and 2 must be complete and green.

Goal: build the Phase 2 pipeline modules, the custom Settings page (with embedded topic editor), the custom Output page, wire the task into the registry, and verify end-to-end with an integration test.

This wave is split into three sub-blocks executed in order:

- **Block A — Phase 2 pipeline modules** (Tasks 7–11): pure-Python; no Qt.
- **Block B — UI wiring** (Tasks 12–16): registry extension, settings page, topic editor, output page, task registration.
- **Block C — Integration + LLM config** (Tasks 17–18): end-to-end test, llm_preferences entry.

---

## Block A — Phase 2 pipeline modules

### Task 7: Per-topic question generation module

**Files:**
- Create: `Scripts/depo_prep_lib/questions.py`
- Create: `tests/test_wizard/test_depo_prep_questions.py`

- [ ] **Step 1: Write the failing test**

```python
# tests/test_wizard/test_depo_prep_questions.py
import json
from unittest.mock import MagicMock

import pytest

from Scripts.depo_prep_lib.questions import generate_questions_for_topic
from Scripts.depo_prep_lib.schemas import Topic


def _topic(refs=None):
    return {
        "id": "t01", "title": "Pre-existing conditions",
        "strategic_note": "Establish chronic LBP since 2019",
        "relevant_digest_refs": refs or ["med.json#factual_anchors[0]"],
        "default_checked": True, "lawyer_added": False,
    }


def _digest():
    return {
        "source_id": "med.json",
        "factual_anchors": [{"fact": "2019-03 PT intake: chronic LBP", "location": "p.12",
                              "topic_tags": ["injury"]}],
        "deponent_statements": [], "inconsistencies": [],
        "source_kind": "medical_records", "summary": "",
    }


def _llm(payload):
    caller = MagicMock()
    caller.call = MagicMock(return_value=json.dumps(payload))
    return caller


def test_generate_questions_returns_topic_questions():
    payload = {"topic_id": "t01", "questions": [
        {"n": 1, "text": "Before 2024, did you have back pain?"}
    ]}
    caller = _llm(payload)
    result = generate_questions_for_topic(
        topic=_topic(), digests_by_source={"med.json": _digest()},
        llm_caller=caller, deponent_name="Jane", deponent_role="P",
        style="lockdown", free_text_notes="",
        flags={"strategic_note": False, "source_facts": False,
               "impeachment_hook": False, "objection_alts": False},
    )
    assert result["topic_id"] == "t01"
    assert len(result["questions"]) == 1
    assert "error" not in result


def test_generate_questions_returns_error_on_llm_failure():
    caller = MagicMock()
    caller.call = MagicMock(side_effect=RuntimeError("LLM timeout"))
    result = generate_questions_for_topic(
        topic=_topic(), digests_by_source={"med.json": _digest()},
        llm_caller=caller, deponent_name="Jane", deponent_role="P",
        style="discovery", free_text_notes="",
        flags={"strategic_note": False, "source_facts": False,
               "impeachment_hook": False, "objection_alts": False},
    )
    assert "error" in result
    assert "LLM timeout" in result["error"]
    assert result["questions"] == []


def test_generate_questions_resolves_relevant_refs():
    """Only the referenced digest entries are sent in the prompt text."""
    payload = {"topic_id": "t01", "questions": [{"n": 1, "text": "Q"}]}
    caller = _llm(payload)
    generate_questions_for_topic(
        topic=_topic(refs=["med.json#factual_anchors[0]"]),
        digests_by_source={
            "med.json": _digest(),
            "other.json": {"source_id": "other.json", "source_kind": "other",
                            "factual_anchors": [{"fact": "irrelevant", "location": "z",
                                                  "topic_tags": []}],
                            "deponent_statements": [], "inconsistencies": [], "summary": ""},
        },
        llm_caller=caller, deponent_name="J", deponent_role="P",
        style="discovery", free_text_notes="",
        flags={"strategic_note": False, "source_facts": False,
               "impeachment_hook": False, "objection_alts": False},
    )
    # The text payload should contain the chronic LBP fact, not the irrelevant one.
    call = caller.call.call_args
    text = call.kwargs.get("text") or call.args[1] if len(call.args) > 1 else ""
    assert "chronic LBP" in text
    assert "irrelevant" not in text


def test_generate_questions_lawyer_added_uses_full_digest():
    """Lawyer-added topics with empty refs see all digest entries."""
    payload = {"topic_id": "t99", "questions": [{"n": 1, "text": "Q"}]}
    caller = _llm(payload)
    topic = {"id": "t99", "title": "Custom", "strategic_note": "s",
             "relevant_digest_refs": [], "default_checked": True, "lawyer_added": True}
    generate_questions_for_topic(
        topic=topic,
        digests_by_source={"med.json": _digest()},
        llm_caller=caller, deponent_name="J", deponent_role="P",
        style="discovery", free_text_notes="",
        flags={"strategic_note": False, "source_facts": False,
               "impeachment_hook": False, "objection_alts": False},
    )
    call = caller.call.call_args
    text = call.kwargs.get("text") or call.args[1] if len(call.args) > 1 else ""
    # Whole digest should be in the payload for lawyer-added topics.
    assert "chronic LBP" in text


def test_generate_questions_passes_flags_to_prompt():
    payload = {"topic_id": "t01", "questions": [{"n": 1, "text": "Q"}]}
    caller = _llm(payload)
    generate_questions_for_topic(
        topic=_topic(), digests_by_source={"med.json": _digest()},
        llm_caller=caller, deponent_name="J", deponent_role="P",
        style="discovery", free_text_notes="",
        flags={"strategic_note": True, "source_facts": True,
               "impeachment_hook": True, "objection_alts": True},
    )
    prompt = caller.call.call_args.kwargs.get("prompt") or caller.call.call_args.args[0]
    assert "purpose" in prompt.lower()
    assert "impeachment" in prompt.lower()
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_wizard/test_depo_prep_questions.py -v`
Expected: ImportError.

- [ ] **Step 3: Write minimal implementation**

```python
# Scripts/depo_prep_lib/questions.py
"""Stage A — per-topic question generation."""
from __future__ import annotations

import json
import re
from typing import Dict

from .prompts import build_per_topic_questions_prompt
from .source_digest import _parse_llm_json


# "med.json#factual_anchors[0]"  →  ("med.json", "factual_anchors", 0)
_REF_RE = re.compile(r"^([^#]+)#([a-z_]+)\[(\d+)\]$")


def _resolve_refs(refs, digests_by_source) -> str:
    """Build text payload listing only the referenced digest entries."""
    if not refs:
        return ""
    blocks = []
    for r in refs:
        m = _REF_RE.match(r.strip())
        if not m:
            continue
        src_id, field, idx = m.group(1), m.group(2), int(m.group(3))
        digest = digests_by_source.get(src_id)
        if not digest:
            continue
        entries = digest.get(field) or []
        if 0 <= idx < len(entries):
            blocks.append(
                f"=== {src_id} :: {field}[{idx}] ===\n"
                + json.dumps(entries[idx], indent=2, ensure_ascii=False)
            )
    return "\n\n".join(blocks)


def _full_digest_payload(digests_by_source) -> str:
    blocks = []
    for src_id, digest in digests_by_source.items():
        blocks.append(f"=== {src_id} ===\n" + json.dumps(digest, indent=2, ensure_ascii=False))
    return "\n\n".join(blocks)


def generate_questions_for_topic(
    *, topic: dict, digests_by_source: Dict[str, dict],
    llm_caller, deponent_name: str, deponent_role: str,
    style: str, free_text_notes: str, flags: Dict[str, bool],
) -> dict:
    """Generate a TopicQuestions-shaped dict for one topic.

    On LLM failure, returns {"topic_id": ..., "questions": [], "error": "..."}.
    """
    is_lawyer_added = bool(topic.get("lawyer_added"))
    refs = topic.get("relevant_digest_refs") or []

    if is_lawyer_added or not refs:
        digest_text = _full_digest_payload(digests_by_source)
    else:
        digest_text = _resolve_refs(refs, digests_by_source)

    prompt, text_payload = build_per_topic_questions_prompt(
        deponent_name=deponent_name, deponent_role=deponent_role, style=style,
        topic_title=topic["title"], strategic_note=topic.get("strategic_note", ""),
        digest_excerpts_text=digest_text, free_text_notes=free_text_notes,
        include_strategic_note=bool(flags.get("strategic_note")),
        include_source_facts=bool(flags.get("source_facts")),
        include_impeachment_hook=bool(flags.get("impeachment_hook")),
        include_objection_alts=bool(flags.get("objection_alts")),
    )

    try:
        raw = llm_caller.call(
            prompt=prompt, text=text_payload,
            task_type="general", agent_id="DepoPrep", pass_name="topic_questions",
        )
        data = _parse_llm_json(raw)
        data["topic_id"] = topic["id"]  # force consistency
        # Ensure required shape.
        if not isinstance(data.get("questions"), list):
            data["questions"] = []
        return data
    except Exception as e:
        return {"topic_id": topic["id"], "questions": [], "error": str(e)}
```

- [ ] **Step 4: Run tests**

Run: `python -m pytest tests/test_wizard/test_depo_prep_questions.py -v`
Expected: 5 passed.

- [ ] **Step 5: Commit**

```bash
git add Scripts/depo_prep_lib/questions.py tests/test_wizard/test_depo_prep_questions.py
git commit -m "feat(depo_prep): questions — per-topic generation with ref resolution + flag-conditional fields"
```

---

### Task 8: Merge module (dedup + coverage)

**Files:**
- Create: `Scripts/depo_prep_lib/merge.py`
- Create: `tests/test_wizard/test_depo_prep_merge.py`

- [ ] **Step 1: Write the failing test**

```python
# tests/test_wizard/test_depo_prep_merge.py
import json
from unittest.mock import MagicMock

import pytest

from Scripts.depo_prep_lib.merge import dedup_and_coverage, apply_dedup


def test_apply_dedup_drops_marked_questions_and_renumbers():
    topic_outputs = [
        {"topic_id": "t01", "questions": [
            {"n": 1, "text": "Q1"}, {"n": 2, "text": "Q2"}, {"n": 3, "text": "Q3"}]},
        {"topic_id": "t02", "questions": [
            {"n": 1, "text": "DupOfT01Q2"}, {"n": 2, "text": "Q5"}]},
    ]
    dedup = {
        "duplicates": [{"keep": "t01.q2", "drop": "t02.q1", "reason": "same"}],
        "coverage_gaps": [], "renumber_after_dedup": True,
    }
    result = apply_dedup(topic_outputs, dedup)
    t02 = next(t for t in result if t["topic_id"] == "t02")
    assert len(t02["questions"]) == 1
    assert t02["questions"][0]["n"] == 1
    assert t02["questions"][0]["text"] == "Q5"


def test_apply_dedup_handles_missing_topic_or_question_gracefully():
    topic_outputs = [{"topic_id": "t01", "questions": [{"n": 1, "text": "Q"}]}]
    dedup = {"duplicates": [{"keep": "t99.q1", "drop": "t01.q99", "reason": "?"}],
             "coverage_gaps": [], "renumber_after_dedup": True}
    # Should not raise; nothing dropped.
    result = apply_dedup(topic_outputs, dedup)
    assert len(result[0]["questions"]) == 1


def test_dedup_and_coverage_returns_parsed_dedup():
    caller = MagicMock()
    caller.call = MagicMock(return_value=json.dumps({
        "duplicates": [{"keep": "t01.q1", "drop": "t02.q1", "reason": "same"}],
        "coverage_gaps": ["Missing X"],
        "renumber_after_dedup": True,
    }))

    topic_outputs = [
        {"topic_id": "t01", "questions": [{"n": 1, "text": "Q1"}]},
        {"topic_id": "t02", "questions": [{"n": 1, "text": "Q1 again"}]},
    ]
    dedup = dedup_and_coverage(
        topic_outputs=topic_outputs,
        digests_by_source={"x.json": {"summary": "..."}},
        llm_caller=caller,
    )
    assert dedup["duplicates"][0]["drop"] == "t02.q1"
    assert dedup["coverage_gaps"] == ["Missing X"]


def test_dedup_returns_empty_on_llm_error():
    """Dedup must not fail Phase 2 if the LLM throws."""
    caller = MagicMock()
    caller.call = MagicMock(side_effect=RuntimeError("boom"))
    result = dedup_and_coverage(
        topic_outputs=[], digests_by_source={}, llm_caller=caller,
    )
    assert result == {"duplicates": [], "coverage_gaps": [],
                       "renumber_after_dedup": False}
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_wizard/test_depo_prep_merge.py -v`
Expected: ImportError.

- [ ] **Step 3: Write minimal implementation**

```python
# Scripts/depo_prep_lib/merge.py
"""Stage B — dedup + coverage check."""
from __future__ import annotations

import re
from typing import Dict, List

from .prompts import build_dedup_prompt
from .source_digest import _parse_llm_json


_REF_RE = re.compile(r"^([^.]+)\.q(\d+)$")


def _summarize_topic_outputs(topic_outputs: List[dict]) -> str:
    lines = []
    for t in topic_outputs:
        tid = t["topic_id"]
        for q in t.get("questions", []):
            lines.append(f"{tid}.q{q['n']}: {q.get('text', '')[:140]}")
    return "\n".join(lines)


def _summarize_digests(digests_by_source: Dict[str, dict]) -> str:
    return "\n".join(
        f"{src_id}: {(d.get('summary') or '').strip()[:140]}"
        for src_id, d in digests_by_source.items()
    )


def dedup_and_coverage(
    *, topic_outputs: List[dict], digests_by_source: Dict[str, dict], llm_caller,
) -> dict:
    """Run the dedup/coverage LLM call. On any failure return an empty result."""
    prompt, text_payload = build_dedup_prompt(
        topic_outputs_summary=_summarize_topic_outputs(topic_outputs),
        digest_summary=_summarize_digests(digests_by_source),
    )
    try:
        raw = llm_caller.call(
            prompt=prompt, text=text_payload,
            task_type="general", agent_id="DepoPrep", pass_name="dedup",
        )
        data = _parse_llm_json(raw)
        if not isinstance(data, dict):
            raise ValueError("dedup payload not a dict")
        data.setdefault("duplicates", [])
        data.setdefault("coverage_gaps", [])
        data.setdefault("renumber_after_dedup", True)
        return data
    except Exception:
        return {"duplicates": [], "coverage_gaps": [], "renumber_after_dedup": False}


def apply_dedup(topic_outputs: List[dict], dedup: dict) -> List[dict]:
    """Return a NEW list of topic_outputs with duplicate-drops applied and renumbered."""
    drops_by_topic: Dict[str, set] = {}
    for d in dedup.get("duplicates", []):
        drop_ref = d.get("drop", "")
        m = _REF_RE.match(drop_ref)
        if not m:
            continue
        topic_id, n = m.group(1), int(m.group(2))
        drops_by_topic.setdefault(topic_id, set()).add(n)

    out = []
    for t in topic_outputs:
        tid = t["topic_id"]
        drops = drops_by_topic.get(tid, set())
        kept = [q for q in t.get("questions", []) if q["n"] not in drops]
        if dedup.get("renumber_after_dedup", True):
            for i, q in enumerate(kept, 1):
                q = dict(q)
                q["n"] = i
                kept[i - 1] = q
        new_topic = dict(t)
        new_topic["questions"] = kept
        out.append(new_topic)
    return out
```

- [ ] **Step 4: Run tests**

Run: `python -m pytest tests/test_wizard/test_depo_prep_merge.py -v`
Expected: 4 passed.

- [ ] **Step 5: Commit**

```bash
git add Scripts/depo_prep_lib/merge.py tests/test_wizard/test_depo_prep_merge.py
git commit -m "feat(depo_prep): merge — dedup + coverage with safe fallback on LLM failure"
```

---

### Task 9: Polish module

**Files:**
- Create: `Scripts/depo_prep_lib/polish.py`
- Create: `tests/test_wizard/test_depo_prep_polish.py`

The polish module's job is to call the LLM with the polish prompt AND to **verify the LLM didn't add/drop questions**. If it did, the polish is rejected and the pre-polish outline is returned unchanged.

- [ ] **Step 1: Write the failing test**

```python
# tests/test_wizard/test_depo_prep_polish.py
import json
from unittest.mock import MagicMock

import pytest

from Scripts.depo_prep_lib.polish import polish_outline


def _outline(topics):
    return {"topics": topics}


def _t(tid, n):
    return {"topic_id": tid, "questions": [{"n": i, "text": f"Q{i}"} for i in range(1, n + 1)]}


def test_polish_accepts_when_question_counts_match():
    orig = _outline([_t("t01", 3), _t("t02", 2)])
    polished_payload = _outline([
        {"topic_id": "t01", "questions": [
            {"n": 1, "text": "Q1 polished"}, {"n": 2, "text": "Q2 polished"},
            {"n": 3, "text": "Q3 polished"}]},
        {"topic_id": "t02", "questions": [
            {"n": 1, "text": "Q1p"}, {"n": 2, "text": "Q2p"}]},
    ])
    caller = MagicMock()
    caller.call = MagicMock(return_value=json.dumps(polished_payload))

    result = polish_outline(outline=orig, llm_caller=caller)
    assert result["topics"][0]["questions"][0]["text"] == "Q1 polished"


def test_polish_rejects_when_topic_added():
    orig = _outline([_t("t01", 2)])
    polished_payload = _outline([_t("t01", 2), _t("t99", 1)])  # added a topic
    caller = MagicMock()
    caller.call = MagicMock(return_value=json.dumps(polished_payload))

    result = polish_outline(outline=orig, llm_caller=caller)
    # Reverts to original.
    assert len(result["topics"]) == 1


def test_polish_rejects_when_question_dropped():
    orig = _outline([_t("t01", 3)])
    polished_payload = _outline([_t("t01", 2)])  # one Q missing
    caller = MagicMock()
    caller.call = MagicMock(return_value=json.dumps(polished_payload))

    result = polish_outline(outline=orig, llm_caller=caller)
    assert len(result["topics"][0]["questions"]) == 3
    # Original Qs preserved.
    assert result["topics"][0]["questions"][2]["text"] == "Q3"


def test_polish_returns_original_on_llm_error():
    orig = _outline([_t("t01", 2)])
    caller = MagicMock()
    caller.call = MagicMock(side_effect=RuntimeError("boom"))
    result = polish_outline(outline=orig, llm_caller=caller)
    assert result == orig
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_wizard/test_depo_prep_polish.py -v`
Expected: ImportError.

- [ ] **Step 3: Write minimal implementation**

```python
# Scripts/depo_prep_lib/polish.py
"""Stage C — phrasing-only polish with structural validation."""
from __future__ import annotations

import json
from typing import Dict, List

from .prompts import build_polish_prompt
from .source_digest import _parse_llm_json


def _shape_sig(outline: dict) -> Dict[str, int]:
    """Return {topic_id: question_count} for structural comparison."""
    return {
        t["topic_id"]: len(t.get("questions", []))
        for t in outline.get("topics", [])
    }


def polish_outline(*, outline: dict, llm_caller) -> dict:
    """Run polish LLM. If the result changes topic_ids or question counts, revert to original."""
    try:
        raw = llm_caller.call(
            prompt=build_polish_prompt(outline_text=json.dumps(outline, indent=2))[0],
            text=json.dumps(outline, indent=2),
            task_type="general", agent_id="DepoPrep", pass_name="polish",
        )
        polished = _parse_llm_json(raw)
        if _shape_sig(polished) != _shape_sig(outline):
            return outline  # structural violation; revert
        return polished
    except Exception:
        return outline
```

- [ ] **Step 4: Run tests**

Run: `python -m pytest tests/test_wizard/test_depo_prep_polish.py -v`
Expected: 4 passed.

- [ ] **Step 5: Commit**

```bash
git add Scripts/depo_prep_lib/polish.py tests/test_wizard/test_depo_prep_polish.py
git commit -m "feat(depo_prep): polish — phrasing pass with structural shape validation"
```

---

### Task 10: Render `.docx` module

**Files:**
- Create: `Scripts/depo_prep_lib/render_docx.py`
- Create: `tests/test_wizard/test_depo_prep_render_docx.py`

- [ ] **Step 1: Write the failing test**

```python
# tests/test_wizard/test_depo_prep_render_docx.py
from pathlib import Path

import pytest
from docx import Document

from Scripts.depo_prep_lib.render_docx import render_outline_docx


def _outline_payload():
    return {
        "deponent_name": "Jane Doe",
        "deponent_role": "Plaintiff",
        "topics": [
            {"topic_id": "t01", "title": "Pre-existing conditions",
             "strategic_note": "Establish baseline.",
             "questions": [
                 {"n": 1, "text": "Before 2024, did you have back pain?",
                  "purpose": "Establish baseline.",
                  "source_facts": ["RFA #7 denied prior pain", "2019 PT intake notes chronic LBP"],
                  "impeachment_hook": "Confront with RFA #7"},
                 {"n": 2, "text": "When did you first see a chiropractor?"},
             ]},
            {"topic_id": "t02", "title": "Treatment timeline",
             "strategic_note": "Highlight gaps.",
             "questions": [{"n": 1, "text": "What treatment did you receive?"}]},
        ],
        "coverage_gaps": ["No question addresses 2019 chiropractor visits."],
    }


def test_render_creates_docx_with_expected_structure(tmp_path):
    out = tmp_path / "outline.docx"
    render_outline_docx(outline=_outline_payload(), output_path=out)
    assert out.exists()

    doc = Document(str(out))
    text = "\n".join(p.text for p in doc.paragraphs)
    assert "Jane Doe" in text
    assert "Pre-existing conditions" in text
    assert "Before 2024" in text
    assert "Coverage" in text  # coverage notes section


def test_render_skips_optional_fields_when_absent(tmp_path):
    payload = {
        "deponent_name": "X", "deponent_role": "Y",
        "topics": [{"topic_id": "t01", "title": "T", "strategic_note": "",
                    "questions": [{"n": 1, "text": "Q only"}]}],
        "coverage_gaps": [],
    }
    out = tmp_path / "outline.docx"
    render_outline_docx(outline=payload, output_path=out)
    doc = Document(str(out))
    text = "\n".join(p.text for p in doc.paragraphs)
    assert "Q only" in text
    assert "Purpose:" not in text
    assert "Source facts:" not in text
    assert "Coverage" not in text  # no gaps → no section


def test_render_no_empty_paragraphs_for_spacing(tmp_path):
    """Spacing must use space_after, not empty paragraphs (MEMORY.md rule)."""
    out = tmp_path / "outline.docx"
    render_outline_docx(outline=_outline_payload(), output_path=out)
    doc = Document(str(out))
    empty_paras = [p for p in doc.paragraphs if not p.text.strip()]
    # Allow a handful of truly structural empties (e.g., spacing line under title);
    # but the bulk must use space_after. Cap loosely at 2.
    assert len(empty_paras) <= 2, (
        f"Found {len(empty_paras)} empty paragraphs — use space_after instead.")
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_wizard/test_depo_prep_render_docx.py -v`
Expected: ImportError.

- [ ] **Step 3: Write minimal implementation**

```python
# Scripts/depo_prep_lib/render_docx.py
"""Stage D — render outline.docx via python-docx."""
from __future__ import annotations

from pathlib import Path
from typing import Union

from docx import Document
from docx.shared import Inches, Pt


def _para(doc, text, *, bold=False, italic=False, size=12, indent_left=0.0,
          first_line=0.0, space_after_pt=6):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Inches(indent_left)
    if first_line:
        p.paragraph_format.first_line_indent = Inches(first_line)
    p.paragraph_format.space_after = Pt(space_after_pt)
    p.paragraph_format.line_spacing = 1.15
    r = p.add_run(text)
    r.bold = bold
    r.italic = italic
    r.font.name = "Times New Roman"
    r.font.size = Pt(size)
    return p


def render_outline_docx(*, outline: dict, output_path: Union[str, Path]) -> None:
    """Write outline.docx at output_path. Overwrites any existing file."""
    output_path = Path(output_path)
    doc = Document()

    style = doc.styles["Normal"]
    style.font.name = "Times New Roman"
    style.font.size = Pt(12)

    deponent_name = outline.get("deponent_name") or "Unknown Deponent"
    deponent_role = outline.get("deponent_role") or ""

    # Title block
    _para(doc, f"Depo Prep Outline — {deponent_name}", bold=True, size=16, space_after_pt=4)
    if deponent_role:
        _para(doc, deponent_role, italic=True, size=11, space_after_pt=12)

    for topic in outline.get("topics", []):
        title = topic.get("title", "(Untitled topic)")
        strat = topic.get("strategic_note", "")
        _para(doc, title.upper(), bold=True, size=13, space_after_pt=4)
        if strat:
            _para(doc, f"Strategic: {strat}", italic=True, size=11,
                  indent_left=0.25, space_after_pt=6)

        for q in topic.get("questions", []):
            _para(doc, f"{q['n']}.  {q.get('text', '')}", size=12,
                  indent_left=0.5, first_line=-0.25, space_after_pt=4)

            if q.get("purpose"):
                _para(doc, f"Purpose: {q['purpose']}", italic=True, size=10,
                      indent_left=0.75, space_after_pt=2)
            if q.get("source_facts"):
                _para(doc, "Source facts:", italic=True, size=10,
                      indent_left=0.75, space_after_pt=2)
                for f in q["source_facts"]:
                    _para(doc, f"• {f}", size=10, indent_left=1.0, space_after_pt=2)
            if q.get("impeachment_hook"):
                _para(doc, f"Impeachment: {q['impeachment_hook']}", italic=True,
                      size=10, indent_left=0.75, space_after_pt=2)
            if q.get("objection_alts"):
                _para(doc, "Objection alts:", italic=True, size=10,
                      indent_left=0.75, space_after_pt=2)
                for a in q["objection_alts"]:
                    _para(doc, f"• {a}", size=10, indent_left=1.0, space_after_pt=2)

    gaps = outline.get("coverage_gaps") or []
    if gaps:
        _para(doc, "Coverage notes from the AI", bold=True, size=12, space_after_pt=4)
        for g in gaps:
            _para(doc, f"• {g}", size=11, indent_left=0.25, space_after_pt=2)

    doc.save(str(output_path))
```

- [ ] **Step 4: Run tests**

Run: `python -m pytest tests/test_wizard/test_depo_prep_render_docx.py -v`
Expected: 3 passed.

- [ ] **Step 5: Commit**

```bash
git add Scripts/depo_prep_lib/render_docx.py tests/test_wizard/test_depo_prep_render_docx.py
git commit -m "feat(depo_prep): render_docx — outline.docx with space_after spacing (no empty paragraphs)"
```

---

### Task 11: Render `.md` module + Phase 2 orchestrator

**Files:**
- Create: `Scripts/depo_prep_lib/render_md.py`
- Create: `Scripts/depo_prep_lib/phase2.py`
- Modify: `Scripts/depo_prep.py` (wire `_cmd_generate` to call `phase2.run_phase2`)
- Create: `tests/test_wizard/test_depo_prep_render_md.py`
- Create: `tests/test_wizard/test_depo_prep_phase2_orchestrator.py`

- [ ] **Step 1: Write the failing tests**

```python
# tests/test_wizard/test_depo_prep_render_md.py
from pathlib import Path

import pytest

from Scripts.depo_prep_lib.render_md import render_outline_md


def test_render_md_includes_questions_and_headers(tmp_path):
    payload = {
        "deponent_name": "Jane Doe", "deponent_role": "Plaintiff",
        "topics": [
            {"topic_id": "t01", "title": "Pre-existing conditions",
             "strategic_note": "Establish baseline.",
             "questions": [
                 {"n": 1, "text": "Before 2024…", "purpose": "Baseline",
                  "source_facts": ["RFA #7 denied"]},
             ]}
        ],
        "coverage_gaps": ["Gap 1"],
    }
    out = tmp_path / "outline.md"
    render_outline_md(outline=payload, output_path=out)
    md = out.read_text(encoding="utf-8")
    assert "# Depo Prep Outline — Jane Doe" in md
    assert "## Pre-existing conditions" in md
    assert "_Strategic: Establish baseline._" in md
    assert "1." in md and "Before 2024" in md
    assert "Purpose" in md
    assert "RFA #7 denied" in md
    assert "## Coverage notes from the AI" in md
    assert "Gap 1" in md
```

```python
# tests/test_wizard/test_depo_prep_phase2_orchestrator.py
import json
from pathlib import Path
from unittest.mock import MagicMock

import pytest


@pytest.fixture
def session_dir(tmp_path):
    sd = tmp_path / "session"
    sd.mkdir()
    (sd / "digests").mkdir()
    return sd


def _make_session(session_dir, deponent_sources=None, context_sources=None):
    payload = {
        "version": 1,
        "phase": "awaiting_input",
        "deponent_name": "Jane Doe", "deponent_role": "Plaintiff",
        "style": "discovery", "free_text_notes": "",
        "per_topic_flags": {"strategic_note": True, "source_facts": False,
                             "impeachment_hook": False, "objection_alts": False},
        "case_root": str(session_dir.parent),
        "deponent_sources": deponent_sources or [], "context_sources": context_sources or [],
        "digests_index": ["src1.pdf"],
        "topics_warning": None,
    }
    (session_dir / "session.json").write_text(json.dumps(payload), encoding="utf-8")


def _make_topics(session_dir, topic_count=2):
    topics = [
        {"id": f"t{i:02d}", "title": f"Topic {i}", "strategic_note": "s",
         "relevant_digest_refs": [], "default_checked": True, "lawyer_added": False}
        for i in range(1, topic_count + 1)
    ]
    (session_dir / "topics.json").write_text(json.dumps({"topics": topics}), encoding="utf-8")


def _make_digest(session_dir, source_id="src1.pdf"):
    (session_dir / "digests" / f"{source_id}.json").write_text(
        json.dumps({"source_id": source_id, "source_kind": "other",
                    "deponent_statements": [], "factual_anchors": [],
                    "inconsistencies": [], "summary": ""}), encoding="utf-8")


def _phase2_llm():
    """Returns valid payloads for question generation, dedup, polish."""
    def call(prompt, text, task_type="general", agent_id=None, pass_name=None, **kwargs):
        if pass_name == "topic_questions":
            # echo topic_id from prompt isn't straightforward; emit a generic shape
            # the orchestrator must reassign topic_id anyway.
            return json.dumps({"topic_id": "t??", "questions": [
                {"n": 1, "text": "Generated Q1"}, {"n": 2, "text": "Generated Q2"}]})
        if pass_name == "dedup":
            return json.dumps({"duplicates": [], "coverage_gaps": ["Gap A"],
                                "renumber_after_dedup": True})
        if pass_name == "polish":
            # Echo the input outline unchanged (well-shaped polish).
            return text
        return ""
    caller = MagicMock()
    caller.call.side_effect = call
    return caller


def test_phase2_writes_outline_docx_and_md(session_dir):
    from Scripts.depo_prep_lib import phase2

    _make_session(session_dir)
    _make_topics(session_dir, topic_count=2)
    _make_digest(session_dir)

    phase2.run_phase2(
        session_path=str(session_dir / "session.json"),
        llm_caller=_phase2_llm(),
        progress=lambda *a, **k: None,
    )

    assert (session_dir / "outline.docx").exists()
    assert (session_dir / "outline.md").exists()
    md = (session_dir / "outline.md").read_text(encoding="utf-8")
    assert "Jane Doe" in md
    assert "Topic 1" in md
    assert "Gap A" in md


def test_phase2_handles_per_topic_failure(session_dir):
    from Scripts.depo_prep_lib import phase2

    _make_session(session_dir)
    _make_topics(session_dir, topic_count=2)
    _make_digest(session_dir)

    # LLM throws on topic_questions, returns valid for dedup/polish.
    def call(prompt, text, task_type="general", agent_id=None, pass_name=None, **kwargs):
        if pass_name == "topic_questions":
            raise RuntimeError("rate limit")
        if pass_name == "dedup":
            return json.dumps({"duplicates": [], "coverage_gaps": [], "renumber_after_dedup": False})
        if pass_name == "polish":
            return text
        return ""
    caller = MagicMock(); caller.call.side_effect = call

    # Should not raise — Phase 2 finishes with empty topic Qs and an error banner.
    phase2.run_phase2(
        session_path=str(session_dir / "session.json"),
        llm_caller=caller,
        progress=lambda *a, **k: None,
    )
    md = (session_dir / "outline.md").read_text(encoding="utf-8")
    assert "Topic 1" in md
    # Render still produces a doc; user sees gaps but topics present.
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `python -m pytest tests/test_wizard/test_depo_prep_render_md.py tests/test_wizard/test_depo_prep_phase2_orchestrator.py -v`
Expected: ImportErrors.

- [ ] **Step 3: Write minimal implementation**

```python
# Scripts/depo_prep_lib/render_md.py
"""Stage D — render outline.md."""
from __future__ import annotations

from pathlib import Path
from typing import Union


def render_outline_md(*, outline: dict, output_path: Union[str, Path]) -> None:
    deponent = outline.get("deponent_name") or "Unknown Deponent"
    role = outline.get("deponent_role") or ""

    lines = [f"# Depo Prep Outline — {deponent}"]
    if role:
        lines.append(f"_{role}_")
    lines.append("")

    for topic in outline.get("topics", []):
        title = topic.get("title", "(Untitled)")
        strat = topic.get("strategic_note", "")
        lines.append(f"## {title}")
        if strat:
            lines.append(f"_Strategic: {strat}_")
        lines.append("")
        for q in topic.get("questions", []):
            lines.append(f"{q['n']}. {q.get('text', '')}")
            if q.get("purpose"):
                lines.append(f"    - *Purpose: {q['purpose']}*")
            if q.get("source_facts"):
                lines.append(f"    - *Source facts:*")
                for f in q["source_facts"]:
                    lines.append(f"        - {f}")
            if q.get("impeachment_hook"):
                lines.append(f"    - *Impeachment: {q['impeachment_hook']}*")
            if q.get("objection_alts"):
                lines.append(f"    - *Objection alts:*")
                for a in q["objection_alts"]:
                    lines.append(f"        - {a}")
            lines.append("")

    gaps = outline.get("coverage_gaps") or []
    if gaps:
        lines.append("## Coverage notes from the AI")
        for g in gaps:
            lines.append(f"- {g}")

    Path(output_path).write_text("\n".join(lines), encoding="utf-8")
```

```python
# Scripts/depo_prep_lib/phase2.py
"""Phase 2 orchestrator — Stage A (parallel) → B (dedup) → C (polish) → D (render)."""
from __future__ import annotations

from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from typing import Callable, Dict, List

from .merge import dedup_and_coverage, apply_dedup
from .polish import polish_outline
from .questions import generate_questions_for_topic
from .render_docx import render_outline_docx
from .render_md import render_outline_md
from .session_io import read_json


def _load_digests(session_dir: Path) -> Dict[str, dict]:
    digests_dir = session_dir / "digests"
    out = {}
    for p in digests_dir.glob("*.json"):
        try:
            data = read_json(p)
            sid = data.get("source_id") or p.stem
            out[sid] = data
        except Exception:
            continue
    return out


def run_phase2(*, session_path: str, llm_caller, progress: Callable[[int, str], None]) -> None:
    session_path = Path(session_path)
    session = read_json(session_path)
    session_dir = session_path.parent

    topics_payload = read_json(session_dir / "topics.json")
    topics: List[dict] = [t for t in topics_payload.get("topics", []) if t.get("default_checked", True)]
    if not topics:
        topics = topics_payload.get("topics", [])

    digests = _load_digests(session_dir)
    flags = session.get("per_topic_flags", {})

    progress(5, f"Generating questions for {len(topics)} topic(s)…")
    topic_outputs: List[dict] = [None] * len(topics)

    def _one(idx_topic):
        idx, topic = idx_topic
        return idx, generate_questions_for_topic(
            topic=topic, digests_by_source=digests, llm_caller=llm_caller,
            deponent_name=session["deponent_name"], deponent_role=session.get("deponent_role", ""),
            style=session.get("style", "discovery"),
            free_text_notes=session.get("free_text_notes", ""),
            flags=flags,
        )

    with ThreadPoolExecutor(max_workers=4) as pool:
        futures = {pool.submit(_one, (i, t)): i for i, t in enumerate(topics)}
        done = 0
        for fut in as_completed(futures):
            idx, payload = fut.result()
            # Carry forward the topic's title + strategic_note for the renderers.
            payload["title"] = topics[idx].get("title", "")
            payload["strategic_note"] = topics[idx].get("strategic_note", "")
            topic_outputs[idx] = payload
            done += 1
            progress(5 + int(60 * done / len(topics)), f"Topic {done}/{len(topics)} done")

    progress(70, "Dedup + coverage check…")
    dedup = dedup_and_coverage(
        topic_outputs=topic_outputs, digests_by_source=digests, llm_caller=llm_caller,
    )
    topic_outputs = apply_dedup(topic_outputs, dedup)

    outline = {
        "deponent_name": session["deponent_name"],
        "deponent_role": session.get("deponent_role", ""),
        "topics": topic_outputs,
        "coverage_gaps": dedup.get("coverage_gaps", []),
    }

    progress(85, "Polish pass…")
    outline = polish_outline(outline=outline, llm_caller=llm_caller)

    progress(95, "Rendering…")
    render_outline_docx(outline=outline, output_path=session_dir / "outline.docx")
    render_outline_md(outline=outline, output_path=session_dir / "outline.md")
    progress(100, "Done")
```

Now modify `Scripts/depo_prep.py` to wire `_cmd_generate`:

```python
# Replace the placeholder _cmd_generate in Scripts/depo_prep.py with:
def _cmd_generate(session_path: str) -> int:
    from Scripts.depo_prep_lib import phase2
    try:
        phase2.run_phase2(
            session_path=session_path, llm_caller=_make_llm_caller(),
            progress=_progress,
        )
    except Exception as e:
        print(f"ERROR: Phase 2 failed: {e}", flush=True)
        import traceback
        traceback.print_exc()
        return 1
    # Print the absolute path of the produced .docx so the wizard can pick it up.
    from pathlib import Path
    docx = Path(session_path).parent / "outline.docx"
    print(f"OUTPUT:{docx}", flush=True)
    return 0
```

- [ ] **Step 4: Run tests**

Run: `python -m pytest tests/test_wizard/test_depo_prep_render_md.py tests/test_wizard/test_depo_prep_phase2_orchestrator.py -v`
Expected: 3 passed.

- [ ] **Step 5: Commit**

```bash
git add Scripts/depo_prep_lib/render_md.py Scripts/depo_prep_lib/phase2.py Scripts/depo_prep.py tests/test_wizard/test_depo_prep_render_md.py tests/test_wizard/test_depo_prep_phase2_orchestrator.py
git commit -m "feat(depo_prep): Phase 2 orchestrator — fan-out questions → dedup → polish → render"
```

---

## Block B — UI wiring

### Task 12: Extend `TaskSpec` with output_page factory + task_tab lookup

**Files:**
- Modify: `icharlotte_core/ui/wizard/registry.py` (add `_output_page_cls_factory` field; add helper getter)
- Modify: `icharlotte_core/ui/wizard/task_tab.py` (construct output page via spec factory if set)
- Create: `tests/test_wizard/test_registry_output_factory.py`

- [ ] **Step 1: Write the failing test**

```python
# tests/test_wizard/test_registry_output_factory.py
from icharlotte_core.ui.wizard.registry import TaskSpec


def test_taskspec_default_output_page_cls_is_OutputPage():
    spec = TaskSpec(task_id="x", title="X", description="x",
                    icon_glyph="X", script_name="x.py")
    from icharlotte_core.ui.wizard.pages.output_page import OutputPage
    assert spec.output_page_cls is OutputPage


def test_taskspec_custom_output_factory_is_used():
    class MyOutputPage:
        pass
    spec = TaskSpec(
        task_id="x", title="X", description="x", icon_glyph="X", script_name="x.py",
        _output_page_cls_factory=lambda: MyOutputPage,
    )
    assert spec.output_page_cls is MyOutputPage
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_wizard/test_registry_output_factory.py -v`
Expected: AttributeError on `output_page_cls`.

- [ ] **Step 3: Modify registry.py**

In `icharlotte_core/ui/wizard/registry.py`, add a default-factory helper near the top, and add the field + property:

```python
# Add near _default_settings_page_cls():
def _default_output_page_cls():
    from .pages.output_page import OutputPage
    return OutputPage
```

In the `TaskSpec` dataclass, add a new field after `_settings_page_cls_factory`:

```python
    _output_page_cls_factory: Optional[object] = field(default=None, repr=False, compare=False)
```

Add a new property after `settings_page_cls`:

```python
    @property
    def output_page_cls(self) -> type:
        """Return the OutputPage subclass for this task."""
        if self._output_page_cls_factory is not None:
            return self._output_page_cls_factory()
        return _default_output_page_cls()
```

- [ ] **Step 4: Modify task_tab.py**

In `icharlotte_core/ui/wizard/task_tab.py`, change the `__init__` line that constructs `OutputPage`:

```python
# Old:
self.output_page = OutputPage()
# New:
self.output_page = spec.output_page_cls()
```

Update the top-of-file import to still cover the default (no change strictly required, but keep `from icharlotte_core.ui.wizard.pages.output_page import OutputPage` for type clarity / fallback).

- [ ] **Step 5: Run tests**

Run: `python -m pytest tests/test_wizard/test_registry_output_factory.py tests/test_wizard/ -v -k "not depo_prep_settings_page and not depo_prep_output_page and not depo_prep_integration"`
Expected: 2 new passed; no regressions among existing wizard tests.

- [ ] **Step 6: Commit**

```bash
git add icharlotte_core/ui/wizard/registry.py icharlotte_core/ui/wizard/task_tab.py tests/test_wizard/test_registry_output_factory.py
git commit -m "refactor(wizard): TaskSpec gains _output_page_cls_factory; task_tab uses it"
```

---

### Task 13: Embedded topic editor widget

**Files:**
- Create: `icharlotte_core/ui/wizard/pages/depo_prep_topic_editor.py`
- Create: `tests/test_wizard/test_depo_prep_topic_editor.py`

The topic editor is a self-contained widget so it can be tested independently. It wraps a `QListWidget` configured with native drag-and-drop and exposes:
- `set_topics(list[dict])` — populate from `topics.json`
- `get_topics() -> list[dict]` — return current state (order, checked, edits)
- `add_topic(title, strategic_note)` — append a lawyer-added topic
- `topics_changed` signal — emitted on any mutation

**Critical (MEMORY.md `qlistwidget_setitemwidget_drag.md`):** native `Qt.ItemFlag` flags only; NO `setItemWidget` (breaks drag-reorder).

- [ ] **Step 1: Write the failing test**

```python
# tests/test_wizard/test_depo_prep_topic_editor.py
import pytest

pytest.importorskip("pytestqt")
pytest.importorskip("PySide6")

from PySide6.QtCore import Qt


def _topics():
    return [
        {"id": "t01", "title": "Pre-existing", "strategic_note": "Establish baseline",
         "relevant_digest_refs": [], "default_checked": True, "lawyer_added": False},
        {"id": "t02", "title": "Treatment", "strategic_note": "Highlight gaps",
         "relevant_digest_refs": [], "default_checked": True, "lawyer_added": False},
    ]


def test_topic_editor_populates_and_returns(qtbot):
    from icharlotte_core.ui.wizard.pages.depo_prep_topic_editor import TopicEditor
    w = TopicEditor()
    qtbot.addWidget(w)
    w.set_topics(_topics())
    out = w.get_topics()
    assert [t["id"] for t in out] == ["t01", "t02"]
    assert all(t["default_checked"] for t in out)


def test_topic_editor_emits_topics_changed_on_uncheck(qtbot):
    from icharlotte_core.ui.wizard.pages.depo_prep_topic_editor import TopicEditor
    w = TopicEditor()
    qtbot.addWidget(w)
    w.set_topics(_topics())
    with qtbot.waitSignal(w.topics_changed, timeout=1000):
        w.set_checked(0, False)
    out = w.get_topics()
    assert out[0]["default_checked"] is False


def test_topic_editor_add_topic_appends_lawyer_added(qtbot):
    from icharlotte_core.ui.wizard.pages.depo_prep_topic_editor import TopicEditor
    w = TopicEditor()
    qtbot.addWidget(w)
    w.set_topics(_topics())
    w.add_topic(title="Custom", strategic_note="My note")
    out = w.get_topics()
    assert len(out) == 3
    assert out[2]["title"] == "Custom"
    assert out[2]["lawyer_added"] is True
    assert out[2]["id"].startswith("t")  # auto-generated id


def test_topic_editor_remove_topic(qtbot):
    from icharlotte_core.ui.wizard.pages.depo_prep_topic_editor import TopicEditor
    w = TopicEditor()
    qtbot.addWidget(w)
    w.set_topics(_topics())
    w.remove_topic_at(0)
    out = w.get_topics()
    assert len(out) == 1
    assert out[0]["id"] == "t02"


def test_topic_editor_does_not_use_setitemwidget(qtbot):
    """MEMORY.md rule: setItemWidget breaks drag-reorder; we must use item flags."""
    from icharlotte_core.ui.wizard.pages.depo_prep_topic_editor import TopicEditor
    w = TopicEditor()
    qtbot.addWidget(w)
    w.set_topics(_topics())
    # Drag-and-drop is enabled at the QListWidget level.
    lw = w._list  # internal handle for testing
    assert lw.dragDropMode() == lw.DragDropMode.InternalMove
    # No items have widgets attached.
    for i in range(lw.count()):
        assert lw.itemWidget(lw.item(i)) is None
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_wizard/test_depo_prep_topic_editor.py -v`
Expected: ImportError.

- [ ] **Step 3: Write minimal implementation**

```python
# icharlotte_core/ui/wizard/pages/depo_prep_topic_editor.py
"""TopicEditor — drag-reorder + checkable + add/remove topic list for Depo Prep.

Uses native Qt.ItemFlag flags; NEVER setItemWidget (MEMORY.md
qlistwidget_setitemwidget_drag.md). The visible row text encodes both title and
strategic note via two lines; on edit, we re-parse.
"""
from __future__ import annotations

from typing import List

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QAbstractItemView, QHBoxLayout, QInputDialog, QListWidget, QListWidgetItem,
    QPushButton, QVBoxLayout, QWidget,
)


def _format_item_text(title: str, strategic_note: str) -> str:
    if strategic_note:
        return f"{title}\n    Strategic: {strategic_note}"
    return title


def _parse_item_text(text: str) -> tuple:
    if "\n    Strategic: " in text:
        title, _, rest = text.partition("\n    Strategic: ")
        return title.strip(), rest.strip()
    return text.strip(), ""


class TopicEditor(QWidget):
    topics_changed = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(4)

        self._list = QListWidget()
        self._list.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self._list.setDragDropMode(QAbstractItemView.DragDropMode.InternalMove)
        self._list.setDefaultDropAction(Qt.DropAction.MoveAction)
        self._list.itemChanged.connect(lambda *_: self.topics_changed.emit())
        self._list.model().rowsMoved.connect(lambda *_: self.topics_changed.emit())
        outer.addWidget(self._list)

        btns = QHBoxLayout()
        self.add_btn = QPushButton("+ Add custom topic")
        self.add_btn.clicked.connect(self._on_add_clicked)
        btns.addWidget(self.add_btn)
        self.remove_btn = QPushButton("Remove selected")
        self.remove_btn.clicked.connect(self._on_remove_clicked)
        btns.addWidget(self.remove_btn)
        btns.addStretch()
        outer.addLayout(btns)

        # Internal: track per-row metadata (id, lawyer_added, refs) keyed by row index.
        self._meta: List[dict] = []

    def set_topics(self, topics: List[dict]) -> None:
        self._list.blockSignals(True)
        try:
            self._list.clear()
            self._meta = []
            for t in topics:
                item = QListWidgetItem(_format_item_text(t.get("title", ""),
                                                          t.get("strategic_note", "")))
                item.setFlags(
                    Qt.ItemFlag.ItemIsSelectable
                    | Qt.ItemFlag.ItemIsEnabled
                    | Qt.ItemFlag.ItemIsUserCheckable
                    | Qt.ItemFlag.ItemIsEditable
                    | Qt.ItemFlag.ItemIsDragEnabled
                )
                item.setCheckState(Qt.CheckState.Checked if t.get("default_checked", True)
                                   else Qt.CheckState.Unchecked)
                self._list.addItem(item)
                self._meta.append({
                    "id": t.get("id"),
                    "lawyer_added": bool(t.get("lawyer_added", False)),
                    "relevant_digest_refs": list(t.get("relevant_digest_refs", [])),
                })
        finally:
            self._list.blockSignals(False)

    def get_topics(self) -> List[dict]:
        out = []
        for row in range(self._list.count()):
            item = self._list.item(row)
            title, strat = _parse_item_text(item.text())
            meta = self._meta[row] if row < len(self._meta) else {}
            out.append({
                "id": meta.get("id") or f"t{row+1:02d}",
                "title": title,
                "strategic_note": strat,
                "relevant_digest_refs": meta.get("relevant_digest_refs", []),
                "default_checked": item.checkState() == Qt.CheckState.Checked,
                "lawyer_added": bool(meta.get("lawyer_added", False)),
            })
        return out

    def set_checked(self, row: int, checked: bool) -> None:
        item = self._list.item(row)
        if item:
            item.setCheckState(Qt.CheckState.Checked if checked else Qt.CheckState.Unchecked)

    def add_topic(self, *, title: str, strategic_note: str) -> None:
        new_id = f"t{(len(self._meta) + 1):02d}"
        item = QListWidgetItem(_format_item_text(title, strategic_note))
        item.setFlags(
            Qt.ItemFlag.ItemIsSelectable
            | Qt.ItemFlag.ItemIsEnabled
            | Qt.ItemFlag.ItemIsUserCheckable
            | Qt.ItemFlag.ItemIsEditable
            | Qt.ItemFlag.ItemIsDragEnabled
        )
        item.setCheckState(Qt.CheckState.Checked)
        self._list.addItem(item)
        self._meta.append({"id": new_id, "lawyer_added": True, "relevant_digest_refs": []})
        self.topics_changed.emit()

    def remove_topic_at(self, row: int) -> None:
        if 0 <= row < self._list.count():
            self._list.takeItem(row)
            if row < len(self._meta):
                self._meta.pop(row)
            self.topics_changed.emit()

    def _on_add_clicked(self) -> None:
        title, ok = QInputDialog.getText(self, "New topic", "Topic title:")
        if not ok or not title.strip():
            return
        note, _ = QInputDialog.getText(self, "Strategic note",
                                        "1–2 sentence strategic note (optional):")
        self.add_topic(title=title.strip(), strategic_note=(note or "").strip())

    def _on_remove_clicked(self) -> None:
        rows = sorted({i.row() for i in self._list.selectedIndexes()}, reverse=True)
        for r in rows:
            self.remove_topic_at(r)
```

- [ ] **Step 4: Run tests**

Run: `python -m pytest tests/test_wizard/test_depo_prep_topic_editor.py -v`
Expected: 5 passed.

- [ ] **Step 5: Commit**

```bash
git add icharlotte_core/ui/wizard/pages/depo_prep_topic_editor.py tests/test_wizard/test_depo_prep_topic_editor.py
git commit -m "feat(depo_prep): TopicEditor widget — drag-reorder + checkable, no setItemWidget"
```

---

### Task 14: DepoPrepSettingsPage

**Files:**
- Create: `icharlotte_core/ui/wizard/pages/depo_prep_settings_page.py`
- Create: `tests/test_wizard/test_depo_prep_settings_page.py`

This page extends `SettingsPage`, hides the default body, and renders the spec layout: deponent dropdown, two file sections, style radios, free-text, content flags, "Analyze Sources" button, and a placeholder for the embedded `TopicEditor` (revealed after Phase 1).

**Important signals (must match TaskTab expectations):**
- `analyze_requested(dict)` — emitted when "Analyze Sources" clicked; carries the config dict that goes to `Scripts/depo_prep.py --phase=analyze`.
- `phase2_requested(str)` — emitted when "Generate Outline" clicked; carries the session.json path.

Because the existing `TaskTab` wires `proceed_requested` → `_start_run` and `phase2_requested` → `advance_to_status_with_phase2`, the cleanest approach is:
- Reuse `proceed_requested` (already wired) for the "Analyze Sources" click — emit it with the analyze config dict; the script will see `--phase=analyze`.
- Add `phase2_requested(str)` for the "Generate Outline" click.

But the existing `TaskTab._start_run` runs Phase 1 via `phase1_args=["--phase=analyze"]` automatically, and Phase 2 via the standard `resume_with_config(session_path)` path. That mechanic already exists for `summarize_depositions`. For Depo Prep, we want a slight variation: the config dict needs to be persisted to disk first (because `--phase=analyze --config=<path>` needs a file).

**Simplest path:** Override `to_dict()` to also write a `config.json` to a temp path in the case folder, and pass the temp path to the subprocess. The `SubprocessWorker` doesn't currently support `--config=<path>` — it appends files as positional arguments.

**Decision:** Adapt `Scripts/depo_prep.py` to accept the config as a positional argument when running `--phase=analyze`. The subprocess will be invoked as:
```
python -u Scripts/depo_prep.py --phase=analyze <config.json path>
```
That matches the existing `SubprocessWorker` argv assembly (one positional file path).

To make that work, the Settings page writes `config.json` into a temp folder under the case root before emitting `proceed_requested`, and the "file" passed to `TaskTab` is a single-element list `[config_json_path]`. The subprocess sees `--phase=analyze <config_path>` because `phase1_args=["--phase=analyze"]` is prepended.

Update `Scripts/depo_prep.py` `main()` to also accept this positional form:

```python
# In Scripts/depo_prep.py main(), after argparse:
# Tolerate the wizard's invocation: --phase=analyze <config_path>
# (SubprocessWorker appends the file path positionally).
if args.phase == "analyze" and not args.config and parser._actions:
    # Re-parse with unknown args allowed; use the first positional as config.
    pass
```

Rather than wedging argparse, switch to manual parsing in `depo_prep.py main()`. Replace `main()` with:

```python
def main():
    argv = sys.argv[1:]
    phase = None
    config = None
    session = None
    positional = []
    i = 0
    while i < len(argv):
        a = argv[i]
        if a.startswith("--phase="):
            phase = a.split("=", 1)[1]
            i += 1
        elif a == "--config" and i + 1 < len(argv):
            config = argv[i + 1]; i += 2
        elif a == "--session" and i + 1 < len(argv):
            session = argv[i + 1]; i += 2
        else:
            positional.append(a); i += 1

    if phase == "analyze":
        # Wizard form: --phase=analyze <config_path>
        if not config and positional:
            config = positional[0]
        if not config:
            print("ERROR: --config (or positional path) required for --phase=analyze", flush=True)
            return 2
        return _cmd_analyze(config)
    if phase == "generate":
        if not session and positional:
            session = positional[0]
        if not session:
            print("ERROR: --session (or positional path) required for --phase=generate", flush=True)
            return 2
        return _cmd_generate(session)
    print("ERROR: --phase=analyze|generate required", flush=True)
    return 2
```

Apply that change to `Scripts/depo_prep.py` before writing the Settings page tests.

- [ ] **Step 1: Update `Scripts/depo_prep.py main()`** with the manual-parsing version above.

- [ ] **Step 2: Write the Settings page failing test**

```python
# tests/test_wizard/test_depo_prep_settings_page.py
import json
from unittest.mock import MagicMock, patch

import pytest

pytest.importorskip("pytestqt")
pytest.importorskip("PySide6")


@pytest.fixture
def spec():
    from icharlotte_core.ui.wizard.registry import TaskSpec
    return TaskSpec(task_id="depo_prep", title="Depo Prep", description="d",
                    icon_glyph="?", script_name="depo_prep.py",
                    phase1_args=["--phase=analyze"], phase2_flag="--phase=generate")


def test_settings_page_renders_all_controls(qtbot, spec, tmp_path):
    from icharlotte_core.ui.wizard.pages.depo_prep_settings_page import DepoPrepSettingsPage
    page = DepoPrepSettingsPage(spec, files=[], case_root=str(tmp_path))
    qtbot.addWidget(page)
    assert page.deponent_name_combo is not None
    assert page.deponent_role_edit is not None
    assert page.add_deponent_files_btn is not None
    assert page.add_context_files_btn is not None
    assert page.style_combo is not None  # or radio_group
    assert page.free_text_edit is not None
    assert page.analyze_btn is not None
    assert page.flag_strategic.isChecked() is True   # default ON
    assert page.flag_source_facts.isChecked() is True  # default ON
    assert page.flag_impeachment.isChecked() is False
    assert page.flag_objection.isChecked() is False
    # Analyze button disabled until deponent + at least one source.
    assert page.analyze_btn.isEnabled() is False


def test_settings_page_analyze_writes_config_and_emits(qtbot, spec, tmp_path):
    from icharlotte_core.ui.wizard.pages.depo_prep_settings_page import DepoPrepSettingsPage
    page = DepoPrepSettingsPage(spec, files=[], case_root=str(tmp_path))
    qtbot.addWidget(page)

    page.set_deponent_name("Jane Doe")
    page.set_deponent_role("Plaintiff")
    page.add_deponent_files([str(tmp_path / "fake_depo.pdf")])
    # Fake-fake the existence check by creating the file.
    (tmp_path / "fake_depo.pdf").write_bytes(b"")
    assert page.analyze_btn.isEnabled() is True

    captured = []
    page.proceed_requested.connect(lambda d: captured.append(d))
    with qtbot.waitSignal(page.proceed_requested, timeout=1000):
        page._on_analyze_clicked()

    # Settings page persists config to disk and reports its path as the single "file".
    files = page.files
    assert len(files) == 1
    assert files[0].endswith("config.json")
    cfg = json.loads(open(files[0], "r", encoding="utf-8").read())
    assert cfg["deponent_name"] == "Jane Doe"
    assert cfg["deponent_role"] == "Plaintiff"
    assert cfg["style"] in ("discovery", "lockdown", "expert", "friendly")
    assert "case_root" in cfg


def test_settings_page_reveals_topic_editor_on_attach_phase1_complete(qtbot, spec, tmp_path):
    from icharlotte_core.ui.wizard.pages.depo_prep_settings_page import DepoPrepSettingsPage

    # Pre-create session+topics files to simulate Phase 1 output.
    session_dir = tmp_path / "NOTES" / "AI Output" / "Depo Prep - Jane - 2026-05-27 1432"
    session_dir.mkdir(parents=True)
    (session_dir / "session.json").write_text(json.dumps({
        "deponent_name": "Jane Doe", "topics_warning": None,
    }), encoding="utf-8")
    (session_dir / "topics.json").write_text(json.dumps({
        "topics": [{"id": "t01", "title": "T1", "strategic_note": "s",
                    "relevant_digest_refs": [], "default_checked": True, "lawyer_added": False}],
    }), encoding="utf-8")

    page = DepoPrepSettingsPage(spec, files=[], case_root=str(tmp_path))
    qtbot.addWidget(page)
    page._on_phase1_complete(str(session_dir / "session.json"))

    assert page.topic_editor.isVisible() is True
    assert len(page.topic_editor.get_topics()) == 1
    assert page.generate_btn.isEnabled() is True


def test_settings_page_generate_emits_phase2_requested(qtbot, spec, tmp_path):
    from icharlotte_core.ui.wizard.pages.depo_prep_settings_page import DepoPrepSettingsPage

    session_dir = tmp_path / "session"
    session_dir.mkdir()
    (session_dir / "session.json").write_text(json.dumps({"deponent_name": "X"}), encoding="utf-8")
    (session_dir / "topics.json").write_text(json.dumps({"topics": [
        {"id": "t01", "title": "T", "strategic_note": "", "relevant_digest_refs": [],
         "default_checked": True, "lawyer_added": False},
    ]}), encoding="utf-8")

    page = DepoPrepSettingsPage(spec, files=[], case_root=str(tmp_path))
    qtbot.addWidget(page)
    page._on_phase1_complete(str(session_dir / "session.json"))

    captured = []
    page.phase2_requested.connect(lambda p: captured.append(p))
    with qtbot.waitSignal(page.phase2_requested, timeout=1000):
        page._on_generate_clicked()

    # Topic editor state was persisted to topics.json before emit.
    assert captured[0] == str(session_dir / "session.json")
    saved_topics = json.loads((session_dir / "topics.json").read_text(encoding="utf-8"))
    assert isinstance(saved_topics["topics"], list)
```

- [ ] **Step 3: Write minimal implementation**

```python
# icharlotte_core/ui/wizard/pages/depo_prep_settings_page.py
"""Settings page for the Depo Prep wizard task."""
from __future__ import annotations

import json
import os
import tempfile
from pathlib import Path
from typing import List, Optional

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QCheckBox, QComboBox, QFileDialog, QGroupBox, QHBoxLayout, QLabel,
    QLineEdit, QListWidget, QListWidgetItem, QPlainTextEdit, QPushButton,
    QSizePolicy, QStackedWidget, QVBoxLayout, QWidget,
)

from ..registry import TaskSpec
from .depo_prep_topic_editor import TopicEditor
from .settings_page import SettingsPage


_STYLES = [
    ("discovery", "Discovery / Fact-gathering"),
    ("lockdown", "Lock-down (leading admissions)"),
    ("expert", "Expert challenge (Daubert-style)"),
    ("friendly", "Friendly (own client prep)"),
]


def _load_case_parties(case_root: str) -> List[str]:
    """Best-effort: return list of party names from CaseDataManager. Empty on any error."""
    try:
        from case_data_manager import CaseDataManager
    except ImportError:
        try:
            from Scripts.case_data_manager import CaseDataManager
        except ImportError:
            return []

    import re
    m = re.search(r"(\d{4})\D+(\d{3})", str(case_root or ""))
    if not m:
        return []
    file_number = f"{m.group(1)}.{m.group(2)}"

    try:
        cdm = CaseDataManager()
        parties = []
        for key in ("plaintiffs", "defendants"):
            v = cdm.get_value(file_number, key)
            if isinstance(v, list):
                parties.extend(str(x) for x in v)
            elif isinstance(v, str) and v.strip():
                parties.append(v.strip())
        return parties
    except Exception:
        return []


class DepoPrepSettingsPage(SettingsPage):
    """Custom settings page for Depo Prep.

    Reuses SettingsPage.proceed_requested(dict) for the "Analyze Sources" click —
    TaskTab's existing wiring runs the subprocess with --phase=analyze and the
    config.json path. Adds phase2_requested(str) for the "Generate Outline" click.
    """

    phase2_requested = Signal(str)

    def __init__(self, spec: TaskSpec, files, case_root: str | None = None, parent=None):
        super().__init__(spec, files=files, case_root=case_root, parent=parent)

        outer = self.layout()
        # Strip the base layout — we rebuild fully.
        while outer.count():
            item = outer.takeAt(0)
            if item.widget():
                item.widget().setParent(None)

        outer.setContentsMargins(24, 24, 24, 24)
        outer.setSpacing(10)

        # ---- Deponent ----
        deponent_box = QGroupBox("Deponent")
        dep_layout = QVBoxLayout(deponent_box)

        row1 = QHBoxLayout()
        row1.addWidget(QLabel("Name:"))
        self.deponent_name_combo = QComboBox()
        self.deponent_name_combo.setEditable(True)
        for name in _load_case_parties(case_root or ""):
            self.deponent_name_combo.addItem(name)
        self.deponent_name_combo.currentTextChanged.connect(self._refresh_buttons)
        row1.addWidget(self.deponent_name_combo, 1)
        dep_layout.addLayout(row1)

        row2 = QHBoxLayout()
        row2.addWidget(QLabel("Role:"))
        self.deponent_role_edit = QLineEdit()
        self.deponent_role_edit.setPlaceholderText("e.g., Plaintiff's treating orthopedist")
        row2.addWidget(self.deponent_role_edit, 1)
        dep_layout.addLayout(row2)

        outer.addWidget(deponent_box)

        # ---- Sources ----
        sources_box = QGroupBox("Source files")
        s_layout = QVBoxLayout(sources_box)

        # Deponent materials section
        s_layout.addWidget(QLabel("Deponent's own materials:"))
        dep_files_row = QHBoxLayout()
        self.add_deponent_files_btn = QPushButton("+ Add files")
        self.add_deponent_files_btn.clicked.connect(self._on_add_deponent_files)
        dep_files_row.addWidget(self.add_deponent_files_btn)
        self.remove_deponent_files_btn = QPushButton("Remove selected")
        self.remove_deponent_files_btn.clicked.connect(
            lambda: self._remove_selected(self.deponent_files_list, self._deponent_files))
        dep_files_row.addWidget(self.remove_deponent_files_btn)
        dep_files_row.addStretch()
        s_layout.addLayout(dep_files_row)
        self.deponent_files_list = QListWidget()
        self.deponent_files_list.setSelectionMode(QListWidget.SelectionMode.ExtendedSelection)
        self.deponent_files_list.setMaximumHeight(80)
        s_layout.addWidget(self.deponent_files_list)
        self._deponent_files: List[str] = []

        # Context section
        s_layout.addWidget(QLabel("Case context:"))
        ctx_files_row = QHBoxLayout()
        self.add_context_files_btn = QPushButton("+ Add files")
        self.add_context_files_btn.clicked.connect(self._on_add_context_files)
        ctx_files_row.addWidget(self.add_context_files_btn)
        self.remove_context_files_btn = QPushButton("Remove selected")
        self.remove_context_files_btn.clicked.connect(
            lambda: self._remove_selected(self.context_files_list, self._context_files))
        ctx_files_row.addWidget(self.remove_context_files_btn)
        ctx_files_row.addStretch()
        s_layout.addLayout(ctx_files_row)
        self.context_files_list = QListWidget()
        self.context_files_list.setSelectionMode(QListWidget.SelectionMode.ExtendedSelection)
        self.context_files_list.setMaximumHeight(80)
        s_layout.addWidget(self.context_files_list)
        self._context_files: List[str] = []

        outer.addWidget(sources_box)

        # ---- Instructions ----
        instr_box = QGroupBox("Instructions")
        i_layout = QVBoxLayout(instr_box)
        style_row = QHBoxLayout()
        style_row.addWidget(QLabel("Style:"))
        self.style_combo = QComboBox()
        for key, label in _STYLES:
            self.style_combo.addItem(label, key)
        style_row.addWidget(self.style_combo, 1)
        i_layout.addLayout(style_row)
        i_layout.addWidget(QLabel("Free-text strategy notes:"))
        self.free_text_edit = QPlainTextEdit()
        self.free_text_edit.setPlaceholderText(
            "Case theory, topics to emphasize, key admissions to extract, things to avoid…")
        self.free_text_edit.setMinimumHeight(80)
        i_layout.addWidget(self.free_text_edit)
        outer.addWidget(instr_box)

        # ---- Per-topic content flags ----
        flags_box = QGroupBox("Per-topic content")
        f_layout = QVBoxLayout(flags_box)
        self.flag_strategic = QCheckBox("Strategic note (\"why this topic\")")
        self.flag_strategic.setChecked(True)
        self.flag_source_facts = QCheckBox("Key source facts (with citations)")
        self.flag_source_facts.setChecked(True)
        self.flag_impeachment = QCheckBox("Impeachment hooks / inconsistencies")
        self.flag_objection = QCheckBox("Anticipated objections + workaround phrasings")
        for cb in (self.flag_strategic, self.flag_source_facts,
                   self.flag_impeachment, self.flag_objection):
            f_layout.addWidget(cb)
        outer.addWidget(flags_box)

        # ---- Action row + topic editor area ----
        action_row = QHBoxLayout()
        action_row.addStretch()
        self.analyze_btn = QPushButton("Analyze Sources")
        self.analyze_btn.setStyleSheet(
            "background-color: #1976D2; color: white; font-weight: 600; padding: 8px 24px;")
        self.analyze_btn.clicked.connect(self._on_analyze_clicked)
        action_row.addWidget(self.analyze_btn)
        outer.addLayout(action_row)

        # Topic editor area (hidden until Phase 1 completes).
        self._phase1_status_label = QLabel("")
        self._phase1_status_label.setStyleSheet("color: #555; font-style: italic;")
        self._phase1_status_label.setVisible(False)
        outer.addWidget(self._phase1_status_label)

        self.topic_editor = TopicEditor()
        self.topic_editor.setVisible(False)
        outer.addWidget(self.topic_editor, 1)

        generate_row = QHBoxLayout()
        generate_row.addStretch()
        self.generate_btn = QPushButton("Generate Outline")
        self.generate_btn.setStyleSheet(
            "background-color: #43A047; color: white; font-weight: 600; padding: 8px 24px;")
        self.generate_btn.setEnabled(False)
        self.generate_btn.setVisible(False)
        self.generate_btn.clicked.connect(self._on_generate_clicked)
        generate_row.addWidget(self.generate_btn)
        outer.addLayout(generate_row)

        self._session_path: Optional[str] = None
        self._refresh_buttons()

    # ---- Compatibility with base class ----
    def _refresh_files_list(self) -> None:
        # Base SettingsPage calls this from __init__; ours does nothing because
        # we manage two separate lists.
        return

    @property
    def files(self) -> list[str]:
        # TaskTab uses self.files as the positional argv passed to the subprocess.
        # For Depo Prep, that's the config.json path after _on_analyze_clicked.
        return list(self._files)

    # ---- Public setters used in tests ----
    def set_deponent_name(self, name: str) -> None:
        self.deponent_name_combo.setEditText(name)

    def set_deponent_role(self, role: str) -> None:
        self.deponent_role_edit.setText(role)

    def add_deponent_files(self, paths: List[str]) -> None:
        for p in paths:
            if p not in self._deponent_files:
                self._deponent_files.append(p)
                self.deponent_files_list.addItem(QListWidgetItem(os.path.basename(p)))
        self._refresh_buttons()

    def add_context_files(self, paths: List[str]) -> None:
        for p in paths:
            if p not in self._context_files:
                self._context_files.append(p)
                self.context_files_list.addItem(QListWidgetItem(os.path.basename(p)))
        self._refresh_buttons()

    # ---- Internals ----
    def _on_add_deponent_files(self) -> None:
        paths, _ = QFileDialog.getOpenFileNames(
            self, "Add deponent's own materials", self._case_root or "", "All files (*.*)")
        if paths:
            self.add_deponent_files(paths)

    def _on_add_context_files(self) -> None:
        paths, _ = QFileDialog.getOpenFileNames(
            self, "Add case context", self._case_root or "", "All files (*.*)")
        if paths:
            self.add_context_files(paths)

    def _remove_selected(self, list_widget: QListWidget, store: List[str]) -> None:
        rows = sorted({i.row() for i in list_widget.selectedIndexes()}, reverse=True)
        for r in rows:
            if 0 <= r < len(store):
                store.pop(r)
                list_widget.takeItem(r)
        self._refresh_buttons()

    def _refresh_buttons(self) -> None:
        has_name = bool(self.deponent_name_combo.currentText().strip())
        has_sources = bool(self._deponent_files or self._context_files)
        self.analyze_btn.setEnabled(has_name and has_sources)

    def _build_config_dict(self) -> dict:
        style = self.style_combo.currentData() or "discovery"
        return {
            "deponent_name": self.deponent_name_combo.currentText().strip(),
            "deponent_role": self.deponent_role_edit.text().strip(),
            "deponent_sources": list(self._deponent_files),
            "context_sources": list(self._context_files),
            "style": style,
            "free_text_notes": self.free_text_edit.toPlainText().strip(),
            "per_topic_flags": {
                "strategic_note": self.flag_strategic.isChecked(),
                "source_facts": self.flag_source_facts.isChecked(),
                "impeachment_hook": self.flag_impeachment.isChecked(),
                "objection_alts": self.flag_objection.isChecked(),
            },
            "case_root": self._case_root or "",
        }

    def _on_analyze_clicked(self) -> None:
        cfg = self._build_config_dict()
        # Persist to a temp config.json — the subprocess reads it.
        tmpdir = tempfile.mkdtemp(prefix="depo_prep_config_")
        cfg_path = Path(tmpdir) / "config.json"
        cfg_path.write_text(json.dumps(cfg, indent=2), encoding="utf-8")
        # Replace _files so TaskTab uses cfg_path as the positional argv.
        self._files = [str(cfg_path)]
        self._phase1_status_label.setText("Analyzing sources…")
        self._phase1_status_label.setVisible(True)
        self.analyze_btn.setEnabled(False)
        self.proceed_requested.emit({})

    def to_dict(self) -> dict:
        return self._build_config_dict()

    # ---- Phase 1 completion hook ----
    def attach_worker(self, worker) -> bool:
        worker.status.connect(self._phase1_status_label.setText)
        worker.awaiting_input.connect(self._on_phase1_complete)
        worker.failed.connect(self._on_phase1_failed)
        return True

    def _on_phase1_complete(self, session_path: str) -> None:
        from .depo_prep_topic_editor import TopicEditor  # noqa
        self._session_path = session_path
        session_dir = Path(session_path).parent
        try:
            topics_payload = json.loads(
                (session_dir / "topics.json").read_text(encoding="utf-8"))
        except Exception as e:
            self._on_phase1_failed(f"Could not load topics.json: {e}")
            return
        self.topic_editor.set_topics(topics_payload.get("topics", []))
        self.topic_editor.setVisible(True)
        self.generate_btn.setVisible(True)
        self.generate_btn.setEnabled(True)
        warning = topics_payload.get("warning")
        if warning:
            self._phase1_status_label.setText(warning)
            self._phase1_status_label.setStyleSheet("color: #E65100; font-style: italic;")
        else:
            self._phase1_status_label.setVisible(False)

    def _on_phase1_failed(self, err: str) -> None:
        self._phase1_status_label.setText(f"Analysis failed: {err}")
        self._phase1_status_label.setStyleSheet("color: #C62828; font-style: italic;")
        self.analyze_btn.setEnabled(True)

    def _on_generate_clicked(self) -> None:
        if not self._session_path:
            return
        # Persist current topic editor state back to topics.json.
        topics = self.topic_editor.get_topics()
        topics_json_path = Path(self._session_path).parent / "topics.json"
        topics_json_path.write_text(
            json.dumps({"topics": topics}, indent=2), encoding="utf-8")
        self.phase2_requested.emit(self._session_path)
```

- [ ] **Step 4: Run tests**

Run: `python -m pytest tests/test_wizard/test_depo_prep_settings_page.py -v`
Expected: 4 passed.

- [ ] **Step 5: Commit**

```bash
git add icharlotte_core/ui/wizard/pages/depo_prep_settings_page.py Scripts/depo_prep.py tests/test_wizard/test_depo_prep_settings_page.py
git commit -m "feat(depo_prep): DepoPrepSettingsPage + embed TopicEditor + manual CLI arg parsing"
```

---

### Task 15: DepoPrepOutputPage

**Files:**
- Create: `icharlotte_core/ui/wizard/pages/depo_prep_output_page.py`
- Create: `tests/test_wizard/test_depo_prep_output_page.py`

The output page extends `OutputPage` and adds a collapsible markdown view side-by-side with the .docx editor. Easiest implementation: read `outline.md` from `<session_dir>/outline.md` (alongside `outline.docx` which the base class loads via `_render_path`), put a `QTextBrowser` with the markdown rendered above the standard editor, and add Copy-Question / Jump-to-source buttons.

For scope control: this task delivers ONLY the markdown viewer + Copy All shortcut for individual questions. The "Jump to source PDF" feature is deferred (just include the citation text inline — clickable jump is a future enhancement).

- [ ] **Step 1: Write the failing test**

```python
# tests/test_wizard/test_depo_prep_output_page.py
from pathlib import Path

import pytest

pytest.importorskip("pytestqt")
pytest.importorskip("PySide6")


def test_output_page_loads_md_alongside_docx(qtbot, tmp_path):
    from icharlotte_core.ui.wizard.pages.depo_prep_output_page import DepoPrepOutputPage
    from Scripts.depo_prep_lib.render_docx import render_outline_docx
    from Scripts.depo_prep_lib.render_md import render_outline_md

    outline = {"deponent_name": "Jane Doe", "deponent_role": "Plaintiff",
               "topics": [{"topic_id": "t01", "title": "T1", "strategic_note": "s",
                            "questions": [{"n": 1, "text": "Q1"}]}],
               "coverage_gaps": []}
    docx = tmp_path / "outline.docx"
    md = tmp_path / "outline.md"
    render_outline_docx(outline=outline, output_path=docx)
    render_outline_md(outline=outline, output_path=md)

    page = DepoPrepOutputPage()
    qtbot.addWidget(page)
    page.load_output(str(docx))
    md_text = page.md_viewer.toPlainText()
    assert "Jane Doe" in md_text
    assert "T1" in md_text


def test_output_page_falls_back_to_docx_only_when_md_missing(qtbot, tmp_path):
    from icharlotte_core.ui.wizard.pages.depo_prep_output_page import DepoPrepOutputPage
    from Scripts.depo_prep_lib.render_docx import render_outline_docx
    outline = {"deponent_name": "X", "deponent_role": "",
               "topics": [{"topic_id": "t01", "title": "T", "strategic_note": "",
                            "questions": [{"n": 1, "text": "Q"}]}], "coverage_gaps": []}
    docx = tmp_path / "outline.docx"
    render_outline_docx(outline=outline, output_path=docx)
    # No outline.md alongside.
    page = DepoPrepOutputPage()
    qtbot.addWidget(page)
    page.load_output(str(docx))
    # Doesn't crash; md_viewer empty or hidden.
    assert page.md_viewer is not None
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_wizard/test_depo_prep_output_page.py -v`
Expected: ImportError.

- [ ] **Step 3: Write minimal implementation**

```python
# icharlotte_core/ui/wizard/pages/depo_prep_output_page.py
"""Custom output page for Depo Prep — adds a markdown view above the .docx editor."""
from __future__ import annotations

from pathlib import Path

from PySide6.QtWidgets import QSplitter, QTextBrowser, QWidget
from PySide6.QtCore import Qt

from .output_page import OutputPage


class DepoPrepOutputPage(OutputPage):
    def __init__(self, parent: QWidget | None = None):
        super().__init__(parent)

        # Insert a QTextBrowser ABOVE the editor by repacking via a splitter.
        outer = self.layout()
        # Find the editor's index; pop it and the splitter that holds it.
        # Simpler: stash a reference, create a splitter, swap the editor into it.

        self.md_viewer = QTextBrowser()
        self.md_viewer.setOpenExternalLinks(True)

        splitter = QSplitter(Qt.Orientation.Vertical)
        # Move the existing editor under the splitter; keep header/buttons in place.
        self.editor.setParent(splitter)
        splitter.addWidget(self.md_viewer)
        splitter.addWidget(self.editor)
        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 1)

        # Insert the splitter where the editor used to sit. We know from
        # OutputPage that the editor was added with addWidget(self.editor, 1)
        # immediately after the header layout. Walk the layout to find that
        # exact position.
        editor_idx = None
        for i in range(outer.count()):
            item = outer.itemAt(i)
            if item is not None and item.widget() is self.editor:
                editor_idx = i
                break
        if editor_idx is not None:
            outer.takeAt(editor_idx)
        outer.insertWidget(editor_idx if editor_idx is not None else 0, splitter, 1)

    def _render_path(self, output_path: str) -> None:
        # Render docx via base class behaviour.
        super()._render_path(output_path)
        md_path = Path(output_path).with_suffix(".md")
        if md_path.exists():
            try:
                self.md_viewer.setMarkdown(md_path.read_text(encoding="utf-8"))
            except Exception:
                self.md_viewer.setPlainText(md_path.read_text(encoding="utf-8"))
        else:
            self.md_viewer.clear()
```

- [ ] **Step 4: Run tests**

Run: `python -m pytest tests/test_wizard/test_depo_prep_output_page.py -v`
Expected: 2 passed.

- [ ] **Step 5: Commit**

```bash
git add icharlotte_core/ui/wizard/pages/depo_prep_output_page.py tests/test_wizard/test_depo_prep_output_page.py
git commit -m "feat(depo_prep): DepoPrepOutputPage — markdown viewer above .docx editor"
```

---

### Task 16: Register the task in `registry.py`

**Files:**
- Modify: `icharlotte_core/ui/wizard/registry.py`
- Create: `tests/test_wizard/test_depo_prep_registry.py`

- [ ] **Step 1: Write the failing test**

```python
# tests/test_wizard/test_depo_prep_registry.py
from icharlotte_core.ui.wizard.registry import TASK_REGISTRY, get_task


def test_depo_prep_task_registered():
    spec = get_task("depo_prep")
    assert spec.title == "Depo Prep"
    assert spec.script_name == "depo_prep.py"
    assert "--phase=analyze" in spec.phase1_args
    assert spec.phase2_flag == "--phase=generate"


def test_depo_prep_settings_page_cls():
    from icharlotte_core.ui.wizard.pages.depo_prep_settings_page import DepoPrepSettingsPage
    spec = get_task("depo_prep")
    assert spec.settings_page_cls is DepoPrepSettingsPage


def test_depo_prep_output_page_cls():
    from icharlotte_core.ui.wizard.pages.depo_prep_output_page import DepoPrepOutputPage
    spec = get_task("depo_prep")
    assert spec.output_page_cls is DepoPrepOutputPage
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_wizard/test_depo_prep_registry.py -v`
Expected: KeyError on `depo_prep`.

- [ ] **Step 3: Modify `icharlotte_core/ui/wizard/registry.py`**

Add factory helpers near the others:

```python
def _depo_prep_settings_page_cls():
    from .pages.depo_prep_settings_page import DepoPrepSettingsPage
    return DepoPrepSettingsPage


def _depo_prep_output_page_cls():
    from .pages.depo_prep_output_page import DepoPrepOutputPage
    return DepoPrepOutputPage
```

Add the entry inside `TASK_REGISTRY`:

```python
    "depo_prep": TaskSpec(
        task_id="depo_prep",
        title="Depo Prep",
        description="Generate a deposition outline with questions grounded in case sources.",
        icon_glyph="❓",  # ❓
        script_name="depo_prep.py",
        default_folders=["DISCOVERY", "PLEADINGS", "RECORDS"],
        phase1_args=["--phase=analyze"],
        phase2_flag="--phase=generate",
        _settings_page_cls_factory=_depo_prep_settings_page_cls,
        _output_page_cls_factory=_depo_prep_output_page_cls,
    ),
```

- [ ] **Step 4: Run tests**

Run: `python -m pytest tests/test_wizard/test_depo_prep_registry.py -v`
Expected: 3 passed.

- [ ] **Step 5: Commit**

```bash
git add icharlotte_core/ui/wizard/registry.py tests/test_wizard/test_depo_prep_registry.py
git commit -m "feat(depo_prep): register task in wizard registry"
```

---

## Block C — LLM config + integration

### Task 17: Add `DepoPrep` agent to `llm_preferences.json`

**Files:**
- Modify: `config/llm_preferences.json`
- Create: `tests/test_wizard/test_depo_prep_llm_config.py`

- [ ] **Step 1: Read the existing config**

```bash
python -c "import json; print(json.dumps(json.load(open('config/llm_preferences.json')), indent=2)[:1500])"
```

Examine the structure — agents are typically keyed by `agent_id`.

- [ ] **Step 2: Write the failing test**

```python
# tests/test_wizard/test_depo_prep_llm_config.py
import json
from pathlib import Path


def test_depo_prep_agent_in_llm_preferences():
    cfg = json.loads(Path("config/llm_preferences.json").read_text(encoding="utf-8"))
    # The shape may be {agents: {DepoPrep: {...}}} or {DepoPrep: {...}} — accept either.
    agents = cfg.get("agents", cfg)
    assert "DepoPrep" in agents, f"DepoPrep not in agents (keys: {list(agents)[:10]})"
    depo = agents["DepoPrep"]
    task_configs = depo.get("task_configs", depo)
    assert "general" in task_configs
    assert "extraction" in task_configs
```

- [ ] **Step 3: Run test to verify it fails**

Run: `python -m pytest tests/test_wizard/test_depo_prep_llm_config.py -v`
Expected: fails on missing `DepoPrep` key.

- [ ] **Step 4: Modify `config/llm_preferences.json`**

Mirror the structure of an existing agent (look at `agent_sum_depo`). Add:

```json
"DepoPrep": {
  "task_configs": {
    "general": {
      "primary_model": "gemini-2.5-pro",
      "fallback_sequence": ["gemini-2.5-pro", "claude-opus-4-7", "gpt-4o"]
    },
    "extraction": {
      "primary_model": "gemini-2.5-flash",
      "fallback_sequence": ["gemini-2.5-flash", "claude-haiku-4-5"]
    }
  }
}
```

Adjust to match the actual file shape — if the file nests agents under `agents:`, place there; if it's flat, place at the root. The model names may also be conventional aliases — keep consistent with `agent_sum_depo`'s entry. If any model name doesn't exist in the file, replace with the closest analog already used by `agent_sum_depo`.

- [ ] **Step 5: Run test**

Run: `python -m pytest tests/test_wizard/test_depo_prep_llm_config.py -v`
Expected: 1 passed.

- [ ] **Step 6: Commit**

```bash
git add config/llm_preferences.json tests/test_wizard/test_depo_prep_llm_config.py
git commit -m "feat(depo_prep): add DepoPrep agent to llm_preferences"
```

---

### Task 18: End-to-end integration test

**Files:**
- Create: `tests/test_wizard/test_depo_prep_integration.py`
- Create: `tests/test_wizard/fixtures/depo_prep_sources/` (a tiny depo .txt + a tiny pleading .txt)

This test exercises the full pipeline by calling `phase1.run_phase1` then `phase2.run_phase2` directly (no subprocess) with a stub LLM caller that returns canned responses keyed by `pass_name`. It produces real `outline.docx` and `outline.md` files and asserts their content.

- [ ] **Step 1: Create fixtures**

```python
# tests/test_wizard/fixtures/depo_prep_sources/__init__.py  (empty marker)
```

Create two small text files used as sources:

```
tests/test_wizard/fixtures/depo_prep_sources/jane_doe_depo_excerpt.txt:
  Q. Before the August 2024 collision, had you ever had back pain?
  A. No. I had no prior back problems.
  Q. What about chiropractic care?
  A. Never.

tests/test_wizard/fixtures/depo_prep_sources/complaint_excerpt.txt:
  Plaintiff Jane Doe alleges injuries to her cervical and lumbar spine
  proximately caused by the August 15, 2024 collision.
```

- [ ] **Step 2: Write the integration test**

```python
# tests/test_wizard/test_depo_prep_integration.py
"""End-to-end Depo Prep integration test (no subprocess)."""
import json
from pathlib import Path
from unittest.mock import MagicMock

import pytest

FIXTURES = Path(__file__).parent / "fixtures" / "depo_prep_sources"


def _stub_llm():
    """Stub LLM keyed by pass_name to return canned, well-shaped payloads."""
    caller = MagicMock()

    def call(prompt, text, task_type="general", agent_id=None, pass_name=None, **kw):
        if pass_name == "source_digest":
            return json.dumps({
                "source_id": "placeholder",  # orchestrator overwrites
                "source_kind": "deposition_transcript",
                "deponent_statements": [
                    {"text": "I had no prior back problems.", "location": "p.1:3",
                     "context": "Direct exam"}
                ],
                "factual_anchors": [
                    {"fact": "Denied chiropractic care", "location": "p.1:5",
                     "topic_tags": ["prior_care"]}
                ],
                "inconsistencies": [],
                "summary": "Witness denies any prior back issues.",
            })
        if pass_name == "topic_clustering":
            return json.dumps({"topics": [
                {"id": "t01", "title": "Prior back issues",
                 "strategic_note": "Establish lack of prior issues for causation.",
                 "relevant_digest_refs": [],
                 "default_checked": True, "lawyer_added": False},
                {"id": "t02", "title": "Mechanism of collision",
                 "strategic_note": "Pin down what plaintiff observed.",
                 "relevant_digest_refs": [],
                 "default_checked": True, "lawyer_added": False},
                {"id": "t03", "title": "Treatment timeline",
                 "strategic_note": "Walk through care, gaps.",
                 "relevant_digest_refs": [],
                 "default_checked": True, "lawyer_added": False},
            ]})
        if pass_name == "topic_questions":
            return json.dumps({"topic_id": "tXX", "questions": [
                {"n": 1, "text": "Before August 15, 2024, did you experience any back pain?"},
                {"n": 2, "text": "Did you ever see a chiropractor before that date?"},
            ]})
        if pass_name == "dedup":
            return json.dumps({"duplicates": [], "coverage_gaps": [
                "No question addresses prior auto accidents."
            ], "renumber_after_dedup": True})
        if pass_name == "polish":
            return text  # echo unchanged
        return ""

    caller.call.side_effect = call
    return caller


def test_full_phase1_phase2_pipeline(tmp_path):
    from Scripts.depo_prep_lib import phase1, phase2

    case_root = tmp_path / "Smith v. Jones"
    case_root.mkdir()

    config = {
        "deponent_name": "Jane Doe", "deponent_role": "Plaintiff",
        "deponent_sources": [str(FIXTURES / "jane_doe_depo_excerpt.txt")],
        "context_sources": [str(FIXTURES / "complaint_excerpt.txt")],
        "style": "lockdown", "free_text_notes": "Focus on causation.",
        "per_topic_flags": {
            "strategic_note": True, "source_facts": True,
            "impeachment_hook": False, "objection_alts": False,
        },
        "case_root": str(case_root),
    }

    llm = _stub_llm()

    # Phase 1
    session_path = phase1.run_phase1(
        config=config, llm_caller=llm, progress=lambda *a, **k: None)
    assert Path(session_path).exists()
    session_dir = Path(session_path).parent
    topics = json.loads((session_dir / "topics.json").read_text(encoding="utf-8"))
    assert len(topics["topics"]) == 3

    # Simulate user editing topics: uncheck one, add a custom topic.
    topics["topics"][1]["default_checked"] = False
    topics["topics"].append({
        "id": "t99", "title": "Social media activity post-accident",
        "strategic_note": "Lawyer-added. Look for inconsistent depictions.",
        "relevant_digest_refs": [], "default_checked": True, "lawyer_added": True,
    })
    (session_dir / "topics.json").write_text(json.dumps(topics), encoding="utf-8")

    # Phase 2
    phase2.run_phase2(session_path=session_path, llm_caller=llm,
                      progress=lambda *a, **k: None)

    docx = session_dir / "outline.docx"
    md = session_dir / "outline.md"
    assert docx.exists()
    assert md.exists()

    md_text = md.read_text(encoding="utf-8")
    assert "Jane Doe" in md_text
    assert "Prior back issues" in md_text
    # The unchecked topic (Mechanism of collision) should NOT appear.
    assert "Mechanism of collision" not in md_text
    # Lawyer-added topic SHOULD appear.
    assert "Social media activity post-accident" in md_text
    # Coverage gap appears.
    assert "prior auto accidents" in md_text


def test_full_pipeline_handles_per_topic_failure(tmp_path):
    """If one topic's LLM call fails, the rest of the outline still renders."""
    from Scripts.depo_prep_lib import phase1, phase2

    case_root = tmp_path / "C"
    case_root.mkdir()
    config = {
        "deponent_name": "X", "deponent_role": "P",
        "deponent_sources": [str(FIXTURES / "jane_doe_depo_excerpt.txt")],
        "context_sources": [],
        "style": "discovery", "free_text_notes": "",
        "per_topic_flags": {"strategic_note": False, "source_facts": False,
                             "impeachment_hook": False, "objection_alts": False},
        "case_root": str(case_root),
    }

    # LLM: fail every other topic_questions call.
    call_count = {"topic_questions": 0}
    def call(prompt, text, task_type="general", agent_id=None, pass_name=None, **kw):
        if pass_name == "source_digest":
            return json.dumps({"source_id": "x", "source_kind": "other",
                               "deponent_statements": [], "factual_anchors": [],
                               "inconsistencies": [], "summary": ""})
        if pass_name == "topic_clustering":
            return json.dumps({"topics": [
                {"id": f"t{i:02d}", "title": f"T{i}", "strategic_note": "s",
                 "relevant_digest_refs": [], "default_checked": True, "lawyer_added": False}
                for i in range(1, 4)
            ]})
        if pass_name == "topic_questions":
            call_count["topic_questions"] += 1
            if call_count["topic_questions"] % 2 == 0:
                raise RuntimeError("rate limit")
            return json.dumps({"topic_id": "tXX",
                                "questions": [{"n": 1, "text": "Question"}]})
        if pass_name == "dedup":
            return json.dumps({"duplicates": [], "coverage_gaps": [], "renumber_after_dedup": False})
        if pass_name == "polish":
            return text
        return ""

    caller = MagicMock()
    caller.call.side_effect = call

    session_path = phase1.run_phase1(config=config, llm_caller=caller,
                                      progress=lambda *a, **k: None)
    phase2.run_phase2(session_path=session_path, llm_caller=caller,
                      progress=lambda *a, **k: None)

    session_dir = Path(session_path).parent
    md = (session_dir / "outline.md").read_text(encoding="utf-8")
    # All three topics should appear; some with questions, some without.
    assert "T1" in md
    assert "T2" in md
    assert "T3" in md
```

- [ ] **Step 3: Run tests**

Run: `python -m pytest tests/test_wizard/test_depo_prep_integration.py -v`
Expected: 2 passed.

- [ ] **Step 4: Commit**

```bash
git add tests/test_wizard/test_depo_prep_integration.py tests/test_wizard/fixtures/depo_prep_sources/
git commit -m "test(depo_prep): end-to-end integration with stub LLM"
```

---

### Wave 3 self-check

Final verification before declaring the feature done:

- [ ] Full test pass: `python -m pytest tests/test_wizard/ -v` → all green
- [ ] No regressions in adjacent tests: `python -m pytest tests/ -v -x --ignore=tests/test_wizard` → still passes (excluding flaky pre-existing tests noted in `MEMORY.md` such as `use_all_text_check`)
- [ ] All 18 commits on branch; rebase onto main if desired
- [ ] Launch iCharlotte locally, open a case, switch to wizard mode, confirm:
  - Depo Prep card appears with the ❓ icon
  - Clicking it opens the new settings page
  - File pickers, deponent dropdown, style, free-text, content flags all work
  - "Analyze Sources" disabled until name + at least one source
  - Clicking "Analyze Sources" runs Phase 1 (will hit real LLM — requires API keys)
  - Topic editor appears; drag-reorder works; add/remove works
  - "Generate Outline" runs Phase 2 and lands on the custom output page with the markdown viewer above the .docx editor

- [ ] Run `word_validator.validate_after_edit` on a produced `outline.docx` and confirm no errors (manual check via Python REPL).

Apply the resulting code to `C:\geminiterminal2\` (the live iCharlotte checkout) after the worktree is merged. iCharlotte must be restarted after the changes land.
