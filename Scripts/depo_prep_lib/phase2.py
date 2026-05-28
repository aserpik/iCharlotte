"""Phase 2 orchestrator - Stage A (parallel) -> B (dedup) -> C (polish) -> D (render)."""
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

    progress(5, f"Generating questions for {len(topics)} topic(s)...")
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

    progress(70, "Dedup + coverage check...")
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

    progress(85, "Polish pass...")
    outline = polish_outline(outline=outline, llm_caller=llm_caller)

    progress(95, "Rendering...")
    render_outline_docx(outline=outline, output_path=session_dir / "outline.docx")
    render_outline_md(outline=outline, output_path=session_dir / "outline.md")
    progress(100, "Done")
