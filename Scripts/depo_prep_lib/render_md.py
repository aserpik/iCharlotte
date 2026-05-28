"""Stage D - render outline.md."""
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
