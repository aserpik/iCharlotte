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
    id: str
    title: str
    description: str
    uses_tables: bool
    prompt_file: str
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
