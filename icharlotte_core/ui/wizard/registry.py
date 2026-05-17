"""Task registry — single source of truth for Wizard Mode task cards.

Each task contributes:
  - task_id            : stable identifier used in persistence and code paths
  - title              : human-readable card title
  - description        : one-line card description
  - icon_glyph         : single emoji-ish character used as the card icon
  - default_folders    : ordered list of relative subfolders (under case root)
                         tried in order when opening the pre-Settings file
                         dialog; empty list means default to case root.
"""
from dataclasses import dataclass, field
from typing import List


@dataclass(frozen=True)
class TaskSpec:
    task_id: str
    title: str
    description: str
    icon_glyph: str
    script_name: str
    default_folders: List[str] = field(default_factory=list)


TASK_REGISTRY: dict[str, TaskSpec] = {
    "summarize_documents": TaskSpec(
        task_id="summarize_documents",
        title="Summarize Documents",
        description="Produce a concise summary of one or more case documents.",
        icon_glyph="\U0001F4C4",  # 📄
        script_name="summarize.py",
        default_folders=[],
    ),
    "summarize_discovery": TaskSpec(
        task_id="summarize_discovery",
        title="Summarize Discovery",
        description="Summarize discovery responses with structure and citations.",
        icon_glyph="\U0001F4CB",  # 📋
        script_name="summarize_discovery.py",
        default_folders=["DISCOVERY/RESPONSES", "DISCOVERY"],
    ),
    "summarize_depositions": TaskSpec(
        task_id="summarize_depositions",
        title="Summarize Depositions",
        description="Generate a structured summary of one or more depositions.",
        icon_glyph="\U0001F399",  # 🎙
        script_name="summarize_deposition.py",
        default_folders=["DISCOVERY/TRANSCRIPTS", "DISCOVERY"],
    ),
    "medical_records": TaskSpec(
        task_id="medical_records",
        title="Medical Records Review",
        description="Extract and summarize medical records into a chronology.",
        icon_glyph="\U0001F3E5",  # 🏥
        script_name="med_record.py",
        default_folders=["RECORDS"],
    ),
}


def get_task(task_id: str) -> TaskSpec:
    """Return the TaskSpec for `task_id`. Raises KeyError if unknown."""
    return TASK_REGISTRY[task_id]


def list_tasks() -> list[TaskSpec]:
    """Return all registered tasks in registry-insertion order."""
    return list(TASK_REGISTRY.values())
