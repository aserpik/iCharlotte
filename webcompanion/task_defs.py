# -*- coding: utf-8 -*-
"""Web-companion task catalogue + script/session bridging.

Mirrors the script-based subset of icharlotte_core/ui/wizard/registry.py.
Session edits reuse the SAME session_manager modules the desktop forms use,
so semantics stay identical.
"""
import json
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import List

REPO_ROOT = Path(__file__).resolve().parent.parent

PDF = (".pdf",)
DOCS = (".pdf", ".docx", ".doc", ".txt")


@dataclass(frozen=True)
class TaskDef:
    task_id: str
    title: str
    glyph: str
    script_name: str
    description: str
    default_folders: tuple = ()
    file_exts: tuple = PDF
    two_phase: bool = False
    phase1_args: tuple = ()
    phase2_flag: str = "--phase=summary"
    awaiting_kind: str = ""    # '' | 'deposition' | 'med_chron' | 'depo_prep'
    pre_settings: str = ""     # '' | 'depo_prep'


TASKS = {
    "summarize_documents": TaskDef(
        task_id="summarize_documents", title="Summarize Documents", glyph="\U0001F4C4",
        script_name="summarize.py",
        description="Concise summary of one or more case documents.",
        file_exts=DOCS),
    "summarize_discovery": TaskDef(
        task_id="summarize_discovery", title="Summarize Discovery", glyph="\U0001F4CB",
        script_name="summarize_discovery.py",
        description="Summarize discovery responses with structure and citations.",
        default_folders=("DISCOVERY/RESPONSES", "DISCOVERY"), file_exts=DOCS),
    "summarize_depositions": TaskDef(
        task_id="summarize_depositions", title="Summarize Depositions", glyph="\U0001F399",
        script_name="summarize_deposition.py",
        description="Structured deposition summary (you pick topics mid-run).",
        default_folders=("DISCOVERY/TRANSCRIPTS", "DISCOVERY"), file_exts=DOCS,
        two_phase=True, awaiting_kind="deposition"),
    "depo_prep": TaskDef(
        task_id="depo_prep", title="Depo Prep", glyph="❔",
        script_name="depo_prep.py",
        description="Deposition outline with questions grounded in case sources.",
        default_folders=("DISCOVERY", "PLEADINGS", "RECORDS"), file_exts=DOCS,
        two_phase=True, phase1_args=("--phase=analyze",),
        phase2_flag="--phase=generate", awaiting_kind="depo_prep",
        pre_settings="depo_prep"),
    "medical_records": TaskDef(
        task_id="medical_records", title="Medical Records Review", glyph="\U0001F3E5",
        script_name="med_record.py",
        description="Extract and summarize medical records into a chronology.",
        default_folders=("RECORDS",)),
    "med_chron_analysis": TaskDef(
        task_id="med_chron_analysis", title="Med Chron Analysis", glyph="\U0001FA7A",
        script_name="med_chron.py",
        description="Selectable analyses on a medical chronology (you pick mid-run).",
        default_folders=("RECORDS",), file_exts=DOCS,
        two_phase=True, phase1_args=("--phase=prep",),
        phase2_flag="--phase=run", awaiting_kind="med_chron"),
    "separate": TaskDef(
        task_id="separate", title="Separate Documents", glyph="\U0001F4D1",
        script_name="separate.py",
        description="Split a combined PDF into individually-named documents."),
}


def script_path(script_name: str) -> str:
    return str(REPO_ROOT / "Scripts" / script_name)


def build_phase1_argv(task: TaskDef, input_path: str) -> List[str]:
    return [script_path(task.script_name), *task.phase1_args, input_path]


def build_phase2_argv(task: TaskDef, session_path: str) -> List[str]:
    return [script_path(task.script_name), task.phase2_flag, session_path]


def read_session_json(session_path: str) -> dict:
    return json.loads(Path(session_path).read_text(encoding="utf-8"))


# ---- Session bridging (same modules the desktop forms use) ----

def apply_deposition_user_config(session_path: str, cfg: dict) -> None:
    from icharlotte_core.deposition import session_manager
    session_manager.update_user_config(Path(session_path), cfg)


def apply_med_chron_user_config(session_path: str, cfg: dict) -> None:
    from icharlotte_core.med_chron import session_manager
    session_manager.update_user_config(Path(session_path), cfg)


def read_depo_prep_topics(session_path: str) -> list:
    topics_path = Path(session_path).parent / "topics.json"
    return json.loads(topics_path.read_text(encoding="utf-8")).get("topics", [])


def write_depo_prep_topics(session_path: str, topics: list) -> None:
    topics_path = Path(session_path).parent / "topics.json"
    topics_path.write_text(json.dumps({"topics": topics}, indent=2), encoding="utf-8")


def write_depo_prep_config(cfg: dict) -> str:
    """Persist a depo-prep config.json to a temp dir; return its path.

    Mirrors DepoPrepSettingsPage._on_analyze_clicked() -- the config path is
    the positional argv for ``depo_prep.py --phase=analyze``.
    """
    tmpdir = tempfile.mkdtemp(prefix="depo_prep_config_")
    cfg_path = Path(tmpdir) / "config.json"
    cfg_path.write_text(json.dumps(cfg, indent=2), encoding="utf-8")
    return str(cfg_path)


# ---- Form option lists (mirror the desktop combos exactly) ----

DEPO_PREP_STYLES = [
    ("discovery", "Discovery / Fact-gathering"),
    ("lockdown", "Lock-down (leading admissions)"),
    ("expert", "Expert challenge (Daubert-style)"),
    ("friendly", "Friendly (own client prep)"),
]

DEPO_AUDIENCES = [
    ("neutral", "Neutral"),
    ("pro_plaintiff", "Plaintiff's Counsel"),
    ("pro_defense", "Defense Counsel"),
    ("custom", "Custom…"),
]

DEPO_TONES = [
    ("recitation", "Recitation (no editorializing)"),
    ("editorial", "Editorial (allow analysis)"),
    ("custom", "Custom…"),
]
