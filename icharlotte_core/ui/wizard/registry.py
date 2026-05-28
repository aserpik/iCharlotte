"""Task registry — single source of truth for Wizard Mode task cards.

Each task contributes:
  - task_id            : stable identifier used in persistence and code paths
  - title              : human-readable card title
  - description        : one-line card description
  - icon_glyph         : single emoji-ish character used as the card icon
  - default_folders    : ordered list of relative subfolders (under case root)
                         tried in order when opening the pre-Settings file
                         dialog; empty list means default to case root.
  - settings_page_cls  : SettingsPage subclass to instantiate for this task.
                         Defaults to base SettingsPage (placeholder).
"""
from dataclasses import dataclass, field
from typing import List, Optional


def _default_settings_page_cls():
    from .pages.settings_page import SettingsPage
    return SettingsPage


def _deposition_settings_page_cls():
    from .pages.deposition_settings_page import DepositionSettingsPage
    return DepositionSettingsPage


def _med_chron_settings_page_cls():
    from .pages.med_chron_settings_page import MedChronSettingsPage
    return MedChronSettingsPage


def _depo_prep_settings_page_cls():
    from .pages.depo_prep_settings_page import DepoPrepSettingsPage
    return DepoPrepSettingsPage


def _depo_prep_output_page_cls():
    from .pages.depo_prep_output_page import DepoPrepOutputPage
    return DepoPrepOutputPage


def _default_output_page_cls():
    from .pages.output_page import OutputPage
    return OutputPage


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
    # None means "use default SettingsPage"; resolved lazily in TaskTab.
    # Set to a factory callable that returns the class to avoid circular imports.
    _settings_page_cls_factory: Optional[object] = field(default=None, repr=False, compare=False)
    _output_page_cls_factory: Optional[object] = field(default=None, repr=False, compare=False)

    @property
    def settings_page_cls(self) -> type:
        """Return the SettingsPage subclass for this task."""
        if self._settings_page_cls_factory is not None:
            return self._settings_page_cls_factory()
        return _default_settings_page_cls()

    @property
    def output_page_cls(self) -> type:
        """Return the OutputPage subclass for this task."""
        if self._output_page_cls_factory is not None:
            return self._output_page_cls_factory()
        return _default_output_page_cls()


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
        _settings_page_cls_factory=_deposition_settings_page_cls,
    ),
    "depo_prep": TaskSpec(
        task_id="depo_prep",
        title="Depo Prep",
        description="Generate a deposition outline with questions grounded in case sources.",
        icon_glyph="❔",  # white question mark ornament
        script_name="depo_prep.py",
        default_folders=["DISCOVERY", "PLEADINGS", "RECORDS"],
        phase1_args=["--phase=analyze"],
        phase2_flag="--phase=generate",
        _settings_page_cls_factory=_depo_prep_settings_page_cls,
        _output_page_cls_factory=_depo_prep_output_page_cls,
    ),
    "medical_records": TaskSpec(
        task_id="medical_records",
        title="Medical Records Review",
        description="Extract and summarize medical records into a chronology.",
        icon_glyph="\U0001F3E5",  # 🏥
        script_name="med_record.py",
        default_folders=["RECORDS"],
    ),
    "med_chron_analysis": TaskSpec(
        task_id="med_chron_analysis",
        title="Med Chron Analysis",
        description="Run selectable analyses on a medical chronology.",
        icon_glyph="\U0001FA7A",  # 🩺
        script_name="med_chron.py",
        # Picker prefers RECORDS/<medical summary subfolder> via find_medical_summary_folder()
        # in iCharlotte._open_task_tab; this list is the fallback if no such subfolder exists.
        default_folders=["RECORDS"],
        phase1_args=["--phase=prep"],
        phase2_flag="--phase=run",
        _settings_page_cls_factory=_med_chron_settings_page_cls,
    ),
    "med_record_extractor": TaskSpec(
        task_id="med_record_extractor",
        title="Med Record Extractor",
        description="Pull pages from medical record PDFs based on pasted chronology entries.",
        icon_glyph="\U0001F4D1",  # 📑
        script_name="",  # in-process QThread worker
        default_folders=[],
    ),
    "subpoena_tracker": TaskSpec(
        task_id="subpoena_tracker",
        title="Subpoena Tracker",
        description="Cross-reference issued subpoenas with received records and chronologies.",
        icon_glyph="\U0001F4DC",  # 📜
        script_name="",  # in-process QThread worker
        default_folders=[],
    ),
    "respond_to_discovery": TaskSpec(
        task_id="respond_to_discovery",
        title="Respond to Discovery",
        description="Draft objections and responses to incoming written discovery.",
        icon_glyph="\U0001F4DD",  # 📝
        script_name="",  # in-process QThread worker
        default_folders=["DISCOVERY/PROPOUNDED", "DISCOVERY"],
    ),
    "oppose_motion": TaskSpec(
        task_id="oppose_motion",
        title="Oppose a Motion",
        description="Draft and verify an opposition memorandum for a California civil motion.",
        icon_glyph="\U0001F4DD",  # 📝
        script_name="",
        default_folders=["MOTIONS", "PLEADINGS", "DISCOVERY"],
    ),
    "chat": TaskSpec(
        task_id="chat",
        title="Chat",
        description="Open an AI chat session for this case.",
        icon_glyph="\U0001F4AC",  # 💬
        script_name="",  # no script — opens an interactive ChatTab
        default_folders=[],
    ),
}


def get_task(task_id: str) -> TaskSpec:
    """Return the TaskSpec for `task_id`. Raises KeyError if unknown."""
    return TASK_REGISTRY[task_id]


def list_tasks() -> list[TaskSpec]:
    """Return all registered tasks in registry-insertion order."""
    return list(TASK_REGISTRY.values())
