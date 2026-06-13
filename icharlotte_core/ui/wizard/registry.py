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
from typing import Dict, List, Optional


# Fixed display order for launcher category sections.
CATEGORY_ORDER: List[str] = [
    "General",
    "Summarize",
    "Discovery",
    "Medical",
    "Motions",
]


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
    category: str = "General"
    keywords: List[str] = field(default_factory=list)
    default_folders: List[str] = field(default_factory=list)
    phase1_args: List[str] = field(default_factory=list)
    phase2_flag: str = "--phase=summary"
    hidden: bool = False
    # Optional launcher-card corner button (e.g. Separate → open the Index).
    # When card_action_id is set, TaskCard renders a small QToolButton that
    # emits action_requested(card_action_id). None = no button (default).
    card_action_id: Optional[str] = None
    card_action_glyph: Optional[str] = None
    card_action_tooltip: Optional[str] = None
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
        category="Summarize",
        keywords=["summary", "summarize", "document", "general"],
        default_folders=[],
        card_action_id="open_summarize_documents_outputs",
        card_action_glyph="\U0001F5C2",  # 🗂 card index dividers
        card_action_tooltip="Browse prior document summaries for this case",
    ),
    "summarize_discovery": TaskSpec(
        task_id="summarize_discovery",
        title="Summarize Discovery",
        description="Summarize discovery responses with structure and citations.",
        icon_glyph="\U0001F4CB",  # 📋
        script_name="summarize_discovery.py",
        category="Summarize",
        keywords=[
            "responses", "RFP", "RFA", "interrogatory",
            "form interrogatories", "special interrogatories", "production",
        ],
        default_folders=["DISCOVERY/RESPONSES", "DISCOVERY"],
        card_action_id="open_summarize_discovery_outputs",
        card_action_glyph="\U0001F5C2",  # 🗂 card index dividers
        card_action_tooltip="Browse prior discovery summaries for this case",
    ),
    "summarize_depositions": TaskSpec(
        task_id="summarize_depositions",
        title="Summarize Depositions",
        description="Generate a structured summary of one or more depositions.",
        icon_glyph="\U0001F399",  # 🎙
        script_name="summarize_deposition.py",
        category="Summarize",
        keywords=["depo", "transcript", "testimony", "witness"],
        default_folders=["DISCOVERY/TRANSCRIPTS", "DISCOVERY"],
        card_action_id="open_summarize_depositions_outputs",
        card_action_glyph="\U0001F5C2",  # 🗂 card index dividers
        card_action_tooltip="Browse prior deposition summaries for this case",
        _settings_page_cls_factory=_deposition_settings_page_cls,
    ),
    "depo_prep": TaskSpec(
        task_id="depo_prep",
        title="Depo Prep",
        description="Generate a deposition outline with questions grounded in case sources.",
        icon_glyph="❔",  # ❔ white question mark ornament
        script_name="depo_prep.py",
        category="Discovery",
        keywords=["depo", "outline", "questions", "examination", "prepare"],
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
        category="Medical",
        keywords=["medical", "records", "chronology", "IME", "billing", "MRI", "treatment"],
        default_folders=["RECORDS"],
        card_action_id="open_medical_records_outputs",
        card_action_glyph="\U0001F5C2",  # 🗂 card index dividers
        card_action_tooltip="Browse prior medical record reviews for this case",
    ),
    "med_chron_analysis": TaskSpec(
        task_id="med_chron_analysis",
        title="Med Chron Analysis",
        description="Run selectable analyses on a medical chronology.",
        icon_glyph="\U0001FA7A",  # 🩺
        script_name="med_chron.py",
        category="Medical",
        keywords=["chronology", "chron", "analysis", "medical", "gaps", "billing"],
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
        category="Medical",
        keywords=["extract", "pages", "PDF", "records", "Bates", "exhibit"],
        default_folders=[],
    ),
    "separate": TaskSpec(
        task_id="separate",
        title="Separate Documents",
        description="Split a combined PDF into individually-named documents using AI.",
        icon_glyph="\U0001F4D1",  # 📑
        script_name="separate.py",
        default_folders=[],
        category="General",
        card_action_id="open_separate_index",
        card_action_glyph="\U0001F5C2",  # 🗂 card index dividers
        card_action_tooltip="Open the document Index for this case",
    ),
    "subpoena_tracker": TaskSpec(
        task_id="subpoena_tracker",
        title="Subpoena Tracker",
        description="Cross-reference issued subpoenas with received records and chronologies.",
        icon_glyph="\U0001F4DC",  # 📜
        script_name="",  # in-process QThread worker
        category="Medical",
        keywords=["subpoena", "SDT", "records", "deposition subpoena", "tracker"],
        default_folders=[],
    ),
    "respond_to_discovery": TaskSpec(
        task_id="respond_to_discovery",
        title="Respond to Discovery",
        description="Draft objections and responses to incoming written discovery.",
        icon_glyph="\U0001F4DD",  # 📝
        script_name="",  # in-process QThread worker
        category="Discovery",
        keywords=[
            "RFP", "RFA", "interrogatory", "form interrogatories",
            "objections", "propounded", "responses",
        ],
        default_folders=["DISCOVERY/PROPOUNDED", "DISCOVERY"],
    ),
    "motion_drafting": TaskSpec(
        task_id="motion_drafting",
        title="Motion Drafting",
        description="Draft a motion, opposition, or reply using Motion Database categories.",
        icon_glyph="⚖️",
        script_name="",
        category="Motions",
        keywords=[
            "motion",
            "draft",
            "opposition",
            "oppose",
            "reply",
            "compel",
            "demurrer",
            "strike",
            "summary judgment",
        ],
        default_folders=["MOTIONS", "PLEADINGS", "DISCOVERY"],
    ),
    "oppose_motion": TaskSpec(
        task_id="oppose_motion",
        title="Oppose a Motion",
        description="Draft and verify an opposition memorandum for a California civil motion.",
        icon_glyph="\U0001F4DD",  # 📝
        script_name="",
        category="Motions",
        keywords=["motion", "opposition", "oppose", "MSJ", "brief", "memorandum", "demurrer"],
        default_folders=["MOTIONS", "PLEADINGS", "DISCOVERY"],
        hidden=True,
    ),
    "generate_motion": TaskSpec(
        task_id="generate_motion",
        title="Generate a Motion",
        description="Draft a California civil motion from scratch from your target documents.",
        icon_glyph="⚖️",  # ⚖️
        script_name="",  # in-process worker
        category="Motions",
        keywords=[
            "motion", "draft", "compel", "demurrer", "strike",
            "memorandum", "points and authorities", "notice of motion",
        ],
        default_folders=["MOTIONS", "PLEADINGS", "DISCOVERY"],
        hidden=True,
    ),
    "mediation_brief": TaskSpec(
        task_id="mediation_brief",
        title="Mediation Brief",
        description="Generate a defense-side mediation brief from case documents.",
        icon_glyph="\U0001F91D",  # 🤝
        script_name="",  # in-process worker
        category="Motions",
        keywords=[
            "mediation", "brief", "settlement", "mediator",
            "defense", "confidential",
        ],
        default_folders=[],
    ),
    "case_intake_docket": TaskSpec(
        task_id="case_intake_docket",
        title="Case Intake & Docket",
        description="Extract complaint metadata, review case details, then download and process the court docket.",
        icon_glyph="\U0001F5C2",
        script_name="",
        category="General",
        keywords=["complaint", "docket", "case number", "venue", "intake", "hearing", "trial"],
        default_folders=[],
    ),
    "chat": TaskSpec(
        task_id="chat",
        title="Chat",
        description="Open an AI chat session for this case.",
        icon_glyph="\U0001F4AC",  # 💬
        script_name="",  # no script — opens an interactive ChatTab
        category="General",
        keywords=["chat", "ask", "assistant", "question"],
        default_folders=[],
    ),
}


def _matches(spec: TaskSpec, query_lower: str) -> bool:
    """True if the (already-lowercased) query is a substring of the task's
    searchable text: title + description + keyword aliases."""
    haystack = " ".join(
        [spec.title, spec.description, *spec.keywords]
    ).lower()
    return query_lower in haystack


def filter_tasks(specs: List[TaskSpec], query: str) -> "Dict[str, List[TaskSpec]]":
    """Group `specs` by category (in CATEGORY_ORDER), optionally filtered by
    `query`. Empty/whitespace query returns every task. Categories with no
    matching tasks are omitted. Task order within a category follows the order
    of `specs`.
    """
    q = (query or "").strip().lower()
    grouped: Dict[str, List[TaskSpec]] = {}
    for category in CATEGORY_ORDER:
        matched = [
            spec
            for spec in specs
            if spec.category == category and (not q or _matches(spec, q))
        ]
        if matched:
            grouped[category] = matched
    return grouped


def get_task(task_id: str) -> TaskSpec:
    """Return the TaskSpec for `task_id`. Raises KeyError if unknown."""
    return TASK_REGISTRY[task_id]


def list_tasks() -> list[TaskSpec]:
    """Return all registered tasks in registry-insertion order."""
    return [spec for spec in TASK_REGISTRY.values() if not spec.hidden]
