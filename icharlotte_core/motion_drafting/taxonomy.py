"""Motion Database folder taxonomy for Motion Drafting."""

from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path


DRAFT_KIND_MOTION = "motion"
DRAFT_KIND_OPPOSITION = "opposition"
DRAFT_KIND_REPLY = "reply"

DRAFT_KIND_LABELS = {
    DRAFT_KIND_MOTION: "Generate a motion",
    DRAFT_KIND_OPPOSITION: "Generate an opposition to a motion",
    DRAFT_KIND_REPLY: "Generate a reply in support of a motion",
}

_MOTION_ROOT_PREFIXES = ("Motion -", "Motions -")
_MOTION_ROOT_NAMES = {"Ex Parte Applications"}
_EXCLUDED_PREFIXES = ("_Support", "_Other")


@dataclass(frozen=True)
class MotionTypeOption:
    """A selectable motion taxonomy entry."""

    label: str
    draft_kind: str
    source_path: str = ""
    engine_type_id: str = "generic"


def motion_database_root() -> Path:
    """Return the configured Motion Database root."""
    override = os.environ.get("ICHARLOTTE_MOTION_DATABASE_ROOT", "").strip()
    if override:
        return Path(override)
    repo_root = Path(__file__).resolve().parents[2]
    return repo_root / "MOTION DATABASE"


def list_motion_type_options(draft_kind: str, root: str | os.PathLike[str] | None = None) -> list[MotionTypeOption]:
    """Return motion-type options derived from Motion Database folders."""
    root_path = Path(root) if root is not None else motion_database_root()
    labels: dict[str, str] = {}
    if root_path.exists():
        for dataset_dir in sorted((p for p in root_path.iterdir() if p.is_dir()), key=lambda p: p.name.lower()):
            for taxonomy_root in _taxonomy_roots(dataset_dir, draft_kind):
                for folder in _iter_selectable_folders(taxonomy_root, draft_kind):
                    label = _label_for(folder, taxonomy_root, draft_kind)
                    if label and not _is_excluded_label(label):
                        labels.setdefault(label, str(folder))

    if not labels:
        return _fallback_options(draft_kind)

    return [
        MotionTypeOption(
            label=label,
            draft_kind=draft_kind,
            source_path=labels[label],
            engine_type_id=_engine_type_id_for_label(label),
        )
        for label in sorted(labels, key=str.lower)
    ]


def _taxonomy_roots(dataset_dir: Path, draft_kind: str) -> list[Path]:
    if draft_kind == DRAFT_KIND_OPPOSITION:
        path = dataset_dir / "Oppositions"
        return [path] if path.is_dir() else []
    if draft_kind == DRAFT_KIND_REPLY:
        path = dataset_dir / "Replies"
        return [path] if path.is_dir() else []
    return [
        child
        for child in dataset_dir.iterdir()
        if child.is_dir()
        and (
            child.name.startswith(_MOTION_ROOT_PREFIXES)
            or child.name in _MOTION_ROOT_NAMES
        )
    ]


def _iter_selectable_folders(root: Path, draft_kind: str):
    if draft_kind == DRAFT_KIND_MOTION:
        yield root
    for folder in root.rglob("*"):
        if folder.is_dir():
            yield folder


def _label_for(folder: Path, taxonomy_root: Path, draft_kind: str) -> str:
    if draft_kind == DRAFT_KIND_MOTION:
        parts = folder.relative_to(taxonomy_root.parent).parts
    else:
        parts = folder.relative_to(taxonomy_root).parts
    return " / ".join(part for part in parts if part)


def _is_excluded_label(label: str) -> bool:
    first = label.split("/", 1)[0].strip()
    return first.startswith(_EXCLUDED_PREFIXES)


def _engine_type_id_for_label(label: str) -> str:
    normalized = label.lower()
    mappings = [
        ("summary judgment", "msj"),
        ("summary adjudication", "msj"),
        ("msj", "msj"),
        ("msa", "msj"),
        ("compel", "compel"),
        ("demurrer", "demurrer"),
        ("strike", "strike"),
        ("quash", "quash"),
        ("sanction", "sanctions"),
        ("continue trial", "continue_trial"),
        ("trial continuance", "continue_trial"),
        ("ex parte", "ex_parte"),
        ("medical exam", "ime"),
        ("ime", "ime"),
        ("dme", "ime"),
        ("good faith", "gfs"),
        ("dismiss", "dismiss"),
        ("leave", "leave"),
        ("consolidate", "consolidate"),
        ("protective order", "protective_order"),
    ]
    for needle, type_id in mappings:
        if needle in normalized:
            return type_id
    return "generic"


def _fallback_options(draft_kind: str) -> list[MotionTypeOption]:
    labels_by_kind = {
        DRAFT_KIND_MOTION: [
            ("Motion to Compel", "compel"),
            ("Demurrer", "demurrer"),
            ("Motion to Strike", "strike"),
            ("Motion for Summary Judgment/Adjudication", "msj"),
            ("Motion to Quash", "quash"),
            ("Ex Parte Application", "ex_parte"),
            ("Other", "generic"),
        ],
        DRAFT_KIND_OPPOSITION: [
            ("Motion to Compel", "compel"),
            ("Demurrer", "demurrer"),
            ("Motion to Strike", "strike"),
            ("MSJ-MSA", "msj"),
            ("Other", "generic"),
        ],
        DRAFT_KIND_REPLY: [
            ("Motion to Compel", "compel"),
            ("Demurrer", "demurrer"),
            ("Motion to Strike", "strike"),
            ("MSJ-MSA", "msj"),
            ("Other", "generic"),
        ],
    }
    return [
        MotionTypeOption(label=label, draft_kind=draft_kind, engine_type_id=engine_type_id)
        for label, engine_type_id in labels_by_kind.get(draft_kind, labels_by_kind[DRAFT_KIND_MOTION])
    ]
