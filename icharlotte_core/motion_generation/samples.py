"""Style-sample loading for the Generate Motion drafter.

Reuses the opposition StyleExampleRegistry: .docx samples tagged by motion-type
id, with an active flag. Samples live next to the motion-types registry at
``Scripts/prompts/generate_motion/style_examples.json`` and are editable via the
Workbench Style Examples tab.
"""
from __future__ import annotations

import os
from typing import List, Optional

from .config import motion_types_path


def style_examples_path() -> str:
    return os.path.join(os.path.dirname(motion_types_path()), "style_examples.json")


def load_exemplars(
    motion_type_id: str,
    *,
    registry_path: Optional[str] = None,
    cache_dir: Optional[str] = None,
) -> List[str]:
    """Return extracted text of active style samples matching ``motion_type_id``."""
    from icharlotte_core.opposition.style_examples import (
        StyleExampleRegistry,
        extract_exemplar_text,
    )

    registry_path = registry_path or style_examples_path()
    cache_dir = cache_dir or os.path.join(
        os.path.dirname(registry_path), ".cache", "style_examples"
    )
    registry = StyleExampleRegistry.load(registry_path)
    exemplars: List[str] = []
    for match in registry.matches_for_motion_type(motion_type_id or ""):
        text = extract_exemplar_text(match.path, cache_dir=cache_dir)
        if text.strip():
            exemplars.append(text)
    return exemplars
