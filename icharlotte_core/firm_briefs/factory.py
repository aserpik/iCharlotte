"""Shared paths + availability + index construction for firm-brief features."""
from __future__ import annotations

import os
from typing import Optional

from icharlotte_core import config

DATA_DIR = config.FIRM_BRIEFS_DATA_DIR


def index_paths():
    return (os.path.join(DATA_DIR, "firm_briefs.db"),
            os.path.join(DATA_DIR, "profiles.f16"))


def index_available() -> bool:
    db, vec = index_paths()
    return os.path.exists(db) and os.path.exists(vec)


def make_index(*, embedder=None):
    if not index_available():
        return None
    from .index import FirmBriefIndex
    db, vec = index_paths()
    return FirmBriefIndex(db_path=db, vectors_path=vec, embedder=embedder)
