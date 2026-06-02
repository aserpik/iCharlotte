"""Per-case persistent library of extracted source text.

Stored on the case drive, alongside wizard_state.json:
    <case_root>/NOTES/AI OUTPUT/.icharlotte/doc_library/
        index.json          # catalog of LibraryEntry
        blobs/<sha1>.txt     # extracted text, deduped by file content hash

Atomic writes via tmp + os.replace (same pattern as WizardStatePersistence).
"""
from __future__ import annotations

import hashlib
import json
import os
import uuid
from datetime import datetime
from typing import Callable, Optional

from .models import LibraryEntry, MemberFile

SCHEMA_VERSION = 1


class DocumentLibrary:
    def __init__(self, case_root: str):
        self.case_root = case_root

    # ---- paths ----
    @property
    def folder(self) -> str:
        return os.path.join(self.case_root, "NOTES", "AI OUTPUT",
                            ".icharlotte", "doc_library")

    @property
    def index_path(self) -> str:
        return os.path.join(self.folder, "index.json")

    @property
    def blobs_dir(self) -> str:
        return os.path.join(self.folder, "blobs")

    # ---- index load/save ----
    def _load(self) -> dict:
        if not os.path.isfile(self.index_path):
            return {"version": SCHEMA_VERSION, "entries": []}
        try:
            with open(self.index_path, "r", encoding="utf-8") as f:
                raw = json.load(f)
            if not isinstance(raw, dict) or not isinstance(raw.get("entries"), list):
                raise ValueError("bad index shape")
            return raw
        except (json.JSONDecodeError, OSError, ValueError):
            try:
                os.replace(self.index_path, self.index_path + ".corrupt")
            except OSError:
                pass
            return {"version": SCHEMA_VERSION, "entries": []}

    def _save(self, data: dict) -> None:
        os.makedirs(self.folder, exist_ok=True)
        tmp = self.index_path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2)
        os.replace(tmp, self.index_path)

    def _save_entries(self, entries: list) -> None:
        self._save({"version": SCHEMA_VERSION,
                    "entries": [e.to_dict() for e in entries]})

    # ---- read ----
    def list_entries(self) -> list:
        return [LibraryEntry.from_dict(d) for d in self._load().get("entries", [])]
