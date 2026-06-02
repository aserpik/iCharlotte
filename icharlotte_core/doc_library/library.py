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
from .extract import extract_any as _default_extractor
from .labels import auto_label
from ..chat.token_counter import TokenCounter

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

    # ---- hashing / blobs ----
    @staticmethod
    def _hash_file(path: str) -> str:
        h = hashlib.sha1()
        with open(path, "rb") as f:
            for chunk in iter(lambda: f.read(1 << 20), b""):
                h.update(chunk)
        return h.hexdigest()

    def get_member_text(self, blob: Optional[str]) -> str:
        if not blob:
            return ""
        path = os.path.join(self.blobs_dir, blob)
        if not os.path.isfile(path):
            return ""
        with open(path, "r", encoding="utf-8", errors="replace") as f:
            return f.read()

    # ---- write ----
    def add_entry(self, task_type: str, source_paths: list,
                  metadata: Optional[dict] = None,
                  extractor: Callable = _default_extractor) -> LibraryEntry:
        os.makedirs(self.blobs_dir, exist_ok=True)
        entries = self.list_entries()

        # Idempotency: drop any prior entry with the same task_type + file set.
        target_set = {os.path.normcase(os.path.abspath(p)) for p in source_paths}
        kept = []
        for e in entries:
            e_set = {os.path.normcase(os.path.abspath(m.source_path)) for m in e.members}
            if not (e.task_type == task_type and e_set == target_set):
                kept.append(e)
        entries = kept

        members = []
        for path in source_paths:
            name = os.path.basename(path)
            try:
                digest = self._hash_file(path)
            except OSError as ex:
                members.append(MemberFile(path, name, None, 0, 0, "failed", str(ex)))
                continue
            blob_name = f"{digest}.txt"
            blob_path = os.path.join(self.blobs_dir, blob_name)
            if os.path.isfile(blob_path):
                text = self.get_member_text(blob_name)
                members.append(MemberFile(
                    path, name, blob_name, len(text),
                    TokenCounter.estimate_tokens(text), "cached", None))
                continue
            result = extractor(path)
            if result.error or not result.text:
                members.append(MemberFile(
                    path, name, None, 0, 0, result.extract_method,
                    result.error or "empty extraction"))
                continue
            tmp = blob_path + ".tmp"
            with open(tmp, "w", encoding="utf-8") as f:
                f.write(result.text)
            os.replace(tmp, blob_path)
            members.append(MemberFile(
                path, name, blob_name, len(result.text),
                TokenCounter.estimate_tokens(result.text),
                result.extract_method, None))

        existing_labels = [e.label for e in entries]
        label = auto_label(task_type, source_paths, metadata or {}, existing_labels)
        entry = LibraryEntry(
            id=uuid.uuid4().hex,
            label=label,
            auto_label=label,
            task_type=task_type,
            created_at=datetime.now().isoformat(timespec="seconds"),
            members=members,
        )
        entries.append(entry)
        self._save_entries(entries)
        return entry
