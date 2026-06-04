"""Shared per-case document-index store.

Both Advanced Mode's IndexTab and Wizard Mode's Separate task read/write one
JSON per case so a run in either mode is visible in both.

File:  <GEMINI_DATA_DIR>/<file_number>_index.json
Shape: {pdf_path: [ {id, title, date, start, end}, ... ]}

GEMINI_DATA_DIR is read from icharlotte_core.config at call time (attribute
access, not a frozen import binding) so tests can monkeypatch it.
"""
import json
import os

from icharlotte_core import config


def index_path(file_number: str) -> str:
    """Absolute path of the per-case index JSON."""
    return os.path.join(config.GEMINI_DATA_DIR, f"{file_number}_index.json")


def load_index(file_number: str) -> dict:
    """Return the {pdf_path: [docs]} map, or {} if missing/corrupt."""
    path = index_path(file_number)
    if not os.path.exists(path):
        return {}
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
    except (OSError, ValueError):
        return {}
    return data if isinstance(data, dict) else {}


def upsert_pdf(file_number: str, pdf_path: str, docs: list) -> None:
    """Set data[pdf_path] = docs and persist (matches IndexTab's format)."""
    data = load_index(file_number)
    data[pdf_path] = docs
    os.makedirs(config.GEMINI_DATA_DIR, exist_ok=True)
    with open(index_path(file_number), "w", encoding="utf-8") as f:
        json.dump(data, f, indent=4)
