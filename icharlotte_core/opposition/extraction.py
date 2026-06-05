"""File validation and extraction wrappers for opposition drafting."""

from __future__ import annotations

import os

from icharlotte_core.doc_library.library import DocumentLibrary
from icharlotte_core.document_processor import (
    DocumentProcessor,
    ExtractionMethod,
    ExtractResult,
)

MOTION_EXTENSIONS = {".pdf", ".docx"}
CONTEXT_EXTENSIONS = {".pdf", ".docx", ".txt", ".msg"}


def is_supported_motion_file(path: str) -> bool:
    return os.path.splitext(path or "")[1].lower() in MOTION_EXTENSIONS


def is_supported_context_file(path: str) -> bool:
    return os.path.splitext(path or "")[1].lower() in CONTEXT_EXTENSIONS


def validate_motion_file(path: str) -> tuple[bool, str]:
    if not path:
        return False, "Motion file path is required."
    if not os.path.isfile(path):
        return False, f"Motion file not found: {path}"

    extension = os.path.splitext(path)[1].lower()
    if not is_supported_motion_file(path):
        allowed = ", ".join(sorted(MOTION_EXTENSIONS))
        return False, f"Unsupported motion file extension {extension}; expected {allowed}."

    return True, ""


def validate_context_files(paths: list[str]) -> tuple[list[str], list[str]]:
    valid_paths: list[str] = []
    warnings: list[str] = []

    for path in paths:
        if not path:
            warnings.append("Skipped blank context file path.")
            continue
        if not os.path.isfile(path):
            warnings.append(f"Skipped context file not found: {path}")
            continue

        extension = os.path.splitext(path)[1].lower()
        if not is_supported_context_file(path):
            allowed = ", ".join(sorted(CONTEXT_EXTENSIONS))
            warnings.append(
                f"Skipped unsupported context file extension {extension}: {path}. "
                f"Expected one of {allowed}."
            )
            continue

        valid_paths.append(path)

    return valid_paths, warnings


def _extract_from_case_cache(path: str, case_root: str | None) -> ExtractResult | None:
    if not case_root:
        return None
    text, _method, error = DocumentLibrary(case_root).get_or_extract_text(path)
    if error:
        return ExtractResult(
            text=text or "",
            page_count=0,
            extraction_method=ExtractionMethod.FAILED,
            char_count=len(text or ""),
            file_path=path,
            error=error,
        )
    return ExtractResult(
        text=text,
        page_count=0,
        extraction_method=ExtractionMethod.NATIVE,
        char_count=len(text or ""),
        file_path=path,
    )


def extract_document_text(
    path: str,
    *,
    ocr_enabled: bool = True,
    case_root: str | None = None,
) -> ExtractResult:
    cached = _extract_from_case_cache(path, case_root)
    if cached is not None:
        return cached
    return DocumentProcessor().extract_text(path, ocr_enabled=ocr_enabled)


def extract_context_bundle(
    paths: list[str],
    case_root: str | None = None,
) -> tuple[str, list[str]]:
    valid_paths, warnings = validate_context_files(paths)
    sections: list[str] = []

    for path in valid_paths:
        kwargs = {"case_root": case_root} if case_root else {}
        result = extract_document_text(path, **kwargs)
        label = os.path.basename(path) or path

        if not result.success:
            reason = result.error or "no extractable text"
            warnings.append(f"Skipped context file with no extracted text: {label} ({reason})")
            continue

        sections.append(f"Document: {label}\n\n{result.text.strip()}")

    return "\n\n---\n\n".join(sections), warnings
