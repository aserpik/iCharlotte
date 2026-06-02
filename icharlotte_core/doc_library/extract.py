"""Single-call text extraction reused by the document library.

Dispatches by extension, reusing the app's existing extractors so the library
never re-implements OCR/Word logic:
  - .docx -> extract_docx_text (tables + headers, document order)
  - .doc  -> Word COM read (read-only attach; never Quit / never set Visible)
  - everything else (.pdf/.txt/.msg/...) -> DocumentProcessor.extract_text
"""
from __future__ import annotations

import os
from dataclasses import dataclass
from typing import Optional


@dataclass
class Extracted:
    text: str
    page_count: int
    extract_method: str
    error: Optional[str]


def extract_any(path: str) -> Extracted:
    if not os.path.exists(path):
        return Extracted("", 0, "failed", f"File not found: {path}")

    ext = os.path.splitext(path)[1].lower()
    try:
        if ext == ".docx":
            from ..document_processor import extract_docx_text
            text = extract_docx_text(path)
            return Extracted(text, 0, "docx", None)
        if ext == ".doc":
            from ..ui.tabs import ChatTab  # reuse the read-only COM helper
            text = ChatTab._extract_doc_text(path)
            err = None
            if text.strip().startswith("[Error reading .doc"):
                err = text.strip()
            return Extracted(text, 0, "doc_com", err)
        # PDF / TXT / MSG / other -> DocumentProcessor
        from ..document_processor import DocumentProcessor, ExtractionMethod
        result = DocumentProcessor().extract_text(path, ocr_enabled=True)
        err = result.error
        if result.extraction_method == ExtractionMethod.FAILED and not err:
            err = "extraction failed"
        return Extracted(
            text=result.text or "",
            page_count=result.page_count or 0,
            extract_method=result.extraction_method.value,
            error=err,
        )
    except Exception as e:  # never raise into a task-completion handler
        return Extracted("", 0, "failed", f"{type(e).__name__}: {e}")
