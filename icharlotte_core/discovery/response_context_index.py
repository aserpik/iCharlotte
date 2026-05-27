"""Request-scoped context selection for discovery response drafting."""
from __future__ import annotations

import os
import re
from dataclasses import dataclass
from typing import Iterable, Mapping

from icharlotte_core.discovery.response_parser import ParsedRequest


@dataclass(frozen=True)
class ContextChunk:
    source_path: str
    sequence: int
    text: str
    heading: str = ""


_DISCOVERY_CUES = {
    "admit", "admission", "communications", "documents", "identify",
    "incident", "produce", "witness", "witnesses",
}

_LOW_SIGNAL_TERMS = {
    "all", "and", "any", "are", "each", "for", "from", "identify",
    "request", "response", "state", "that", "the", "this", "with", "you",
    "your",
}


def build_context_chunks(text_by_path: Mapping[str, str]) -> list[ContextChunk]:
    chunks: list[ContextChunk] = []
    for source_path, text in text_by_path.items():
        for sequence, raw_chunk_text in enumerate(_split_text(text), start=1):
            heading = _detect_heading(raw_chunk_text)
            chunk_text = _collapse_text(raw_chunk_text)
            chunks.append(
                ContextChunk(
                    source_path=source_path,
                    sequence=sequence,
                    text=chunk_text,
                    heading=heading,
                )
            )
    return chunks


def select_context_packet(
    request: ParsedRequest,
    chunks: Iterable[ContextChunk],
    max_chunks: int = 5,
    min_score: int = 2,
) -> list[ContextChunk]:
    scored = [
        (_score_chunk(request, chunk), chunk)
        for chunk in chunks
    ]
    scored = [(score, chunk) for score, chunk in scored if score >= min_score]
    scored.sort(key=lambda item: (-item[0], item[1].source_path, item[1].sequence))
    return [chunk for _score, chunk in scored[:max_chunks]]


def format_context_packet(chunks: Iterable[ContextChunk]) -> str:
    parts: list[str] = []
    for chunk in chunks:
        label = f"[{os.path.basename(chunk.source_path)} #{chunk.sequence}]"
        parts.append(f"{label}\n{chunk.text.strip()}")
    return "\n\n".join(parts)


def _split_text(text: str) -> list[str]:
    normalized = re.sub(r"\r\n?", "\n", text or "")
    raw_parts = re.split(r"\n\s*\n+", normalized)
    chunks = [part.strip() for part in raw_parts]
    return [chunk for chunk in chunks if len(_collapse_text(chunk)) >= 20]


def _collapse_text(text: str) -> str:
    return re.sub(r"\s+", " ", text or "").strip()


def _detect_heading(text: str) -> str:
    first_line = (text or "").splitlines()[0].strip() if text else ""
    word_count = len(re.findall(r"[A-Za-z0-9][A-Za-z0-9'-]*", first_line))
    has_terminal_sentence_punctuation = bool(re.search(r"[.!?]$", first_line))
    if (
        0 < len(first_line) <= 80
        and 1 <= word_count <= 4
        and not has_terminal_sentence_punctuation
    ):
        return first_line
    return ""


def _score_chunk(request: ParsedRequest, chunk: ContextChunk) -> int:
    request_terms = _terms(request.text)
    chunk_text = f"{chunk.heading} {chunk.text}".lower()
    score = 0
    for term in request_terms:
        if term in chunk_text:
            score += 3 if term in _DISCOVERY_CUES else 2
    for term in getattr(request, "defined_terms_used", []) or []:
        if term.lower() in chunk_text:
            score += 4
    if re.search(r"\b\d{1,2}/\d{1,2}/\d{2,4}\b", request.text or "") and re.search(
        r"\b\d{1,2}/\d{1,2}/\d{2,4}\b", chunk.text
    ):
        score += 2
    return score


def _terms(text: str) -> set[str]:
    words = {word.lower() for word in re.findall(r"[A-Za-z][A-Za-z0-9'-]{2,}", text or "")}
    return {word for word in words if word not in _LOW_SIGNAL_TERMS}
