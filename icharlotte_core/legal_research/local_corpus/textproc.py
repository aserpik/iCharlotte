"""Text normalization + passage chunking for the corpus."""
from __future__ import annotations

import re

# CAP JSON occasionally carries U+FFFD where the section symbol belongs.
_REPLACEMENT = "�"


def normalize_text(text: str) -> str:
    if not text:
        return ""
    t = text.replace("\r\n", "\n").replace("\r", "\n")
    # The most common mojibake in CAP CA opinions is the section symbol.
    t = t.replace(_REPLACEMENT, "§")
    t = t.replace(" ", " ")                 # NBSP -> space
    t = re.sub(r"[ \t]+", " ", t)                # collapse intra-line runs
    t = re.sub(r"\n{3,}", "\n\n", t)             # collapse blank-line runs
    t = re.sub(r" *\n *", "\n", t)               # trim around newlines
    return t.strip()


def _approx_tokens(s: str) -> int:
    # Cheap heuristic: ~4 chars/token. Good enough for chunk budgeting.
    return max(1, len(s) // 4)


def chunk_passages(text: str, *, target_tokens: int = 512) -> list[str]:
    """Split normalized text into ~target_tokens passages on paragraph
    boundaries. Oversized paragraphs are hard-split by sentence."""
    text = normalize_text(text)
    if not text:
        return []
    paras = [p.strip() for p in text.split("\n\n") if p.strip()]
    chunks: list[str] = []
    buf: list[str] = []
    buf_tok = 0
    for para in paras:
        ptok = _approx_tokens(para)
        if ptok > target_tokens:
            # Flush buffer, then hard-split the big paragraph by sentences.
            if buf:
                chunks.append("\n\n".join(buf)); buf, buf_tok = [], 0
            chunks.extend(_split_sentences(para, target_tokens))
            continue
        if buf_tok + ptok > target_tokens and buf:
            chunks.append("\n\n".join(buf)); buf, buf_tok = [], 0
        buf.append(para); buf_tok += ptok
    if buf:
        chunks.append("\n\n".join(buf))
    return chunks


def _split_sentences(para: str, target_tokens: int) -> list[str]:
    sentences = re.split(r"(?<=[.;:])\s+", para)
    out: list[str] = []
    buf: list[str] = []
    buf_tok = 0
    for s in sentences:
        stok = _approx_tokens(s)
        if buf_tok + stok > target_tokens and buf:
            out.append(" ".join(buf)); buf, buf_tok = [], 0
        buf.append(s); buf_tok += stok
    if buf:
        out.append(" ".join(buf))
    return out
