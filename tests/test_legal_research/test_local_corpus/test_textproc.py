from icharlotte_core.legal_research.local_corpus.textproc import (
    normalize_text, chunk_passages,
)


def test_normalize_fixes_section_and_whitespace():
    raw = "Pen. Code, � 187;  multiple   spaces\n\n\n\nand  breaks"
    out = normalize_text(raw)
    assert "�" not in out
    assert "§ 187" in out          # replacement char -> section symbol
    assert "multiple spaces" in out      # runs collapsed
    assert "\n\n\n" not in out           # >2 newlines collapsed


def test_chunk_passages_splits_on_paragraphs_within_budget():
    para = "Sentence one. " * 30          # ~ a chunk-ish paragraph
    text = "\n\n".join([para, para, para])
    chunks = chunk_passages(text, target_tokens=120)
    assert len(chunks) >= 2
    assert all(c.strip() for c in chunks)
    # No chunk wildly exceeds the budget (token ~ chars/4 heuristic, allow 2x)
    assert all(len(c) <= 120 * 4 * 2 for c in chunks)


def test_chunk_passages_empty_returns_empty():
    assert chunk_passages("   ") == []
