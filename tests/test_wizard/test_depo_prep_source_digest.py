import json
from pathlib import Path
from unittest.mock import MagicMock

import pytest

from Scripts.depo_prep_lib.source_digest import (
    digest_single_source,
    DigestResult,
)


def _mock_llm_returning(payload: dict):
    caller = MagicMock()
    caller.call = MagicMock(return_value=json.dumps(payload))
    return caller


def _good_payload(source_id="x.pdf"):
    return {
        "source_id": source_id,
        "source_kind": "medical_records",
        "deponent_statements": [],
        "factual_anchors": [{"fact": "x", "location": "p.1", "topic_tags": ["x"]}],
        "inconsistencies": [],
        "summary": "test",
    }


def test_digest_single_source_writes_json(tmp_path):
    src = tmp_path / "test.pdf"
    src.write_bytes(b"some content")
    digests_dir = tmp_path / "digests"
    digests_dir.mkdir()
    (digests_dir / "raw").mkdir()
    (digests_dir / "raw" / "test.txt").write_text("...extracted text...", encoding="utf-8")

    caller = _mock_llm_returning(_good_payload(source_id="test.pdf"))

    result = digest_single_source(
        source_path=src,
        extracted_text_path=digests_dir / "raw" / "test.txt",
        digests_dir=digests_dir,
        llm_caller=caller,
        deponent_name="Jane Doe",
        deponent_role="Plaintiff",
    )
    assert isinstance(result, DigestResult)
    assert result.from_cache is False
    out = digests_dir / "test.pdf.json"
    assert out.exists()
    data = json.loads(out.read_text(encoding="utf-8"))
    assert data["source_id"] == "test.pdf"


def test_digest_single_source_uses_cache_on_same_hash(tmp_path):
    src = tmp_path / "test.pdf"
    src.write_bytes(b"same bytes")
    digests_dir = tmp_path / "digests"
    digests_dir.mkdir()
    (digests_dir / "raw").mkdir()
    (digests_dir / "raw" / "test.txt").write_text("...", encoding="utf-8")

    caller = _mock_llm_returning(_good_payload(source_id="test.pdf"))

    # First call - LLM is invoked.
    r1 = digest_single_source(
        source_path=src,
        extracted_text_path=digests_dir / "raw" / "test.txt",
        digests_dir=digests_dir, llm_caller=caller,
        deponent_name="Jane", deponent_role="P",
    )
    assert r1.from_cache is False
    assert caller.call.call_count == 1

    # Second call with same file - should hit cache, no LLM call.
    r2 = digest_single_source(
        source_path=src,
        extracted_text_path=digests_dir / "raw" / "test.txt",
        digests_dir=digests_dir, llm_caller=caller,
        deponent_name="Jane", deponent_role="P",
    )
    assert r2.from_cache is True
    assert caller.call.call_count == 1  # unchanged


def test_digest_single_source_invalidates_when_file_changes(tmp_path):
    src = tmp_path / "test.pdf"
    src.write_bytes(b"original")
    digests_dir = tmp_path / "digests"
    digests_dir.mkdir()
    (digests_dir / "raw").mkdir()
    (digests_dir / "raw" / "test.txt").write_text("v1", encoding="utf-8")

    caller = _mock_llm_returning(_good_payload(source_id="test.pdf"))

    digest_single_source(source_path=src, extracted_text_path=digests_dir/"raw"/"test.txt",
                         digests_dir=digests_dir, llm_caller=caller,
                         deponent_name="Jane", deponent_role="P")

    # Change file contents - hash changes - cache invalidates.
    src.write_bytes(b"changed")
    (digests_dir / "raw" / "test.txt").write_text("v2", encoding="utf-8")

    r2 = digest_single_source(source_path=src, extracted_text_path=digests_dir/"raw"/"test.txt",
                              digests_dir=digests_dir, llm_caller=caller,
                              deponent_name="Jane", deponent_role="P")
    assert r2.from_cache is False
    assert caller.call.call_count == 2


def test_digest_single_source_strips_json_fences(tmp_path):
    src = tmp_path / "test.pdf"
    src.write_bytes(b"x")
    digests_dir = tmp_path / "digests"
    digests_dir.mkdir()
    (digests_dir / "raw").mkdir()
    (digests_dir / "raw" / "test.txt").write_text("...", encoding="utf-8")

    fenced = "```json\n" + json.dumps(_good_payload()) + "\n```"
    caller = MagicMock()
    caller.call = MagicMock(return_value=fenced)

    result = digest_single_source(source_path=src, extracted_text_path=digests_dir/"raw"/"test.txt",
                                  digests_dir=digests_dir, llm_caller=caller,
                                  deponent_name="J", deponent_role="P")
    assert result.from_cache is False
    # The JSON inside the fence is parsed; the digest file is written.
    assert (digests_dir / "test.pdf.json").exists() or any(p.suffix == ".json" for p in digests_dir.glob("*.json"))


def test_digest_single_source_raises_on_invalid_json(tmp_path):
    src = tmp_path / "test.pdf"
    src.write_bytes(b"x")
    digests_dir = tmp_path / "digests"
    digests_dir.mkdir()
    (digests_dir / "raw").mkdir()
    (digests_dir / "raw" / "test.txt").write_text("...", encoding="utf-8")

    caller = MagicMock()
    caller.call = MagicMock(return_value="not json at all")

    with pytest.raises(ValueError, match="JSON"):
        digest_single_source(source_path=src, extracted_text_path=digests_dir/"raw"/"test.txt",
                             digests_dir=digests_dir, llm_caller=caller,
                             deponent_name="J", deponent_role="P")
