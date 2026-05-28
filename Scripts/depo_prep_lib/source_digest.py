"""Stage 1.2 - per-source structured digest with file-hash caching."""
from __future__ import annotations

import json
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Union

from .prompts import build_per_source_digest_prompt
from .schemas import validate_source_digest_dict
from .session_io import file_sha256, write_json, read_json


_FENCE_RE = re.compile(r"^```(?:json)?\s*\n(.*)\n```\s*$", re.DOTALL)


@dataclass
class DigestResult:
    digest_path: Path
    digest_data: dict
    from_cache: bool


def _strip_fences(s: str) -> str:
    s = (s or "").strip()
    m = _FENCE_RE.match(s)
    return m.group(1).strip() if m else s


def _parse_llm_json(raw: str) -> dict:
    text = _strip_fences(raw)
    try:
        return json.loads(text)
    except json.JSONDecodeError as e:
        raise ValueError(f"LLM returned invalid JSON: {e}") from e


def _cache_hash_path(digest_path: Path) -> Path:
    return digest_path.with_suffix(digest_path.suffix + ".sha256")


def digest_single_source(
    *,
    source_path: Union[str, Path],
    extracted_text_path: Union[str, Path],
    digests_dir: Union[str, Path],
    llm_caller,
    deponent_name: str,
    deponent_role: str,
) -> DigestResult:
    """Produce (or load from cache) the structured digest for one source file.

    Cache key = sha256 of source_path. Side-by-side .sha256 file holds the hash
    of the source the digest was produced from; if the source's current hash
    matches, we reuse the digest.
    """
    source_path = Path(source_path)
    digests_dir = Path(digests_dir)
    digests_dir.mkdir(parents=True, exist_ok=True)

    digest_path = digests_dir / f"{source_path.name}.json"
    hash_path = _cache_hash_path(digest_path)

    current_hash = file_sha256(source_path)

    if digest_path.exists() and hash_path.exists():
        cached_hash = hash_path.read_text(encoding="utf-8").strip()
        if cached_hash == current_hash:
            return DigestResult(
                digest_path=digest_path,
                digest_data=read_json(digest_path),
                from_cache=True,
            )

    extracted_text = Path(extracted_text_path).read_text(encoding="utf-8", errors="replace")

    prompt, text_payload = build_per_source_digest_prompt(
        deponent_name=deponent_name,
        deponent_role=deponent_role,
        source_text=extracted_text,
        source_filename=source_path.name,
    )

    raw = llm_caller.call(
        prompt=prompt,
        text=text_payload,
        task_type="extraction",
        agent_id="DepoPrep",
        pass_name="source_digest",
    )

    data = _parse_llm_json(raw)
    # Force source_id to match the actual filename in case the LLM hallucinated.
    data["source_id"] = source_path.name
    validate_source_digest_dict(data)

    write_json(digest_path, data)
    hash_path.write_text(current_hash, encoding="utf-8")

    return DigestResult(digest_path=digest_path, digest_data=data, from_cache=False)
