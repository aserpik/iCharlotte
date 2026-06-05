"""Phase 1 orchestrator — runs in-process; depo_prep.py calls run_phase1()."""
from __future__ import annotations

import os
import shutil
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from pathlib import Path
from typing import Callable, List

from .session_io import compute_session_paths, compute_source_cache_paths, write_json
from .source_digest import DigestResult, digest_single_source
from .topics import cluster_topics


_TEXT_EXTS = {".txt", ".md"}


def _extract_text_to_raw(source_path: Path, raw_dir: Path, logger=None) -> Path:
    """Extract text from source_path into raw_dir/<source>.txt and return that path.

    Uses icharlotte_core.document_processor for PDFs/DOCX; plain text is copied.
    """
    raw_dir.mkdir(parents=True, exist_ok=True)
    out = raw_dir / f"{source_path.name}.txt"
    ext = source_path.suffix.lower()
    if ext in _TEXT_EXTS:
        out.write_text(source_path.read_text(encoding="utf-8", errors="replace"),
                       encoding="utf-8")
        return out
    if ext == ".pdf":
        from icharlotte_core.document_processor import DocumentProcessor, OCRConfig
        processor = DocumentProcessor(ocr_config=OCRConfig(adaptive=True), logger=logger)
        result = processor.extract_with_dynamic_ocr(str(source_path))
        text = result.text if result.success else ""
        out.write_text(text, encoding="utf-8")
        return out
    if ext == ".docx":
        from icharlotte_core.document_processor import extract_docx_text
        text = extract_docx_text(str(source_path))
        out.write_text(text or "", encoding="utf-8")
        return out
    # Unsupported types: write empty text and let the LLM see an empty source.
    out.write_text("", encoding="utf-8")
    return out


def _session_raw_path(raw_dir: Path, source_path: Path) -> Path:
    return raw_dir / f"{source_path.name}.txt"


def _copy_text_file(source: Path, target: Path) -> None:
    target.parent.mkdir(parents=True, exist_ok=True)
    if source == target:
        return
    shutil.copyfile(source, target)


def _get_or_extract_cached_raw(
    source_path: Path,
    cache_raw_dir: Path,
    cache_raw_path: Path,
    logger=None,
) -> tuple[Path, bool]:
    if cache_raw_path.exists():
        return cache_raw_path, True
    extracted_path = _extract_text_to_raw(source_path, cache_raw_dir, logger=logger)
    return extracted_path, False


def _mirror_digest_to_session(result: DigestResult, session_digest_path: Path) -> None:
    write_json(session_digest_path, result.digest_data)
    cache_hash_path = result.digest_path.with_suffix(result.digest_path.suffix + ".sha256")
    if cache_hash_path.exists():
        session_digest_path.with_suffix(session_digest_path.suffix + ".sha256").write_text(
            cache_hash_path.read_text(encoding="utf-8"),
            encoding="utf-8",
        )


def run_phase1(*, config: dict, llm_caller, progress: Callable[[int, str], None]) -> str:
    """Execute Phase 1: ingest → per-source digest → topic clustering → persist.

    Returns the absolute path to session.json.
    """
    deponent_name = config["deponent_name"]
    deponent_role = config.get("deponent_role", "")
    style = config.get("style", "discovery")
    free_text = config.get("free_text_notes", "")
    case_root = config["case_root"]

    all_sources = list(config.get("deponent_sources", [])) + list(config.get("context_sources", []))
    if not all_sources:
        raise ValueError("Phase 1 requires at least one source file")

    paths = compute_session_paths(
        case_root=case_root, deponent_name=deponent_name,
        when_iso=datetime.now().isoformat(timespec="minutes"),
    )
    paths.session_dir.mkdir(parents=True, exist_ok=True)
    paths.digests_dir.mkdir(parents=True, exist_ok=True)
    paths.raw_dir.mkdir(parents=True, exist_ok=True)

    progress(5, "Extracting source text…")
    extracted_map: dict = {}
    source_cache_map: dict = {}
    for i, src_str in enumerate(all_sources, 1):
        src = Path(src_str)
        cache_paths = compute_source_cache_paths(case_root, src)
        source_cache_map[src] = cache_paths
        raw_path, from_cache = _get_or_extract_cached_raw(
            src,
            cache_paths.raw_dir,
            cache_paths.raw_text,
        )
        _copy_text_file(raw_path, _session_raw_path(paths.raw_dir, src))
        extracted_map[src] = raw_path
        verb = "Reused cached text for" if from_cache else "Extracted"
        progress(5 + int(15 * i / len(all_sources)), f"{verb} {src.name}")

    progress(25, "Building per-source digests…")
    digests: List[dict] = []

    def _one(src_path: Path):
        result = digest_single_source(
            source_path=src_path,
            extracted_text_path=extracted_map[src_path],
            digests_dir=source_cache_map[src_path].digests_dir,
            llm_caller=llm_caller,
            deponent_name=deponent_name,
            deponent_role=deponent_role,
        )
        _mirror_digest_to_session(result, paths.digests_dir / f"{src_path.name}.json")
        return result

    with ThreadPoolExecutor(max_workers=4) as pool:
        futures = {pool.submit(_one, Path(s)): s for s in all_sources}
        done = 0
        for fut in as_completed(futures):
            done += 1
            try:
                result = fut.result()
                digests.append(result.digest_data)
            except Exception as e:
                digests.append({
                    "source_id": Path(futures[fut]).name,
                    "source_kind": "other",
                    "deponent_statements": [], "factual_anchors": [],
                    "inconsistencies": [],
                    "summary": f"DIGEST FAILED: {e}",
                })
            progress(25 + int(50 * done / len(all_sources)),
                     f"Digested {done}/{len(all_sources)}")

    progress(80, "Clustering topics…")
    topics_result = cluster_topics(
        digests=digests, llm_caller=llm_caller,
        deponent_name=deponent_name, deponent_role=deponent_role,
        style=style, free_text_notes=free_text,
    )

    write_json(paths.topics_json, {
        "topics": topics_result.topics,
        "warning": topics_result.warning,
    })

    write_json(paths.session_json, {
        "version": 1,
        "phase": "awaiting_input",
        "deponent_name": deponent_name,
        "deponent_role": deponent_role,
        "style": style,
        "free_text_notes": free_text,
        "per_topic_flags": config.get("per_topic_flags", {}),
        "case_root": case_root,
        "deponent_sources": list(config.get("deponent_sources", [])),
        "context_sources": list(config.get("context_sources", [])),
        "digests_index": [d.get("source_id") for d in digests],
        "topics_warning": topics_result.warning,
    })

    progress(95, f"Discovered {len(topics_result.topics)} topics")
    return str(paths.session_json)
