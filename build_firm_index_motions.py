"""One-off: build the firm-brief index for MOTIONS, OPPOSITIONS, REPLIES only.

Excludes Pleadings, _Support*, _Other. "Motions" includes Ex Parte Applications
(moving papers). Reads the LOCAL sorted library (no Egnyte). OCR disabled for
speed (briefs with a native text layer only; rare image-only briefs are skipped).
Reuses the shipped firm_briefs helpers so the index format matches normal ingest.
"""
import os
import sys

# Resolve imports relative to this file (works from main checkout OR a worktree).
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from icharlotte_core.firm_briefs import factory
from icharlotte_core.firm_briefs.index import FirmBriefIndex
from icharlotte_core.firm_briefs.embedding import get_embedder
from icharlotte_core.firm_briefs.path_meta import meta_for_path
from icharlotte_core.firm_briefs.citation_harvest import harvest_cites
from icharlotte_core.firm_briefs.profile import extract_headings, compose_profile, profile_from_text
from icharlotte_core.firm_briefs.ingest import content_hash, _ocr_ratio
from icharlotte_core.document_processor import DocumentProcessor

ROOT = r"C:\geminiterminal2\5800_AMTRUST_Pleadings_PDFs"
ALLOWED_SIDES = {"moving", "opposition", "reply"}

# Pass 1 (default): OCR off, fast, indexes briefs with a native text layer.
# Pass 2 (`--ocr`): OCR on, slow, fills in image-only briefs that pass 1 left
# empty (skipped via has_current, so only the gaps get the expensive Tesseract).
OCR = "--ocr" in sys.argv


def extract(path: str) -> str:
    try:
        return DocumentProcessor().extract_text(path, ocr_enabled=OCR).text or ""
    except Exception as e:
        print("  extract error:", os.path.basename(path), e)
        return ""


def main() -> int:
    os.makedirs(factory.DATA_DIR, exist_ok=True)
    db, vec = factory.index_paths()
    idx = FirmBriefIndex(db_path=db, vectors_path=vec)
    idx.create_schema()
    print("Loading embedder (BGE-small)...")
    emb = get_embedder(fake=False)

    added = skipped = failed = 0
    cites_total = 0
    i = 0
    for dirpath, _dirs, files in os.walk(ROOT):
        for name in files:
            if not name.lower().endswith(".pdf"):
                continue
            path = os.path.join(dirpath, name)
            meta = meta_for_path(path, ROOT)
            if meta is None:
                continue
            mtype, side = meta
            if side not in ALLOWED_SIDES:
                continue
            i += 1
            h = content_hash(path)
            if idx.has_current(path, h):
                skipped += 1
                continue
            text = extract(path)
            if not text.strip():
                failed += 1
                print(f"  no-text (skipped): {name}")
                continue
            cites = harvest_cites(text)
            headings = extract_headings(text)
            profile = compose_profile("", headings, [c.proposition for c in cites]) or profile_from_text(text)
            vecrow = emb.encode([profile])[0]
            prop_vecs = list(emb.encode([c.proposition for c in cites])) if cites else None
            idx.upsert_brief(
                path=path, content_hash=h, motion_type=mtype, side=side,
                heading=headings[0] if headings else "", profile=profile,
                profile_vec=vecrow, char_len=len(text), ocr_ratio=_ocr_ratio(text),
                cites=cites, full_text=text, prop_vecs=prop_vecs,
            )
            cites_total += len(cites)
            added += 1
            if i % 25 == 0:
                print(f"[{i}] added={added} skipped={skipped} failed={failed} cites={cites_total}")

    print(f"=== BUILD DONE added={added} skipped={skipped} failed={failed} "
          f"cites={cites_total} stats={idx.stats()} ===")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
