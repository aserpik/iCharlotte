"""Fill the firm index for image-only briefs the native pass left empty.

The DocumentProcessor OCR path needs pdf2image (not installed in this venv), so
we OCR directly via PyMuPDF (fitz) -> Tesseract, which ARE installed. Only
processes allowed-side files not already in the index (the ~23 gaps). Caps very
large files (exhibits/records, not briefs) to bound runtime.
"""
import gc
import io
import os
import sys

# Resolve imports relative to this file (works from main checkout OR a worktree).
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import fitz  # PyMuPDF
import pytesseract
from PIL import Image

pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

from icharlotte_core.firm_briefs import factory
from icharlotte_core.firm_briefs.index import FirmBriefIndex
from icharlotte_core.firm_briefs.embedding import get_embedder
from icharlotte_core.firm_briefs.path_meta import meta_for_path
from icharlotte_core.firm_briefs.citation_harvest import harvest_cites
from icharlotte_core.firm_briefs.profile import extract_headings, compose_profile, profile_from_text
from icharlotte_core.firm_briefs.ingest import content_hash, _ocr_ratio

ROOT = r"C:\geminiterminal2\5800_AMTRUST_Pleadings_PDFs"
ALLOWED_SIDES = {"moving", "opposition", "reply"}
MAX_PAGES = 40   # skip giant exhibits/records (not briefs)
DPI = 150        # ample for text OCR; higher DPI caused leptonica malloc failures


def ocr_pdf(path: str) -> str:
    doc = fitz.open(path)
    try:
        if doc.page_count > MAX_PAGES:
            print(f"  skip-large ({doc.page_count} pages): {os.path.basename(path)}")
            return ""
        mat = fitz.Matrix(DPI / 72, DPI / 72)
        parts = []
        for i, page in enumerate(doc):
            try:
                pix = page.get_pixmap(matrix=mat)
                data = pix.tobytes("png")
                pix = None
                img = Image.open(io.BytesIO(data))
                parts.append(pytesseract.image_to_string(img) or "")
                img.close()
            except Exception as e:
                print(f"  page {i + 1} OCR failed ({type(e).__name__}); skipping page")
            finally:
                gc.collect()
        return "\n".join(parts)
    finally:
        doc.close()


def main() -> int:
    db, vec = factory.index_paths()
    idx = FirmBriefIndex(db_path=db, vectors_path=vec)
    idx.create_schema()
    _threads = int(os.environ.get("FB_EMBED_THREADS", "0") or "0")
    if _threads > 0:
        from icharlotte_core.legal_research.local_corpus.embedder import OnnxEmbedder
        emb = OnnxEmbedder(threads=_threads)
    else:
        emb = get_embedder(fake=False)
    _skip_prop = os.environ.get("FB_SKIP_PROP_VECS") == "1"
    added = failed = 0
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
            h = content_hash(path)
            if idx.has_current(path, h):
                continue  # already indexed by the native pass
            print(f"OCR: {name}")
            text = ocr_pdf(path)
            if not text.strip():
                failed += 1
                print(f"  still-empty: {name}")
                continue
            cites = harvest_cites(text)
            headings = extract_headings(text)
            profile = compose_profile("", headings, [c.proposition for c in cites]) or profile_from_text(text)
            vecrow = emb.encode([profile])[0]
            prop_vecs = None if _skip_prop else (list(emb.encode([c.proposition for c in cites])) if cites else None)
            idx.upsert_brief(
                path=path, content_hash=h, motion_type=mtype, side=side,
                heading=headings[0] if headings else "", profile=profile,
                profile_vec=vecrow, char_len=len(text), ocr_ratio=_ocr_ratio(text),
                cites=cites, full_text=text, prop_vecs=prop_vecs,
            )
            added += 1
            print(f"  indexed ({mtype}/{side}, {len(text)} chars, {len(cites)} cites)")
    print(f"=== OCR-FILL DONE added={added} still_empty={failed} stats={idx.stats()} ===")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
