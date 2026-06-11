from icharlotte_core.discovery import _io
from icharlotte_core.doc_library.extract import Extracted
from icharlotte_core.doc_library.library import DocumentLibrary


def test_read_document_text_uses_case_library_cache(tmp_path, monkeypatch):
    pdf = tmp_path / "context.pdf"
    pdf.write_bytes(b"%PDF cached")
    DocumentLibrary(str(tmp_path)).get_or_extract_text(
        str(pdf),
        extractor=lambda path: Extracted("DISCOVERY CONTEXT", 1, "pdf_native", None),
    )

    class FailingFitz:
        @staticmethod
        def open(path):
            raise AssertionError("fitz should not run on cached text")

    monkeypatch.setattr(_io, "fitz", FailingFitz)

    text = _io.read_document_text(str(pdf), case_root=str(tmp_path))

    assert text == "DISCOVERY CONTEXT"
