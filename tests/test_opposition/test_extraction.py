from icharlotte_core.document_processor import ExtractResult
from icharlotte_core.opposition import extraction


def test_validate_motion_file_accepts_pdf_and_docx(tmp_path):
    pdf = tmp_path / "motion.pdf"
    docx = tmp_path / "motion.docx"
    txt = tmp_path / "motion.txt"
    missing = tmp_path / "missing.pdf"
    for path in (pdf, docx, txt):
        path.write_text("content")

    assert extraction.validate_motion_file(str(pdf)) == (True, "")
    assert extraction.validate_motion_file(str(docx)) == (True, "")

    txt_valid, txt_warning = extraction.validate_motion_file(str(txt))
    assert txt_valid is False
    assert ".txt" in txt_warning

    missing_valid, missing_warning = extraction.validate_motion_file(str(missing))
    assert missing_valid is False
    assert "not found" in missing_warning.lower()


def test_validate_context_files_accepts_supported_and_warns_for_skipped(tmp_path):
    supported = []
    for suffix in (".pdf", ".docx", ".txt", ".msg"):
        path = tmp_path / f"context{suffix}"
        path.write_text("content")
        supported.append(str(path))

    unsupported = tmp_path / "context.xlsx"
    unsupported.write_text("content")
    missing = tmp_path / "missing.txt"

    valid, warnings = extraction.validate_context_files(
        supported + [str(unsupported), str(missing)]
    )

    assert valid == supported
    assert len(warnings) == 2
    assert any(".xlsx" in warning for warning in warnings)
    assert any("not found" in warning.lower() for warning in warnings)


def test_extract_context_bundle_labels_documents_without_citation_strings(
    tmp_path, monkeypatch
):
    first = tmp_path / "first.txt"
    second = tmp_path / "second.docx"
    skipped = tmp_path / "skipped.xlsx"
    first.write_text("ignored")
    second.write_text("ignored")
    skipped.write_text("ignored")

    def fake_extract_document_text(path, *, ocr_enabled=True):
        return ExtractResult(
            text=f"extracted text from {path}",
            page_count=1,
            file_path=path,
        )

    monkeypatch.setattr(extraction, "extract_document_text", fake_extract_document_text)

    bundle, warnings = extraction.extract_context_bundle(
        [str(first), str(second), str(skipped)]
    )

    assert "Document: first.txt" in bundle
    assert "Document: second.docx" in bundle
    assert "extracted text from" in bundle
    assert "skipped.xlsx" not in bundle
    assert any(".xlsx" in warning for warning in warnings)
    assert "[first.txt]" not in bundle
