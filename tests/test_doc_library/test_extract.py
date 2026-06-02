from icharlotte_core.doc_library.extract import extract_any, Extracted


def test_extract_txt(tmp_path):
    p = tmp_path / "note.txt"
    p.write_text("hello world", encoding="utf-8")
    out = extract_any(str(p))
    assert isinstance(out, Extracted)
    assert "hello world" in out.text
    assert out.error is None


def test_extract_docx(tmp_path):
    from docx import Document
    p = tmp_path / "d.docx"
    doc = Document()
    doc.add_paragraph("First para")
    doc.save(str(p))
    out = extract_any(str(p))
    assert "First para" in out.text
    assert out.extract_method == "docx"
    assert out.error is None


def test_extract_missing_file_returns_error():
    out = extract_any(r"Z:\nope\missing.pdf")
    assert out.text == ""
    assert out.error  # non-empty error string
