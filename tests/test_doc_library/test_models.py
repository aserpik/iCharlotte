from icharlotte_core.doc_library.models import MemberFile, LibraryEntry


def test_memberfile_roundtrips_through_dict():
    m = MemberFile(
        source_path=r"Z:\case\Depo Vol 1.pdf",
        source_name="Depo Vol 1.pdf",
        blob="abc123.txt",
        char_count=100,
        est_tokens=25,
        extract_method="pdf_native",
        error=None,
    )
    assert MemberFile.from_dict(m.to_dict()) == m


def test_entry_roundtrips_and_defaults_members():
    e = LibraryEntry(
        id="id1",
        label="Plaintiff's Deposition Transcript",
        auto_label="Plaintiff's Deposition Transcript",
        task_type="summarize_depositions",
        created_at="2026-06-02T10:00:00",
        members=[MemberFile("p", "p", "h.txt")],
    )
    d = e.to_dict()
    assert d["members"][0]["blob"] == "h.txt"
    assert LibraryEntry.from_dict(d) == e


def test_entry_from_dict_tolerates_missing_members():
    e = LibraryEntry.from_dict({
        "id": "x", "label": "L", "auto_label": "L",
        "task_type": "manual", "created_at": "t",
    })
    assert e.members == []
