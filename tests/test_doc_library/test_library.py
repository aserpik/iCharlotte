import os
from icharlotte_core.doc_library.library import DocumentLibrary


def test_folder_path_under_icharlotte(tmp_path):
    lib = DocumentLibrary(str(tmp_path))
    assert lib.folder == os.path.join(
        str(tmp_path), "NOTES", "AI OUTPUT", ".icharlotte", "doc_library")


def test_empty_library_lists_nothing(tmp_path):
    assert DocumentLibrary(str(tmp_path)).list_entries() == []


def test_save_then_reload_persists_entries(tmp_path):
    from icharlotte_core.doc_library.models import LibraryEntry, MemberFile
    lib = DocumentLibrary(str(tmp_path))
    entry = LibraryEntry("id1", "L", "L", "manual", "t",
                         [MemberFile("p", "p", "h.txt", 10, 3, "docx")])
    lib._save_entries([entry])
    reloaded = DocumentLibrary(str(tmp_path)).list_entries()
    assert len(reloaded) == 1
    assert reloaded[0].id == "id1"
    assert reloaded[0].members[0].blob == "h.txt"


def test_corrupt_index_is_recovered_not_raised(tmp_path):
    lib = DocumentLibrary(str(tmp_path))
    os.makedirs(lib.folder, exist_ok=True)
    with open(lib.index_path, "w", encoding="utf-8") as f:
        f.write("{ this is not valid json")
    assert lib.list_entries() == []  # backed up + reinitialized
    assert os.path.exists(lib.index_path + ".corrupt")


from icharlotte_core.doc_library.extract import Extracted


def _fake_extractor(text="DEPO TEXT"):
    return lambda path: Extracted(text=text, page_count=2,
                                  extract_method="pdf_native", error=None)


def test_add_entry_creates_blob_and_entry(tmp_path):
    src = tmp_path / "depo.pdf"
    src.write_bytes(b"%PDF-1.4 fake bytes")
    lib = DocumentLibrary(str(tmp_path))
    entry = lib.add_entry("summarize_depositions", [str(src)],
                          {"party": "Plaintiff"}, extractor=_fake_extractor())
    assert entry.label == "Plaintiff's Deposition Transcript"
    assert entry.members[0].char_count == len("DEPO TEXT")
    assert entry.members[0].est_tokens > 0
    blob = os.path.join(lib.blobs_dir, entry.members[0].blob)
    assert os.path.isfile(blob)
    assert lib.get_member_text(entry.members[0].blob) == "DEPO TEXT"
    assert len(lib.list_entries()) == 1


def test_same_file_in_two_entries_extracts_once(tmp_path):
    src = tmp_path / "depo.pdf"
    src.write_bytes(b"same bytes")
    calls = {"n": 0}

    def counting_extractor(path):
        calls["n"] += 1
        return Extracted("X", 1, "pdf_native", None)

    lib = DocumentLibrary(str(tmp_path))
    lib.add_entry("manual", [str(src)], {}, extractor=counting_extractor)
    lib.add_entry("summarize_documents", [str(src)], {}, extractor=counting_extractor)
    assert calls["n"] == 1  # second add reused the blob


def test_rerun_same_task_same_files_replaces_entry(tmp_path):
    src = tmp_path / "depo.pdf"
    src.write_bytes(b"bytes")
    lib = DocumentLibrary(str(tmp_path))
    lib.add_entry("summarize_depositions", [str(src)], {"party": "Plaintiff"},
                  extractor=_fake_extractor())
    lib.add_entry("summarize_depositions", [str(src)], {"party": "Plaintiff"},
                  extractor=_fake_extractor())
    assert len(lib.list_entries()) == 1  # replaced, not duplicated


def test_extraction_failure_records_error_member_no_blob(tmp_path):
    src = tmp_path / "bad.pdf"
    src.write_bytes(b"bytes")
    lib = DocumentLibrary(str(tmp_path))
    fail = lambda path: Extracted("", 0, "failed", "boom")
    entry = lib.add_entry("manual", [str(src)], {}, extractor=fail)
    assert entry.members[0].error == "boom"
    assert entry.members[0].blob is None


def test_rename_and_reset(tmp_path):
    src = tmp_path / "d.pdf"
    src.write_bytes(b"b")
    lib = DocumentLibrary(str(tmp_path))
    e = lib.add_entry("summarize_depositions", [str(src)], {"party": "Plaintiff"},
                      extractor=_fake_extractor())
    lib.rename_entry(e.id, "My Custom Name")
    assert lib.list_entries()[0].label == "My Custom Name"
    lib.reset_label(e.id)
    assert lib.list_entries()[0].label == "Plaintiff's Deposition Transcript"


def test_delete_entry_gcs_unreferenced_blob(tmp_path):
    src = tmp_path / "d.pdf"
    src.write_bytes(b"b")
    lib = DocumentLibrary(str(tmp_path))
    e = lib.add_entry("manual", [str(src)], {}, extractor=_fake_extractor())
    blob_path = os.path.join(lib.blobs_dir, e.members[0].blob)
    assert os.path.isfile(blob_path)
    lib.delete_entry(e.id)
    assert lib.list_entries() == []
    assert not os.path.isfile(blob_path)  # GC'd: no other entry referenced it


def test_delete_keeps_blob_referenced_by_another_entry(tmp_path):
    src = tmp_path / "d.pdf"
    src.write_bytes(b"b")
    lib = DocumentLibrary(str(tmp_path))
    e1 = lib.add_entry("manual", [str(src)], {}, extractor=_fake_extractor())
    e2 = lib.add_entry("summarize_documents", [str(src)], {}, extractor=_fake_extractor())
    blob_path = os.path.join(lib.blobs_dir, e1.members[0].blob)
    lib.delete_entry(e1.id)
    assert os.path.isfile(blob_path)  # still referenced by e2
