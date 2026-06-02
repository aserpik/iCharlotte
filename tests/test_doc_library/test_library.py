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
