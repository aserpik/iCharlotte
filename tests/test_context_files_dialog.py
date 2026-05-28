import pytest

pytest.importorskip("pytestqt")

from PySide6.QtWidgets import QDialog

from icharlotte_core.ui.context_files_dialog import ContextFilesDialog


def test_add_files_accumulates_across_multiple_folders(qtbot, monkeypatch):
    dlg = ContextFilesDialog(start_dir="/start")
    qtbot.addWidget(dlg)

    calls = [
        ([r"C:\case\DISCOVERY\smith.pdf"], ""),
        ([r"C:\case\RECORDS\med.pdf", r"C:\case\PLEADINGS\complaint.docx"], ""),
    ]

    def fake_get_open(*_args, **_kwargs):
        return calls.pop(0)

    monkeypatch.setattr(
        "icharlotte_core.ui.context_files_dialog.QFileDialog.getOpenFileNames",
        fake_get_open,
    )

    dlg._on_add_files()
    dlg._on_add_files()

    assert dlg.selected_files() == [
        r"C:\case\DISCOVERY\smith.pdf",
        r"C:\case\RECORDS\med.pdf",
        r"C:\case\PLEADINGS\complaint.docx",
    ]


def test_add_files_dedupes_case_insensitively(qtbot, monkeypatch):
    dlg = ContextFilesDialog()
    qtbot.addWidget(dlg)

    calls = [
        ([r"C:\case\smith.pdf"], ""),
        ([r"c:\CASE\Smith.pdf"], ""),  # same file, different case
    ]
    monkeypatch.setattr(
        "icharlotte_core.ui.context_files_dialog.QFileDialog.getOpenFileNames",
        lambda *_a, **_k: calls.pop(0),
    )

    dlg._on_add_files()
    dlg._on_add_files()

    assert len(dlg.selected_files()) == 1


def test_remove_selected_drops_rows(qtbot):
    dlg = ContextFilesDialog(initial=[r"C:\a\one.pdf", r"C:\b\two.pdf", r"C:\c\three.pdf"])
    qtbot.addWidget(dlg)

    dlg.list_widget.item(1).setSelected(True)
    dlg._on_remove_selected()

    assert dlg.selected_files() == [r"C:\a\one.pdf", r"C:\c\three.pdf"]


def test_remove_selected_with_nothing_selected_is_noop(qtbot):
    dlg = ContextFilesDialog(initial=[r"C:\a\one.pdf"])
    qtbot.addWidget(dlg)

    dlg._on_remove_selected()  # nothing selected

    assert dlg.selected_files() == [r"C:\a\one.pdf"]


def test_initial_paths_are_listed(qtbot):
    dlg = ContextFilesDialog(initial=[r"C:\a\one.pdf", r"C:\b\two.pdf"])
    qtbot.addWidget(dlg)

    assert dlg.selected_files() == [r"C:\a\one.pdf", r"C:\b\two.pdf"]
    assert dlg.list_widget.count() == 2


def test_get_files_returns_none_on_cancel(qtbot, monkeypatch):
    monkeypatch.setattr(
        ContextFilesDialog, "exec", lambda self: QDialog.DialogCode.Rejected
    )
    result = ContextFilesDialog.get_files(None, start_dir="/x")
    assert result is None


def test_get_files_returns_list_on_accept(qtbot, monkeypatch):
    def fake_exec(self):
        self._add_paths([r"C:\a\one.pdf"])
        return QDialog.DialogCode.Accepted

    monkeypatch.setattr(ContextFilesDialog, "exec", fake_exec)
    result = ContextFilesDialog.get_files(None)
    assert result == [r"C:\a\one.pdf"]
