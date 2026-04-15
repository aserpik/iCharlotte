"""Tests for the Import Reports feature in ChatTab."""
import os
import tempfile
import unittest
from types import SimpleNamespace
from unittest.mock import patch


class TestCarrierReportRegex(unittest.TestCase):
    """Verify CARRIER_REPORT_RE matches the spec's must/must-not lists."""

    def setUp(self):
        from icharlotte_core.ui.tabs import CARRIER_REPORT_RE
        self.pattern = CARRIER_REPORT_RE

    def _matches(self, name: str) -> bool:
        return self.pattern.match(name) is not None

    # --- Must match -----------------------------------------------------
    def test_basic_carrier001_docx(self):
        self.assertTrue(self._matches("carrier001.docx"))

    def test_carrier015_doc_lowercase(self):
        self.assertTrue(self._matches("carrier015.doc"))

    def test_trailing_space_and_parens(self):
        self.assertTrue(self._matches("carrier002 (FSR).docx"))

    def test_trailing_parens_no_space(self):
        self.assertTrue(self._matches("carrier003(lit plan).docx"))

    def test_uppercase_carrier_and_extension(self):
        self.assertTrue(self._matches("Carrier007.DOCX"))

    def test_trailing_dash_suffix(self):
        self.assertTrue(self._matches("carrier010 - Final.docx"))

    def test_all_caps_carrier(self):
        self.assertTrue(self._matches("CARRIER005.docx"))

    # --- Must NOT match -------------------------------------------------
    def test_bracket_prefix_rejected(self):
        self.assertFalse(self._matches("[draft]carrier001.docx"))

    def test_word_prefix_rejected(self):
        self.assertFalse(self._matches("draft_carrier001.docx"))

    def test_carrier000_below_range(self):
        self.assertFalse(self._matches("carrier000.docx"))

    def test_carrier016_above_range(self):
        self.assertFalse(self._matches("carrier016.docx"))

    def test_four_digit_run_rejected(self):
        self.assertFalse(self._matches("carrier0011.docx"))

    def test_wrong_extension_pdf(self):
        self.assertFalse(self._matches("carrier001.pdf"))

    def test_no_number(self):
        self.assertFalse(self._matches("carrier.docx"))

    def test_two_digit_number(self):
        self.assertFalse(self._matches("carrier01.docx"))

    def test_carrier100_out_of_range(self):
        self.assertFalse(self._matches("carrier100.docx"))


class TestImportCarrierReportsIntegration(unittest.TestCase):
    """End-to-end behavior of ChatTab.import_carrier_reports against a real temp FS.

    Calls the method as an unbound function against a minimal fake ``self`` with
    just the attributes the method touches (``attached_files``, ``window``,
    ``add_file``). Qt popup calls are captured via ``unittest.mock.patch``.
    """

    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="icharlotte_import_test_")
        self.case_path = os.path.join(self.tmp, "case_folder")
        os.makedirs(self.case_path)
        self.status_dir = os.path.join(self.case_path, "STATUS")

    def tearDown(self):
        import shutil
        shutil.rmtree(self.tmp, ignore_errors=True)

    def _make_files(self, names):
        os.makedirs(self.status_dir, exist_ok=True)
        for name in names:
            with open(os.path.join(self.status_dir, name), "w") as f:
                f.write("")

    _UNSET = object()

    def _fake_self(self, case_path=_UNSET):
        """Build a minimal object that looks like ChatTab for the method's purposes."""
        ns = SimpleNamespace()
        ns.attached_files = []
        effective = self.case_path if case_path is self._UNSET else case_path
        window_obj = SimpleNamespace(case_path=effective)
        ns.window = lambda: window_obj
        ns.add_file = lambda path: ns.attached_files.append(path)
        return ns

    def _run(self, fake_self):
        """Invoke import_carrier_reports, collecting popup calls."""
        from icharlotte_core.ui.tabs import ChatTab
        info_calls = []
        warn_calls = []
        with patch("icharlotte_core.ui.tabs.QMessageBox") as mock_box:
            mock_box.information.side_effect = lambda *a, **kw: info_calls.append(a)
            mock_box.warning.side_effect = lambda *a, **kw: warn_calls.append(a)
            ChatTab.import_carrier_reports(fake_self)
        return info_calls, warn_calls

    # --- Scenarios mirroring the manual test plan -----------------------

    def test_import_into_empty_list(self):
        """Plan step 4: attach 2 carrier files from an empty list."""
        self._make_files([
            "carrier001.docx",
            "carrier002 (FSR).docx",
            "[draft]carrier003.docx",   # rejected: prefix
            "random.docx",              # rejected: no carrier
            "notes.txt",                # rejected: wrong extension
        ])
        s = self._fake_self()
        infos, warns = self._run(s)

        self.assertEqual(len(s.attached_files), 2)
        basenames = sorted(os.path.basename(p) for p in s.attached_files)
        self.assertEqual(basenames, ["carrier001.docx", "carrier002 (FSR).docx"])
        self.assertEqual(len(infos), 1)
        self.assertIn("Imported 2 carrier report(s)", infos[0][2])
        self.assertEqual(warns, [])

    def test_second_click_all_already_attached(self):
        """Plan step 5: second click → 'already attached' popup, no duplicates."""
        self._make_files(["carrier001.docx", "carrier004.doc"])
        s = self._fake_self()
        self._run(s)  # first import
        # Second click
        infos, warns = self._run(s)

        self.assertEqual(len(s.attached_files), 2)  # no duplicates
        self.assertEqual(len(infos), 1)
        self.assertIn("All 2 matching report(s) were already attached", infos[0][2])

    def test_mixed_existing_and_new(self):
        """Plan step 6: unrelated file already in list → carrier files added alongside."""
        self._make_files(["carrier001.docx", "carrier002.docx"])
        s = self._fake_self()
        s.attached_files.append(r"C:\some\unrelated.pdf")

        infos, warns = self._run(s)

        self.assertEqual(len(s.attached_files), 3)
        self.assertIn(r"C:\some\unrelated.pdf", s.attached_files)

    def test_no_status_folder(self):
        """Plan step 7: STATUS missing → warning popup, no attachments."""
        # Intentionally do not create self.status_dir
        s = self._fake_self()

        infos, warns = self._run(s)

        self.assertEqual(s.attached_files, [])
        self.assertEqual(infos, [])
        self.assertEqual(len(warns), 1)
        self.assertIn("No STATUS folder found", warns[0][2])

    def test_no_case_loaded(self):
        """Plan step 8: no case selected → warning popup."""
        s = self._fake_self(case_path=None)

        infos, warns = self._run(s)

        self.assertEqual(s.attached_files, [])
        self.assertEqual(len(warns), 1)
        self.assertIn("No case is currently selected", warns[0][2])

    def test_zero_matches_in_status(self):
        """Plan step 9: STATUS exists but has no carrier files."""
        self._make_files(["random.docx", "notes.txt", "[draft]carrier001.docx"])
        s = self._fake_self()

        infos, warns = self._run(s)

        self.assertEqual(s.attached_files, [])
        self.assertEqual(len(infos), 1)
        self.assertIn("No carrier reports", infos[0][2])

    def test_rejects_out_of_range_carriers(self):
        """carrier000, 016, 100, 0011 must all be skipped even if extension matches."""
        self._make_files([
            "carrier000.docx",
            "carrier016.docx",
            "carrier100.docx",
            "carrier0011.docx",
            "carrier015.docx",  # the only one that should match
        ])
        s = self._fake_self()

        self._run(s)

        self.assertEqual(len(s.attached_files), 1)
        self.assertTrue(s.attached_files[0].endswith("carrier015.docx"))

    def test_normcase_dedup_against_preattached(self):
        """Issue #1 fix: pre-attached file with different case is detected as duplicate on Windows."""
        self._make_files(["carrier001.docx"])
        s = self._fake_self()
        # Simulate a pre-attached file with different case
        pre_attached = os.path.join(self.status_dir, "CARRIER001.DOCX")
        s.attached_files.append(pre_attached)

        infos, warns = self._run(s)

        # On Windows (normcase lowercases), this should be detected as already attached.
        # On Linux (normcase is identity), it would be treated as new. Only assert
        # the Windows behavior if we're on Windows.
        if os.name == "nt":
            self.assertEqual(len(s.attached_files), 1)  # no new addition
            self.assertIn("already attached", infos[0][2].lower())
        else:
            self.assertEqual(len(s.attached_files), 2)

    def test_add_file_raises_oserror_is_collected(self):
        """Issue #2 fix: one bad file doesn't abort the batch; failure is surfaced."""
        self._make_files(["carrier001.docx", "carrier002.docx", "carrier003.docx"])
        s = self._fake_self()

        # Make add_file raise on carrier002 only
        def flaky_add_file(path):
            if path.endswith("carrier002.docx"):
                raise OSError("simulated disk error")
            s.attached_files.append(path)

        s.add_file = flaky_add_file

        infos, warns = self._run(s)

        # The two good files should still be imported
        self.assertEqual(len(s.attached_files), 2)
        # Failure should be surfaced via warning popup
        self.assertEqual(len(warns), 1)
        self.assertIn("Failed to import 1 file(s)", warns[0][2])
        self.assertIn("carrier002.docx", warns[0][2])
        self.assertIn("simulated disk error", warns[0][2])
        self.assertIn("Imported 2 carrier report(s)", warns[0][2])


if __name__ == "__main__":
    unittest.main()
