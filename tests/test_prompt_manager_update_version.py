"""Tests for PromptManager.update_version() — in-place overwrite of a version.

Regression coverage for the "Save to Current" workbench button, which must
overwrite the selected version in place rather than create a new version.
"""
import os
import shutil
import tempfile
import unittest

from icharlotte_core.prompt_manager import PromptManager


class TestUpdateVersion(unittest.TestCase):

    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        self.pm = PromptManager(prompts_dir=self.tmp)

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)

    def test_update_overwrites_content_in_place(self):
        self.pm.create_version("word_assistant", "system_prompt", "ORIGINAL",
                               version="v1", set_as_current=True)

        ok = self.pm.update_version("word_assistant", "system_prompt", "v1", "EDITED")

        self.assertTrue(ok)
        self.assertEqual(
            self.pm.get_prompt("word_assistant", "system_prompt", "v1"), "EDITED")

    def test_update_does_not_create_a_new_version(self):
        self.pm.create_version("word_assistant", "system_prompt", "ORIGINAL",
                               version="v1", set_as_current=True)

        self.pm.update_version("word_assistant", "system_prompt", "v1", "EDITED")

        versions = self.pm.list_versions("word_assistant", "system_prompt")
        self.assertEqual(len(versions), 1)
        self.assertEqual(versions[0].version, "v1")

    def test_update_refreshes_current_pointer_when_version_is_active(self):
        self.pm.create_version("word_assistant", "system_prompt", "ORIGINAL",
                               version="v1", set_as_current=True)

        self.pm.update_version("word_assistant", "system_prompt", "v1", "EDITED")

        # The runtime reads "current"; it must reflect the in-place edit.
        self.assertEqual(
            self.pm.get_prompt("word_assistant", "system_prompt", "current"), "EDITED")

    def test_update_of_non_active_version_leaves_current_pointer_untouched(self):
        self.pm.create_version("word_assistant", "system_prompt", "ONE",
                               version="v1", set_as_current=True)
        self.pm.create_version("word_assistant", "system_prompt", "TWO",
                               version="v2", set_as_current=True)  # v2 is now current

        ok = self.pm.update_version("word_assistant", "system_prompt", "v1", "ONE-EDITED")

        self.assertTrue(ok)
        self.assertEqual(
            self.pm.get_prompt("word_assistant", "system_prompt", "v1"), "ONE-EDITED")
        # Current pointer still points at v2's content.
        self.assertEqual(
            self.pm.get_prompt("word_assistant", "system_prompt", "current"), "TWO")

    def test_update_unknown_version_returns_false(self):
        self.pm.create_version("word_assistant", "system_prompt", "ORIGINAL",
                               version="v1", set_as_current=True)

        ok = self.pm.update_version("word_assistant", "system_prompt", "v99", "X")

        self.assertFalse(ok)
        # No stray file created for the unknown version.
        self.assertFalse(os.path.exists(
            os.path.join(self.tmp, "word_assistant", "system_prompt_v99.txt")))

    def test_update_unknown_prompt_returns_false(self):
        ok = self.pm.update_version("no_agent", "no_pass", "v1", "X")
        self.assertFalse(ok)


if __name__ == "__main__":
    unittest.main()
