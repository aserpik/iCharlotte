import unittest

from icharlotte_core.ui.wizard.task_routing import (
    get_in_process_task_builder_name,
    is_in_process_task,
    requires_initial_file_picker,
)


class WizardTaskRoutingTests(unittest.TestCase):
    def test_subpoena_tracker_uses_visible_in_process_route_without_file_picker(self):
        self.assertEqual(
            get_in_process_task_builder_name("subpoena_tracker"),
            "build_subpoena_tab",
        )
        self.assertFalse(requires_initial_file_picker("subpoena_tracker"))

    def test_document_summary_still_uses_initial_file_picker(self):
        self.assertIsNone(get_in_process_task_builder_name("summarize_documents"))
        self.assertTrue(requires_initial_file_picker("summarize_documents"))

    def test_oppose_motion_is_in_process_task(self):
        self.assertTrue(is_in_process_task("oppose_motion"))
        self.assertFalse(is_in_process_task("summarize_documents"))


if __name__ == "__main__":
    unittest.main()
