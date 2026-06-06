import unittest

from icharlotte_core.ui.wizard.task_routing import (
    get_in_process_task_builder_name,
    is_in_process_task,
    opens_settings_without_picker,
    requires_initial_file_picker,
)


class WizardTaskRoutingTests(unittest.TestCase):
    def test_subpoena_tracker_uses_visible_in_process_route_without_file_picker(self):
        self.assertEqual(
            get_in_process_task_builder_name("subpoena_tracker"),
            "build_subpoena_tab",
        )
        self.assertFalse(requires_initial_file_picker("subpoena_tracker"))

    def test_med_record_extractor_uses_visible_in_process_route_without_file_picker(self):
        self.assertEqual(
            get_in_process_task_builder_name("med_record_extractor"),
            "build_med_extractor_tab",
        )
        self.assertTrue(is_in_process_task("med_record_extractor"))
        self.assertFalse(requires_initial_file_picker("med_record_extractor"))

    def test_summarize_tasks_open_settings_without_picker(self):
        # These tasks all have a Files box on Settings, so the launcher should
        # open Settings directly and let Add Files handle source selection.
        self.assertIsNone(get_in_process_task_builder_name("summarize_documents"))
        for task_id in (
            "summarize_documents",
            "summarize_discovery",
            "summarize_depositions",
        ):
            self.assertTrue(opens_settings_without_picker(task_id))
            self.assertFalse(requires_initial_file_picker(task_id))

    def test_oppose_motion_is_in_process_task(self):
        self.assertTrue(is_in_process_task("oppose_motion"))
        self.assertFalse(is_in_process_task("summarize_documents"))

    def test_case_intake_docket_uses_in_process_route_without_file_picker(self):
        self.assertEqual(
            get_in_process_task_builder_name("case_intake_docket"),
            "build_case_intake_docket_tab",
        )
        self.assertTrue(is_in_process_task("case_intake_docket"))
        self.assertFalse(requires_initial_file_picker("case_intake_docket"))

    def test_depo_prep_opens_settings_without_picker(self):
        # Depo Prep's Settings page has its own source pickers, so the generic
        # pre-Settings file picker must be skipped (otherwise its selection is
        # silently dropped).
        self.assertTrue(opens_settings_without_picker("depo_prep"))
        self.assertFalse(requires_initial_file_picker("depo_prep"))
        # It is NOT an in-process task — it still runs the subprocess agent.
        self.assertFalse(is_in_process_task("depo_prep"))

    def test_other_tasks_do_not_skip_picker(self):
        self.assertFalse(opens_settings_without_picker("medical_records"))


if __name__ == "__main__":
    unittest.main()
