import unittest

from icharlotte_core.ui.wizard.registry import get_task, list_tasks
from icharlotte_core.ui.wizard.task_routing import (
    get_in_process_task_builder_name,
    requires_initial_file_picker,
)


class RespondToDiscoveryRegistryTests(unittest.TestCase):
    def test_task_registered(self):
        ids = {task.task_id for task in list_tasks()}
        self.assertIn("respond_to_discovery", ids)
        spec = get_task("respond_to_discovery")
        self.assertEqual(spec.title, "Respond to Discovery")
        self.assertEqual(spec.default_folders, ["DISCOVERY/PROPOUNDED", "DISCOVERY"])

    def test_task_uses_in_process_builder_without_generic_picker(self):
        self.assertEqual(
            get_in_process_task_builder_name("respond_to_discovery"),
            "build_respond_to_discovery_tab",
        )
        self.assertFalse(requires_initial_file_picker("respond_to_discovery"))


if __name__ == "__main__":
    unittest.main()
