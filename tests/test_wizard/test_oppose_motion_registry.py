import unittest

from icharlotte_core.ui.wizard.registry import get_task, list_tasks
from icharlotte_core.ui.wizard.task_routing import (
    get_in_process_task_builder_name,
    requires_initial_file_picker,
)


class OpposeMotionRegistryTests(unittest.TestCase):
    def test_task_registered(self):
        ids = {task.task_id for task in list_tasks()}
        self.assertIn("oppose_motion", ids)
        spec = get_task("oppose_motion")
        self.assertEqual(spec.title, "Oppose a Motion")
        self.assertEqual(spec.default_folders, ["MOTIONS", "PLEADINGS", "DISCOVERY"])
        self.assertEqual(spec.script_name, "")

    def test_task_uses_in_process_builder_without_generic_picker(self):
        self.assertEqual(
            get_in_process_task_builder_name("oppose_motion"),
            "build_oppose_motion_tab",
        )
        self.assertFalse(requires_initial_file_picker("oppose_motion"))


if __name__ == "__main__":
    unittest.main()
