"""Routing rules for Wizard Mode task selection."""


_IN_PROCESS_TASK_BUILDERS = {
    "subpoena_tracker": "build_subpoena_tab",
    "respond_to_discovery": "build_respond_to_discovery_tab",
    "oppose_motion": "build_oppose_motion_tab",
}


def get_in_process_task_builder_name(task_id: str) -> str | None:
    """Return the builder name for wizard tasks backed by in-process workers."""
    return _IN_PROCESS_TASK_BUILDERS.get(task_id)


def is_in_process_task(task_id: str) -> bool:
    """Return whether a task is backed by an in-process wizard tab."""
    return get_in_process_task_builder_name(task_id) is not None


def requires_initial_file_picker(task_id: str) -> bool:
    """Whether selecting this wizard task should first show the file picker."""
    if task_id == "chat":
        return False
    return get_in_process_task_builder_name(task_id) is None
