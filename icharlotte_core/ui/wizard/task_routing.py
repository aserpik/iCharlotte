"""Routing rules for Wizard Mode task selection."""


_IN_PROCESS_TASK_BUILDERS = {
    "subpoena_tracker": "build_subpoena_tab",
    "med_record_extractor": "build_med_extractor_tab",
    "respond_to_discovery": "build_respond_to_discovery_tab",
    "motion_drafting": "build_motion_drafting_tab",
    "oppose_motion": "build_oppose_motion_tab",
    "separate": "build_separate_tab",
    "generate_motion": "build_generate_motion_tab",
    "mediation_brief": "build_mediation_brief_tab",
    "case_intake_docket": "build_case_intake_docket_tab",
}


# Tasks whose Settings page manages its own source selection. For these we skip
# the generic pre-Settings file picker and open the Settings screen directly —
# the page provides its own (often multi-bucket) file pickers, so a single
# pre-Settings picker would be redundant and its selection would be dropped.
_SETTINGS_MANAGED_SOURCE_TASKS = frozenset({
    "depo_prep",
    "medical_records",
    "motion_drafting",
    "oppose_motion",
    "summarize_documents",
    "summarize_discovery",
    "summarize_depositions",
})


def get_in_process_task_builder_name(task_id: str) -> str | None:
    """Return the builder name for wizard tasks backed by in-process workers."""
    return _IN_PROCESS_TASK_BUILDERS.get(task_id)


def is_in_process_task(task_id: str) -> bool:
    """Return whether a task is backed by an in-process wizard tab."""
    return get_in_process_task_builder_name(task_id) is not None


def opens_settings_without_picker(task_id: str) -> bool:
    """Whether this task opens its Settings page directly, with no pre-Settings
    file picker (the Settings page manages its own source selection)."""
    return task_id in _SETTINGS_MANAGED_SOURCE_TASKS


def requires_initial_file_picker(task_id: str) -> bool:
    """Whether selecting this wizard task should first show the file picker."""
    if task_id == "chat":
        return False
    if opens_settings_without_picker(task_id):
        return False
    return get_in_process_task_builder_name(task_id) is None
