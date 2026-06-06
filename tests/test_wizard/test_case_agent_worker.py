import os

from icharlotte_core.ui.wizard.runners.case_agent_worker import CaseAgentWorker


def test_case_agent_worker_builds_file_number_command(monkeypatch):
    worker = CaseAgentWorker(
        script_name="complaint.py",
        case_path=r"C:\cases\1234",
        file_number="1234.001",
    )
    monkeypatch.setattr(worker, "_script_path", lambda: r"C:\repo\Scripts\complaint.py")

    assert worker.command_argv() == [
        r"C:\repo\Scripts\complaint.py",
        "1234.001",
        "--headless",
    ]


def test_case_agent_worker_preserves_extra_flags_after_headless(monkeypatch):
    worker = CaseAgentWorker(
        script_name="docket.py",
        case_path=r"C:\cases\1234",
        file_number="1234.001",
        extra_flags=["--headful"],
    )
    monkeypatch.setattr(worker, "_script_path", lambda: r"C:\repo\Scripts\docket.py")

    assert worker.command_argv() == [
        r"C:\repo\Scripts\docket.py",
        "1234.001",
        "--headless",
        "--headful",
    ]


def test_case_agent_worker_keeps_recent_lines_bounded():
    worker = CaseAgentWorker(
        script_name="docket.py",
        case_path=r"C:\cases\1234",
        file_number="1234.001",
        recent_line_limit=3,
    )

    for line in ["one", "two", "three", "four"]:
        worker._handle_line(line)

    assert worker.recent_lines == ["two", "three", "four"]


def test_case_agent_worker_script_path_points_to_scripts_folder():
    worker = CaseAgentWorker(
        script_name="docket.py",
        case_path=r"C:\cases\1234",
        file_number="1234.001",
    )

    path = worker._script_path()

    assert path.endswith(os.path.join("Scripts", "docket.py"))
