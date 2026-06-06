import json


def test_start_emit_finish_records_buffer_and_jsonl(tmp_path):
    import icharlotte_core.task_debug as td

    td.reset_for_tests(trace_dir=tmp_path)

    run_id = td.start_run(
        task_id="chat_legal_research",
        task_title="Chat Legal Research",
        source="ChatTab",
        details={"api_token": "secret", "file_count": 2},
    )
    td.emit_event(
        run_id=run_id,
        task_id="chat_legal_research",
        task_title="Chat Legal Research",
        phase="search",
        level="info",
        message="Searching local corpus",
        source="Local California corpus",
        details={"hits": 3},
    )
    td.finish_run(run_id, status="success", message="Research complete")

    events = td.get_events()
    assert [e.phase for e in events] == ["start", "search", "finish"]
    assert events[0].details["api_token"] == "[REDACTED]"
    assert events[0].elapsed_ms == 0
    assert events[1].elapsed_ms >= 0
    trace_files = list(tmp_path.glob("*.jsonl"))
    assert len(trace_files) == 1
    text = trace_files[0].read_text(encoding="utf-8")
    assert "Searching local corpus" in text
    assert "secret" not in text


def test_recorder_coerces_non_json_details_and_bounds_buffer(tmp_path):
    import icharlotte_core.task_debug as td

    td.reset_for_tests(trace_dir=tmp_path, max_events=2)
    run_id = td.start_run(task_id="t", task_title="Task")
    td.emit_event(
        run_id=run_id,
        task_id="t",
        task_title="Task",
        phase="one",
        message="one",
        details={"bad": object()},
    )
    td.emit_event(
        run_id=run_id,
        task_id="t",
        task_title="Task",
        phase="two",
        message="two",
    )
    td.emit_event(
        run_id=run_id,
        task_id="t",
        task_title="Task",
        phase="three",
        message="three",
    )
    assert [event.phase for event in td.get_events()] == ["two", "three"]


def test_bridge_emits_recorded_event(qtbot, tmp_path):
    import icharlotte_core.task_debug as td

    td.reset_for_tests(trace_dir=tmp_path)
    bridge = td.get_bridge()
    seen = []
    bridge.event_emitted.connect(seen.append)
    run_id = td.start_run(task_id="task", task_title="Task")
    td.emit_event(
        run_id=run_id,
        task_id="task",
        task_title="Task",
        phase="status",
        message="working",
    )
    assert any(event.message == "working" for event in seen)


def test_details_are_strict_json_safe_and_redact_nested_sensitive_values(tmp_path):
    import icharlotte_core.task_debug as td

    td.reset_for_tests(trace_dir=tmp_path)
    run_id = td.start_run(
        task_id="strict_json",
        task_title="Strict JSON",
        details={
            "nan_value": float("nan"),
            "inf_value": float("inf"),
            "nested": {
                "password": "secret",
                "items": [{"api_key": "hidden"}],
            },
        },
    )

    event = td.get_events()[0]
    assert event.details["nested"]["password"] == "[REDACTED]"
    assert event.details["nested"]["items"][0]["api_key"] == "[REDACTED]"
    assert isinstance(event.details["nan_value"], str)
    assert isinstance(event.details["inf_value"], str)

    text = next(tmp_path.glob("*.jsonl")).read_text(encoding="utf-8")
    assert "NaN" not in text
    assert "Infinity" not in text
    assert "secret" not in text
    assert "hidden" not in text
    for line in text.splitlines():
        json.loads(
            line,
            parse_constant=lambda constant: (_ for _ in ()).throw(ValueError(constant)),
        )

    td.finish_run(run_id)


def test_cyclic_details_do_not_drop_event(tmp_path):
    import icharlotte_core.task_debug as td

    td.reset_for_tests(trace_dir=tmp_path)
    details = {"name": "cycle"}
    details["self"] = details

    run_id = td.start_run(task_id="cycle", task_title="Cycle")
    td.emit_event(
        run_id=run_id,
        task_id="cycle",
        task_title="Cycle",
        phase="cyclic",
        message="cyclic details",
        details=details,
    )

    events = td.get_events()
    assert [event.phase for event in events] == ["start", "cyclic"]
    assert events[1].details["self"] == "[CYCLE]"


def test_write_failures_do_not_raise_or_drop_buffered_event(tmp_path):
    import icharlotte_core.task_debug as td

    not_a_dir = tmp_path / "task_debug_file"
    not_a_dir.write_text("not a directory", encoding="utf-8")
    td.reset_for_tests(trace_dir=not_a_dir)

    run_id = td.start_run(task_id="write_failure", task_title="Write Failure")
    td.emit_event(
        run_id=run_id,
        task_id="write_failure",
        task_title="Write Failure",
        phase="still-buffered",
        message="write failed but buffer works",
    )

    assert [event.phase for event in td.get_events()] == ["start", "still-buffered"]


def test_finish_details_cannot_override_status(tmp_path):
    import icharlotte_core.task_debug as td

    td.reset_for_tests(trace_dir=tmp_path)
    run_id = td.start_run(task_id="finish", task_title="Finish")
    td.finish_run(
        run_id,
        status="success",
        message="done",
        details={"status": "failure", "note": "caller detail"},
    )

    finish_event = td.get_events()[-1]
    assert finish_event.details["status"] == "success"
    assert finish_event.details["note"] == "caller detail"
