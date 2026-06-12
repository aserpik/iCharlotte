"""Headless runner for the in-process subpoena tracker worker."""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from icharlotte_core.subpoena_tracker import SubpoenaTrackerWorker  # noqa: E402


def run_tracker(case_path: str, file_number: str = "", output_root: str | None = None) -> dict:
    worker = SubpoenaTrackerWorker(
        case_path,
        file_number=file_number,
        output_root=output_root,
    )
    progress = []
    warnings = []
    finished = []

    worker.progress.connect(progress.append)
    worker.warning.connect(warnings.append)
    worker.finished_result.connect(
        lambda success, result: finished.append(
            {"success": bool(success), "result": str(result)}
        )
    )

    try:
        worker._run_phases()
    except Exception as exc:
        return {
            "success": False,
            "result": f"{type(exc).__name__}: {exc}",
            "progress": progress,
            "warnings": warnings,
        }

    if finished:
        payload = dict(finished[-1])
    else:
        payload = {"success": False, "result": "Subpoena tracker did not finish."}
    payload["progress"] = progress
    payload["warnings"] = warnings
    return payload


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Run the iCharlotte Subpoena Tracker.")
    parser.add_argument("--case-path", required=True, help="Path to the case folder.")
    parser.add_argument("--file-number", default="", help="Case file number for the report title.")
    parser.add_argument(
        "--output-root",
        default=None,
        help="Optional sandbox case root for NOTES/AI OUTPUT instead of writing to the case folder.",
    )
    parser.add_argument(
        "--json",
        action="store_true",
        help="Emit a JSON payload instead of human-readable status lines.",
    )
    args = parser.parse_args(argv)

    payload = run_tracker(args.case_path, args.file_number, args.output_root)
    if args.json:
        print(json.dumps(payload, indent=2))
    else:
        for message in payload["progress"]:
            print(message)
        for message in payload["warnings"]:
            print(f"WARNING: {message}", file=sys.stderr)
        label = "OK" if payload["success"] else "FAIL"
        print(f"{label}: {payload['result']}")

    return 0 if payload["success"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
