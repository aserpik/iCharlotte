"""Depo Prep CLI agent.

Two phases:
  --phase=analyze --config=<path>    Reads config.json, runs Phase 1, emits
                                     AWAITING_INPUT:<session.json path>.
  --phase=generate --session=<path>  Reads session.json + topics.json (mutated
                                     by the wizard during topic editing), runs
                                     Phase 2, writes outline.docx + outline.md.

Prints PROGRESS:<int>:<msg> lines for the wizard's status page.
"""
from __future__ import annotations

import os
import sys
import traceback

# MUST come BEFORE any icharlotte_core import.
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


def _progress(n: int, msg: str = "") -> None:
    if msg:
        print(f"PROGRESS:{n}:{msg}", flush=True)
    else:
        print(f"PROGRESS:{n}", flush=True)


def _make_llm_caller():
    from icharlotte_core.llm_config import LLMCaller
    return LLMCaller()


def _cmd_analyze(config_path: str) -> int:
    from Scripts.depo_prep_lib import phase1
    from Scripts.depo_prep_lib.session_io import read_json

    config = read_json(config_path)
    try:
        session_json_path = phase1.run_phase1(
            config=config, llm_caller=_make_llm_caller(), progress=_progress,
        )
    except Exception as e:
        print(f"ERROR: Phase 1 failed: {e}", flush=True)
        traceback.print_exc()
        return 1

    print(f"AWAITING_INPUT:{session_json_path}", flush=True)
    return 0


def _cmd_generate(session_path: str) -> int:
    from Scripts.depo_prep_lib import phase2
    try:
        phase2.run_phase2(
            session_path=session_path, llm_caller=_make_llm_caller(),
            progress=_progress,
        )
    except Exception as e:
        print(f"ERROR: Phase 2 failed: {e}", flush=True)
        traceback.print_exc()
        return 1
    # Print the absolute path of the produced .docx so the wizard can pick it up.
    from pathlib import Path
    docx = Path(session_path).parent / "outline.docx"
    print(f"OUTPUT:{docx}", flush=True)
    return 0


def main():
    argv = sys.argv[1:]
    phase = None
    config = None
    session = None
    positional = []
    i = 0
    while i < len(argv):
        a = argv[i]
        if a.startswith("--phase="):
            phase = a.split("=", 1)[1]
            i += 1
        elif a == "--config" and i + 1 < len(argv):
            config = argv[i + 1]; i += 2
        elif a == "--session" and i + 1 < len(argv):
            session = argv[i + 1]; i += 2
        else:
            positional.append(a); i += 1

    if phase == "analyze":
        # Wizard form: --phase=analyze <config_path>
        if not config and positional:
            config = positional[0]
        if not config:
            print("ERROR: --config (or positional path) required for --phase=analyze", flush=True)
            return 2
        return _cmd_analyze(config)
    if phase == "generate":
        if not session and positional:
            session = positional[0]
        if not session:
            print("ERROR: --session (or positional path) required for --phase=generate", flush=True)
            return 2
        return _cmd_generate(session)
    print("ERROR: --phase=analyze|generate required", flush=True)
    return 2


if __name__ == "__main__":
    sys.exit(main())
