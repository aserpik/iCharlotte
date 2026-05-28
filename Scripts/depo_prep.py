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

import argparse
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
    parser = argparse.ArgumentParser(description="Depo Prep agent")
    parser.add_argument("--phase", required=True, choices=("analyze", "generate"))
    parser.add_argument("--config", default=None,
                        help="Path to config.json (required for --phase=analyze)")
    parser.add_argument("--session", default=None,
                        help="Path to session.json (required for --phase=generate)")
    args = parser.parse_args()

    if args.phase == "analyze":
        if not args.config:
            print("ERROR: --config is required for --phase=analyze", flush=True)
            return 2
        return _cmd_analyze(args.config)
    if args.phase == "generate":
        if not args.session:
            print("ERROR: --session is required for --phase=generate", flush=True)
            return 2
        return _cmd_generate(args.session)
    return 2


if __name__ == "__main__":
    sys.exit(main())
