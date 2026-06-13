"""FastAPI app for the wizard web companion.

Entry point: ``python -m webcompanion.server`` (binds the Tailscale IP by
default; ``--lan`` binds 0.0.0.0 for same-network development).
"""
import argparse
import os
import shutil
import subprocess
import sys
from pathlib import Path

from fastapi import FastAPI, Request
from fastapi.responses import (FileResponse, HTMLResponse, JSONResponse,
                               RedirectResponse)
from fastapi.templating import Jinja2Templates

from . import cases
from . import jobs as J
from . import task_defs as T
from . import chat
from .job_manager import JobManager
from .jobs import JobStore, new_job

_CHAT_MANAGER = chat.ChatTurnManager()

_TEMPLATES_DIR = Path(__file__).parent / "templates"
DEFAULT_JOBS_PATH = T.REPO_ROOT / "logs" / "webcompanion" / "jobs.json"
_DOCX_MEDIA_TYPE = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document")


def _crumbs(path: str):
    """[('DISCOVERY', 'DISCOVERY'), ('RESPONSES', 'DISCOVERY/RESPONSES')]"""
    out, acc = [], []
    for part in (path or "").replace("\\", "/").split("/"):
        if not part:
            continue
        acc.append(part)
        out.append((part, "/".join(acc)))
    return out


def create_app(manager: JobManager) -> FastAPI:
    app = FastAPI(title="iCharlotte Web Companion")
    templates = Jinja2Templates(directory=str(_TEMPLATES_DIR))

    def _render(name, request, **ctx):
        return templates.TemplateResponse(request, name, ctx)

    # ---- home / case / picker ----

    @app.get("/", response_class=HTMLResponse)
    def home(request: Request, q: str = ""):
        return _render("home.html", request, jobs=manager.store.all()[:20],
                       cases=cases.list_cases(q), q=q, tasks=T.TASKS)

    @app.get("/case/{file_number}", response_class=HTMLResponse)
    def case_page(request: Request, file_number: str):
        case = cases.get_case(file_number)
        if case is None:
            return HTMLResponse("Case not found", status_code=404)
        return _render("case.html", request, case=case,
                       tasks=list(T.TASKS.values()))

    @app.get("/case/{file_number}/task/{task_id}", response_class=HTMLResponse)
    def picker(request: Request, file_number: str, task_id: str,
               path: str = None):
        case = cases.get_case(file_number)
        if case is None or task_id not in T.TASKS:
            return HTMLResponse("Not found", status_code=404)
        task = T.TASKS[task_id]
        if path is None:
            path = cases.resolve_start_folder(case["case_path"],
                                              task.default_folders)
        try:
            dirs, files = cases.browse(case["case_path"], path, task.file_exts)
        except ValueError:
            return HTMLResponse("Invalid path", status_code=400)
        return _render("picker.html", request, case=case, task=task,
                       path=path, dirs=dirs, files=files,
                       crumbs=_crumbs(path))

    @app.post("/case/{file_number}/task/{task_id}/start")
    async def start(request: Request, file_number: str, task_id: str):
        case = cases.get_case(file_number)
        if case is None or task_id not in T.TASKS:
            return HTMLResponse("Not found", status_code=404)
        task = T.TASKS[task_id]
        form = await request.form()
        rel_files = [f for f in form.getlist("files") if f]
        if not rel_files:
            return HTMLResponse("Pick at least one file.", status_code=400)
        try:
            abs_files = [str(cases.safe_resolve(case["case_path"], rf))
                         for rf in rel_files]
        except ValueError:
            return HTMLResponse("Invalid path", status_code=400)
        if task.two_phase and len(abs_files) > 1:
            return HTMLResponse(
                "This task accepts exactly one input file.", status_code=400)
        if task.pre_settings == "depo_prep":
            return _render("depo_prep_settings.html", request, case=case,
                           task=task, rel_files=rel_files,
                           styles=T.DEPO_PREP_STYLES)
        job = new_job(task_id, case["case_path"], file_number, abs_files)
        manager.submit(job)
        return RedirectResponse(f"/job/{job.id}", status_code=303)

    _register_job_routes(app, manager, templates)
    _register_depo_prep_routes(app, manager, templates)
    _register_awaiting_routes(app, manager, templates)
    _register_chat_routes(app, templates)
    return app


def _register_chat_routes(app, templates):
    def _render(name, request, **ctx):
        return templates.TemplateResponse(request, name, ctx)

    @app.get("/case/{file_number}/chat", response_class=HTMLResponse)
    def chat_list(request: Request, file_number: str):
        case = cases.get_case(file_number)
        if case is None:
            return HTMLResponse("Case not found", status_code=404)
        return _render("chat_conversations.html", request, case=case,
                       conversations=chat.list_conversations(file_number))

    @app.post("/case/{file_number}/chat/new")
    def chat_new(file_number: str):
        case = cases.get_case(file_number)
        if case is None:
            return HTMLResponse("Case not found", status_code=404)
        conv_id = chat.create_conversation(file_number)
        return RedirectResponse(f"/case/{file_number}/chat/{conv_id}",
                                status_code=303)

    @app.get("/case/{file_number}/chat/{conv_id}", response_class=HTMLResponse)
    def chat_view(request: Request, file_number: str, conv_id: str,
                  turn: str = None):
        case = cases.get_case(file_number)
        if case is None:
            return HTMLResponse("Case not found", status_code=404)
        conv = chat.get_conversation(file_number, conv_id)
        if conv is None:
            return HTMLResponse("Conversation not found", status_code=404)
        return _render("chat_conversation.html", request, case=case, conv=conv,
                       turn_id=turn)

    @app.post("/case/{file_number}/chat/{conv_id}/send")
    async def chat_send(request: Request, file_number: str, conv_id: str):
        case = cases.get_case(file_number)
        if case is None:
            return HTMLResponse("Case not found", status_code=404)
        conv = chat.get_conversation(file_number, conv_id)
        if conv is None:
            return HTMLResponse("Conversation not found", status_code=404)
        form = await request.form()
        message = (form.get("message") or "").strip()
        if not message:
            return HTMLResponse("Type a message.", status_code=400)
        provider = form.get("provider") or conv.provider
        model = form.get("model") or conv.model
        rel_files = [f for f in form.getlist("attach") if f]
        research_on = form.get("research") == "on"
        try:
            turn_id = _CHAT_MANAGER.start_turn(
                file_number, conv_id, user_text=message, provider=provider,
                model=model, attach_rel_files=rel_files, research_on=research_on)
        except ValueError as exc:
            return HTMLResponse(str(exc), status_code=409)
        return RedirectResponse(
            f"/case/{file_number}/chat/{conv_id}?turn={turn_id}",
            status_code=303)

    @app.get("/api/chat/{conv_id}/turn/{turn_id}")
    def chat_turn_status(conv_id: str, turn_id: str):
        t = _CHAT_MANAGER.get_turn(turn_id)
        if t is None:
            return JSONResponse({"error": "not found"}, status_code=404)
        return JSONResponse({"status": t["status"], "log": t["log"],
                             "done": t["done"], "error": t["error"]})


def _register_job_routes(app, manager, templates):
    import os

    def _render(name, request, **ctx):
        return templates.TemplateResponse(request, name, ctx)

    @app.get("/job/{job_id}", response_class=HTMLResponse)
    def job_page(request: Request, job_id: str):
        job = manager.store.get(job_id)
        if job is None:
            return HTMLResponse("Job not found", status_code=404)
        task = T.TASKS.get(job.task_id)
        return _render("job.html", request, job=job, task=task)

    @app.get("/api/job/{job_id}")
    def job_state(job_id: str):
        job = manager.store.get(job_id)
        if job is None:
            return JSONResponse({"error": "not found"}, status_code=404)
        return JSONResponse({
            "state": job.state,
            "progress": job.progress,
            "log": job.log[-50:],
            "has_output": bool(job.output_path),
            "error": job.error,
        })

    @app.post("/job/{job_id}/cancel")
    def cancel_job(job_id: str):
        manager.cancel(job_id)
        return RedirectResponse(f"/job/{job_id}", status_code=303)

    @app.get("/job/{job_id}/output")
    def job_output(job_id: str):
        job = manager.store.get(job_id)
        if job is None or not job.output_path \
                or not os.path.isfile(job.output_path):
            return HTMLResponse("Output not found", status_code=404)
        return FileResponse(
            job.output_path,
            media_type=_DOCX_MEDIA_TYPE,
            filename=os.path.basename(job.output_path),
        )


def _register_depo_prep_routes(app, manager, templates):
    @app.post("/case/{file_number}/task/depo_prep/submit")
    async def depo_prep_submit(request: Request, file_number: str):
        case = cases.get_case(file_number)
        if case is None:
            return HTMLResponse("Case not found", status_code=404)
        form = await request.form()
        rel_files = [f for f in form.getlist("files") if f]
        if not rel_files:
            return HTMLResponse("No source files.", status_code=400)
        try:
            abs_files = [str(cases.safe_resolve(case["case_path"], rf))
                         for rf in rel_files]
        except ValueError:
            return HTMLResponse("Invalid path", status_code=400)
        cfg = {
            "deponent_name": (form.get("deponent_name") or "").strip(),
            "deponent_role": (form.get("deponent_role") or "").strip(),
            "deponent_sources": abs_files,
            "context_sources": [],
            "style": form.get("style") or "discovery",
            "free_text_notes": (form.get("free_text_notes") or "").strip(),
            "per_topic_flags": {
                "strategic_note": form.get("flag_strategic") == "on",
                "source_facts": form.get("flag_source_facts") == "on",
                "impeachment_hook": form.get("flag_impeachment") == "on",
                "objection_alts": form.get("flag_objection") == "on",
            },
            "case_root": case["case_path"],
        }
        cfg_path = T.write_depo_prep_config(cfg)
        job = new_job("depo_prep", case["case_path"], file_number, [cfg_path])
        manager.submit(job)
        return RedirectResponse(f"/job/{job.id}", status_code=303)


def _register_awaiting_routes(app, manager, templates):
    def _render(name, request, **ctx):
        return templates.TemplateResponse(request, name, ctx)

    def _get_awaiting(job_id):
        job = manager.store.get(job_id)
        if job is None or job.state != J.AWAITING_INPUT \
                or not job.session_path:
            return None
        return job

    @app.get("/job/{job_id}/awaiting", response_class=HTMLResponse)
    def awaiting_form(request: Request, job_id: str):
        job = _get_awaiting(job_id)
        if job is None:
            return HTMLResponse("Not awaiting input", status_code=404)
        kind = T.TASKS[job.task_id].awaiting_kind
        if kind == "deposition":
            session = T.read_session_json(job.session_path)
            return _render("awaiting_deposition.html", request, job=job,
                           session=session, audiences=T.DEPO_AUDIENCES,
                           tones=T.DEPO_TONES)
        if kind == "med_chron":
            session = T.read_session_json(job.session_path)
            return _render("awaiting_med_chron.html", request, job=job,
                           session=session)
        if kind == "depo_prep":
            topics = T.read_depo_prep_topics(job.session_path)
            return _render("awaiting_depo_prep.html", request, job=job,
                           topics=topics)
        return HTMLResponse("Unknown task kind", status_code=400)

    @app.post("/job/{job_id}/resume")
    async def resume_job(request: Request, job_id: str):
        job = _get_awaiting(job_id)
        if job is None:
            return HTMLResponse("Not awaiting input", status_code=404)
        form = await request.form()
        kind = T.TASKS[job.task_id].awaiting_kind

        if kind == "deposition":
            cfg = {
                "selected_topics": [t for t in form.getlist("topic")
                                    if t.strip()],
                "added_topics": [
                    ln.strip()
                    for ln in (form.get("added_topics") or "").splitlines()
                    if ln.strip()],
                "bullets_per_topic": int(form.get("bullets") or 5),
                "deponent_label": (form.get("deponent_label") or "").strip()
                                  or "Deponent",
                "custom_rules": (form.get("custom_rules") or "").strip(),
                "cross_check_enabled": form.get("cross_check") == "on",
                "context_doc_paths": [],
                "audience": form.get("audience") or "neutral",
                "audience_custom": (
                    (form.get("audience_custom") or "").strip()
                    if form.get("audience") == "custom" else ""),
                "tone": form.get("tone") or "recitation",
                "tone_custom": (
                    (form.get("tone_custom") or "").strip()
                    if form.get("tone") == "custom" else ""),
            }
            if not cfg["selected_topics"] and not cfg["added_topics"]:
                return HTMLResponse("Select or add at least one topic.",
                                    status_code=400)
            T.apply_deposition_user_config(job.session_path, cfg)

        elif kind == "med_chron":
            selected = [a for a in form.getlist("analysis") if a]
            custom = []
            for i in (1, 2, 3):
                lbl = (form.get(f"custom_label_{i}") or "").strip()
                instr = (form.get(f"custom_instruction_{i}") or "").strip()
                if lbl and instr:
                    custom.append({"label": lbl, "instruction": instr,
                                   "context_files": []})
            if not selected and not custom:
                return HTMLResponse(
                    "Select or add at least one analysis.", status_code=400)
            T.apply_med_chron_user_config(job.session_path, {
                "selected_catalog_ids": selected,
                "custom_analyses": custom,
            })

        elif kind == "depo_prep":
            existing = T.read_depo_prep_topics(job.session_path)
            topics = []
            for i, t in enumerate(existing):
                topics.append({
                    "id": t.get("id") or f"t{i + 1:02d}",
                    "title": (form.get(f"title_{i}")
                              or t.get("title", "")).strip(),
                    "strategic_note": (form.get(f"note_{i}") or "").strip(),
                    "relevant_digest_refs": t.get("relevant_digest_refs", []),
                    "default_checked": form.get(f"keep_{i}") == "on",
                    "lawyer_added": bool(t.get("lawyer_added", False)),
                })
            for i in (1, 2, 3):
                title = (form.get(f"new_title_{i}") or "").strip()
                if title:
                    topics.append({
                        "id": f"t{len(topics) + 1:02d}", "title": title,
                        "strategic_note": (form.get(f"new_note_{i}")
                                           or "").strip(),
                        "relevant_digest_refs": [],
                        "default_checked": True, "lawyer_added": True,
                    })
            if not any(t["default_checked"] for t in topics):
                return HTMLResponse("Keep or add at least one topic.",
                                    status_code=400)
            T.write_depo_prep_topics(job.session_path, topics)

        else:
            return HTMLResponse("Unknown task kind", status_code=400)

        manager.resume(job_id)
        return RedirectResponse(f"/job/{job_id}", status_code=303)


# ---- entry point ----

def _tailscale_exe() -> str:
    """Return a runnable tailscale path: bare command if on PATH, else the
    standard Windows install location (the MSI does not add it to PATH)."""
    found = shutil.which("tailscale")
    if found:
        return found
    program_files = os.environ.get("ProgramFiles", r"C:\Program Files")
    fallback = os.path.join(program_files, "Tailscale", "tailscale.exe")
    return fallback if os.path.isfile(fallback) else "tailscale"


def detect_tailscale_ip() -> str | None:
    try:
        out = subprocess.run([_tailscale_exe(), "ip", "-4"],
                             capture_output=True, text=True, timeout=10)
        if out.returncode == 0:
            lines = [ln.strip() for ln in out.stdout.splitlines() if ln.strip()]
            if lines:
                return lines[0]
    except (OSError, subprocess.TimeoutExpired):
        pass
    return None


def main() -> None:
    parser = argparse.ArgumentParser(description="iCharlotte web companion")
    parser.add_argument("--lan", action="store_true",
                        help="Bind 0.0.0.0 instead of the Tailscale IP")
    parser.add_argument("--port", type=int, default=8765)
    args = parser.parse_args()

    if args.lan:
        host = "0.0.0.0"
    else:
        host = detect_tailscale_ip()
        if not host:
            print("ERROR: Could not detect a Tailscale IPv4 address. "
                  "Is Tailscale installed and running? "
                  "Use --lan to bind the local network instead.")
            sys.exit(1)

    app = create_app(JobManager(JobStore(DEFAULT_JOBS_PATH)))
    import uvicorn
    print(f"iCharlotte web companion: http://{host}:{args.port}")
    uvicorn.run(app, host=host, port=args.port)


if __name__ == "__main__":
    main()
