"""
Therefore Documentation Generator — FastAPI web application.

Endpoints:
  GET  /                       — serve the frontend
  POST /generate               — start a generation job, returns {job_id}
  GET  /progress/{job_id}      — SSE stream of log lines + completion/error event
  GET  /download/{job_id}      — download the generated .docx
  POST /validate-wrapper       — check a .docx contains the placeholder
"""

import asyncio
import json
import os
import queue
import shutil
import sys
import tempfile
import threading
import time
import uuid
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles

ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(ROOT))

from themes import list_themes, load_theme

app = FastAPI(title="Therefore Documentation Generator")

TEMPLATES_DIR = ROOT / "templates"


def list_templates() -> list[dict]:
    """Return available base Word templates from the templates/ directory."""
    if not TEMPLATES_DIR.exists():
        return []
    templates = []
    for f in sorted(TEMPLATES_DIR.glob("*.docx")):
        templates.append({
            "id": f.stem,
            "name": f.stem.replace("_", " ").replace("-", " ").title(),
            "path": str(f),
        })
    return templates

STATIC_DIR = Path(__file__).parent / "static"
app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")

# ---------------------------------------------------------------------------
# Job state
# ---------------------------------------------------------------------------
@dataclass
class Job:
    id:          str
    status:      str = "running"   # running | done | error
    output_path: Optional[str] = None
    output_name: Optional[str] = None
    error:       Optional[str] = None
    warnings:    list = field(default_factory=list)
    log_queue:   queue.Queue = field(default_factory=queue.Queue)
    created_at:  float = field(default_factory=time.time)

_jobs: dict[str, Job] = {}
_jobs_lock = threading.Lock()

def _cleanup_old_jobs():
    """Remove jobs and their temp files older than 30 minutes. Caller must NOT hold _jobs_lock."""
    cutoff = time.time() - 1800
    with _jobs_lock:
        stale = {jid: j for jid, j in _jobs.items() if j.created_at < cutoff}
        for jid in stale:
            _jobs.pop(jid)
    for j in stale.values():
        if j.output_path:
            shutil.rmtree(os.path.dirname(j.output_path), ignore_errors=True)


# ---------------------------------------------------------------------------
# Routes
# ---------------------------------------------------------------------------
@app.get("/", response_class=HTMLResponse)
async def index():
    return (STATIC_DIR / "index.html").read_text()


@app.get("/themes")
async def get_themes():
    return {"themes": list_themes()}


@app.get("/templates")
async def get_templates():
    return {"templates": list_templates()}


@app.post("/generate")
async def start_generate(
    xml_file:      UploadFile = File(...),
    wrapper_file:  UploadFile = File(None),
    template_id:   str        = Form(""),
    theme_file:    UploadFile = File(None),
    theme_id:      str        = Form(""),
    sections:      str        = Form(""),
    start_section: int        = Form(0),
    img_format:    str        = Form("png"),
    api_url:       str        = Form(""),
    api_tenant:    str        = Form(""),
    api_username:  str        = Form(""),
    api_password:  str        = Form(""),
):
    xml_bytes     = await xml_file.read()
    wrapper_bytes = await wrapper_file.read() if wrapper_file and wrapper_file.filename else None
    theme_bytes   = await theme_file.read() if theme_file and theme_file.filename else None

    xml_name     = xml_file.filename or "config.xml"
    wrapper_name = (wrapper_file.filename or "wrapper.docx") if wrapper_bytes else None
    theme_name   = theme_file.filename or "theme.yaml" if theme_bytes else None

    sections_list = [s.strip() for s in sections.split(",") if s.strip()] or None

    job = Job(id=str(uuid.uuid4()))
    job.start_section = start_section
    with _jobs_lock:
        _jobs[job.id] = job
    _cleanup_old_jobs()

    thread = threading.Thread(
        target=_run_job,
        args=(job, xml_bytes, xml_name, wrapper_bytes, wrapper_name,
              sections_list, start_section, theme_bytes, theme_name, img_format, theme_id, template_id,
              api_url.strip(), api_tenant.strip(), api_username.strip(), api_password),
        daemon=True,
    )
    thread.start()

    return {"job_id": job.id}


@app.get("/progress/{job_id}")
async def progress(job_id: str):
    job = _jobs.get(job_id)
    if not job:
        raise HTTPException(status_code=404, detail="Job not found")

    async def stream():
        while True:
            # Drain all queued log lines
            try:
                while True:
                    line = job.log_queue.get_nowait()
                    yield f"data: {json.dumps({'type': 'log', 'text': line})}\n\n"
            except queue.Empty:
                pass

            if job.status == "done":
                payload = {
                    "type":     "done",
                    "filename": job.output_name,
                    "warnings": job.warnings,
                }
                yield f"data: {json.dumps(payload)}\n\n"
                break
            elif job.status == "error":
                yield f"data: {json.dumps({'type': 'error', 'message': job.error})}\n\n"
                break

            await asyncio.sleep(0.25)

    return StreamingResponse(
        stream(),
        media_type="text/event-stream",
        headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"},
    )


@app.get("/download/{job_id}")
async def download(job_id: str):
    job = _jobs.get(job_id)
    if not job:
        raise HTTPException(status_code=404, detail="Job not found")
    if job.status != "done":
        raise HTTPException(status_code=400, detail="Job not complete")
    if not job.output_path or not os.path.exists(job.output_path):
        raise HTTPException(status_code=410, detail="File no longer available")

    return FileResponse(
        path=job.output_path,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename=job.output_name,
    )


@app.post("/test-connection")
async def test_connection(
    api_url:      str = Form(""),
    api_tenant:   str = Form(""),
    api_username: str = Form(""),
    api_password: str = Form(""),
):
    if not api_url.strip() or not api_username.strip():
        return JSONResponse({"ok": False, "error": "API URL and username are required."})
    try:
        from fetch_server_info import fetch_server_info, derive_tenant
        info = fetch_server_info(api_url.strip(), api_tenant.strip(), api_username.strip(), api_password)
        tenant = derive_tenant(api_url.strip(), api_tenant.strip())
        return {
            "ok":              True,
            "service_version": info.get("service_version", ""),
            "server_name":     info.get("server_name", ""),
            "customer_id":     info.get("customer_id", ""),
            "tenant":          info.get("tenant_name", "") or tenant,
            "region":          info.get("region", ""),
        }
    except Exception as exc:
        return JSONResponse({"ok": False, "error": str(exc)})


@app.post("/validate-wrapper")
async def validate_wrapper(wrapper_file: UploadFile = File(...)):
    data = await wrapper_file.read()
    placeholder = "{{THEREFORE_CONTENT}}"
    try:
        from docx import Document
        import io
        doc = Document(io.BytesIO(data))
        found = any(p.text.strip() == placeholder for p in doc.paragraphs)
        return {"found": found, "placeholder": placeholder}
    except Exception as e:
        return {"found": False, "placeholder": placeholder, "error": str(e)}


# ---------------------------------------------------------------------------
# Job worker
# ---------------------------------------------------------------------------
def _run_job(job: Job, xml_bytes, xml_name, wrapper_bytes, wrapper_name,
             sections_list, start_section=1, theme_bytes=None, theme_name=None,
             img_format="png", theme_id="", template_id="",
             api_url="", api_tenant="", api_username="", api_password=""):
    tmpdir = tempfile.mkdtemp(prefix="therefore_web_")
    try:
        from build_doc import generate
        from merge_docs import merge as merge_docs
        from themes import load_theme, list_themes

        xml_path = os.path.join(tmpdir, xml_name)
        with open(xml_path, "wb") as f:
            f.write(xml_bytes)

        theme = None
        if theme_bytes:
            theme_path = os.path.join(tmpdir, theme_name or "theme.yaml")
            with open(theme_path, "wb") as f:
                f.write(theme_bytes)
            theme = load_theme(theme_path)
        elif theme_id:
            for t in list_themes():
                if t["id"] == theme_id:
                    theme = load_theme(t["path"])
                    break

        # Resolve template selection
        actual_wrapper_bytes = wrapper_bytes
        actual_wrapper_name = wrapper_name
        if template_id and not wrapper_bytes:
            for t in list_templates():
                if t["id"] == template_id:
                    actual_wrapper_bytes = open(t["path"], "rb").read()
                    actual_wrapper_name = os.path.basename(t["path"])
                    break

        stem = os.path.splitext(xml_name)[0]
        doc_title = (stem.replace("TheConfiguration-", "")
                         .replace("TheConfiguration", "")
                         .strip(" -_") or "Therefore Documentation")
        output_name = f"{doc_title}_Documentation.docx"
        output_path = os.path.join(tmpdir, output_name)

        server_info = None
        if api_url and api_username:
            try:
                import sys as _sys
                _sys.path.insert(0, str(ROOT))
                from fetch_server_info import fetch_server_info
                job.log_queue.put(f"Connecting to Therefore API at {api_url} ...")
                server_info = fetch_server_info(api_url, api_tenant, api_username, api_password)
                job.log_queue.put("Server configuration retrieved.")
            except Exception as exc:
                job.log_queue.put(f"Warning: could not fetch server info — {exc}")
                job.warnings.append(f"Server info unavailable: {exc}")

        warnings = generate(
            xml_path, output_path,
            sections      = sections_list,
            body_only     = actual_wrapper_bytes is not None,
            start_section = start_section,
            log_fn        = job.log_queue.put,
            theme         = theme,
            img_format    = img_format,
            server_info   = server_info,
        )
        job.warnings = warnings

        if actual_wrapper_bytes:
            wrapper_path = os.path.join(tmpdir, actual_wrapper_name)
            with open(wrapper_path, "wb") as f:
                f.write(actual_wrapper_bytes)
            merged_path = os.path.join(tmpdir, f"merged_{output_name}")
            job.log_queue.put("Merging into wrapper document...")
            merge_docs(wrapper_path, output_path, merged_path)
            final_path = merged_path
        else:
            final_path = output_path

        # Move result out of tmpdir to a dedicated result dir (tmpdir will be cleaned later)
        result_dir  = tempfile.mkdtemp(prefix="therefore_result_")
        result_path = os.path.join(result_dir, output_name)
        shutil.copy2(final_path, result_path)

        job.output_path = result_path
        job.output_name = output_name
        job.status      = "done"

    except Exception as exc:
        import traceback
        traceback.print_exc()
        job.error  = str(exc)
        job.status = "error"
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)
