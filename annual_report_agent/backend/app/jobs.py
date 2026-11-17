"""In-process job queue: each upload becomes a job that runs the LangGraph pipeline in a worker."""
from __future__ import annotations

import concurrent.futures
import datetime as dt
import json
import logging
import threading
import time
import traceback
import uuid
from pathlib import Path
from typing import Any

from . import config
from .graph import run_job

_EXECUTOR = concurrent.futures.ThreadPoolExecutor(max_workers=max(1, config.JOB_WORKERS))
_LOCK = threading.RLock()
log = logging.getLogger("jobs")
_JOBS: dict[str, dict[str, Any]] = {}


def _now() -> str:
    return dt.datetime.now(dt.timezone.utc).isoformat(timespec="seconds")


def job_dir(job_id: str) -> Path:
    return config.DATA_DIR / "jobs" / job_id


def _save(job: dict) -> None:
    """Persist job state. Best effort: a failed save must never fail the extraction.

    On Windows, antivirus / search indexer / OneDrive can briefly lock job.json, which makes
    the atomic replace raise PermissionError, so retry a few times and then give up quietly
    (the in-memory state is what the API serves while the server is running).
    """
    path = job_dir(job["id"]) / "job.json"
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        data = json.dumps(job, indent=2, default=str)
        tmp = path.with_name(f"job.{uuid.uuid4().hex[:8]}.tmp")
        tmp.write_text(data, encoding="utf-8")
    except OSError as exc:
        log.warning("could not write job state for %s: %s", job.get("id"), exc)
        return
    for attempt in range(6):
        try:
            tmp.replace(path)
            return
        except PermissionError:
            time.sleep(0.05 * (attempt + 1))
        except OSError:
            break
    try:   # last resort: write in place, then drop the temp file
        path.write_text(data, encoding="utf-8")
    except OSError as exc:
        log.warning("could not save job state for %s: %s", job.get("id"), exc)
    finally:
        tmp.unlink(missing_ok=True)


def get(job_id: str) -> dict | None:
    with _LOCK:
        if job_id in _JOBS:
            return json.loads(json.dumps(_JOBS[job_id], default=str))
    path = job_dir(job_id) / "job.json"
    if path.exists():
        return json.loads(path.read_text(encoding="utf-8"))
    return None


def create(inputs: dict, spec_text: str, files: list[tuple[str, bytes]]) -> str:
    job_id = uuid.uuid4().hex[:12]
    folder = job_dir(job_id)
    (folder / "uploads").mkdir(parents=True, exist_ok=True)
    reports = []
    for i, (name, data) in enumerate(files, 1):
        safe = "".join(ch if ch.isalnum() or ch in "._- " else "_" for ch in name)[:120] or f"report_{i}.pdf"
        path = folder / "uploads" / f"{i:02d}_{safe}"
        path.write_bytes(data)
        reports.append({"filename": name, "path": str(path)})
    job = {
        "id": job_id, "status": "queued", "created_at": _now(), "finished_at": None,
        "inputs": inputs, "stage": "queued", "error": None, "output_file": None,
        "reports": {r["filename"]: {"stage": "waiting", "year": None, "company": None,
                                    "sheets_done": 0, "sheets_total": 0, "error": None}
                    for r in reports},
        "warnings": [], "summary": [],
    }
    with _LOCK:
        _JOBS[job_id] = job
        _save(job)
    (folder / "spec.txt").write_text(spec_text, encoding="utf-8")
    _EXECUTOR.submit(_run, job_id, inputs, spec_text, reports)
    return job_id


def _progress_for(job_id: str):
    def update(filename: str | None, **fields: Any) -> None:
        with _LOCK:
            job = _JOBS[job_id]
            if filename is None:
                if "stage" in fields:
                    job["stage"] = fields["stage"]
            elif filename in job["reports"]:
                job["reports"][filename].update({k: v for k, v in fields.items() if v is not None})
            _save(job)
    return update


def _run(job_id: str, inputs: dict, spec_text: str, reports: list[dict]) -> None:
    with _LOCK:
        _JOBS[job_id]["status"] = "running"
        _save(_JOBS[job_id])
    try:
        out = run_job({**inputs, "output_dir": str(job_dir(job_id) / "output")}, spec_text, reports,
                      _progress_for(job_id))
        with _LOCK:
            job = _JOBS[job_id]
            job["output_file"] = out["output_path"]
            job["warnings"] = out.get("warnings", [])
            job["summary"] = [{
                "filename": r["filename"], "status": r["status"],
                "year": r["profile"].get("financial_year"),
                "rows": sum(len(v) for v in r.get("rows", {}).values()),
                "error": r.get("error"),
            } for r in out["year_results"]]
            failed = all(r["status"] == "failed" for r in out["year_results"])
            job["status"] = "failed" if failed else "done"
            job["stage"] = "failed" if failed else "done"
            if failed:
                job["error"] = "All reports failed: " + "; ".join(
                    str(r.get("error")) for r in out["year_results"])
    except Exception as exc:  # noqa: BLE001
        with _LOCK:
            job = _JOBS[job_id]
            job["status"], job["stage"] = "failed", "failed"
            job["error"] = f"{type(exc).__name__}: {exc}"
            (job_dir(job_id) / "error.txt").write_text(traceback.format_exc(), encoding="utf-8")
    with _LOCK:
        _JOBS[job_id]["finished_at"] = _now()
        _save(_JOBS[job_id])
