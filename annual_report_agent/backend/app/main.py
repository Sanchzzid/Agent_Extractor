"""HTTP API used by the web interface.

POST /api/jobs                upload PDFs (+ company name, optional sheet spec) -> {job_id}
GET  /api/jobs/{id}           progress per PDF / year
GET  /api/jobs/{id}/download  the consolidated Excel workbook
GET  /api/jobs/{id}/json      the same data as JSON
GET  /api/default-spec        the default sheet spec
GET  /api/health
"""
from __future__ import annotations

from pathlib import Path

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, PlainTextResponse

from . import config, jobs
from .pdf_reader import is_pdf
from .schema import parse_spec

app = FastAPI(title="Annual Report Extraction Agent", version="1.0.0")
app.add_middleware(CORSMiddleware, allow_origins=config.ALLOWED_ORIGINS,
                   allow_methods=["GET", "POST"], allow_headers=["*"])


def _default_spec() -> str:
    return config.DEFAULT_SPEC_PATH.read_text(encoding="utf-8-sig") if config.DEFAULT_SPEC_PATH.exists() else ""


@app.get("/api/health")
def health() -> dict:
    return {"ok": True, "model": config.LLM_MODEL, "llm_configured": bool(config.LLM_MODEL and config.LLM_API_KEY),
            "report_workers": config.REPORT_WORKERS, "sheet_workers": config.SHEET_WORKERS}


@app.get("/api/default-spec", response_class=PlainTextResponse)
def default_spec() -> str:
    return _default_spec()


@app.post("/api/jobs")
async def create_job(
    company_name: str = Form(...),
    files: list[UploadFile] = File(...),
    company_id: str = Form(""),
    rating_cycle_year: str = Form(""),
    source_code: str = Form(""),
    spec_text: str = Form(""),
) -> dict:
    company_name = company_name.strip()
    if not company_name:
        raise HTTPException(400, "Company name is required.")
    if not files:
        raise HTTPException(400, "Upload at least one PDF.")
    if len(files) > config.MAX_FILES_PER_JOB:
        raise HTTPException(400, f"Upload at most {config.MAX_FILES_PER_JOB} PDFs per run.")
    spec = spec_text.strip() or _default_spec()
    try:
        sheets = parse_spec(spec)
    except ValueError as exc:
        raise HTTPException(400, f"Sheet spec error: {exc}") from exc

    loaded: list[tuple[str, bytes]] = []
    limit = config.MAX_UPLOAD_MB * 1024 * 1024
    for f in files:
        data = await f.read(limit + 1)
        name = Path(f.filename or "report.pdf").name
        if len(data) > limit:
            raise HTTPException(400, f"{name} is larger than {config.MAX_UPLOAD_MB} MB.")
        if not is_pdf(data):
            raise HTTPException(400, f"{name} is not a PDF file.")
        loaded.append((name, data))
    if len({n for n, _ in loaded}) != len(loaded):
        raise HTTPException(400, "Two uploaded files have the same name.")

    inputs = {"company_name": company_name, "company_id": company_id.strip(),
              "rating_cycle_year": rating_cycle_year.strip(), "source_code": source_code.strip()}
    job_id = jobs.create(inputs, spec, loaded)
    return {"job_id": job_id, "sheets": len(sheets), "files": [n for n, _ in loaded]}


@app.get("/api/jobs/{job_id}")
def job_status(job_id: str) -> dict:
    job = jobs.get(job_id) if job_id.isalnum() else None
    if not job:
        raise HTTPException(404, "Job not found.")
    job["download_ready"] = bool(job.get("output_file"))
    job.pop("output_file", None)
    return job


def _output(job_id: str) -> Path:
    job = jobs.get(job_id) if job_id.isalnum() else None
    if not job or not job.get("output_file"):
        raise HTTPException(404, "Result not ready.")
    return Path(job["output_file"])


@app.get("/api/jobs/{job_id}/download")
def download(job_id: str) -> FileResponse:
    path = _output(job_id)
    return FileResponse(path, filename=path.name,
                        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.get("/api/jobs/{job_id}/json")
def download_json(job_id: str) -> FileResponse:
    path = _output(job_id).parent / "result.json"
    return FileResponse(path, filename="result.json", media_type="application/json")


# Serve the built web interface (frontend/dist) from the same server when it exists,
# so `python run.py` gives the full app at http://localhost:<port>. Must stay after the /api routes.
_UI = Path(__file__).resolve().parents[2] / "frontend" / "dist"
if _UI.is_dir():
    from fastapi.staticfiles import StaticFiles
    app.mount("/", StaticFiles(directory=_UI, html=True), name="ui")
