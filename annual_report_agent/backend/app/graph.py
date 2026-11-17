"""LangGraph workflow that replaces the manual process.

Manual process                               ->  Graph node
---------------------------------------------------------------------------------
download one annual report (one year)        ->  report subgraph: load_pdf
look at scanned pages / pictures             ->  index_scanned_pages (vision captions)
tell ChatGPT which company/year it is        ->  profile_report
paste the prompt, get the year's data        ->  extract_sheets   (per sheet, per page batch)
check the answer against the PDF             ->  verify_sheets    (second LLM pass, same model)
tidy numbers / ids / source pages            ->  finalize
repeat for every year                        ->  main graph fans out with Send (parallel workers)
"now join all the years"                     ->  merge_years
copy into Excel                              ->  build_excel
"""
from __future__ import annotations

import concurrent.futures
import json
import operator
import traceback
from pathlib import Path
from typing import Annotated, Any, Callable, TypedDict

from langchain_core.runnables import RunnableConfig
from langgraph.graph import END, START, StateGraph
from langgraph.types import Send

from . import config as settings
from . import excel, extraction, llm, prompts
from .pdf_reader import read_pdf, scale_snippets, year_from_filename
from .schema import Sheet, parse_spec


# ---------------------------------------------------------------------------
# Progress reporting (wired to the job store by the API)
# ---------------------------------------------------------------------------
def _progress(config: RunnableConfig) -> Callable[..., None]:
    return (config.get("configurable") or {}).get("progress") or (lambda *a, **k: None)


# ===========================================================================
# Report subgraph: ONE PDF = ONE financial year
# ===========================================================================
class ReportState(TypedDict, total=False):
    job: dict
    sheets: list[Sheet]
    filename: str
    path: str
    pdf_bytes: bytes
    pages: list
    profile: dict
    extracted: dict      # sheet -> {"rows": [...], "batches": [...]}
    final: dict          # sheet -> [final rows]
    log: list
    usage: dict
    error: str


def load_pdf(state: ReportState, config: RunnableConfig) -> dict:
    _progress(config)(state["filename"], stage="reading PDF")
    data = Path(state["path"]).read_bytes()
    pages = read_pdf(data)
    if not any(p.text.strip() or p.scanned for p in pages):
        raise ValueError("PDF has no readable text or images")
    return {"pdf_bytes": data, "pages": pages, "log": [], "usage": {}}


def index_scanned_pages(state: ReportState, config: RunnableConfig) -> dict:
    scanned = sum(p.scanned for p in state["pages"])
    if scanned:
        _progress(config)(state["filename"], stage=f"reading {scanned} scanned page(s)")
        extraction.caption_scanned_pages(state["pages"], state["pdf_bytes"], state["usage"], state["log"])
    return {"pages": state["pages"], "log": state["log"], "usage": state["usage"]}


def profile_report(state: ReportState, config: RunnableConfig) -> dict:
    _progress(config)(state["filename"], stage="detecting company and year")
    pages = state["pages"]
    front = "\n\n".join(f"=== PAGE {p.number} ===\n{p.text[:2500]}" for p in pages[:3])
    prompt = prompts.profile_prompt(state["filename"], front[:9000], scale_snippets(pages)[:5000])
    images = []
    if settings.USE_VISION and not pages[0].text.strip():
        from .pdf_reader import render_png
        images = [render_png(state["pdf_bytes"], 1, 100)]
    try:
        raw = llm.call_json(prompt, images=images, usage=state["usage"], max_output=600)
    except llm.LLMError as exc:
        raw = {}
        state["log"].append({"sheet": "-", "check": "PROFILE FAILED", "detail": str(exc)})

    year = raw.get("financial_year")
    try:
        year = int(year)
        if not 1990 <= year <= 2100:
            year = None
    except (TypeError, ValueError):
        year = None
    file_year = year_from_filename(state["filename"])
    if year is None:
        year = file_year
    elif file_year and file_year != year:
        state["log"].append({"sheet": "-", "check": "YEAR CHECK",
                             "detail": f"report says FY{year}, filename says {file_year}; using FY{year}"})
    profile = {
        "filename": state["filename"],
        "company_name": str(raw.get("company_name") or "N/A").strip(),
        "financial_year": year,
        "fiscal_year_end": raw.get("fiscal_year_end", "N/A"),
        "currency": raw.get("currency", "N/A"),
        "amounts_scale": raw.get("amounts_scale", "UNKNOWN"),
        "report_title": str(raw.get("report_title") or Path(state["filename"]).stem).strip()[:120],
        "page_count": len(pages),
    }
    _progress(config)(state["filename"], year=year, company=profile["company_name"])
    return {"profile": profile, "log": state["log"], "usage": state["usage"]}


def _parallel(items: list, fn: Callable, workers: int) -> list:
    with concurrent.futures.ThreadPoolExecutor(max_workers=max(1, workers)) as pool:
        return list(pool.map(fn, items))


def extract_sheets(state: ReportState, config: RunnableConfig) -> dict:
    progress = _progress(config)
    sheets = [s for s in state["sheets"] if not s.empty and s.model_columns]
    total = len(sheets)
    done = {"n": 0}
    year = state["profile"].get("financial_year") or "unknown"
    company = state["job"].get("company_name") or state["profile"]["company_name"]
    progress(state["filename"], stage="extracting", sheets_total=total, sheets_done=0)

    def run(sheet: Sheet) -> tuple[str, dict]:
        try:
            nums, how = extraction.locate_pages(sheet, state["pages"], year, state["usage"])
            if how != "llm":
                state["log"].append({"sheet": sheet.name, "check": "PAGE LOCATOR", "detail": how})
            rows, batches = (extraction.extract_sheet(sheet, nums, state["pages"], state["pdf_bytes"],
                                                      state["profile"], company, state["usage"],
                                                      state["log"]) if nums else ([], []))
            result = {"rows": rows, "batches": batches, "pages": nums}
        except Exception as exc:  # noqa: BLE001 - one bad sheet must not stop the report
            state["log"].append({"sheet": sheet.name, "check": "SHEET ERROR",
                                 "detail": f"{type(exc).__name__}: {exc}"})
            result = {"rows": [], "batches": [], "pages": []}
        done["n"] += 1
        progress(state["filename"], sheets_done=done["n"])
        return sheet.name, result

    extracted = dict(_parallel(sheets, run, settings.SHEET_WORKERS))
    return {"extracted": extracted, "log": state["log"], "usage": state["usage"]}


def verify_sheets(state: ReportState, config: RunnableConfig) -> dict:
    progress = _progress(config)
    year = state["profile"].get("financial_year") or "unknown"
    work = [(s, state["extracted"][s.name]) for s in state["sheets"]
            if s.name in state["extracted"] and state["extracted"][s.name]["rows"]]
    progress(state["filename"], stage="verifying", sheets_total=len(work), sheets_done=0)
    done = {"n": 0}

    def run(item):
        sheet, ext = item
        for batch in ext["batches"]:
            rows = [r for r in ext["rows"] if r["_batch"] == batch.pages]
            for i in range(0, len(rows), 12):   # keep each check small enough for the context
                extraction.verify_batch(sheet, batch, rows[i:i + 12], year, state["usage"], state["log"])
        done["n"] += 1
        progress(state["filename"], sheets_done=done["n"])

    _parallel(work, run, settings.SHEET_WORKERS)
    return {"extracted": state["extracted"], "log": state["log"], "usage": state["usage"]}


def finalize(state: ReportState, config: RunnableConfig) -> dict:
    _progress(config)(state["filename"], stage="finalizing")
    final = {}
    for sheet in state["sheets"]:
        ext = state["extracted"].get(sheet.name)
        final[sheet.name] = (extraction.finalize_rows(sheet, ext["rows"], state["profile"], state["job"],
                                                      state["pages"], state["log"]) if ext else [])
    return {"final": final, "log": state["log"]}


def build_report_graph():
    g = StateGraph(ReportState)
    g.add_node("load_pdf", load_pdf)
    g.add_node("index_scanned_pages", index_scanned_pages)
    g.add_node("profile_report", profile_report)
    g.add_node("extract_sheets", extract_sheets)
    g.add_node("verify_sheets", verify_sheets)
    g.add_node("finalize", finalize)
    g.add_edge(START, "load_pdf")
    g.add_edge("load_pdf", "index_scanned_pages")
    g.add_edge("index_scanned_pages", "profile_report")
    g.add_edge("profile_report", "extract_sheets")
    g.add_edge("extract_sheets", "verify_sheets")
    g.add_edge("verify_sheets", "finalize")
    g.add_edge("finalize", END)
    return g.compile()


REPORT_GRAPH = build_report_graph()


# ===========================================================================
# Main graph: all PDFs of one company -> one workbook
# ===========================================================================
class JobState(TypedDict, total=False):
    job: dict                 # company_name, company_id, rating_cycle_year, source_code, output_dir
    spec_text: str
    sheets: list[Sheet]
    reports: list[dict]       # [{"filename", "path"}]
    year_results: Annotated[list, operator.add]
    merged: dict
    warnings: list
    output_path: str


def prepare(state: JobState, config: RunnableConfig) -> dict:
    sheets = parse_spec(state["spec_text"])
    _progress(config)(None, stage="started", message=f"{len(state['reports'])} report(s), {len(sheets)} sheet(s)")
    return {"sheets": sheets}


def fan_out(state: JobState) -> list[Send]:
    return [Send("process_report", {"job": state["job"], "sheets": state["sheets"],
                                    "filename": r["filename"], "path": r["path"]})
            for r in state["reports"]]


def process_report(payload: dict, config: RunnableConfig) -> dict:
    progress = _progress(config)
    try:
        out = REPORT_GRAPH.invoke(payload, {"recursion_limit": 25,
                                            "configurable": {"progress": progress}})
        result = {"filename": payload["filename"], "status": "done", "profile": out["profile"],
                  "rows": out["final"], "log": out["log"], "usage": out["usage"]}
        progress(payload["filename"], stage="done")
    except Exception as exc:  # noqa: BLE001 - report failures are recorded, not fatal
        result = {"filename": payload["filename"], "status": "failed",
                  "error": f"{type(exc).__name__}: {exc}", "trace": traceback.format_exc()[-2000:],
                  "profile": {"filename": payload["filename"]}, "rows": {}, "log": [], "usage": {}}
        progress(payload["filename"], stage="failed", error=result["error"])
    return {"year_results": [result]}


def merge_years(state: JobState, config: RunnableConfig) -> dict:
    _progress(config)(None, stage="merging years")
    ok = [r for r in state["year_results"] if r["status"] == "done"]
    ok.sort(key=lambda r: (r["profile"].get("financial_year") or 0, r["filename"]))
    warnings = []
    for r in state["year_results"]:
        if r["status"] != "done":
            warnings.append({"check": "REPORT FAILED", "detail": f"{r['filename']}: {r.get('error')}"})
    years = [r["profile"].get("financial_year") for r in ok]
    for y in sorted({y for y in years if y}):
        files = [r["filename"] for r in ok if r["profile"].get("financial_year") == y]
        if len(files) > 1:
            warnings.append({"check": "DUPLICATE YEAR", "detail": f"FY{y}: {', '.join(files)}"})
    known = sorted({y for y in years if y})
    if known:
        missing = [y for y in range(known[0], known[-1] + 1) if y not in known]
        if missing:
            warnings.append({"check": "MISSING YEARS", "detail": ", ".join(map(str, missing))})
    for r in ok:
        if not r["profile"].get("financial_year"):
            warnings.append({"check": "YEAR NOT DETECTED", "detail": r["filename"]})
    names = {extraction.norm_company(r["profile"].get("company_name", "")) for r in ok} - {""}
    if len(names) > 1:
        warnings.append({"check": "COMPANY MISMATCH",
                         "detail": "; ".join(f"{r['filename']}: {r['profile'].get('company_name')}" for r in ok)})

    merged: dict[str, list] = {s.name: [] for s in state["sheets"]}
    seen: dict[str, set] = {s.name: set() for s in state["sheets"]}
    for r in ok:
        for sheet_name, rows in r["rows"].items():
            for row in rows:
                key = json.dumps(row["values"], sort_keys=True, default=str)
                if key not in seen[sheet_name]:
                    seen[sheet_name].add(key)
                    merged[sheet_name].append({**row, "filename": r["filename"]})
    return {"merged": merged, "warnings": warnings}


def build_excel(state: JobState, config: RunnableConfig) -> dict:
    _progress(config)(None, stage="building Excel")
    out_dir = Path(state["job"]["output_dir"])
    out_dir.mkdir(parents=True, exist_ok=True)
    path = out_dir / excel.output_filename(state["job"].get("company_name") or "company",
                                           [r["profile"].get("financial_year") for r in state["year_results"]])
    excel.write_workbook(path, state["sheets"], state["merged"], state["year_results"], state["warnings"])
    (out_dir / "result.json").write_text(json.dumps({
        "sheets": {k: [r["values"] for r in v] for k, v in state["merged"].items()},
        "reports": [{k: r.get(k) for k in ("filename", "status", "profile", "usage", "error")}
                    for r in state["year_results"]],
        "warnings": state["warnings"],
    }, indent=2, default=str), encoding="utf-8")
    return {"output_path": str(path)}


def build_job_graph():
    g = StateGraph(JobState)
    g.add_node("prepare", prepare)
    g.add_node("process_report", process_report)
    g.add_node("merge_years", merge_years)
    g.add_node("build_excel", build_excel)
    g.add_edge(START, "prepare")
    g.add_conditional_edges("prepare", fan_out, ["process_report"])
    g.add_edge("process_report", "merge_years")
    g.add_edge("merge_years", "build_excel")
    g.add_edge("build_excel", END)
    return g.compile()


JOB_GRAPH = build_job_graph()


def run_job(job: dict, spec_text: str, reports: list[dict],
            progress: Callable[..., None] | None = None) -> dict[str, Any]:
    """Run the whole pipeline. REPORT_WORKERS PDFs are processed at the same time."""
    return JOB_GRAPH.invoke(
        {"job": job, "spec_text": spec_text, "reports": reports, "year_results": []},
        {"max_concurrency": settings.REPORT_WORKERS,
         "recursion_limit": 50,
         "configurable": {"progress": progress or (lambda *a, **k: None)}},
    )
