"""Write the merged JSON into one Excel workbook (spec sheets + audit sheets)."""
from __future__ import annotations

import re
from pathlib import Path
from typing import Any

from openpyxl import Workbook
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from .schema import Sheet

HEADER_FILL = PatternFill("solid", fgColor="1F3864")
HEADER_FONT = Font(color="FFFFFF", bold=True)
STATUS_FILL = {"CORRECTED": "FFF2CC", "NEEDS_REVIEW": "F8CBAD", "UNVERIFIED": "EDEDED"}


def output_filename(company: str, years: list[Any]) -> str:
    base = re.sub(r"[^A-Za-z0-9]+", "_", company).strip("_")[:40] or "company"
    ys = sorted({int(y) for y in years if isinstance(y, int)})
    span = f"{ys[0]}-{ys[-1]}" if len(ys) > 1 else (str(ys[0]) if ys else "unknown_year")
    return f"{base}_{span}_extraction.xlsx"


def _set(cell, value: Any) -> None:
    if isinstance(value, str):
        value = ILLEGAL_CHARACTERS_RE.sub("", value)[:32000]
        cell.value = value
        if value.startswith(("=", "+", "@")):   # never let report text become a formula
            cell.data_type = "s"
    else:
        cell.value = value


def _header(ws, headers: list[str]) -> None:
    for i, h in enumerate(headers, 1):
        c = ws.cell(row=1, column=i)
        _set(c, h)
        c.fill, c.font = HEADER_FILL, HEADER_FONT
        c.alignment = Alignment(wrap_text=True, vertical="center")
        ws.column_dimensions[get_column_letter(i)].width = max(12, min(40, len(h) + 4))
    ws.freeze_panes = "A2"


def _table(ws, headers: list[str], rows: list[list[Any]]) -> None:
    _header(ws, headers)
    for r, values in enumerate(rows, 2):
        for c, v in enumerate(values, 1):
            _set(ws.cell(row=r, column=c), v)


def write_workbook(path: Path, sheets: list[Sheet], merged: dict[str, list[dict]],
                   reports: list[dict], warnings: list[dict]) -> None:
    wb = Workbook()
    wb.remove(wb.active)

    for sheet in sheets:
        ws = wb.create_sheet(sheet.name)
        _header(ws, sheet.headers)
        for r, row in enumerate(merged.get(sheet.name, []), 2):
            fill = STATUS_FILL.get(row.get("status", ""))
            for c, name in enumerate(sheet.headers, 1):
                cell = ws.cell(row=r, column=c)
                _set(cell, row["values"].get(name, "N/A"))
                if fill:
                    cell.fill = PatternFill("solid", fgColor=fill)

    # ---- audit sheets ----
    summary = []
    for rep in sorted(reports, key=lambda x: (x.get("profile", {}).get("financial_year") or 0)):
        p, u = rep.get("profile", {}), rep.get("usage", {})
        rows = rep.get("rows", {})
        summary.append([rep["filename"], rep["status"], p.get("financial_year", "N/A"),
                        p.get("company_name", "N/A"), p.get("currency", "N/A"),
                        p.get("amounts_scale", "N/A"), p.get("page_count", "N/A"),
                        sum(len(v) for v in rows.values()),
                        sum(1 for v in rows.values() for x in v if x["status"] == "CORRECTED"),
                        sum(1 for v in rows.values() for x in v if x["status"] == "NEEDS_REVIEW"),
                        u.get("calls", 0), u.get("input_tokens", 0), u.get("output_tokens", 0),
                        rep.get("error", "")])
    _table(wb.create_sheet("_RUN_SUMMARY"),
           ["File", "Status", "Financial Year", "Company Detected", "Currency", "Scale", "Pages",
            "Rows Extracted", "Rows Corrected", "Rows Need Review", "LLM Calls", "Input Tokens",
            "Output Tokens", "Error"], summary)

    log_rows = []
    for rep in reports:
        year = rep.get("profile", {}).get("financial_year", "N/A")
        for e in rep.get("log", []):
            log_rows.append([rep["filename"], year, e.get("sheet", ""), e.get("check", ""),
                             e.get("column", ""), _txt(e.get("extracted", "")), _txt(e.get("verifier", "")),
                             e.get("action", ""), e.get("reason", "") or e.get("detail", ""),
                             ", ".join(map(str, e.get("pages", [])))])
    _table(wb.create_sheet("_VERIFICATION_LOG"),
           ["File", "Financial Year", "Sheet", "Check", "Column", "Extracted Value", "Verifier Value",
            "Action", "Reason / Detail", "PDF Pages"], log_rows)

    _table(wb.create_sheet("_WARNINGS"), ["Check", "Detail"],
           [[w["check"], w["detail"]] for w in warnings] or [["NONE", "No cross-year warnings"]])

    wb.save(path)


def _txt(v: Any) -> Any:
    return v if isinstance(v, (int, float)) else str(v)
