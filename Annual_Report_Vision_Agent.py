# Requirements: streamlit, openai, google-generativeai, anthropic, PyMuPDF,
# openpyxl, python-dotenv, Pillow, requests
# Python 3.10+
"""Annual Report Vision Agent: one company, up to six annual reports, one workbook.

The sheet specification is authoritative. Main workbook columns never come from
model-generated field names. An independent audit workbook holds provenance and QA.
"""

from __future__ import annotations

import base64
import concurrent.futures
import datetime as dt
import decimal
import difflib
import hashlib
import io
import json
import os
import queue
import re
import shutil
import tempfile
import threading
import time
from collections import Counter, defaultdict
from pathlib import Path
from typing import Any
from urllib.parse import urlparse

import fitz
import requests
from dotenv import load_dotenv
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from PIL import Image

load_dotenv()

# =========================== EDITABLE CONFIG ===========================
STATUS_DEFAULTS = {
    "data_status": "EXTRACTED",
    "verification_status": "UNVERIFIED",
    "publication_status": "DRAFT",
    "verified_at": "N/A",
}
DERIVED_COLUMNS = {
    "employee_turnover": {
        "average": "Average Headcount Used",
        "calculated": "Calculated Exit Rate (%)",
        "final": "Final Turnover / Attrition Rate (%)",
        "method": "Calculation / Selection Method",
    }
}
PAGE_NUMBERING = "printed"  # "printed" or "pdf_index"
MAX_WORKERS = 3
MAX_PAGES_PER_SHEET = 12
PAGES_PER_VISION_BATCH = 5
MAX_REPORTS = 6
MAX_PDF_DOWNLOAD_BYTES = 150 * 1024 * 1024
MAX_API_OUTPUT_TOKENS = 8192
CACHE_DIRECTORY = Path(".annual_report_vision_cache")
TEMP_IMAGE_DIRECTORY = Path("temp_images")
SHEET_SCHEMAS: dict[str, dict[str, Any]] = {}

# Python assigns only row-primary identifiers; related/source/taxonomy IDs
# remain model columns unless explicitly added to ROW_IDENTIFIERS.
CONFIG = {
    "system_aliases": {
        "staticcompanyid": "static_company_id",
        "companyname": "company_name",
        "company": "company_name",
        "financialyear": "financial_year",
        "ratingcycleyear": "rating_cycle_year",
        "sourcecode": "source_code",
        "retrievaldate": "retrieval_date",
        "sourcelocation": "source_location",
        "sourcepage": "source_location",
        "sourcepageorsection": "source_location",
        "datastatus": "data_status",
        "verificationstatus": "verification_status",
        "publicationstatus": "publication_status",
        "verifiedat": "verified_at",
    },
    "row_identifiers": {
        "recordkey", "linecode", "objectid", "objectobservationid", "eventid",
        "relationshipid", "attributeid", "policycode", "certificationcode",
        "targetcode", "adaptationcode", "programmecode", "riskcode",
        "scenariocode", "suppliercode", "assessmentcode", "traceabilitycode",
        "updatecode", "disclosurecode", "enforcementcode", "transitioncode",
        "factcode",
    },
    "reference_identifiers": {
        "regulatorcode", "metriccode", "objecttypecode", "parentobjectid",
        "subjectobjectid", "subjectobjecttype", "relatedcompanyid",
        "relatedobjectid", "relatedobjecttype", "unitcode", "currencycode",
        "evidencesourcecode", "relationshipmetriccode", "eventmetriccode",
        "attributemetriccode", "documentcode", "underlyingupdatecode",
        "parentrecordid", "parentrecordclass", "parentobjecttype",
        "sourcecode", "isin", "ticker", "sectorid", "issuerid",
        "frameworkcode", "regulatorcode", "suppliercode_ref",
    },
    "money_names": {
        "revenue", "operatingprofit", "profitbeforetax", "profitaftertax",
        "totalassets", "totalliabilities", "totalequity", "totaldebt",
        "cashandequivalents", "capitalexpenditure", "dividendspaid",
        "procurementspend", "annualspendamount", "climatecapex",
        "expectedfinancialloss", "penaltyamount", "amountinvested", "amountnumeric",
        "cashamount", "financialeffectvalue", "marketcapitalisation",
        "freefloatmarketcap", "valuetraded", "cashbalance",
    },
    "non_money_markers": {
        "percent", "rate", "ratio", "pershare", "peremployee", "percapita",
        "count", "headcount", "number", "frequency", "score", "intensity",
        "price", "volume", "shares", "latitude", "longitude", "year",
    },
}

BASE_EXTRACTION_PROMPT = (
    "You are an expert data extraction assistant. Zero hallucination. "
    "Extract ONLY what is visibly present in the supplied report pages, for "
    "the stated financial year only. If a value is not found, return 'N/A'. "
    "Never guess, infer, estimate, round or fill gaps from outside knowledge. "
    "Keep exact names and wording as printed. Return RAW numbers exactly as "
    "printed; do NOT apply any scale conversion yourself. If the source shows "
    "brackets, e.g. (1,245,000), return -1245000. Numbers must be plain JSON "
    "numbers with no commas, brackets or currency symbols. Replace the bullet "
    "character with a hyphen in all text. Dates must be YYYY-MM-DD. "
    "Return JSON only. Ignore all prior-year comparative columns, including "
    "tables that print current and prior years side by side. Never include "
    "irrelevant company/subsidiary data where the boundary is different."
)

_DECIMAL = decimal.Decimal
_SCALE_FACTORS = {
    "units": 1, "thousands": 1000, "millions": 1_000_000,
    "billions": 1_000_000_000,
}
_SCALE_KEYS = ("financial_position", "income", "cash_flow", "other")
_CURRENCY_RE = re.compile(r"\b(NGN|USD|EUR|GBP|GHS|ZAR|KES|XOF|XAF)\b", re.I)
_lock = threading.RLock()


def norm(name: str) -> str:
    """Normalise a column heading solely for rule lookup; never for output."""
    return re.sub(r"[^a-z0-9]", "", str(name).lower())


def parse_sheet_spec(text: str) -> dict[str, dict[str, Any]]:
    """Parse an ordered, numbered spec without altering sheet names or headers."""
    result: dict[str, dict[str, Any]] = {}
    pattern = re.compile(r"^\s*(\d+)\.\s+(.+?)\s*:\s*(.*?)\s*$")
    seen_casefold: set[str] = set()
    for number, original_line in enumerate(text.splitlines(), 1):
        line = original_line.strip()
        if not line:
            continue
        match = pattern.fullmatch(line)
        if not match:
            raise ValueError(f"Spec line {number} must be '1. SHEET: col1, col2'.")
        name = match.group(2).strip()
        empty = bool(re.search(r"\s*\(EMPTY\)\s*$", name, flags=re.I))
        if empty:
            name = re.sub(r"\s*\(EMPTY\)\s*$", "", name, flags=re.I).strip()
        if not name or len(name) > 31:
            raise ValueError(f"Invalid sheet '{name}': Excel names must be 1–31 chars.")
        if re.search(r"[\\/*?:\[\]]", name):
            raise ValueError(f"Invalid Excel character in sheet '{name}'.")
        if name.casefold() in seen_casefold:
            raise ValueError(f"Duplicate sheet name '{name}'.")
        cols = [x.strip() for x in match.group(3).split(",")]
        if any(not col for col in cols):
            raise ValueError(f"Missing column name on spec line {number}.")
        if len({c.casefold() for c in cols}) != len(cols):
            raise ValueError(f"Duplicate column heading on spec line {number}.")
        seen_casefold.add(name.casefold())
        result[name] = {"empty": empty, "columns": cols, "number": int(match.group(1))}
    if not result:
        raise ValueError("SHEET SPEC is empty.")
    return result


def _is_money(name: str) -> bool:
    """Identify likely money fields; never multiply a count, rate or share price."""
    k = norm(name)
    if any(marker in k for marker in CONFIG["non_money_markers"]):
        return False
    return (k in CONFIG["money_names"] or
            any(term in k for term in ("amount", "spend", "expenditure", "capex",
                                      "financialloss", "cashflow", "profit", "revenue")))


def _is_numeric(name: str) -> bool:
    """Infer whether an ordinary model field should hold an Excel number."""
    k = norm(name)
    if (k in ("financialyear", "ratingcycleyear", "fiscalyearend") or
            "boolean" in k or k.endswith("flag") or k.startswith("is")):
        return False
    if _is_money(name):
        return True
    return any(token in k for token in (
        "percent", "rate", "ratio", "amount", "count", "headcount", "numeric",
        "sharesheld", "sharesoutstanding", "numberof", "meeting", "latitude",
        "longitude", "score", "volume", "exits", "price", "openprice",
        "closingheadcount", "openingheadcount", "value", "total", "attended",
        "eligible", "frequency", "intensity", "capitalisation", "shareholding",
        "pershare", "earnings",
    )) and not any(x in k for x in ("name", "description", "status", "date",
                                         "method", "code", "note", "type", "text"))


def _is_date(name: str) -> bool:
    """Recognise date columns without treating a year field as a full date."""
    k = norm(name)
    return ("date" in k or k.endswith("at") or k in ("fiscalyearend",)) and k not in (
        "datastatus", "reportingdateformat")


def classify_columns(schema: dict[str, Any]) -> dict[str, dict[str, Any]]:
    """Classify each original column by cN, keeping the exact original name."""
    classes = {}
    # Only one row-primary ID per sheet: secondary IDs/codes are references
    # or reported identifiers, not synthetic replacements.
    row_identifier_assigned = False
    for i, name in enumerate(schema["columns"]):
        k = norm(name)
        alias = CONFIG["system_aliases"].get(k)
        if alias:
            kind = "SYSTEM"
        elif k in CONFIG["row_identifiers"] and not row_identifier_assigned:
            kind, alias = "SYSTEM", "row_id"
            row_identifier_assigned = True
        elif k in CONFIG["row_identifiers"]:
            kind = "MODEL"
        elif k in CONFIG["reference_identifiers"]:
            kind = "MODEL"
        elif k in (norm(v) for cfg in DERIVED_COLUMNS.values() for v in cfg.values()):
            kind = "DERIVED"
        else:
            kind = "MODEL"
        classes[f"c{i}"] = {
            "name": name, "kind": kind, "alias": alias,
            "money": _is_money(name), "numeric": _is_numeric(name),
            "date": _is_date(name),
        }
    return classes


def _sha(data: bytes) -> str:
    """Return a stable content hash for caching and filenames."""
    return hashlib.sha256(data).hexdigest()


def fetch_pdf_from_url(url: str) -> tuple[str, bytes]:
    """Download only a directly served, size-limited PDF; reject HTML and .docx."""
    parsed = urlparse(url.strip())
    if parsed.scheme not in {"http", "https"} or not parsed.netloc:
        raise ValueError(f"Invalid direct PDF link: {url}")
    try:
        response = requests.get(
            url, stream=True, timeout=(15, 45), allow_redirects=True,
            headers={"Accept": "application/pdf"},
        )
        response.raise_for_status()
        data = bytearray()
        for chunk in response.iter_content(chunk_size=1024 * 1024):
            if chunk:
                data.extend(chunk)
                if len(data) > MAX_PDF_DOWNLOAD_BYTES:
                    raise ValueError("PDF exceeds the 150 MB download limit.")
        if not bytes(data[:5]) == b"%PDF-":
            raise ValueError("This link did not return a PDF. Upload the file instead.")
        name = Path(urlparse(response.url).path).name or "linked_report.pdf"
        if not name.lower().endswith(".pdf"):
            name += ".pdf"
        return name, bytes(data)
    except requests.RequestException as exc:
        raise ValueError(f"Cannot download PDF; upload the file instead: {exc}") from exc


def _printed_label(page: Any, index: int, text: str) -> str:
    """Use PDF's explicit page label or a plausible printed footer, else index."""
    try:
        label = str(page.get_label()).strip()
        if label and label != str(index) and len(label) < 12:
            return label
    except (AttributeError, ValueError):
        label = ""
    lines = [s.strip() for s in text.splitlines() if s.strip()]
    for line in reversed(lines[-5:]):
        match = re.fullmatch(r"(?:page|pg\.?)[ :]*([0-9ivxlcdm]{1,5})", line, re.I)
        if match:
            return match.group(1)
        if re.fullmatch(r"[0-9]{1,4}", line):
            value = int(line)
            if 0 < value < 2000:
                return str(value)
    return str(index)


def _looks_like_table(text: str, drawing_count: int, image_count: int) -> bool:
    """Select vision for numeric tables, scanned pages and image-heavy pages."""
    if len(text.strip()) < 140:
        return True
    numbered = sum(bool(re.search(r"\d[\d,.()]*\s+\d", line))
                   for line in text.splitlines())
    financial_heading = bool(re.search(
        r"(?:statement of (?:financial position|profit|income|cash flows)|"
        r"balance sheet|consolidated statement|financial statement)", text, re.I))
    return ((financial_heading and (numbered >= 5 or drawing_count >= 35)) or
            (drawing_count > 115 and numbered >= 7) or
            (image_count > 2 and len(text) < 1000))


def _page_index_title(text: str, number: int, section: str = "") -> str:
    """Generate a compact page-topic line used for low-cost retrieval."""
    lines = [re.sub(r"\s+", " ", x).strip(" -|•") for x in text.splitlines()]
    titles = [x for x in lines if 4 <= len(x) <= 105 and not re.fullmatch(r"[0-9 ]+", x)]
    titles = titles[:3]
    return f"PAGE {number}: {section + ' | ' if section else ''}{' / '.join(titles)[:170]}"


def _read_pdf_input(data: bytes) -> list[dict[str, Any]]:
    """Extract all available PDF text and tag pages needing 300-DPI vision."""
    doc = fitz.open(stream=data, filetype="pdf")
    pages = []
    try:
        for i, page in enumerate(doc):
            text = page.get_text("text", sort=True)
            drawing_count = len(page.get_drawings())
            image_count = len(page.get_images(full=False))
            pages.append({
                "number": i + 1, "printed": _printed_label(page, i + 1, text),
                "text": text, "visual": _looks_like_table(
                    text, drawing_count, image_count),
                "unreadable": not text.strip() and image_count == 0 and drawing_count < 3,
                "section": "", "index": _page_index_title(text, i + 1),
            })
    finally:
        doc.close()
    return pages


def read_excel_input(data: bytes) -> list[dict[str, Any]]:
    """Read actual workbook cell values, with worksheet/row references and chunks."""
    wb = load_workbook(io.BytesIO(data), read_only=True, data_only=True)
    pages: list[dict[str, Any]] = []
    try:
        for ws in wb.worksheets:
            lines = []
            first_row = 1
            for row in ws.iter_rows():
                fragments = []
                for cell in row:
                    if cell.value is None:
                        continue
                    cell_text = str(cell.value).replace("\n", " ")
                    fragments.append(f"{cell.coordinate}={cell_text[:4000]}")
                if fragments:
                    lines.append(" | ".join(fragments))
                if len(lines) >= 65:
                    number = len(pages) + 1
                    body = f"WORKSHEET {ws.title} (starting at row {first_row})\n" + "\n".join(lines)
                    pages.append({"number": number, "printed": str(number),
                                  "text": body, "visual": False,
                                  "section": ws.title,
                                  "index": _page_index_title(body, number, ws.title)})
                    lines, first_row = [], row[0].row + 1 if row else first_row
            if lines or not any(p["section"] == ws.title for p in pages):
                number = len(pages) + 1
                body = f"WORKSHEET {ws.title} (starting at row {first_row})\n" + "\n".join(lines)
                pages.append({"number": number, "printed": str(number),
                              "text": body, "visual": False,
                              "section": ws.title,
                              "index": _page_index_title(body, number, ws.title)})
    finally:
        wb.close()
    return pages


def convert_pdf_to_images(
    data: bytes, page_numbers: list[int], report_hash: str, dpi: int = 300
) -> dict[int, Path]:
    """Render full pages at the requested DPI, each report in its own folder."""
    folder = TEMP_IMAGE_DIRECTORY / report_hash[:20]
    folder.mkdir(parents=True, exist_ok=True)
    output: dict[int, Path] = {}
    doc = fitz.open(stream=data, filetype="pdf")
    try:
        for n in sorted(set(page_numbers)):
            if not 1 <= n <= len(doc):
                continue
            path = folder / f"page_{n:04d}_{dpi}.jpg"
            if not path.exists():
                pix = doc[n - 1].get_pixmap(matrix=fitz.Matrix(dpi / 72, dpi / 72),
                                           alpha=False)
                img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
                img.save(path, format="JPEG", quality=86, optimize=True)
                img.close()
            output[n] = path
    finally:
        doc.close()
    return output


def _extract_json(response: str) -> dict[str, Any]:
    """Reject non-JSON model responses rather than silently accepting prose."""
    text = response.strip()
    if text.startswith("```"):
        text = re.sub(r"^```(?:json)?\s*|\s*```$", "", text, flags=re.I).strip()
    parsed = json.loads(text)
    if not isinstance(parsed, dict):
        raise ValueError("AI response must be a JSON object.")
    return parsed


def _provider_settings() -> tuple[str, str, str]:
    """Read model selection and API key from environment, never hardcoding them."""
    provider = os.getenv("VISION_PROVIDER", "").strip().lower()
    model = os.getenv("VISION_MODEL", "").strip()
    keys = {"openai": "OPENAI_API_KEY", "gemini": "GEMINI_API_KEY",
            "anthropic": "ANTHROPIC_API_KEY"}
    if provider not in keys:
        raise ValueError("Set VISION_PROVIDER=openai, gemini or anthropic in .env.")
    if not model:
        raise ValueError("Set VISION_MODEL in .env.")
    key = os.getenv(keys[provider], "").strip()
    if not key:
        raise ValueError(f"Set {keys[provider]} in .env.")
    return provider, model, key


def _call_provider(provider: str, model: str, key: str,
                   parts: list[tuple[str, Any]]) -> tuple[dict[str, Any], dict[str, int]]:
    """Use the selected official SDK only, with JSON-only/temperature-zero calls."""
    prompt_only = "\n\n".join(value for typ, value in parts if typ == "text")
    if provider == "openai":
        from openai import OpenAI
        client = OpenAI(api_key=key)
        content = []
        for typ, value in parts:
            if typ == "text":
                content.append({"type": "text", "text": value})
            else:
                encoded = base64.b64encode(value).decode("ascii")
                content.append({"type": "image_url", "image_url": {
                    "url": f"data:image/jpeg;base64,{encoded}", "detail": "high"}})
        response = client.chat.completions.create(
            model=model, temperature=0, max_completion_tokens=MAX_API_OUTPUT_TOKENS,
            response_format={"type": "json_object"},
            messages=[{"role": "user", "content": content}],
        )
        choice = response.choices[0]
        if choice.finish_reason == "length":
            raise ValueError("Model output was truncated; reduce page batch size.")
        usage = {"input_tokens": getattr(response.usage, "prompt_tokens", 0) or 0,
                 "output_tokens": getattr(response.usage, "completion_tokens", 0) or 0}
        return _extract_json(choice.message.content or ""), usage
    if provider == "gemini":
        import google.generativeai as genai
        genai.configure(api_key=key)
        gemini_parts: list[Any] = []
        for typ, value in parts:
            if typ == "text":
                gemini_parts.append(value)
            else:
                picture = Image.open(io.BytesIO(value)).convert("RGB")
                gemini_parts.append(picture)
        client = genai.GenerativeModel(model_name=model)
        response = client.generate_content(
            gemini_parts, generation_config={"temperature": 0,
                "response_mime_type": "application/json",
                "max_output_tokens": MAX_API_OUTPUT_TOKENS},
        )
        usage_meta = getattr(response, "usage_metadata", None)
        usage = {"input_tokens": getattr(usage_meta, "prompt_token_count", 0) or 0,
                 "output_tokens": getattr(usage_meta, "candidates_token_count", 0) or 0}
        return _extract_json(response.text), usage
    if provider == "anthropic":
        import anthropic
        client = anthropic.Anthropic(api_key=key)
        content = []
        for typ, value in parts:
            if typ == "text":
                content.append({"type": "text", "text": value})
            else:
                content.append({"type": "image", "source": {
                    "type": "base64", "media_type": "image/jpeg",
                    "data": base64.b64encode(value).decode("ascii")}})
        response = client.messages.create(
            model=model, max_tokens=MAX_API_OUTPUT_TOKENS, temperature=0,
            system="Return exactly one valid JSON object, with no markdown or preamble.",
            messages=[{"role": "user", "content": content}],
        )
        if response.stop_reason == "max_tokens":
            raise ValueError("Model output was truncated; reduce page batch size.")
        answer = "".join(block.text for block in response.content if block.type == "text")
        usage = {"input_tokens": getattr(response.usage, "input_tokens", 0) or 0,
                 "output_tokens": getattr(response.usage, "output_tokens", 0) or 0}
        return _extract_json(answer), usage
    raise ValueError(f"Unknown provider: {provider}")


def _vision_json(parts: list[tuple[str, Any]], report_hash: str,
                 namespace: str) -> tuple[dict[str, Any], dict[str, int], bool]:
    """Retry transient SDK failures and cache successful responses by content hash."""
    provider, model, key = _provider_settings()
    fingerprint = hashlib.sha256()
    fingerprint.update(f"v2|{provider}|{model}|{report_hash}|{namespace}".encode())
    for typ, value in parts:
        fingerprint.update(typ.encode())
        fingerprint.update(value.encode() if typ == "text" else value)
    digest = fingerprint.hexdigest()
    folder = CACHE_DIRECTORY / report_hash[:20]
    path = folder / f"{digest}.json"
    with _lock:
        if path.exists():
            try:
                with path.open("r", encoding="utf-8") as handle:
                    cached = json.load(handle)
                return cached["result"], {"input_tokens": 0, "output_tokens": 0}, True
            except (ValueError, OSError, KeyError):
                # Corrupt or partial caches are ignored and refreshed.
                cached = None
    error = None
    for attempt in range(3):
        try:
            result, usage = _call_provider(provider, model, key, parts)
            folder.mkdir(parents=True, exist_ok=True)
            tmp_path = folder / f"{digest}.{threading.get_ident()}.tmp"
            with tmp_path.open("w", encoding="utf-8") as handle:
                json.dump({"result": result, "usage": usage}, handle,
                          ensure_ascii=False)
            os.replace(tmp_path, path)
            return result, usage, False
        except Exception as exc:
            error = exc
            if attempt < 2:
                is_rate = any(word in str(exc).lower() for word in (
                    "rate", "429", "too many requests", "quota"))
                time.sleep(5 * (2 ** attempt) if is_rate else 5)
    raise RuntimeError(f"AI request failed after three attempts: {error}") from error


def _detect_visible_scale(text: str) -> str:
    """Accept explicit unit annotations only; an absent scale stays UNKNOWN."""
    s = text.lower()
    if re.search(r"\b(?:in|figures in|expressed in|amounts in)\s+"
                 r"(?:[a-z$₦ ]{0,18})?\b(?:billion|billions)\b", s):
        return "billions"
    if re.search(r"\b(?:in|figures in|expressed in|amounts in)\s+"
                 r"(?:[a-z$₦ ]{0,18})?\b(?:million|millions)\b", s):
        return "millions"
    if re.search(r"\b(?:in|figures in|expressed in|amounts in)\s+"
                 r"(?:[a-z$₦ ]{0,18})?\b(?:thousand|thousands)\b|"
                 r"(?:ngn|n|₦|usd|ghs|zar)\s*['’]000\b", s):
        return "thousands"
    if re.search(r"\b(?:figures|amounts)\s+(?:are\s+)?in\s+(?:whole\s+)?units\b", s):
        return "units"
    return "UNKNOWN"


def _scale_context(pages: list[dict[str, Any]]) -> list[str]:
    """Collect source headings and statement unit annotations for report profiling."""
    snippets = []
    for page in pages:
        lines = page["text"].splitlines()
        for i, line in enumerate(lines):
            if re.search(r"(?:thousands|millions|billions|['’]000|"
                         r"statement of financial position|statement of profit|"
                         r"cash flows|reporting currency)", line, re.I):
                excerpt = " ".join(lines[max(0, i - 1):i + 2]).strip()
                snippets.append(f"PAGE {page['number']}: {excerpt[:330]}")
                if len(snippets) >= 45:
                    return snippets
    return snippets


def _year_from_text(text: str) -> int | None:
    """Find a plausible current year from a title, not comparative table columns."""
    matches = re.findall(
        r"\b(?:FY\s*|financial\s+year\s*|annual\s+report\s*|"
        r"year\s+ended\s+(?:\d{1,2}\s+[a-z]+\s+|\d{1,2}[/\-]\d{1,2}[/\-]))"
        r"(20\d{2})\b", text, flags=re.I)
    if matches:
        return int(matches[0])
    return None


def _enrich_scanned_page_index(report: dict[str, Any]) -> dict[str, int]:
    """Caption textless PDF pages using low-DPI vision so retrieval can find them."""
    totals = {"input_tokens": 0, "output_tokens": 0}
    if report["kind"] != "pdf":
        return totals
    candidates = [p["number"] for p in report["pages"]
                  if len(p["text"].strip()) < 140 and not p.get("unreadable")]
    for i in range(0, len(candidates), 5):
        group = candidates[i:i + 5]
        images = convert_pdf_to_images(report["bytes"], group, report["hash"], dpi=110)
        parts: list[tuple[str, Any]] = [("text",
            "Describe the visible topic of each labelled PDF page in <=14 words. "
            "No guesses; never transcribe values. Return JSON object "
            "{\"pages\":[{\"number\":1,\"topic\":\"...\"},...]}. "
            "If visually blank, say BLANK. Only use supplied PAGE labels.")]
        for n in group:
            parts.append(("text", f"PAGE {n}"))
            parts.append(("image", images[n].read_bytes()))
        try:
            response, usage, _ = _vision_json(parts, report["hash"],
                                              f"scan_page_index|{group}")
            totals["input_tokens"] += usage.get("input_tokens", 0)
            totals["output_tokens"] += usage.get("output_tokens", 0)
            for item in response.get("pages", []):
                if not isinstance(item, dict):
                    continue
                try:
                    n = int(item.get("number"))
                except (TypeError, ValueError):
                    continue
                if n in group:
                    topic = str(item.get("topic", "")).strip()[:170]
                    if topic:
                        report["pages"][n - 1]["index"] = f"PAGE {n}: {topic}"
                        if topic.upper() == "BLANK":
                            report["pages"][n - 1]["unreadable"] = True
        except Exception as exc:
            report.setdefault("index_warnings", []).append(
                f"Page-index vision failed for pages {group}: {exc}")
    return totals


def detect_report_profile(report: dict[str, Any]) -> tuple[dict[str, Any], dict[str, int]]:
    """Profile a PDF/XLSX cheaply; retain model uncertainty for human confirmation."""
    pages = report["pages"]
    index_usage = _enrich_scanned_page_index(report)
    index = "\n".join(p["index"] for p in pages)
    opening = "\n".join(
        f"PAGE {p['number']}: {p['text'][:1900]}" for p in pages[:4])
    statement_snippets = "\n".join(_scale_context(pages))
    prompt = (
        "Profile ONE annual report. Return JSON object ONLY with keys: "
        "report_name (string), company_name (string), financial_year (integer "
        "or null), fiscal_year_end (YYYY-MM-DD or N/A), "
        "scales (object with financial_position, income, cash_flow, other; "
        "each 'units','thousands','millions','billions' or 'UNKNOWN'), "
        "currencies (object with the same four keys, each ISO currency code or N/A). "
        "Select CURRENT YEAR, not prior-year comparative; never make up a scale. "
        "Scale for 'other' should be UNKNOWN unless explicit other-disclosure "
        "units are shown. Return N/A for unsupported fields. "
        f"Filename: {report['filename']}\n"
        f"PAGE INDEX:\n{index[:28000]}\n"
        f"FRONT PAGES:\n{opening[:8000]}\n"
        f"STATEMENT AND SCALE EVIDENCE:\n{statement_snippets[:12000]}"
    )
    parts: list[tuple[str, Any]] = [("text", prompt)]
    if report["kind"] == "pdf":
        visual = [p["number"] for p in pages[:4] if p["visual"]]
        images = convert_pdf_to_images(report["bytes"], visual[:2], report["hash"], 120)
        for n, path in images.items():
            parts.append(("text", f"PAGE {n} (report-front image):"))
            parts.append(("image", path.read_bytes()))
    raw, usage, cached = _vision_json(parts, report["hash"], "profile")
    usage["input_tokens"] += index_usage["input_tokens"]
    usage["output_tokens"] += index_usage["output_tokens"]
    year = raw.get("financial_year")
    try:
        year = int(year) if year not in (None, "N/A", "") else None
        if year is not None and not 1900 <= year <= dt.date.today().year + 1:
            year = None
    except (TypeError, ValueError):
        year = None
    if year is None:
        year = _year_from_text(opening)
    raw_scales = raw.get("scales", {}) if isinstance(raw.get("scales"), dict) else {}
    raw_currencies = raw.get("currencies", {}) if isinstance(raw.get("currencies"), dict) else {}
    scales = {}
    currencies = {}
    for key in _SCALE_KEYS:
        scale = str(raw_scales.get(key, "UNKNOWN")).strip().lower()
        scales[key] = scale if scale in _SCALE_FACTORS else "UNKNOWN"
        currency = str(raw_currencies.get(key, "N/A")).strip().upper()
        currencies[key] = currency if re.fullmatch(r"[A-Z]{3}", currency) else "N/A"
    report_name = str(raw.get("report_name", "")).strip()
    company = str(raw.get("company_name", "")).strip()
    report_name = report_name if report_name and report_name != "N/A" else Path(report["filename"]).stem
    company = company if company and company != "N/A" else "N/A"
    end_date = str(raw.get("fiscal_year_end", "N/A")).strip()
    if not re.fullmatch(r"\d{4}-\d{2}-\d{2}", end_date):
        end_date = "N/A"
    profile = {"filename": report["filename"], "report_name": report_name,
               "company_detected": company, "financial_year": year,
               "fiscal_year_end": end_date, "scales": scales,
               "currencies": currencies, "cached_profile": cached}
    return profile, usage


def _company_key(name: str) -> str:
    """Remove punctuation and common legal suffixes for company checks only."""
    words = re.findall(r"[a-z0-9]+", name.lower())
    legal = {"plc", "ltd", "limited", "inc", "incorporated", "llc", "co",
             "company", "holdings", "holding", "group"}
    return " ".join(word for word in words if word not in legal)


def check_company_consistency(profiles: list[dict[str, Any]],
                              company_name: str) -> list[dict[str, Any]]:
    """Flag files unlike the majority or requested company; untick them by default."""
    keys = [_company_key(p.get("company_detected", "")) for p in profiles]
    nonempty = [k for k in keys if k and k != "n a"]
    majority = Counter(nonempty).most_common(1)[0][0] if nonempty else ""
    expected = _company_key(company_name)
    output = []
    for item, key in zip(profiles, keys):
        anchor = majority or expected
        similarity = (difflib.SequenceMatcher(None, key, anchor).ratio()
                      if key and anchor else 0)
        expected_sim = (difflib.SequenceMatcher(None, key, expected).ratio()
                        if key and expected else 0)
        flagged = bool(key and key != "n a" and anchor and
                       similarity < 0.68 and expected_sim < 0.68)
        p = dict(item)
        p["company_warning"] = "COMPANY MISMATCH" if flagged else ""
        p["include"] = not flagged
        output.append(p)
    return output


def check_year_coverage(years: list[int]) -> list[int]:
    """Return missing years between the minimum and maximum selected FYs."""
    clean = sorted(set(int(y) for y in years))
    if len(clean) < 2:
        return []
    return [y for y in range(clean[0], clean[-1] + 1) if y not in clean]


def retrieve_pages_for_sheet(sheet: str, schema: dict[str, Any],
                             report: dict[str, Any], year: int,
                             stats: dict[str, Any]) -> tuple[list[int], list[str]]:
    """Retrieve evidence pages from the index with no report images or table output."""
    indexed = "\n".join(page["index"] for page in report["pages"])
    model_cols = [item["name"] for item in classify_columns(schema).values()
                  if item["kind"] == "MODEL"]
    prompt = (
        "You are a precise evidence-page locator for an annual report. "
        "Respond ONLY with JSON object {\"pages\":[integer,...]}. "
        f"Report current financial year: {year}. Sheet: {sheet}. "
        f"Fields: {', '.join(model_cols)}. "
        "Select page numbers in the PAGE INDEX most likely to contain explicit "
        "CURRENT-YEAR evidence for those fields. Do not select general filler "
        "pages and do not favour prior-year comparatives. Choose up to 12 pages; "
        "if no likely evidence exists, return []. All numbers are 1-based PAGE "
        "indices, not years. The page index follows:\n" + indexed[:50000]
    )
    response, usage, cached = _vision_json(
        [("text", prompt)], report["hash"], f"retrieve|{sheet}|{year}")
    _add_usage(stats, usage, cached)
    raw_nums = response.get("pages", [])
    if not isinstance(raw_nums, list):
        raw_nums = []
    valid = []
    alerts = []
    for candidate in raw_nums:
        try:
            n = int(candidate)
        except (TypeError, ValueError):
            continue
        if 1 <= n <= len(report["pages"]) and n not in valid:
            valid.append(n)
        else:
            alerts.append(f"INVALID RETRIEVED PAGE {candidate}")
    if len(valid) > MAX_PAGES_PER_SHEET:
        alerts.append(f"PAGE RETRIEVAL TRUNCATED {len(valid)} to {MAX_PAGES_PER_SHEET}")
        valid = valid[:MAX_PAGES_PER_SHEET]
    return valid, alerts


def build_extraction_prompt(sheet: str, schema: dict[str, Any],
                            year: int, report_name: str) -> str:
    """Request cN keys only for model columns and require page evidence metadata."""
    classes = classify_columns(schema)
    column_list = "\n".join(
        f"{key}: {item['name']}" for key, item in classes.items()
        if item["kind"] == "MODEL")
    money_keys = [key for key, item in classes.items()
                  if item["kind"] == "MODEL" and item["money"]]
    return (
        BASE_EXTRACTION_PROMPT + "\n\n"
        f"Report: {report_name}. CURRENT financial year: {year}.\n"
        f"SHEET: {sheet}\nMODEL COLUMNS (only these cN keys):\n{column_list}\n"
        "Return a JSON object with exactly a 'rows' array. Each array entry "
        "may have the listed cN MODEL keys plus the following helper keys:\n"
        "_pages: list of PAGE numbers cited for THIS row; "
        "_scale: the figure-unit description exactly observed, one of "
        "units/thousands/millions/billions/UNKNOWN; "
        "_scale_fields: list of cN keys that are MONETARY amounts and visibly "
        "expressed in report units; "
        "_raw: object of each such cN -> original printed numeric value; "
        "_scale_by_field: object cN -> scale visibly used by that field; "
        "_statement_types: object cN -> financial_position/income/cash_flow/other, "
        "to identify the source financial statement of each monetary field; "
        "_evidence_extract: a verbatim quote of at most 25 words from the page, "
        "supporting the specific row. If the schema has an evidence_extract or "
        "Evidence / Notes MODEL field, populate it from the report only.\n"
        f"The most likely monetary keys are {money_keys}. Only include a key "
        "in _scale_fields when the source explicitly uses a scaled unit; if "
        "monetary amounts are given in unscaled units, include them with "
        "_scale='units'. NEVER include counts, ratios, percentages, share "
        "prices, total shares, earnings per share, or headcount in _scale_fields.\n"
        "PAGE n labels are 1-based input-page indices. Cite the actual PAGE "
        "numbers in the supplied material, never fabricate them. Each distinct "
        "director, subsidiary, policy, incident, emission scope, shareholder "
        "or other separate record requires its own row. A field that is not "
        "explicitly supported is 'N/A'. No aggregate or copied comparative "
        "values. If no evidence exists return {\"rows\":[]}."
    )


def extract_sheet_with_vision(sheet: str, schema: dict[str, Any],
                              report: dict[str, Any], year: int,
                              report_name: str, pages: list[int],
                              stats: dict[str, Any],
                              issues: list[dict[str, Any]] | None = None) -> list[dict[str, Any]]:
    """Send full pages in five-page batches, mixing text and 300-DPI images."""
    if not pages:
        return []
    found: list[dict[str, Any]] = []
    for offset in range(0, len(pages), PAGES_PER_VISION_BATCH):
        group = pages[offset:offset + PAGES_PER_VISION_BATCH]
        parts: list[tuple[str, Any]] = [
            ("text", build_extraction_prompt(sheet, schema, year, report_name))]
        to_image = [n for n in group if report["kind"] == "pdf"
                    and report["pages"][n - 1]["visual"]]
        images = (convert_pdf_to_images(report["bytes"], to_image, report["hash"], 300)
                  if to_image else {})
        for n in group:
            page = report["pages"][n - 1]
            parts.append(("text", f"PAGE {n}\n" +
                          (page["text"][:30000] if n not in images
                           else "Full-page image follows; do not guess text.")))
            if n in images:
                parts.append(("image", images[n].read_bytes()))
        try:
            response, usage, cached = _vision_json(
                parts, report["hash"], f"extract|{sheet}|{year}|{group}")
            _add_usage(stats, usage, cached)
            rows = response.get("rows", [])
            if not isinstance(rows, list):
                raise ValueError(f"{sheet}: model JSON 'rows' must be a list.")
            for entry in rows:
                if isinstance(entry, dict):
                    entry["_provided_pages"] = list(group)
                    found.append(entry)
        except Exception as exc:
            if issues is None:
                raise
            issues.append({"check_name": "BATCH EXTRACTION ERROR", "sheet": sheet,
                           "notes": f"PAGE {group}: {type(exc).__name__}: {exc}"})
            # Already extracted batches remain in the output and are audited.
            continue
    return found


def _normalise_model_value(value: Any) -> Any:
    """Convert missing JSON values and bullet glyphs to safe source-preserving text."""
    if value is None or value == "":
        return "N/A"
    if isinstance(value, (dict, list)):
        return json.dumps(value, ensure_ascii=False)
    if isinstance(value, str):
        value = value.replace("•", "-").strip()
        return value if value and value.lower() not in ("null", "none", "unknown") else "N/A"
    return value


def validate_rows_against_schema(raw_rows: list[dict[str, Any]],
                                 schema: dict[str, Any], sheet: str,
                                 issues: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Drop unknown model keys, fill missing originals and preserve audit helpers."""
    classes = classify_columns(schema)
    model_keys = {key for key, col in classes.items() if col["kind"] == "MODEL"}
    output = []
    for index, raw in enumerate(raw_rows, 1):
        model_provided = {k for k in raw if re.fullmatch(r"c\d+", k)}
        unknown = sorted(model_provided - model_keys)
        if unknown:
            issues.append({"check_name": "UNKNOWN MODEL KEYS", "sheet": sheet,
                           "notes": f"Row {index}: dropped {unknown}"})
        missing = model_keys - model_provided
        if missing:
            issues.append({"check_name": "MISSING MODEL KEYS", "sheet": sheet,
                           "notes": f"Row {index}: filled {len(missing)} absent keys with N/A"})
        values = {col: "N/A" for col in schema["columns"]}
        for key in model_keys:
            values[classes[key]["name"]] = _normalise_model_value(raw.get(key))
        cited = raw.get("_pages", [])
        cited = cited if isinstance(cited, list) else []
        pages = []
        for n in cited:
            try:
                value = int(n)
                if value not in pages:
                    pages.append(value)
            except (TypeError, ValueError):
                issues.append({"check_name": "INVALID PAGE NUMBER", "sheet": sheet,
                               "notes": f"Row {index}: bad page token {n!r}"})
        metadata = {
            "pages": pages, "scale": str(raw.get("_scale", "UNKNOWN")),
            "scale_fields": raw.get("_scale_fields", []),
            "raw": raw.get("_raw", {}),
            "scale_by_field": raw.get("_scale_by_field", {}),
            "statement_types": raw.get("_statement_types", {}),
            "evidence": str(raw.get("_evidence_extract", ""))[:1100],
            "provided_pages": raw.get("_provided_pages", []),
            "conversion": {}, "quote_verified": "NOT TESTABLE",
        }
        output.append({"values": values, "meta": metadata})
    return output


def apply_negative_sign_rule(value: Any) -> Any:
    """Interpret accounting parentheses as negatives without scaling."""
    if isinstance(value, bool) or value in (None, "N/A"):
        return "N/A" if value is None else value
    if isinstance(value, (int, float, decimal.Decimal)):
        return value
    string = str(value).strip()
    sign = -1 if re.fullmatch(r"\([\d, ]+(?:\.\d+)?\)", string) else 1
    if sign == -1:
        string = string[1:-1]
    if re.fullmatch(r"-?\d[\d, ]*(?:\.\d+)?", string):
        raw = string.replace(",", "").replace(" ", "")
        try:
            number = _DECIMAL(raw) * sign
            return int(number) if number == int(number) else float(number)
        except (decimal.InvalidOperation, ValueError):
            return value
    return value


def _format_date(value: Any) -> str:
    """Normalise unambiguous dates; reject ambiguous day/month strings."""
    if value in (None, "", "N/A"):
        return "N/A"
    if isinstance(value, (dt.datetime, dt.date)):
        return value.date().isoformat() if isinstance(value, dt.datetime) else value.isoformat()
    text = str(value).strip()
    for fmt in ("%Y-%m-%d", "%Y-%m-%d %H:%M:%S", "%d %B %Y",
                "%d %b %Y", "%d-%b-%Y", "%B %d, %Y", "%b %d, %Y"):
        try:
            return dt.datetime.strptime(text, fmt).date().isoformat()
        except ValueError:
            continue
    if re.fullmatch(r"\d{1,2}[/-]\d{1,2}[/-]\d{4}", text):
        first, second, _ = re.split(r"[/-]", text)
        if int(first) <= 12 and int(second) <= 12:
            return "N/A"  # source gives no unambiguous DD/MM or MM/DD order
        for fmt in ("%d/%m/%Y", "%m/%d/%Y", "%d-%m-%Y", "%m-%d-%Y"):
            try:
                return dt.datetime.strptime(text, fmt).date().isoformat()
            except ValueError:
                continue
    return "N/A"


def _statement_for_field(name: str, model_choice: str) -> str:
    """Map monetary evidence to a separately confirmed financial-statement scale."""
    if model_choice in _SCALE_KEYS:
        return model_choice
    k = norm(name)
    if any(x in k for x in ("asset", "liabilit", "equity", "totaldebt",
                           "cashandequivalent")):
        return "financial_position"
    if any(x in k for x in ("revenue", "profit", "earning", "dividend")):
        return "income"
    if any(x in k for x in ("capitalexpenditure", "cashflow", "cashpaid")):
        return "cash_flow"
    return "other"


def _normalise_scale(raw: Any) -> str:
    """Match an explicitly named magnitude; unknown never becomes units."""
    s = str(raw).strip().lower()
    for scale in ("billions", "millions", "thousands", "units"):
        if s == scale or re.search(r"\b" + scale[:-1] + r"s?\b", s):
            return scale
    if re.search(r"['’]000", s):
        return "thousands"
    return "UNKNOWN"


def _numeric_excel(value: decimal.Decimal) -> int | float:
    """Produce native Excel numeric types without stringifying scaled values."""
    return int(value) if value == value.to_integral_value() else float(value)


def apply_conversion(row: dict[str, Any], sheet: str,
                     schema: dict[str, Any], profile: dict[str, Any],
                     issues: list[dict[str, Any]]) -> dict[str, Any]:
    """Scale ONLY declared money cN fields, using confirmed statement scales."""
    classes = classify_columns(schema)
    raw_values = row["values"]
    meta = row["meta"]
    scale_keys = meta["scale_fields"]
    scale_keys = scale_keys if isinstance(scale_keys, list) else []
    raw_meta = meta["raw"] if isinstance(meta["raw"], dict) else {}
    statement_meta = (meta["statement_types"] if isinstance(meta["statement_types"], dict)
                      else {})
    scale_by_field = (meta["scale_by_field"] if isinstance(meta["scale_by_field"], dict)
                      else {})
    permitted = set()
    for key in scale_keys:
        if key not in classes or classes[key]["kind"] != "MODEL" or not classes[key]["money"]:
            issues.append({"check_name": "INVALID SCALE FIELD", "sheet": sheet,
                           "notes": f"Refused to scale {key!r}; not a monetary MODEL field."})
        else:
            permitted.add(key)
    for key, col in classes.items():
        if col["kind"] != "MODEL":
            continue
        name = col["name"]
        value = raw_values[name]
        if value == "N/A":
            continue
        if col["date"]:
            converted = _format_date(value)
            if converted == "N/A":
                issues.append({"check_name": "AMBIGUOUS OR INVALID DATE", "sheet": sheet,
                               "notes": f"{name}={value!r} cannot be safely normalised."})
            raw_values[name] = converted
            continue
        if col["numeric"]:
            parsed = apply_negative_sign_rule(value)
            if isinstance(parsed, (int, float, decimal.Decimal)) and not isinstance(parsed, bool):
                raw_values[name] = parsed
            elif value != "N/A":
                issues.append({"check_name": "NUMBER NOT PARSED", "sheet": sheet,
                               "notes": f"{name}={str(value)[:120]!r}"})
                raw_values[name] = "N/A"
    for key in permitted:
        col = classes[key]
        name = col["name"]
        value = raw_values[name]
        if value == "N/A":
            continue
        if not isinstance(value, (int, float, decimal.Decimal)) or isinstance(value, bool):
            issues.append({"check_name": "SCALE INPUT NONNUMERIC", "sheet": sheet,
                           "notes": f"{name} is not a number; cannot scale."})
            continue
        statement = _statement_for_field(name, str(statement_meta.get(key, "")))
        confirmed_scale = profile["scales"].get(statement, "UNKNOWN")
        if confirmed_scale not in _SCALE_FACTORS:
            issues.append({"check_name": "SCALE NOT CONFIRMED", "sheet": sheet,
                           "notes": f"{statement}: {name} was not converted."})
            continue
        declared = _normalise_scale(scale_by_field.get(key, meta["scale"]))
        if declared != "UNKNOWN" and declared != confirmed_scale:
            issues.append({"check_name": "SCALE MISMATCH", "sheet": sheet,
                           "notes": (f"{name}: page scale {declared} != user "
                                     f"confirmation {confirmed_scale}; confirm source.")})
        original = apply_negative_sign_rule(raw_meta.get(key, value))
        if isinstance(original, (int, float, decimal.Decimal)) and not isinstance(original, bool):
            if _DECIMAL(str(original)) != _DECIMAL(str(value)):
                issues.append({"check_name": "RAW VALUE MISMATCH", "sheet": sheet,
                               "notes": f"{name}: _raw differs from extracted value."})
        number = _DECIMAL(str(value)) * _SCALE_FACTORS[confirmed_scale]
        raw_values[name] = _numeric_excel(number)
        meta["conversion"][name] = {
            "raw": raw_meta.get(key, value), "scale": confirmed_scale,
            "model_scale": declared, "factor": _SCALE_FACTORS[confirmed_scale],
            "statement": statement,
        }
    for key, col in classes.items():
        if (col["kind"] == "MODEL" and col["money"] and key not in permitted
                and raw_values[col["name"]] != "N/A"):
            issues.append({"check_name": "MONEY SCALE NOT DECLARED", "sheet": sheet,
                           "notes": f"{col['name']} kept raw; check source scale."})
    return row


def _source_string(report_name: str, pages: list[int],
                   report_pages: list[dict[str, Any]]) -> str:
    """Format traceability exactly as ReportName (pg X, Y, Z)."""
    valid = sorted(set(n for n in pages if 1 <= n <= len(report_pages)))
    if not valid:
        return "Source on Website - No Report Page"
    labels = [(report_pages[n - 1]["printed"] if PAGE_NUMBERING == "printed"
               else str(n)) for n in valid]
    return f"{report_name} (pg {', '.join(labels)})"


def fill_system_columns(row: dict[str, Any], sheet: str,
                        schema: dict[str, Any], profile: dict[str, Any],
                        company_name: str, static_company_id: str,
                        rating_cycle_year: str, source_code: str,
                        row_number: int, pages: list[dict[str, Any]]) -> dict[str, Any]:
    """Fill contextual and deterministic system values without calling the model."""
    values = row["values"]
    context = {
        "static_company_id": static_company_id,
        "company_name": company_name,
        "financial_year": profile["financial_year"],
        "rating_cycle_year": rating_cycle_year or "N/A",
        "source_code": source_code or "N/A",
        "retrieval_date": dt.date.today().isoformat(),
        "source_location": _source_string(profile["report_name"], row["meta"]["pages"], pages),
    }
    for key, col in classify_columns(schema).items():
        if col["kind"] != "SYSTEM":
            continue
        alias = col["alias"]
        if alias == "row_id":
            value = f"{static_company_id}-{sheet}-{profile['financial_year']}-{row_number}"
        else:
            value = context.get(alias, STATUS_DEFAULTS.get(alias, "N/A"))
        values[col["name"]] = value
    return row


def _decimal_or_none(value: Any) -> decimal.Decimal | None:
    """Return a valid decimal or None for a missing/non-numeric operand."""
    raw = apply_negative_sign_rule(value)
    if raw == "N/A" or isinstance(raw, bool):
        return None
    try:
        return _DECIMAL(str(raw))
    except (decimal.InvalidOperation, TypeError, ValueError):
        return None


def compute_derived_columns(row: dict[str, Any], schema: dict[str, Any],
                            sheet: str) -> dict[str, Any]:
    """Compute turnover averages/rates exclusively from already populated fields."""
    if norm(sheet) != "employeeturnover":
        return row
    vals = row["values"]
    lookup = {norm(name): name for name in schema["columns"]}

    def read(name: str) -> Any:
        return vals.get(lookup.get(norm(name), ""), "N/A")

    def write(name: str, value: Any) -> None:
        key = lookup.get(norm(name))
        if key is not None:
            vals[key] = value

    reported_avg = _decimal_or_none(read("Reported Average Headcount"))
    opening = _decimal_or_none(read("Opening Headcount"))
    closing = _decimal_or_none(read("Closing Headcount"))
    avg = (reported_avg if reported_avg is not None else
           (opening + closing) / 2 if opening is not None and closing is not None else None)
    write("Average Headcount Used", _numeric_excel(avg) if avg is not None else "N/A")
    exits = _decimal_or_none(read("Employee Exits"))
    calc = exits / avg * 100 if exits is not None and avg not in (None, 0) else None
    write("Calculated Exit Rate (%)", _numeric_excel(calc) if calc is not None else "N/A")
    reported = _decimal_or_none(read("Reported Turnover Rate (%)"))
    final = reported if reported is not None else calc
    write("Final Turnover / Attrition Rate (%)",
          _numeric_excel(final) if final is not None else "N/A")
    if reported is not None:
        method = "Reported turnover rate selected from source"
    elif calc is not None:
        basis = "reported average headcount" if reported_avg is not None else \
            "(opening + closing) / 2"
        method = f"Employee exits / {basis} x 100"
    else:
        method = "N/A"
    write("Calculation / Selection Method", method)
    return row


def _normalise_quote(text: str) -> str:
    """Normalise whitespace and bullets to compare source quotations."""
    return re.sub(r"\s+", " ", text.replace("•", "-")).strip().casefold()


def verify_evidence(row: dict[str, Any], report: dict[str, Any]) -> str:
    """Verify verbatim evidence against text-layer pages, retaining failures."""
    pages = row["meta"]["pages"]
    if not pages:
        return "NOT TESTABLE"
    quote = row["meta"]["evidence"].strip()
    if not quote:
        for name, val in row["values"].items():
            if norm(name) in ("evidenceextract", "evidencenotes") and val != "N/A":
                quote = str(val)
                break
    if not quote:
        return "UNVERIFIED QUOTE"
    quote = " ".join(quote.split()[:25])
    for n in pages:
        if 1 <= n <= len(report["pages"]):
            body = report["pages"][n - 1]["text"]
            if body.strip() and _normalise_quote(quote) in _normalise_quote(body):
                return "VERIFIED"
    if all(1 <= n <= len(report["pages"]) and
           not report["pages"][n - 1]["text"].strip() for n in pages):
        return "NOT TESTABLE - SCANNED"
    return "UNVERIFIED QUOTE"


def _add_usage(target: dict[str, Any], usage: dict[str, int], cached: bool) -> None:
    """Accumulate only actually billed token usage, separately counting cache hits."""
    target["input_tokens"] += int(usage.get("input_tokens", 0))
    target["output_tokens"] += int(usage.get("output_tokens", 0))
    if cached:
        target["cache_hits"] += 1
    else:
        target["api_requests"] += 1


def _missing_row(schema: dict[str, Any]) -> dict[str, Any]:
    """Build the mandatory one-row N/A placeholder for a checked empty result."""
    return {"values": {name: "N/A" for name in schema["columns"]},
            "meta": {"pages": [], "scale": "UNKNOWN", "scale_fields": [],
                     "raw": {}, "scale_by_field": {}, "statement_types": {},
                     "evidence": "", "provided_pages": [], "conversion": {},
                     "quote_verified": "NOT TESTABLE"}}


def process_one_report(report: dict[str, Any], profile: dict[str, Any],
                       schemas: dict[str, dict[str, Any]],
                       company_name: str, static_company_id: str,
                       rating_cycle_year: str, source_code: str,
                       events: queue.Queue) -> dict[str, Any]:
    """Extract every nonempty sheet; isolate failures so other sheets proceed."""
    year = profile["financial_year"]
    stats = {"input_tokens": 0, "output_tokens": 0, "api_requests": 0,
             "cache_hits": 0}
    issues: list[dict[str, Any]] = []
    for page in report["pages"]:
        if page.get("unreadable"):
            issues.append({"check_name": "UNREADABLE PAGE", "sheet": "REPORT",
                           "notes": f"PAGE {page['number']} has no readable text or images"})
    completed: dict[str, list[dict[str, Any]]] = {}
    for sheet, schema in schemas.items():
        if schema["empty"]:
            completed[sheet] = []
            events.put((report["source_order"], sheet, 0, "EMPTY - headers only"))
            continue
        try:
            page_numbers, lookup_issues = retrieve_pages_for_sheet(
                sheet, schema, report, year, stats)
            for message in lookup_issues:
                issues.append({"check_name": "PAGE RETRIEVAL", "sheet": sheet,
                               "notes": message})
            raw_rows = extract_sheet_with_vision(
                sheet, schema, report, year, profile["report_name"],
                page_numbers, stats, issues)
            checked = validate_rows_against_schema(raw_rows, schema, sheet, issues)
            if not checked:
                checked = [_missing_row(schema)]
            for index, row in enumerate(checked, 1):
                row["sheet"] = sheet
                row["year"] = year
                row["source_file"] = report["filename"]
                row["report_index"] = report["source_order"]
                for n in row["meta"]["pages"]:
                    if n < 1 or n > len(report["pages"]):
                        issues.append({"check_name": "INVALID CITED PAGE", "sheet": sheet,
                                       "notes": f"Row {index}: PAGE {n} not in source"})
                    if n not in row["meta"]["provided_pages"]:
                        issues.append({"check_name": "UNSENT CITED PAGE", "sheet": sheet,
                                       "notes": f"Row {index}: model cited PAGE {n} not provided"})
                row["meta"]["input_values"] = dict(row["values"])
                apply_conversion(row, sheet, schema, profile, issues)
                fill_system_columns(row, sheet, schema, profile, company_name,
                                    static_company_id, rating_cycle_year,
                                    source_code, index, report["pages"])
                compute_derived_columns(row, schema, sheet)
                row["meta"]["quote_verified"] = verify_evidence(row, report)
                if row["meta"]["quote_verified"] == "UNVERIFIED QUOTE":
                    issues.append({"check_name": "UNVERIFIED QUOTE", "sheet": sheet,
                                   "notes": f"Row {index}: evidence quote not in text page"})
            completed[sheet] = checked
            events.put((report["source_order"], sheet, len(checked),
                        f"{len(checked)} rows; {len(page_numbers)} source pages"))
        except Exception as exc:
            issues.append({"check_name": "SHEET EXTRACTION ERROR", "sheet": sheet,
                           "notes": f"{type(exc).__name__}: {exc}"})
            placeholder = _missing_row(schema)
            placeholder.update({"sheet": sheet, "year": year,
                                "source_file": report["filename"],
                                "report_index": report["source_order"]})
            fill_system_columns(placeholder, sheet, schema, profile, company_name,
                                static_company_id, rating_cycle_year,
                                source_code, 1, report["pages"])
            completed[sheet] = [placeholder]
            events.put((report["source_order"], sheet, 1, f"ERROR - {str(exc)[:110]}"))
    return {"year": year, "source_file": report["filename"],
            "report_index": report["source_order"], "rows": completed,
            "issues": issues, "stats": stats, "profile": profile}


def consolidate_reports(results: list[dict[str, Any]],
                        schemas: dict[str, dict[str, Any]]) -> dict[str, list[dict[str, Any]]]:
    """Concatenate each sheet by FY and file order; never deduplicate across FYs."""
    combined: dict[str, list[dict[str, Any]]] = {name: [] for name in schemas}
    for result in results:
        for sheet, rows in result["rows"].items():
            if sheet in combined and not schemas[sheet]["empty"]:
                combined[sheet].extend(rows)
    for sheet in combined:
        combined[sheet].sort(key=lambda r: (r["year"], r["report_index"]))
    return combined


def _test_result(name: str, year: int | str, file: str, expected: Any,
                 actual: Any, note: str = "") -> dict[str, Any]:
    """Return one standard validation record with the specified numeric tolerance."""
    a = _decimal_or_none(expected)
    b = _decimal_or_none(actual)
    if a is None or b is None:
        status, difference = "NOT TESTABLE", "N/A"
    else:
        difference = _numeric_excel(abs(a - b))
        limit = max(_DECIMAL("1"), _DECIMAL("0.005") * max(abs(a), abs(b)))
        status = "PASS" if abs(a - b) <= limit else "CHECK FAILED"
    return {"check_name": name, "financial_year": year, "source_file": file,
            "expected_value": expected, "actual_value": actual,
            "difference": difference, "status": status, "notes": note}


def _numeric_in_year(combined: dict[str, list[dict[str, Any]]],
                     year: int, sheet_name: str, names: list[str]) -> Any:
    """Find the first supported numeric field for one sheet/year."""
    sheet_key = next((k for k in combined if norm(k) == norm(sheet_name)), "")
    for row in combined.get(sheet_key, []):
        if row["year"] != year:
            continue
        by_normal = {norm(k): v for k, v in row["values"].items()}
        for name in names:
            value = by_normal.get(norm(name), "N/A")
            if _decimal_or_none(value) is not None:
                return value
    return "N/A"


def run_qa_checks(combined: dict[str, list[dict[str, Any]]],
                  schemas: dict[str, dict[str, Any]],
                  results: list[dict[str, Any]], missing_years: list[int]) -> list[dict[str, Any]]:
    """Audit financial balances, cross-table values, continuity and source errors."""
    checks: list[dict[str, Any]] = []
    years = sorted({r["year"] for group in combined.values() for r in group})
    files_by_year = {r["year"]: r["source_file"] for r in results}
    for year in years:
        fn = files_by_year.get(year, "N/A")
        assets = _numeric_in_year(combined, year, "FINANCIALS", ["total_assets"])
        liabilities = _numeric_in_year(combined, year, "FINANCIALS", ["total_liabilities"])
        equity = _numeric_in_year(combined, year, "FINANCIALS", ["total_equity"])
        l, e = _decimal_or_none(liabilities), _decimal_or_none(equity)
        checks.append(_test_result(
            "Total Assets = Total Equity + Total Liabilities", year, fn,
            assets, _numeric_excel(l + e) if l is not None and e is not None else "N/A"))
        profile_pat = _numeric_in_year(combined, year, "COMPANY_PROFILES",
                                      ["profit_after_tax"])
        fin_pat = _numeric_in_year(combined, year, "FINANCIALS",
                                  ["profit_after_tax"])
        checks.append(_test_result("Profit After Tax agrees across sheets", year,
                                   fn, profile_pat, fin_pat))
        profile_cash = _numeric_in_year(combined, year, "COMPANY_PROFILES",
                                       ["cash_and_equivalents", "cash_equivalents"])
        financial_cash = _numeric_in_year(combined, year, "FINANCIALS",
                                         ["cash_and_equivalents", "cash_equivalents"])
        checks.append(_test_result("Cash equivalents agrees across sheets", year,
                                   fn, profile_cash, financial_cash))
    for first, second in zip(years, years[1:]):
        if second != first + 1:
            continue
        # Look for available opening/closing balances in any requested sheet.
        close_cash = "N/A"
        open_cash = "N/A"
        for group in combined.values():
            for row in group:
                values = {norm(k): v for k, v in row["values"].items()}
                if row["year"] == first:
                    close_cash = next((values[k] for k in (
                        "closingcash", "cashandcashequivalentsatendofyear",
                        "closingcashbalance") if k in values), close_cash)
                if row["year"] == second:
                    open_cash = next((values[k] for k in (
                        "openingcash", "cashandcashequivalentsatbeginningofyear",
                        "openingcashbalance") if k in values), open_cash)
        check = _test_result("Cross-year closing/opening cash continuity", second,
                             files_by_year.get(second, "N/A"), close_cash, open_cash,
                             "A mismatch may reflect restatements; review, do not edit.")
        if check["status"] == "CHECK FAILED":
            check["status"] = "REVIEW"
        checks.append(check)
    for year in missing_years:
        checks.append({"check_name": "MISSING YEAR", "financial_year": year,
                       "source_file": "N/A", "expected_value": "Report",
                       "actual_value": "Missing", "difference": "N/A",
                       "status": "REVIEW", "notes": "Selected report FY gap."})
    for result in results:
        for issue in result["issues"]:
            checks.append({
                "check_name": issue.get("check_name", "QA ISSUE"),
                "financial_year": result["year"],
                "source_file": result["source_file"],
                "expected_value": "N/A", "actual_value": "N/A",
                "difference": "N/A", "status": "CHECK FAILED",
                "notes": f"{issue.get('sheet', 'REPORT')}: {issue.get('notes', '')}",
            })
    # Explicit workbook-level check of all rows/columns, without repairs.
    for sheet, rows in combined.items():
        expected = set(schemas[sheet]["columns"])
        for i, row in enumerate(rows, 1):
            actual = set(row["values"])
            if actual != expected:
                checks.append({"check_name": "SCHEMA MISMATCH",
                               "financial_year": row["year"],
                               "source_file": row["source_file"],
                               "expected_value": len(expected),
                               "actual_value": len(actual), "difference": "N/A",
                               "status": "CHECK FAILED",
                               "notes": f"{sheet} row {i}: column mismatch"})
    return checks


def _safe_excel_cell(cell: Any, value: Any) -> None:
    """Write text as data, not as executable spreadsheet formulas."""
    if value is None:
        value = "N/A"
    cell.value = value
    if isinstance(value, str) and value.startswith("="):
        cell.data_type = "s"
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        cell.number_format = "#,##0.##"


def _style_headers(ws: Any, columns: list[str]) -> None:
    """Apply restrained headers while keeping exactly one header row."""
    for i, name in enumerate(columns, 1):
        cell = ws.cell(row=1, column=i, value=name)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="17365D")
        cell.alignment = Alignment(wrap_text=True, vertical="center")
        ws.column_dimensions[get_column_letter(i)].width = min(36, max(13, len(name) + 2))
    ws.row_dimensions[1].height = 32
    ws.freeze_panes = "A2"
    if columns:
        ws.auto_filter.ref = f"A1:{get_column_letter(len(columns))}1"


def _build_workbook_object(schemas: dict[str, dict[str, Any]],
                           combined: dict[str, list[dict[str, Any]]]) -> Workbook:
    """Construct a strict workbook containing only ordered schema sheets/columns."""
    wb = Workbook()
    wb.remove(wb.active)
    for sheet, schema in schemas.items():
        ws = wb.create_sheet(title=sheet)
        _style_headers(ws, schema["columns"])
        if schema["empty"]:
            continue
        for n, row in enumerate(combined.get(sheet, []), 2):
            for c, header in enumerate(schema["columns"], 1):
                _safe_excel_cell(ws.cell(n, c), row["values"].get(header, "N/A"))
    return wb


def build_workbook(schemas: dict[str, dict[str, Any]],
                   combined: dict[str, list[dict[str, Any]]]) -> bytes:
    """Serialise the exact-sheet, exact-header consolidated main XLSX."""
    wb = _build_workbook_object(schemas, combined)
    stream = io.BytesIO()
    wb.save(stream)
    wb.close()
    return stream.getvalue()


def _row_key(row: dict[str, Any], sheet: str, sequential: int) -> str:
    """Find a deterministic row ID for audit joins even if spec omits one."""
    for col, val in row["values"].items():
        if norm(col) in ("recordkey", "objectid", "eventid", "relationshipid",
                         "attributeid", "programme code", "policycode") and val != "N/A":
            return str(val)
    return f"{sheet}-{row['year']}-{row['report_index']}-{sequential}"


def build_audit_workbook(schemas: dict[str, dict[str, Any]],
                         combined: dict[str, list[dict[str, Any]]],
                         validations: list[dict[str, Any]],
                         years: list[int]) -> bytes:
    """Export cell provenance, QA results and sheet/year counts to three tabs."""
    wb = Workbook()
    extract_ws = wb.active
    extract_ws.title = "Extraction_Audit"
    audit_headers = ["sheet", "row_key", "column", "raw_value_as_displayed",
                     "scale_found", "conversion_factor", "final_value",
                     "source_pages", "quote_verified"]
    _style_headers(extract_ws, audit_headers)
    next_audit_row = 2
    no_page = []
    for sheet, schema in schemas.items():
        for seq, row in enumerate(combined.get(sheet, []), 1):
            meta = row["meta"]
            rk = _row_key(row, sheet, seq)
            source = _source_from_row(row)
            if source == "Source on Website - No Report Page":
                no_page.append(f"{sheet}/{rk}")
            raw_snapshot = meta.get("input_values", {})
            for col in schema["columns"]:
                convert = meta["conversion"].get(col, {})
                raw = convert.get("raw", raw_snapshot.get(col, "N/A"))
                info = [sheet, rk, col, raw, convert.get("scale", "N/A"),
                        convert.get("factor", "N/A"), row["values"].get(col, "N/A"),
                        source, meta["quote_verified"]]
                for i, value in enumerate(info, 1):
                    _safe_excel_cell(extract_ws.cell(next_audit_row, i), value)
                next_audit_row += 1
    validate_ws = wb.create_sheet("Validation_Report")
    validate_headers = ["check_name", "financial_year", "source_file",
                        "expected_value", "actual_value", "difference", "status", "notes"]
    _style_headers(validate_ws, validate_headers)
    for j, result in enumerate(validations, 2):
        for i, name in enumerate(validate_headers, 1):
            _safe_excel_cell(validate_ws.cell(j, i), result.get(name, "N/A"))
    summary_ws = wb.create_sheet("Summary")
    _style_headers(summary_ws, ["sheet", "financial_year", "record_count", "empty_by_spec"])
    n = 2
    for sheet, schema in schemas.items():
        for year in years:
            count = sum(r["year"] == year for r in combined.get(sheet, []))
            values = [sheet, year, count, "YES" if schema["empty"] else "NO"]
            for i, value in enumerate(values, 1):
                _safe_excel_cell(summary_ws.cell(n, i), value)
            n += 1
    n += 1
    summary_ws.cell(n, 1, "Source page conclusion")
    summary_ws.cell(n, 2, "All sources are report pages" if not no_page else
                    "Rows without report-page evidence")
    if no_page:
        for missing in no_page:
            n += 1
            summary_ws.cell(n, 2, missing)
    n += 2
    summary_ws.cell(n, 1, "PAGE_NUMBERING")
    summary_ws.cell(n, 2, PAGE_NUMBERING)
    summary_ws.cell(n + 1, 1, "PDF page indexes are 1-based")
    summary_ws.cell(n + 2, 1, "Spreadsheet input pages refer to worksheet/row chunks")
    out = io.BytesIO()
    wb.save(out)
    wb.close()
    return out.getvalue()


def _source_from_row(row: dict[str, Any]) -> str:
    """Read a formatted source column if available, else fall back safely."""
    for name, value in row["values"].items():
        if norm(name) in ("sourcelocation", "sourcepage", "sourcepageorsection"):
            return str(value)
    return "Source on Website - No Report Page"


def append_to_workbook(existing_bytes: bytes,
                       schemas: dict[str, dict[str, Any]],
                       combined: dict[str, list[dict[str, Any]]]) -> bytes:
    """Append only wholly new sheets; preserve existing sheet data and names."""
    wb = load_workbook(io.BytesIO(existing_bytes))
    existing = {name.casefold() for name in wb.sheetnames}
    collisions = [name for name in schemas if name.casefold() in existing]
    if collisions:
        wb.close()
        raise ValueError("APPEND MODE: these sheets already exist and cannot "
                         f"be overwritten: {', '.join(collisions)}")
    for sheet, schema in schemas.items():
        ws = wb.create_sheet(sheet)
        _style_headers(ws, schema["columns"])
        if not schema["empty"]:
            for n, row in enumerate(combined.get(sheet, []), 2):
                for c, header in enumerate(schema["columns"], 1):
                    _safe_excel_cell(ws.cell(n, c), row["values"].get(header, "N/A"))
    stream = io.BytesIO()
    wb.save(stream)
    wb.close()
    return stream.getvalue()


def _safe_file_base(company_name: str, years: list[int]) -> str:
    """Generate a cross-platform filename with the exact selected FY range."""
    company = re.sub(r"[^A-Za-z0-9]+", "_", company_name).strip("_") or "Company"
    year_part = f"FY{years[0]}" if len(years) == 1 else f"FY{years[0]}-FY{years[-1]}"
    return f"{company}_{year_part}"


def _pricing_rates() -> tuple[float | None, float | None]:
    """Use explicitly provided $/million-token rates, never invented prices."""
    try:
        a = float(os.getenv("VISION_INPUT_USD_PER_M", ""))
        b = float(os.getenv("VISION_OUTPUT_USD_PER_M", ""))
        if a < 0 or b < 0:
            return None, None
        return a, b
    except ValueError:
        return None, None


def _usage_cost(stats: dict[str, int]) -> str:
    """Price returned token usage only when the operator supplied token rates."""
    inp, out = _pricing_rates()
    if inp is None or out is None:
        return "N/A (API does not supply billing cost; set token rates in .env)"
    price = (stats["input_tokens"] * inp + stats["output_tokens"] * out) / 1_000_000
    return f"${price:,.4f} (from actual SDK tokens × configured rates)"


def _preflight_estimate(reports: list[dict[str, Any]],
                        schemas: dict[str, dict[str, Any]]) -> dict[str, Any]:
    """Give a conspicuously rough token/cost estimate before the paid extraction."""
    active = sum(not s["empty"] for s in schemas.values())
    text_tokens = sum(min(35000, sum(len(p["index"]) for p in r["pages"])) // 4
                      for r in reports)
    projected_pages = min(MAX_PAGES_PER_SHEET, 6)
    estimated_input = len(reports) * active * (text_tokens // max(1, len(reports)) +
                         projected_pages * 1900 + 1300)
    estimated_output = len(reports) * active * 2200
    inp, out = _pricing_rates()
    cost = ((estimated_input * inp + estimated_output * out) / 1_000_000
            if inp is not None and out is not None else None)
    return {"input_tokens": estimated_input, "output_tokens": estimated_output,
            "cost": cost, "active_sheets": active}


def _format_schema_preview(schemas: dict[str, dict[str, Any]]) -> list[dict[str, Any]]:
    """Give users an exact, readable specification preview before confirmation."""
    return [{"Order": i, "Sheet": name, "Columns": len(spec["columns"]),
             "EMPTY (headers only)": "YES" if spec["empty"] else "NO"}
            for i, (name, spec) in enumerate(schemas.items(), 1)]


def _init_report(filename: str, data: bytes, source_order: int) -> dict[str, Any]:
    """Hash and profile source bytes without ever depending on a filename extension alone."""
    ext = Path(filename).suffix.lower()
    if ext not in (".pdf", ".xlsx"):
        raise ValueError(f"{filename}: upload only PDF or XLSX (not DOCX).")
    if ext == ".pdf" and data[:5] != b"%PDF-":
        raise ValueError(f"{filename}: PDF signature was not found.")
    try:
        pages = _read_pdf_input(data) if ext == ".pdf" else read_excel_input(data)
    except Exception as exc:
        raise ValueError(f"Cannot read {filename}: {exc}") from exc
    if not pages:
        raise ValueError(f"{filename}: contains no pages or worksheets.")
    return {"filename": filename, "bytes": data, "hash": _sha(data),
            "kind": ext.lstrip("."), "source_order": source_order,
            "pages": pages}


def _edited_profile_rows(profiles: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Flatten statement-specific scale/currency confirmations into one edit grid."""
    output = []
    for p in profiles:
        row = {"Include": p.get("include", True), "File": p["filename"],
               "ReportName": p["report_name"],
               "Company detected": p["company_detected"],
               "financial_year": p["financial_year"],
               "fiscal_year_end": p["fiscal_year_end"],
               "Company warning": p.get("company_warning", "")}
        for category in _SCALE_KEYS:
            row[f"{category} scale"] = p["scales"].get(category, "UNKNOWN")
            row[f"{category} currency"] = p["currencies"].get(category, "N/A")
        output.append(row)
    return output


def _confirmed_inputs(edited: list[dict[str, Any]],
                      originals: list[dict[str, Any]],
                      reports: list[dict[str, Any]]) -> tuple[list[dict[str, Any]],
                                                              list[dict[str, Any]],
                                                              list[str]]:
    """Validate selected years, per-statement scale and unedited file identities."""
    selected_p, selected_r, errors = [], [], []
    for i, row in enumerate(edited):
        if not row.get("Include", False):
            continue
        if i >= len(originals) or row.get("File") != originals[i]["filename"]:
            errors.append("File order changed. Re-profile the sources before extraction.")
            continue
        p = dict(originals[i])
        year = row.get("financial_year")
        try:
            year = int(year)
            if not 1900 <= year <= dt.date.today().year + 1:
                raise ValueError()
        except (TypeError, ValueError, OverflowError):
            errors.append(f"{p['filename']}: enter a valid current financial_year.")
            continue
        p["financial_year"] = year
        p["report_name"] = str(row.get("ReportName") or "").strip()
        if not p["report_name"]:
            errors.append(f"{p['filename']}: ReportName cannot be empty.")
        p["fiscal_year_end"] = str(row.get("fiscal_year_end") or "N/A").strip()
        p["company_detected"] = str(row.get("Company detected") or "N/A")
        p["scales"] = {}
        p["currencies"] = {}
        for category in _SCALE_KEYS:
            raw_scale = row.get(f"{category} scale")
            scale = str(raw_scale or "UNKNOWN").lower().strip()
            p["scales"][category] = scale
            p["currencies"][category] = str(
                row.get(f"{category} currency") or "N/A").strip().upper()
            if scale not in _SCALE_FACTORS:
                errors.append(f"{p['filename']}: confirm '{category} scale'; "
                              "select units only if the source actually uses units.")
        selected_p.append(p)
        selected_r.append(reports[i])
    if not selected_p:
        errors.append("Select at least one report.")
    if len(selected_p) > MAX_REPORTS:
        errors.append(f"Select no more than {MAX_REPORTS} reports.")
    counts = Counter(p["financial_year"] for p in selected_p)
    duplicates = sorted(y for y, n in counts.items() if n > 1)
    if duplicates:
        errors.append("Duplicate fiscal years: " + ", ".join(map(str, duplicates)) +
                      ". Untick all but one file per year.")
    return selected_p, selected_r, errors


def main() -> None:
    """Render the four-step Streamlit UI and coordinate bounded report workers."""
    import streamlit as st

    st.set_page_config(page_title="Annual Report Vision Agent", layout="wide")
    st.title("Annual Report Vision Agent")
    st.caption("One company · 1–6 annual reports · user-defined sheets · one consolidated XLSX + QA audit")
    with st.sidebar:
        st.header("Step 1 · Company and settings")
        company_name = st.text_input("company_name", placeholder="e.g. MTN Nigeria Plc")
        static_company_id = st.text_input("static_company_id", placeholder="e.g. IPMC/ESG/070")
        rating_cycle_year = st.text_input("rating_cycle_year (optional)")
        source_code = st.text_input("source_code (optional)")
        st.caption(f"Source-page numbering: {PAGE_NUMBERING} (editable at top of script)")
        try:
            provider, model, _ = _provider_settings()
            st.success(f"Configured provider: {provider} | model: {model}")
        except ValueError as exc:
            st.warning(str(exc))

    st.subheader("Step 1 · Sheet specification and sources")
    col_left, col_right = st.columns([2, 1])
    with col_left:
        spec_upload = st.file_uploader("Upload SHEET SPEC (.txt)", type=["txt"],
                                       key="schema_text_upload")
        spec_text = st.text_area("Or paste numbered SHEET SPEC lines",
                                 height=135, key="schema_text_area")
        if spec_upload is not None:
            spec_text = spec_upload.getvalue().decode("utf-8-sig")
            st.caption("The uploaded .txt takes precedence over the pasted spec.")
    with col_right:
        existing_workbook = st.file_uploader(
            "Optional existing output .xlsx (APPEND MODE)", type=["xlsx"],
            key="append_original")
        if existing_workbook:
            st.info("Append mode adds only new sheets. Existing sheet names, data and headers remain untouched.")
    uploads = st.file_uploader("Annual report files (.pdf/.xlsx, up to six)",
                               type=["pdf", "xlsx"],
                               accept_multiple_files=True, key="annual_reports")
    links_text = st.text_area("Direct PDF URLs (one per line; optional)",
                              height=80, key="report_direct_urls")
    links = [line.strip() for line in links_text.splitlines() if line.strip()]
    source_count = len(uploads) + len(links)
    if source_count > MAX_REPORTS:
        st.error(f"There are {source_count} sources; upload/link at most {MAX_REPORTS}.")
        st.stop()

    schemas = None
    if spec_text.strip():
        try:
            schemas = parse_sheet_spec(spec_text)
            with st.expander(f"Step 2 · Parsed SHEET SPEC: {len(schemas)} sheets", expanded=True):
                st.dataframe(_format_schema_preview(schemas), hide_index=True,
                             use_container_width=True)
                empty_count = sum(s["empty"] for s in schemas.values())
                st.caption(f"{empty_count} sheets are header-only and will not be sent to AI. "
                           "All headers retain exact case, spaces, slashes and order.")
        except ValueError as exc:
            st.error(str(exc))
    confirm_spec = st.checkbox("I have checked and confirm the parsed sheet names and exact columns",
                               key="confirm_schema")

    input_signature = hashlib.sha256(
        ("|".join(f"{f.name}:{_sha(f.getvalue())}" for f in uploads)
         + "|" + "|".join(links)).encode()).hexdigest()
    st.subheader("Step 3 · Profile and confirm reports")
    if st.button("Profile Reports", disabled=(source_count == 0), type="secondary"):
        try:
            _provider_settings()
            with st.spinner("Reading sources and detecting report metadata …"):
                source_reports: list[dict[str, Any]] = []
                for file in uploads:
                    source_reports.append(_init_report(
                        file.name, file.getvalue(), len(source_reports)))
                for url in links:
                    filename, data = fetch_pdf_from_url(url)
                    source_reports.append(_init_report(
                        filename, data, len(source_reports)))
                source_profiles = []
                profile_usages = []
                # Profiling can be concurrent, but each report remains independent.
                with concurrent.futures.ThreadPoolExecutor(max_workers=MAX_WORKERS) as pool:
                    futures = [pool.submit(detect_report_profile, r) for r in source_reports]
                    for file, future in zip(source_reports, futures):
                        try:
                            profile, usage = future.result()
                            source_profiles.append(profile)
                            profile_usages.append(usage)
                        except Exception as exc:
                            st.warning(f"{file['filename']}: profile incomplete: {exc}. "
                                       "Fill the confirmation table manually.")
                            source_profiles.append({
                                "filename": file["filename"],
                                "report_name": Path(file["filename"]).stem,
                                "company_detected": "N/A", "financial_year": None,
                                "fiscal_year_end": "N/A",
                                "scales": {k: "UNKNOWN" for k in _SCALE_KEYS},
                                "currencies": {k: "N/A" for k in _SCALE_KEYS},
                                "cached_profile": False})
                            profile_usages.append({"input_tokens": 0, "output_tokens": 0})
                checked_profiles = check_company_consistency(source_profiles, company_name)
                st.session_state["report_store"] = source_reports
                st.session_state["profile_store"] = checked_profiles
                st.session_state["profile_usages"] = profile_usages
                st.session_state["profile_usage_pending"] = True
                st.session_state["report_signature"] = input_signature
                # A new source set must not reuse an old report-edit grid.
                st.session_state.pop("report_editor", None)
                st.session_state.pop("extraction_result", None)
        except (ValueError, OSError) as exc:
            st.error(str(exc))

    valid_session = (st.session_state.get("report_signature") == input_signature
                     and "report_store" in st.session_state)
    if not valid_session:
        st.info("Upload the sources and click 'Profile Reports' to prepare the editable confirmation table.")
        return
    profiles = st.session_state["profile_store"]
    stored_reports = st.session_state["report_store"]
    st.caption("Review and edit the current FY, report title, fiscal year end, and "
               "every statement-specific scale/currency. UNKNOWN scales block extraction.")
    edited_grid = st.data_editor(
        _edited_profile_rows(profiles), key="report_editor", hide_index=True,
        disabled=["File", "Company warning"], use_container_width=True,
        column_config={
            "Include": st.column_config.CheckboxColumn("Include", default=True),
            "financial_year": st.column_config.NumberColumn("financial_year", format="%d"),
            **{f"{x} scale": st.column_config.SelectboxColumn(
                f"{x} scale", options=["UNKNOWN", "units", "thousands", "millions", "billions"])
               for x in _SCALE_KEYS},
        },
    )
    confirmed_profiles, confirmed_reports, confirm_errors = _confirmed_inputs(
        edited_grid, profiles, stored_reports)
    selected_years = sorted(p["financial_year"] for p in confirmed_profiles)
    missing_years = check_year_coverage(selected_years)
    if missing_years:
        st.warning("MISSING YEAR in selected reports: " + ", ".join(map(str, missing_years)) +
                   ". This will be logged in the audit workbook.")
    duplicates = [year for year, total in Counter(selected_years).items() if total > 1]
    if duplicates:
        st.error("Duplicate-year guard: " + ", ".join(map(str, duplicates)) +
                 ". Keep only one current-year report per FY.")
    if confirm_errors:
        st.warning("Confirmation required: " + " | ".join(confirm_errors[:8]))

    st.subheader("Step 4 · Extract All Reports")
    if schemas:
        estimate = _preflight_estimate(confirmed_reports, schemas)
        cost_label = (f"${estimate['cost']:,.2f} with your configured rates"
                      if estimate["cost"] is not None else
                      "N/A without operator-supplied pricing rates")
        st.caption(f"Pre-extraction estimate (rough upper-workload planning only): "
                   f"{estimate['input_tokens']:,} input / "
                   f"{estimate['output_tokens']:,} output tokens; USD cost {cost_label}. "
                   "Actual tokens will be read from provider usage, not from this estimate.")
    else:
        st.caption("Supply a valid spec to enable extraction.")
    append_collision = []
    if existing_workbook and schemas:
        try:
            old = load_workbook(io.BytesIO(existing_workbook.getvalue()), read_only=True)
            existing_names = {x.casefold() for x in old.sheetnames}
            append_collision = [x for x in schemas if x.casefold() in existing_names]
            old.close()
            if append_collision:
                st.error("APPEND MODE - existing sheet names conflict: " +
                         ", ".join(append_collision) + ". Supply only new sheets.")
        except Exception as exc:
            st.error(f"Cannot inspect existing output workbook: {exc}")
            append_collision = ["INVALID WORKBOOK"]
    ready = (bool(schemas) and confirm_spec and not confirm_errors and
             not append_collision and bool(company_name.strip()) and
             bool(static_company_id.strip()) and bool(confirmed_reports))
    if not company_name.strip() or not static_company_id.strip():
        st.info("Enter company_name and static_company_id in the sidebar.")
    output_fingerprint = _sha(json.dumps({
        "spec": spec_text, "edited": edited_grid, "company": company_name,
        "company_id": static_company_id, "rating_cycle_year": rating_cycle_year,
        "source_code": source_code, "source_signature": input_signature,
        "append_hash": (_sha(existing_workbook.getvalue()) if existing_workbook else ""),
    }, sort_keys=True, default=str).encode())
    if st.button("Extract All Reports", type="primary", disabled=not ready):
        SHEET_SCHEMAS.clear()
        SHEET_SCHEMAS.update(schemas)
        progress = st.progress(0, text="Preparing extraction workers")
        events: queue.Queue = queue.Queue()
        total_tasks = len(confirmed_reports) * len(schemas)
        completed_tasks = 0
        statuses = {}
        status_labels = {}
        for report, p in zip(confirmed_reports, confirmed_profiles):
            source_id = report["source_order"]
            status_labels[source_id] = f"{source_id + 1}. {p['filename']}"
            statuses[source_id] = st.status(
                f"{status_labels[source_id]}: starting", state="running", expanded=False)
        recorded_rows: dict[tuple[str, str], int] = {}
        live_counts = st.empty()
        results: list[dict[str, Any]] = []
        with concurrent.futures.ThreadPoolExecutor(max_workers=MAX_WORKERS) as pool:
            pending = {}
            for report, profile in zip(confirmed_reports, confirmed_profiles):
                future = pool.submit(
                    process_one_report, report, profile, schemas,
                    company_name.strip(), static_company_id.strip(),
                    rating_cycle_year.strip(), source_code.strip(), events)
                pending[future] = (report, profile)
            while pending:
                try:
                    file, sheet, rows_n, note = events.get(timeout=0.3)
                    completed_tasks += 1
                    recorded_rows[(file, sheet)] = rows_n
                    progress.progress(min(1.0, completed_tasks / max(1, total_tasks)),
                                      text=f"{completed_tasks}/{total_tasks} sheets processed")
                    if file in statuses:
                        statuses[file].update(label=f"{status_labels[file]}: {sheet} - {note}")
                    live_counts.caption(
                        f"Rows reported so far: {sum(recorded_rows.values()):,} "
                        f"across {completed_tasks} completed sheet tasks")
                except queue.Empty:
                    # No progress event is expected while an API call is in flight.
                    empty_cycle = True
                finished = [future for future in pending if future.done()]
                for future in finished:
                    report, profile = pending.pop(future)
                    try:
                        results.append(future.result())
                        statuses[report["source_order"]].update(
                            label=f"{status_labels[report['source_order']]}: complete", state="complete")
                    except Exception as exc:
                        statuses[report["source_order"]].update(
                            label=f"{status_labels[report['source_order']]}: failed - see audit", state="error")
                        results.append({
                            "year": profile["financial_year"],
                            "source_file": report["filename"],
                            "report_index": report["source_order"],
                            "rows": {sheet: [] for sheet in schemas},
                            "issues": [{"check_name": "REPORT FAILURE", "sheet": "REPORT",
                                        "notes": str(exc)}],
                            "stats": {"input_tokens": 0, "output_tokens": 0,
                                      "api_requests": 0, "cache_hits": 0},
                            "profile": profile,
                        })
        progress.progress(1.0, text="Validating and assembling workbooks")
        try:
            combined = consolidate_reports(results, schemas)
            validations = run_qa_checks(combined, schemas, results, missing_years)
            if existing_workbook:
                main_bytes = append_to_workbook(existing_workbook.getvalue(),
                                                schemas, combined)
                main_filename = existing_workbook.name
            else:
                main_bytes = build_workbook(schemas, combined)
                main_filename = _safe_file_base(company_name, selected_years) + ".xlsx"
            audit_bytes = build_audit_workbook(schemas, combined, validations, selected_years)
            audit_filename = Path(main_filename).stem + "_AUDIT.xlsx"
            statistics = []
            for result in sorted(results, key=lambda x: x["report_index"]):
                s = dict(result["stats"])
                if st.session_state.get("profile_usage_pending", False):
                    profile_usage = st.session_state["profile_usages"][result["report_index"]]
                    s["input_tokens"] += profile_usage.get("input_tokens", 0)
                    s["output_tokens"] += profile_usage.get("output_tokens", 0)
                    if profiles[result["report_index"]].get("cached_profile"):
                        s["cache_hits"] += 1
                    else:
                        s["api_requests"] += 1
                statistics.append({"report": result["source_file"], **s,
                                   "cost": _usage_cost(s)})
            st.session_state["extraction_result"] = {
                "main_bytes": main_bytes, "audit_bytes": audit_bytes,
                "main_filename": main_filename, "audit_filename": audit_filename,
                "validations": validations, "combined": combined,
                "schemas": schemas, "statistics": statistics,
                "years": selected_years, "fingerprint": output_fingerprint,
            }
            st.session_state["profile_usage_pending"] = False
        except Exception as exc:
            st.error(f"Workbook generation failed: {type(exc).__name__}: {exc}")
        finally:
            # Delete the temporary 120/300 DPI files, retaining only JSON cache.
            for report in confirmed_reports:
                shutil.rmtree(TEMP_IMAGE_DIRECTORY / report["hash"][:20],
                              ignore_errors=True)
    final = st.session_state.get("extraction_result")
    if not final:
        return
    if final.get("fingerprint") != output_fingerprint:
        st.info("Inputs changed after the previous extraction. Run extraction again "
                "before downloading new workbooks.")
        return
    failed = sum(v.get("status") == "CHECK FAILED" for v in final["validations"])
    if failed:
        st.warning(f"Extraction finished with {failed} QA findings. Review the "
                   "audit workbook before approving or importing any records.")
    else:
        st.success("Extraction run complete. Review validation results before approving data.")
    st.subheader("Actual model usage")
    st.dataframe(final["statistics"], hide_index=True, use_container_width=True)
    total_input = sum(s["input_tokens"] for s in final["statistics"])
    total_output = sum(s["output_tokens"] for s in final["statistics"])
    total_cache = sum(s["cache_hits"] for s in final["statistics"])
    st.caption(f"Run total: {total_input:,} input tokens, "
               f"{total_output:,} output tokens, {total_cache} cache hits. "
               f"Cost: {_usage_cost({'input_tokens': total_input, 'output_tokens': total_output})}")
    st.subheader("Validation Report")
    st.dataframe(final["validations"][:2000], hide_index=True,
                 use_container_width=True)
    st.subheader("Per-sheet preview")
    sheet_choice = st.selectbox("Choose sheet", list(final["schemas"]),
                                key="preview_sheet_choice")
    spec = final["schemas"][sheet_choice]
    rows = final["combined"].get(sheet_choice, [])
    st.caption(f"{sheet_choice}: {len(rows)} rows; {len(spec['columns'])} exact columns"
               + (" (EMPTY: header-only)" if spec["empty"] else ""))
    if rows:
        st.dataframe([r["values"] for r in rows[:100]], hide_index=True,
                     use_container_width=True)
    else:
        st.info("This sheet contains only column headers.")
    st.subheader("Summary")
    summary = [
        {"sheet": name, "financial_year": year,
         "records": sum(r["year"] == year for r in final["combined"].get(name, [])),
         "EMPTY by spec": spec["empty"]}
        for name, spec in final["schemas"].items() for year in final["years"]
    ]
    st.dataframe(summary, hide_index=True, use_container_width=True)
    left, right = st.columns(2)
    with left:
        st.download_button("Download MAIN workbook", final["main_bytes"],
                           file_name=final["main_filename"],
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           type="primary")
    with right:
        st.download_button("Download AUDIT workbook", final["audit_bytes"],
                           file_name=final["audit_filename"],
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ============================= HOW TO RUN =============================
# 1. Create a .env file in the directory alongside this script:
#       VISION_PROVIDER=openai             # or gemini or anthropic
#       VISION_MODEL=<vision-capable-model-you-choose>
#       OPENAI_API_KEY=<your-key>           # or GEMINI_API_KEY / ANTHROPIC_API_KEY
#    Optional operator-supplied price rates (USD per 1 million tokens):
#       VISION_INPUT_USD_PER_M=<current-input-rate>
#       VISION_OUTPUT_USD_PER_M=<current-output-rate>
#    No actual API key, model identifier, or price is included in this file.
# 2. Install dependencies: pip install -r requirements.txt
# 3. Start: streamlit run Annual_Report_Vision_Agent.py
# 4. Upload SCHEMA_SPEC.txt (or paste any conforming sheet specification),
#    add 1–6 reports, profile/confirm the report table, then Extract All Reports.
# 5. Review the audit workbook and manually validate flagged values.

if __name__ == "__main__":
    main()
