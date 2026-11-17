"""Building blocks used by the graph nodes: locate pages, extract, verify, clean rows.

Kept separate from graph.py so each step can be unit-tested without LangGraph.
"""
from __future__ import annotations

import datetime as dt
import decimal
import re
from dataclasses import dataclass, field
from typing import Any

from . import config, llm, prompts
from .pdf_reader import Page, render_png
from .schema import Sheet, norm

SCALE_FACTORS = {"units": 1, "thousands": 1_000, "millions": 1_000_000, "billions": 1_000_000_000}
NA = "N/A"
_STOP = {"name", "type", "date", "code", "status", "value", "text", "numeric", "boolean", "source",
         "year", "period", "notes", "with", "from", "rating", "cycle", "static", "company", "record",
         "financial", "description", "location", "page", "unit", "category", "flag", "count"}


# ---------------------------------------------------------------------------
# Page selection
# ---------------------------------------------------------------------------
def sheet_keywords(sheet: Sheet) -> set[str]:
    words = set()
    for text in [sheet.name] + [c.name for c in sheet.model_columns]:
        for w in re.split(r"[^a-zA-Z]+", text.lower()):
            if len(w) >= 4 and w not in _STOP:
                words.add(w)
    return words


def keyword_scores(sheet: Sheet, pages: list[Page]) -> dict[int, int]:
    words = sheet_keywords(sheet)
    scores = {}
    for p in pages:
        low = (p.text + " " + p.caption).lower()
        scores[p.number] = sum(low.count(w) for w in words)
    return scores


def build_page_index(sheet: Sheet, pages: list[Page]) -> str:
    lines = [p.index_line() for p in pages]
    if llm.estimate_tokens("\n".join(lines)) <= config.INDEX_TOKEN_BUDGET:
        return "\n".join(lines)
    # Too long for the context window: keep the pages with the most keyword hits.
    scores = keyword_scores(sheet, pages)
    ranked = sorted(pages, key=lambda p: (-scores[p.number], p.number))
    keep: list[Page] = []
    used = 0
    for p in ranked:
        cost = llm.estimate_tokens(p.index_line()) + 1
        if used + cost > config.INDEX_TOKEN_BUDGET:
            break
        keep.append(p)
        used += cost
    return "\n".join(p.index_line() for p in sorted(keep, key=lambda p: p.number))


def caption_scanned_pages(pages: list[Page], pdf_bytes: bytes, usage: dict, log: list) -> int:
    """Give image-only pages a short topic (vision) so the page locator can find them."""
    targets = [p for p in pages if p.scanned]
    if not targets or not config.USE_VISION:
        return 0

    def validate(data: dict) -> dict:
        if not isinstance(data.get("pages"), list):
            raise ValueError('reply must be {"pages": [...]}')
        return data

    done = 0
    for i in range(0, len(targets), 4):
        group = targets[i:i + 4]
        images = [render_png(pdf_bytes, p.number, 80) for p in group]
        try:
            data = llm.call_json(prompts.caption_pages_prompt([p.number for p in group]),
                                 validate, images, usage, max_output=800)
        except llm.LLMError as exc:
            log.append({"sheet": "-", "check": "SCANNED PAGE CAPTION FAILED",
                        "detail": f"pages {[p.number for p in group]}: {exc}"})
            continue
        by_num = {p.number: p for p in group}
        for item in data["pages"]:
            try:
                p = by_num[int(item.get("page"))]
            except (KeyError, TypeError, ValueError, AttributeError):
                continue
            p.caption = str(item.get("topic", ""))[:120]
            done += 1
    return done


def locate_pages(sheet: Sheet, pages: list[Page], year: Any, usage: dict) -> tuple[list[int], str]:
    """Return (page numbers, note). Falls back to keyword ranking if the LLM call fails."""
    if len(pages) <= 4:
        return [p.number for p in pages], "small document: all pages"
    index = build_page_index(sheet, pages)
    prompt = prompts.locate_pages_prompt(sheet.name, [c.name for c in sheet.model_columns], year,
                                         index, config.MAX_PAGES_PER_SHEET)

    def validate(data: dict) -> dict:
        raw = data.get("pages")
        if not isinstance(raw, list):
            raise ValueError('"pages" must be a list of integers')
        nums: list[int] = []
        for v in raw:
            try:
                n = int(str(v).lstrip("Pp"))
            except ValueError:
                continue
            if 1 <= n <= len(pages) and n not in nums:
                nums.append(n)
        return {"pages": nums[:config.MAX_PAGES_PER_SHEET]}

    try:
        return llm.call_json(prompt, validate, usage=usage, max_output=400)["pages"], "llm"
    except llm.LLMError as exc:
        scores = keyword_scores(sheet, pages)
        best = [n for n, s in sorted(scores.items(), key=lambda kv: -kv[1]) if s > 0]
        return best[:config.MAX_PAGES_PER_SHEET], f"keyword fallback ({exc})"


# ---------------------------------------------------------------------------
# Batching pages to fit the context window
# ---------------------------------------------------------------------------
@dataclass
class Batch:
    pages: list[int]
    text: str                      # concatenated page text for the prompt
    images: list[bytes] = field(default_factory=list)
    has_images: bool = False

    def numbers_in_text(self) -> set[decimal.Decimal]:
        return numbers_in(self.text)


def make_batches(page_numbers: list[int], pages: list[Page], pdf_bytes: bytes,
                 budget: int | None = None) -> list[Batch]:
    budget = budget or config.PAGE_TOKEN_BUDGET
    batches: list[Batch] = []
    cur_pages: list[int] = []
    cost = 0
    for n in sorted(set(page_numbers)):
        p = pages[n - 1]
        use_image = config.USE_VISION and p.scanned
        page_cost = config.IMAGE_TOKEN_COST if use_image else llm.estimate_tokens(p.text) + 10
        page_cost = min(page_cost, budget)
        if cur_pages and cost + page_cost > budget:
            batches.append(_build_batch(cur_pages, pages, pdf_bytes, budget))
            cur_pages, cost = [], 0
        cur_pages.append(n)
        cost += page_cost
    if cur_pages:
        batches.append(_build_batch(cur_pages, pages, pdf_bytes, budget))
    return batches


def _build_batch(nums: list[int], pages: list[Page], pdf_bytes: bytes, budget: int) -> Batch:
    parts, images = [], []
    char_cap = int(budget * 2.8)  # a single huge page is truncated rather than overflowing
    for n in nums:
        p = pages[n - 1]
        if config.USE_VISION and p.scanned:
            images.append(render_png(pdf_bytes, n, config.IMAGE_DPI))
            parts.append(f"=== PAGE {n} === (scanned page: see image {len(images)})")
        else:
            parts.append(f"=== PAGE {n} ===\n{p.text[:char_cap]}")
    return Batch(nums, "\n\n".join(parts), images, bool(images))


# ---------------------------------------------------------------------------
# Extraction
# ---------------------------------------------------------------------------
def _rows_validator(sheet: Sheet):
    keys = [c.key for c in sheet.model_columns]

    def validate(data: dict) -> dict:
        rows = data.get("rows")
        if rows is None and len(data) == 1 and isinstance(next(iter(data.values())), list):
            rows = next(iter(data.values()))  # model used a different top-level key
        if not isinstance(rows, list):
            raise ValueError('reply must be {"rows": [ ... ]}')
        clean = []
        for r in rows:
            if not isinstance(r, dict):
                raise ValueError("every item in rows must be an object")
            fixed = _map_keys(r, sheet)
            if not any(k in fixed for k in keys):
                continue
            clean.append(fixed)
        return {"rows": clean}

    return validate


def _map_keys(row: dict, sheet: Sheet) -> dict:
    """Accept c-keys, or the real column names if the model used those instead."""
    by_name = {norm(c.name): c.key for c in sheet.model_columns}
    out = {}
    for k, v in row.items():
        if k.startswith("_") or re.fullmatch(r"c\d+", k):
            out[k] = v
        elif norm(k) in by_name:
            out[by_name[norm(k)]] = v
    return out


def extract_batch(sheet: Sheet, batch: Batch, profile: dict, company: str,
                  usage: dict) -> list[dict]:
    prompt = (prompts.extraction_prompt(sheet.name, sheet.legend, company,
                                        profile.get("financial_year") or "unknown",
                                        profile.get("report_title", ""))
              + "\n\nREPORT PAGES:\n" + batch.text)
    rows = llm.call_json(prompt, _rows_validator(sheet), batch.images, usage)["rows"]
    for r in rows:
        r["_batch"] = batch.pages
    return rows


def extract_sheet(sheet: Sheet, page_numbers: list[int], pages: list[Page], pdf_bytes: bytes,
                  profile: dict, company: str, usage: dict, log: list) -> tuple[list[dict], list[Batch]]:
    """Extract all batches; a batch that is too large is split in half and retried."""
    rows: list[dict] = []
    done: list[Batch] = []
    queue = make_batches(page_numbers, pages, pdf_bytes)
    while queue:
        batch = queue.pop(0)
        try:
            rows.extend(extract_batch(sheet, batch, profile, company, usage))
            done.append(batch)
        except llm.ContextTooLarge as exc:
            if len(batch.pages) > 1:
                half = len(batch.pages) // 2
                queue[:0] = [_build_batch(batch.pages[:half], pages, pdf_bytes, config.PAGE_TOKEN_BUDGET),
                             _build_batch(batch.pages[half:], pages, pdf_bytes, config.PAGE_TOKEN_BUDGET)]
            else:
                smaller = _build_batch(batch.pages, pages, pdf_bytes, config.PAGE_TOKEN_BUDGET // 2)
                if smaller.text != batch.text:
                    queue.insert(0, smaller)
                else:
                    log.append({"sheet": sheet.name, "check": "EXTRACTION FAILED",
                                "detail": f"pages {batch.pages}: {exc}"})
        except llm.LLMError as exc:
            log.append({"sheet": sheet.name, "check": "EXTRACTION FAILED",
                        "detail": f"pages {batch.pages}: {exc}"})
    return rows, done


# ---------------------------------------------------------------------------
# Verification (second LLM pass) + deterministic guard against a wrong verifier
# ---------------------------------------------------------------------------
_NUM_RE = re.compile(r"\(?-?\d[\d,]*(?:\.\d+)?\)?")


def to_decimal(value: Any) -> decimal.Decimal | None:
    if isinstance(value, bool) or value is None:
        return None
    s = str(value).strip().replace("−", "-")
    neg = s.startswith("(") and s.endswith(")")
    s = s.strip("()").replace(",", "").replace(" ", "").rstrip("%")
    for sym in ("₦", "$", "€", "£", "N", "NGN", "USD"):
        if s.startswith(sym):
            s = s[len(sym):]
    if not re.fullmatch(r"-?\d+(?:\.\d+)?", s):
        return None
    try:
        d = decimal.Decimal(s)
    except decimal.InvalidOperation:
        return None
    return -d if neg else d


def numbers_in(text: str) -> set[decimal.Decimal]:
    found = set()
    for m in _NUM_RE.findall(text):
        d = to_decimal(m)
        if d is not None:
            found.add(abs(d).normalize())
    return found


def same_value(a: Any, b: Any) -> bool:
    da, db = to_decimal(a), to_decimal(b)
    if da is not None and db is not None:
        return da == db
    norm_text = lambda v: re.sub(r"[\s.,;:]+", " ", str(v)).strip().lower()  # noqa: E731
    return norm_text(a) == norm_text(b)


def scale_only_difference(a: Any, b: Any) -> bool:
    da, db = to_decimal(a), to_decimal(b)
    if not da or not db:
        return False
    ratio = abs(db / da)
    return any(ratio == f or ratio == decimal.Decimal(1) / f for f in (1000, 1_000_000, 1_000_000_000))


def value_in_batch(value: Any, batch: Batch) -> bool | None:
    """True/False if the value can be checked in the page text, None if not checkable."""
    if value in (None, "", NA):
        return None
    d = to_decimal(value)
    if d is not None:
        if batch.has_images:
            return None
        return abs(d).normalize() in batch.numbers_in_text()
    text = re.sub(r"\s+", " ", str(value)).strip().lower()
    if len(text) < 3:
        return None
    page_text = re.sub(r"\s+", " ", batch.text).lower()
    if text in page_text:
        return True
    return None if batch.has_images or len(text) > 60 else False


def _issues_validator(n_rows: int, keys: set[str]):
    def validate(data: dict) -> dict:
        issues = data.get("issues")
        if not isinstance(issues, list):
            raise ValueError('reply must be {"issues": [ ... ]}')
        clean = []
        for it in issues:
            if not isinstance(it, dict):
                continue
            try:
                row = int(it.get("row"))
            except (TypeError, ValueError):
                continue
            key = str(it.get("key", ""))
            verdict = str(it.get("verdict", "")).upper().replace(" ", "_")
            if 0 <= row < n_rows and key in keys and verdict in {"WRONG", "NOT_FOUND"}:
                clean.append({"row": row, "key": key, "verdict": verdict,
                              "correct_value": it.get("correct_value", NA),
                              "reason": str(it.get("reason", ""))[:200]})
        return {"issues": clean}
    return validate


def verify_batch(sheet: Sheet, batch: Batch, rows: list[dict], year: Any, usage: dict,
                 log: list) -> None:
    """Second LLM pass. Mutates rows in place and sets row['_verification']."""
    keys = [c.key for c in sheet.model_columns]
    slim = [{k: r.get(k, NA) for k in keys} for r in rows]
    prompt = (prompts.verification_prompt(sheet.name, sheet.legend, year, slim)
              + "\n\nREPORT PAGES:\n" + batch.text)
    try:
        issues = llm.call_json(prompt, _issues_validator(len(rows), set(keys)),
                               batch.images, usage, max_output=2500)["issues"]
    except llm.LLMError as exc:
        for r in rows:
            r["_verification"] = "UNVERIFIED"
        log.append({"sheet": sheet.name, "check": "VERIFICATION FAILED",
                    "detail": f"pages {batch.pages}: {exc}"})
        return

    for r in rows:
        r["_verification"] = "VERIFIED"
    names = sheet.legend
    for it in issues:
        row, key = rows[it["row"]], it["key"]
        old, new = row.get(key, NA), it["correct_value"]
        if new in (None, ""):
            new = NA
        entry = {"sheet": sheet.name, "pages": batch.pages, "row": it["row"] + 1,
                 "column": names[key], "extracted": old, "verifier": new, "reason": it["reason"]}
        if old == NA and (it["verdict"] == "NOT_FOUND" or new == NA):
            continue                  # nothing to change
        if it["verdict"] == "NOT_FOUND":
            if value_in_batch(old, batch) is True:      # it IS on the page: keep it
                entry.update(check="VERIFIER OVERRULED", action="kept (value is printed on page)")
            else:
                row[key] = NA
                row["_verification"] = "CORRECTED"
                entry.update(check="NOT IN DOCUMENT", action="set to N/A")
        else:  # WRONG
            if same_value(old, new) or scale_only_difference(old, new):
                continue          # only formatting or a thousands/millions unit difference
            found = value_in_batch(new, batch)
            if new == NA:
                row[key] = NA
                row["_verification"] = "CORRECTED"
                entry.update(check="WRONG VALUE", action="set to N/A")
            elif found is False:
                row["_verification"] = "NEEDS_REVIEW"
                entry.update(check="DISPUTED", action="kept original (verifier value not on page)")
            else:
                row[key] = new
                row["_verification"] = "CORRECTED"
                entry.update(check="WRONG VALUE", action="corrected")
        log.append(entry)


# ---------------------------------------------------------------------------
# Row clean-up, merging and system columns
# ---------------------------------------------------------------------------
def norm_company(name: str) -> str:
    words = re.findall(r"[a-z0-9]+", str(name).lower())
    legal = {"plc", "ltd", "limited", "inc", "group", "holdings", "holding", "company", "co", "the", "na", "n"}
    return " ".join(w for w in words if w not in legal)


def clean_value(value: Any) -> Any:
    if value is None:
        return NA
    if isinstance(value, (list, dict)):
        value = "; ".join(map(str, value)) if isinstance(value, list) else str(value)
    if isinstance(value, str):
        v = re.sub(r"\s+", " ", value.replace("•", "-")).strip()
        return NA if v.lower() in {"", "null", "none", "n/a", "na", "unknown", "not available",
                                   "not stated", "not disclosed", "-"} else v
    return value


def merge_batch_rows(rows: list[dict], keys: list[str]) -> list[dict]:
    """Rows from DIFFERENT page batches that never conflict describe the same record
    (e.g. revenue on one page, total assets on another) and are merged."""
    merged: list[dict] = []
    for r in rows:
        target = None
        for m in merged:
            if r["_batch"] in m["_batches"]:
                continue
            conflict = any(
                r.get(k, NA) != NA and m.get(k, NA) != NA
                and str(r.get(k)).strip().lower() != str(m.get(k)).strip().lower()
                for k in keys)
            if not conflict:
                target = m
                break
        if target is None:
            r = dict(r)
            r["_batches"] = [r["_batch"]]
            r["_pages"] = list(r.get("_pages") or [])
            merged.append(r)
            continue
        for k in keys:
            if target.get(k, NA) == NA and r.get(k, NA) != NA:
                target[k] = r[k]
        target["_batches"].append(r["_batch"])
        target["_pages"] = sorted(set(target["_pages"]) | set(r.get("_pages") or []))
        rank = {"NEEDS_REVIEW": 3, "UNVERIFIED": 2, "CORRECTED": 1, "VERIFIED": 0}
        if rank.get(r.get("_verification"), 2) > rank.get(target.get("_verification"), 2):
            target["_verification"] = r.get("_verification")
        if target.get("_scale", "UNKNOWN") == "UNKNOWN":
            target["_scale"] = r.get("_scale", "UNKNOWN")
    return merged


def finalize_rows(sheet: Sheet, rows: list[dict], profile: dict, job: dict,
                  pages: list[Page], log: list) -> list[dict[str, Any]]:
    keys = [c.key for c in sheet.model_columns]
    for r in rows:
        for k in keys:
            r[k] = clean_value(r.get(k))
        pages_cited = []
        for n in r.get("_pages") or []:
            try:
                n = int(n)
            except (TypeError, ValueError):
                continue
            if n in r["_batch"]:
                pages_cited.append(n)
        r["_pages"] = pages_cited or list(r["_batch"])
    rows = merge_batch_rows(rows, keys)
    rows = [r for r in rows if any(r.get(k, NA) != NA for k in keys)]

    year = profile.get("financial_year") or NA
    company_name = job.get("company_name") or profile.get("company_name") or NA
    company_id = job.get("company_id") or NA
    title = profile.get("report_title") or profile.get("filename", "report")
    keep_raw_scale = any(norm(c.name) in {"scale", "unitscale", "amountscale"} for c in sheet.columns)
    today = dt.date.today().isoformat()
    out = []
    for i, r in enumerate(rows, 1):
        values: dict[str, Any] = {}
        for col in sheet.columns:
            if col.kind == "MODEL":
                values[col.name] = _typed_value(col, r.get(col.key, NA), r, sheet, keep_raw_scale, log)
        status = r.get("_verification", "UNVERIFIED")
        labels = [pages[n - 1].label for n in sorted(set(r["_pages"])) if 1 <= n <= len(pages)]
        system = {
            "company_id": company_id,
            "company_name": company_name,
            "financial_year": year,
            "rating_cycle_year": job.get("rating_cycle_year") or NA,
            "source_code": job.get("source_code") or NA,
            "retrieval_date": today,
            "source_location": f"{title} (pg {', '.join(labels)})" if labels else NA,
            "data_status": "EXTRACTED",
            "verification_status": status,
            "publication_status": "DRAFT",
            "verified_at": today if status in {"VERIFIED", "CORRECTED"} else NA,
            "row_id": f"{company_id if company_id != NA else re.sub(r'[^A-Za-z0-9]+', '', str(company_name))[:12]}"
                      f"-{sheet.name}-{year}-{i}",
        }
        for col in sheet.columns:
            if col.kind == "SYSTEM":
                values[col.name] = system.get(col.alias, NA)
        out.append({"values": values, "pages": r["_pages"], "status": status,
                    "evidence": str(r.get("_evidence", ""))[:300]})
    return out


def _typed_value(col, value, row, sheet, keep_raw_scale, log):
    if value == NA or not col.numeric:
        return value
    d = to_decimal(value)
    if d is None:
        log.append({"sheet": sheet.name, "check": "NOT A NUMBER", "column": col.name,
                    "extracted": value, "action": "kept as text"})
        return value
    if col.money and config.CONVERT_SCALE and not keep_raw_scale:
        scale = str(row.get("_scale", "UNKNOWN")).lower()
        if scale in SCALE_FACTORS and scale != "units":
            d = d * SCALE_FACTORS[scale]
            log.append({"sheet": sheet.name, "check": "SCALE APPLIED", "column": col.name,
                        "extracted": value, "action": f"x {SCALE_FACTORS[scale]:,} ({scale})"})
    return int(d) if d == d.to_integral_value() else float(d)
