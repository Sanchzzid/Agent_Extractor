
from __future__ import annotations

import json
from typing import Any

# --------------------------------------------------------------------------
# Shared rules (the same rules the human extractors were given)
# --------------------------------------------------------------------------
SYSTEM_JSON = (
    "You are a meticulous financial-report data extraction engine. "
    "You ALWAYS reply with exactly ONE valid JSON object: no markdown fences, "
    "no comments, no text before or after the JSON."
)

EXTRACTION_RULES = """RULES (follow strictly):
1. Zero hallucination: use ONLY what is printed in the supplied pages. Never use outside knowledge.
2. If a value is not printed, use "N/A". Never guess, estimate, infer or calculate a value that is not shown.
3. CURRENT financial year only ({year}). Ignore prior-year / comparative columns, restated prior periods and forecasts.
4. Keep names and wording exactly as printed (people, companies, policies, titles).
5. Numbers: plain JSON numbers exactly as printed, with no commas, currency symbols or % signs.
   Accounting brackets mean negative: (1,245) -> -1245. Do NOT convert scale yourself.
6. Dates: YYYY-MM-DD. If only a year is printed, use "N/A" for the full-date field.
7. One row per distinct record (each director, subsidiary, shareholder, policy, emission scope, programme...).
   Sheets that describe the company as a whole normally have ONE row.
8. Replace bullet characters with "-". No line breaks inside values.
9. Only extract records that genuinely belong to THIS sheet's subject (its name and columns).
   e.g. a POLICIES sheet takes only named policies, not emissions figures or directors.
   Do not force unrelated data into the sheet: if nothing matches, return no rows."""


# --------------------------------------------------------------------------
# 1. Report profiler: which company / which year is this PDF?
# --------------------------------------------------------------------------
def profile_prompt(filename: str, front_text: str, scale_snippets: str) -> str:
    return f"""Identify this annual report. Reply with JSON:
{{"company_name": string or "N/A",
  "financial_year": integer (the CURRENT reporting year of this report, e.g. 2023) or null,
  "fiscal_year_end": "YYYY-MM-DD" or "N/A",
  "currency": ISO code like "NGN" or "N/A",
  "amounts_scale": "units" | "thousands" | "millions" | "billions" | "UNKNOWN",
  "report_title": string}}

Use the year the report COVERS (e.g. "Annual Report 2023" / "year ended 31 December 2023" -> 2023),
not the publication date and not comparative years. "amounts_scale" is the unit used in the primary
financial statements (e.g. "N'000" or "in thousands" -> "thousands"); UNKNOWN if not stated.

FILENAME: {filename}

FRONT PAGES:
{front_text}

UNIT / CURRENCY EVIDENCE FROM THE FINANCIAL STATEMENTS:
{scale_snippets or "(none found)"}"""


# --------------------------------------------------------------------------
# 1b. Scanned-page captioner: describe image-only pages so they can be found
# --------------------------------------------------------------------------
def caption_pages_prompt(page_numbers: list[int]) -> str:
    labels = ", ".join(f"image {i + 1} = PAGE {n}" for i, n in enumerate(page_numbers))
    return f"""These images are scanned pages of an annual report ({labels}).
For each page give its heading and the topics / table names visible (max 20 words). Do not copy numbers.
Reply with JSON: {{"pages": [{{"page": page number, "topic": "..."}}]}}
Use "BLANK" for an empty or decorative page."""


# --------------------------------------------------------------------------
# 2. Page locator: which pages hold data for this sheet?
# --------------------------------------------------------------------------
def locate_pages_prompt(sheet: str, columns: list[str], year: int | str,
                        page_index: str, max_pages: int) -> str:
    return f"""You are locating evidence pages in an annual report (financial year {year}).

TARGET SHEET: {sheet}
FIELDS NEEDED: {", ".join(columns)}

Below is a PAGE INDEX: one line per page with the page number and the first words on that page.
Pick the pages most likely to contain CURRENT-YEAR values for these fields (statements, notes,
governance sections, sustainability sections, tables). Prefer specific pages over generic overviews.
Return at most {max_pages} pages, best first. If the report clearly has no such data, return [].

Reply with JSON: {{"pages": [page numbers as integers]}}

PAGE INDEX:
{page_index}"""


# --------------------------------------------------------------------------
# 3. Extractor: pull rows for one sheet from a batch of pages
# --------------------------------------------------------------------------
def extraction_prompt(sheet: str, column_legend: dict[str, str], company: str,
                      year: int | str, report_title: str) -> str:
    legend = "\n".join(f'  "{key}": {name}' for key, name in column_legend.items())
    example_row = {key: "..." for key in list(column_legend)[:3]}
    example_row.update({"_pages": [12], "_scale": "millions", "_evidence": "short quote"})
    return f"""Extract data for the sheet "{sheet}" from the annual report pages below.

COMPANY: {company}
REPORT: {report_title}
FINANCIAL YEAR: {year}

{EXTRACTION_RULES.format(year=year)}

COLUMNS (use these short keys exactly; the text after the colon is the column name):
{legend}

Each row object must contain EVERY key above, plus:
  "_pages": list of the PAGE numbers the row's values came from,
  "_scale": "units" | "thousands" | "millions" | "billions" | "UNKNOWN" — the unit printed for money amounts in this row,
  "_evidence": a verbatim quote (max 25 words) from the page that supports the row.

Reply with JSON: {{"rows": [ {json.dumps(example_row)} , ... ]}}
If the pages contain no data for this sheet, reply {{"rows": []}}."""


# --------------------------------------------------------------------------
# 4. Verifier: a second LLM pass that checks the extracted JSON
# --------------------------------------------------------------------------
def verification_prompt(sheet: str, column_legend: dict[str, str], year: int | str,
                        rows: list[dict[str, Any]]) -> str:
    legend = "\n".join(f'  "{key}": {name}' for key, name in column_legend.items())
    rows_json = json.dumps([{"row": i, **r} for i, r in enumerate(rows)], ensure_ascii=False)
    return f"""You are an independent auditor. Another assistant extracted the rows below for the sheet
"{sheet}" (financial year {year}) from the annual report pages that follow. Check every value
against the pages.

For each value decide:
  - correct: it is printed in the pages for financial year {year} -> do NOT report it.
    Numbers were deliberately copied RAW, exactly as printed, WITHOUT applying "thousands/millions"
    units: that is correct, never "fix" it by multiplying. Ignore commas, brackets-as-negative,
    capitalisation and spacing differences.
  - "WRONG": the pages show a DIFFERENT value for this field and year (e.g. a prior-year figure, a
    typo, a wrong person/name, a wrong sign). Give the correct value exactly as printed.
  - "NOT_FOUND": the value is not supported by these pages at all (hallucinated or guessed).
Values that are "N/A" only need reporting if the pages clearly DO show the value (verdict "WRONG").

COLUMN KEYS:
{legend}

EXTRACTED ROWS:
{rows_json}

Reply with JSON listing ONLY the problems:
{{"issues": [{{"row": 0, "key": "c5", "verdict": "WRONG" | "NOT_FOUND",
              "correct_value": value as printed or "N/A", "reason": "max 20 words"}}]}}
If everything is correct reply {{"issues": []}}."""


# --------------------------------------------------------------------------
# 5. JSON repair: sent back to the model when its reply could not be parsed
# --------------------------------------------------------------------------
def json_repair_prompt(error: str) -> str:
    return (f"Your previous reply could not be used: {error}\n"
            "Reply again with ONLY the corrected, complete, valid JSON object in the "
            "required format. No markdown, no explanations.")
