"""Sheet spec parsing and column classification.

Spec format (one sheet per line, blank lines ignored):
    1. COMPANY_PROFILES: static_company_id, financial_year, company_name, ...
    12. FINANCIAL_STATEMENT_LINES (EMPTY): line_code, ...   <- header-only sheet

Columns are split into:
    SYSTEM  - filled by code (company id, year, source page, status columns, row id)
    MODEL   - extracted by the LLM from the PDF
Nothing about sheets or columns is hard-coded: a different spec gives a different extraction.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field


def norm(name: str) -> str:
    return re.sub(r"[^a-z0-9]", "", str(name).lower())


# normalised column name -> what the code fills in
SYSTEM_ALIASES = {
    "staticcompanyid": "company_id",
    "companyid": "company_id",
    "companyname": "company_name",
    "company": "company_name",
    "financialyear": "financial_year",
    "fiscalyear": "financial_year",
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
}
# The first of these in a sheet becomes a generated row id (e.g. MTN-DIRECTORS-2023-4).
ROW_ID_NAMES = {
    "recordkey", "linecode", "objectid", "objectobservationid", "eventid", "relationshipid",
    "attributeid", "policycode", "certificationcode", "targetcode", "adaptationcode",
    "programmecode", "riskcode", "scenariocode", "suppliercode", "assessmentcode",
    "traceabilitycode", "updatecode", "disclosurecode", "enforcementcode", "transitioncode",
    "factcode",
}
MONEY_NAMES = {
    "revenue", "operatingprofit", "profitbeforetax", "profitaftertax", "totalassets",
    "totalliabilities", "totalequity", "totaldebt", "cashandequivalents", "capitalexpenditure",
    "dividendspaid", "procurementspend", "annualspendamount", "climatecapex",
    "expectedfinancialloss", "penaltyamount", "amountinvested", "amountnumeric", "cashamount",
    "financialeffectvalue", "marketcapitalisation", "freefloatmarketcap", "valuetraded",
    "cashbalance", "netincome", "grossprofit", "ebitda", "totalrevenue",
}
NON_MONEY_MARKERS = (
    "percent", "rate", "ratio", "pershare", "peremployee", "percapita", "count", "headcount",
    "number", "frequency", "score", "intensity", "price", "volume", "shares", "latitude",
    "longitude", "year", "pct",
)
NUMERIC_HINTS = (
    "percent", "rate", "ratio", "amount", "count", "headcount", "numeric", "sharesheld",
    "sharesoutstanding", "numberof", "latitude", "longitude", "score", "volume", "exits",
    "price", "value", "total", "attended", "eligible", "frequency", "intensity",
    "capitalisation", "shareholding", "pershare", "earnings", "revenue", "profit", "assets",
    "liabilities", "equity", "spend", "expenditure", "emissions", "consumption",
)
NUMERIC_BLOCKERS = ("name", "description", "status", "date", "method", "code", "note", "type",
                    "text", "boolean", "flag", "unit", "currency", "countr", "account", "summary",
                    "boundary", "title", "role")


@dataclass
class Column:
    key: str          # short key sent to the LLM: c0, c1, ...
    name: str         # exact header from the spec (used in Excel)
    kind: str         # SYSTEM | MODEL
    alias: str = ""   # SYSTEM fill rule
    money: bool = False
    numeric: bool = False


@dataclass
class Sheet:
    name: str
    columns: list[Column]
    empty: bool = False
    number: int = 0
    model_columns: list[Column] = field(default_factory=list)

    @property
    def headers(self) -> list[str]:
        return [c.name for c in self.columns]

    @property
    def legend(self) -> dict[str, str]:
        return {c.key: c.name for c in self.model_columns}


def is_money(name: str) -> bool:
    k = norm(name)
    if any(m in k for m in NON_MONEY_MARKERS):
        return False
    return k in MONEY_NAMES or any(t in k for t in (
        "amount", "spend", "expenditure", "capex", "financialloss", "cashflow", "profit",
        "revenue", "assets", "liabilities", "equity", "ebitda", "income"))


def is_numeric(name: str) -> bool:
    k = norm(name)
    if k in ("financialyear", "ratingcycleyear", "baseyear", "targetyear") or k.startswith("is"):
        return False
    if any(b in k for b in NUMERIC_BLOCKERS):
        return False
    return is_money(name) or any(h in k for h in NUMERIC_HINTS)


def classify(columns: list[str]) -> list[Column]:
    out: list[Column] = []
    row_id_done = False
    for i, name in enumerate(columns):
        k = norm(name)
        if k in SYSTEM_ALIASES:
            out.append(Column(f"c{i}", name, "SYSTEM", SYSTEM_ALIASES[k]))
        elif k in ROW_ID_NAMES and not row_id_done:
            row_id_done = True
            out.append(Column(f"c{i}", name, "SYSTEM", "row_id"))
        else:
            out.append(Column(f"c{i}", name, "MODEL", money=is_money(name), numeric=is_numeric(name)))
    return out


_LINE = re.compile(r"^\s*(\d+)\s*[.)]\s*(.+?)\s*:\s*(.+?)\s*$")


def parse_spec(text: str) -> list[Sheet]:
    sheets: list[Sheet] = []
    seen: set[str] = set()
    for line_no, line in enumerate(text.splitlines(), 1):
        if not line.strip() or line.strip().startswith("#"):
            continue
        m = _LINE.match(line)
        if not m:
            raise ValueError(f"Spec line {line_no} must look like '1. SHEET_NAME: col1, col2, ...'")
        name = m.group(2).strip()
        empty = bool(re.search(r"\(EMPTY\)\s*$", name, re.I))
        name = re.sub(r"\s*\(EMPTY\)\s*$", "", name, flags=re.I).strip()
        if not name or len(name) > 31 or re.search(r"[\\/*?:\[\]]", name):
            raise ValueError(f"Spec line {line_no}: '{name}' is not a valid Excel sheet name (max 31 chars).")
        if name.casefold() in seen:
            raise ValueError(f"Spec line {line_no}: duplicate sheet '{name}'.")
        cols = [c.strip() for c in m.group(3).split(",") if c.strip()]
        if len({c.casefold() for c in cols}) != len(cols):
            raise ValueError(f"Spec line {line_no}: duplicate column in '{name}'.")
        seen.add(name.casefold())
        columns = classify(cols)
        sheets.append(Sheet(name=name, columns=columns, empty=empty, number=int(m.group(1)),
                            model_columns=[c for c in columns if c.kind == "MODEL"]))
    if not sheets:
        raise ValueError("The sheet spec is empty.")
    return sheets
