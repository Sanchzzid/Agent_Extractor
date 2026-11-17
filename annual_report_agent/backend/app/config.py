"""All runtime settings, read from environment variables (.env)."""
from __future__ import annotations

import os
from pathlib import Path

from dotenv import load_dotenv

# Load the backend .env first, then fall back to the repository-root .env.
_HERE = Path(__file__).resolve()
for parent in list(_HERE.parents)[1:4]:   # backend/, annual_report_agent/, repo root
    candidate = parent / ".env"
    if candidate.exists():
        load_dotenv(candidate, override=False)


def _int(name: str, default: int) -> int:
    try:
        return int(os.getenv(name, default))
    except ValueError:
        return default


def _bool(name: str, default: bool) -> bool:
    return os.getenv(name, str(default)).strip().lower() in {"1", "true", "yes", "on"}


def _first(*names: str) -> str:
    for name in names:
        value = os.getenv(name, "").strip()
        if value:
            return value
    return ""


# ---------- LLM: ONE model (any OpenAI-compatible API) does extraction AND verification ----------
# The old PUBLICA_LLM_* names are still accepted so existing .env files keep working.
LLM_ENDPOINT = _first("LLM_BASE_URL", "PUBLICA_LLM_ENDPOINT")   # empty -> api.openai.com
LLM_MODEL = _first("LLM_MODEL", "PUBLICA_LLM_MODEL")
LLM_API_KEY = _first("LLM_API_KEY", "PUBLICA_LLM_API_KEY")
LLM_JSON_MODE = _bool("LLM_JSON_MODE", True)   # set false if the provider rejects response_format
LLM_CONTEXT_TOKENS = _int("LLM_CONTEXT_TOKENS", 30000)      # model's max context window
LLM_MAX_OUTPUT_TOKENS = _int("LLM_MAX_OUTPUT_TOKENS", 8000)
LLM_TIMEOUT_SECONDS = _int("LLM_TIMEOUT_SECONDS", 180)
LLM_JSON_RETRIES = _int("LLM_JSON_RETRIES", 3)              # re-asks on invalid JSON
LLM_NETWORK_RETRIES = _int("LLM_NETWORK_RETRIES", 4)        # retries on HTTP / timeout errors
LLM_MAX_CONCURRENCY = _int("LLM_MAX_CONCURRENCY", 8)        # simultaneous calls to the endpoint
USE_VISION = _bool("USE_VISION", True)                      # send scanned pages as images

# ---------- Workers ----------
REPORT_WORKERS = _int("REPORT_WORKERS", 3)   # PDFs processed at the same time (per job)
SHEET_WORKERS = _int("SHEET_WORKERS", 4)     # sheets processed at the same time (per PDF)
JOB_WORKERS = _int("JOB_WORKERS", 2)         # jobs (uploads) running at the same time

# ---------- Extraction tuning ----------
PAGE_TOKEN_BUDGET = _int("PAGE_TOKEN_BUDGET", 12000)   # report text sent per extraction call
MAX_PAGES_PER_SHEET = _int("MAX_PAGES_PER_SHEET", 10)
INDEX_TOKEN_BUDGET = _int("INDEX_TOKEN_BUDGET", 14000)  # page index sent to the page locator
IMAGE_TOKEN_COST = _int("IMAGE_TOKEN_COST", 1100)       # budget reserved per page image
IMAGE_DPI = _int("IMAGE_DPI", 150)
CONVERT_SCALE = _bool("CONVERT_SCALE", True)           # "in millions" -> absolute units

# ---------- Server ----------
PORT = _int("PORT", 8023)
DATA_DIR = Path(os.getenv("DATA_DIR", str(_HERE.parents[1] / "data")))
MAX_FILES_PER_JOB = _int("MAX_FILES_PER_JOB", 10)
MAX_UPLOAD_MB = _int("MAX_UPLOAD_MB", 150)
ALLOWED_ORIGINS = [o.strip() for o in os.getenv("ALLOWED_ORIGINS", "*").split(",") if o.strip()]
DEFAULT_SPEC_PATH = Path(os.getenv("DEFAULT_SPEC_PATH", str(_HERE.parents[1] / "default_schema.txt")))


def check_llm_settings() -> None:
    missing = [n for n, v in (("LLM_MODEL", LLM_MODEL), ("LLM_API_KEY", LLM_API_KEY)) if not v]
    if missing:
        raise RuntimeError(f"Missing in .env: {', '.join(missing)}")
