"""PDF reading: page text, printed page labels, a compact page index and page images."""
from __future__ import annotations

import re
from dataclasses import dataclass

import pymupdf as fitz


@dataclass
class Page:
    number: int        # 1-based PDF page index
    label: str         # printed page number when detectable, else the index
    text: str
    scanned: bool      # little or no text layer -> send as image
    caption: str = ""  # vision-generated topic for scanned pages

    def index_line(self, width: int = 110) -> str:
        if self.scanned:
            return f"P{self.number}: [scanned] {self.caption or 'image-only page'}"[:width + 15]
        lines = [re.sub(r"\s+", " ", ln).strip(" -|••") for ln in self.text.splitlines()]
        words = [ln for ln in lines if 3 <= len(ln) <= 120 and not re.fullmatch(r"[\d\s.,()%-]+", ln)]
        return f"P{self.number}: " + " / ".join(words[:4])[:width]


def is_pdf(data: bytes) -> bool:
    return data[:5] == b"%PDF-"


def _printed_label(page: fitz.Page, index: int, text: str) -> str:
    try:
        label = str(page.get_label()).strip()
        if label and len(label) < 10:
            return label
    except Exception:  # noqa: BLE001 - labels are optional
        pass
    tail = [s.strip() for s in text.splitlines() if s.strip()][-4:]
    for line in reversed(tail):
        if re.fullmatch(r"(?:page\s*)?\d{1,4}", line, re.I):
            return re.sub(r"\D", "", line)
    return str(index)


def read_pdf(data: bytes) -> list[Page]:
    pages: list[Page] = []
    with fitz.open(stream=data, filetype="pdf") as doc:
        for i, page in enumerate(doc, 1):
            text = page.get_text("text", sort=True) or ""
            text = re.sub(r"[ \t]{2,}", "  ", text)
            text = re.sub(r"\n{3,}", "\n\n", text).strip()
            scanned = len(text) < 120 and bool(page.get_images(full=False))
            pages.append(Page(i, _printed_label(page, i, text), text, scanned))
    return pages


def render_png(data: bytes, page_number: int, dpi: int = 150) -> bytes:
    with fitz.open(stream=data, filetype="pdf") as doc:
        pix = doc[page_number - 1].get_pixmap(matrix=fitz.Matrix(dpi / 72, dpi / 72), alpha=False)
        return pix.tobytes("png")


def scale_snippets(pages: list[Page], limit: int = 25) -> str:
    """Lines that state the currency/unit of the financial statements."""
    out: list[str] = []
    pattern = re.compile(r"(thousand|million|billion|['’]000|\bin\s+[A-Z]{3}\b|statement of financial position|"
                         r"statement of profit|income statement|balance sheet|reporting currency)", re.I)
    for p in pages:
        lines = p.text.splitlines()
        for j, ln in enumerate(lines):
            if pattern.search(ln):
                out.append(f"P{p.number}: " + " ".join(lines[max(0, j - 1):j + 2])[:240])
                if len(out) >= limit:
                    return "\n".join(out)
    return "\n".join(out)


YEAR_RE = re.compile(r"(?<!\d)(20\d{2}|19\d{2})(?!\d)")


def year_from_filename(name: str) -> int | None:
    years = [int(y) for y in YEAR_RE.findall(name)]
    return max(years) if years else None
