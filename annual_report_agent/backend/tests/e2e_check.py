"""End-to-end check against a running stack (real LLM).

    python backend/tests/make_sample_reports.py
    python backend/tests/e2e_check.py http://localhost:8080

Uploads the sample PDFs through the API, waits for the job, downloads the Excel workbook and
compares key values with the known answers in the sample reports.
"""
import io
import json
import sys
import time
import urllib.request
import uuid
from pathlib import Path

from openpyxl import load_workbook

BASE = (sys.argv[1] if len(sys.argv) > 1 else "http://localhost:8080").rstrip("/")
FIX = Path(__file__).parent / "fixtures"

EXPECTED = {   # (sheet, year, column) -> value; FY2024 is printed in millions, others in thousands
    ("FINANCIALS", 2022, "revenue"): 1_845_220_000,
    ("FINANCIALS", 2023, "revenue"): 2_104_876_000,
    ("FINANCIALS", 2024, "revenue"): 2_488_600_000,
    ("FINANCIALS", 2023, "profit_before_tax"): -35_214_000,
    ("FINANCIALS", 2024, "profit_before_tax"): 301_200_000,
    ("FINANCIALS", 2024, "total_assets"): 6_410_300_000,
    ("FINANCIALS", 2022, "total_equity"): 2_119_484_000,
    ("WORKFORCE", 2023, "total_employees"): 2512,
    ("WORKFORCE", 2024, "total_employees"): 2630,
}
EXPECTED_DIRECTORS = {2022: 3, 2023: 4, 2024: 4}
EXPECTED_EMISSIONS = {(2023, "1"): 17880, (2024, "1"): 16940, (2024, "2"): 38870}


def post_job() -> str:
    boundary = uuid.uuid4().hex
    body = io.BytesIO()

    def field(name, value):
        body.write(f'--{boundary}\r\nContent-Disposition: form-data; name="{name}"\r\n\r\n{value}\r\n'.encode())

    field("company_name", "Acme Telecoms Plc")
    field("company_id", "ACME01")
    field("spec_text", (FIX / "test_spec.txt").read_text())
    for pdf in sorted(FIX.glob("*.pdf")):
        body.write(f'--{boundary}\r\nContent-Disposition: form-data; name="files"; filename="{pdf.name}"\r\n'
                   f"Content-Type: application/pdf\r\n\r\n".encode())
        body.write(pdf.read_bytes() + b"\r\n")
    body.write(f"--{boundary}--\r\n".encode())
    req = urllib.request.Request(f"{BASE}/api/jobs", data=body.getvalue(), method="POST",
                                 headers={"Content-Type": f"multipart/form-data; boundary={boundary}"})
    return json.load(urllib.request.urlopen(req))["job_id"]


def main() -> int:
    t0 = time.time()
    job_id = post_job()
    print("job", job_id)
    while True:
        job = json.load(urllib.request.urlopen(f"{BASE}/api/jobs/{job_id}"))
        if job["status"] in ("done", "failed"):
            break
        time.sleep(3)
    print(f"status={job['status']} in {time.time() - t0:.0f}s; warnings={job['warnings']}")
    for s in job["summary"]:
        print("  ", s)
    if job["status"] != "done":
        print("ERROR", job.get("error"))
        return 1

    wb = load_workbook(io.BytesIO(urllib.request.urlopen(f"{BASE}/api/jobs/{job_id}/download").read()))

    def rows(sheet):
        ws = wb[sheet]
        head = [c.value for c in ws[1]]
        return [dict(zip(head, r)) for r in ws.iter_rows(min_row=2, values_only=True)]

    failures = 0

    def check(label, got, want):
        nonlocal failures
        ok = got == want
        failures += not ok
        print(f"  {'PASS' if ok else 'FAIL'}  {label}: got {got!r}, expected {want!r}")

    for (sheet, year, col), want in EXPECTED.items():
        match = [r for r in rows(sheet) if r["financial_year"] == year]
        check(f"{sheet} FY{year} {col}", match[0][col] if match else None, want)
    for year, n in EXPECTED_DIRECTORS.items():
        check(f"DIRECTORS FY{year} count", sum(r["financial_year"] == year for r in rows("DIRECTORS")), n)
    for (year, scope), want in EXPECTED_EMISSIONS.items():
        match = [r for r in rows("CLIMATE_EMISSIONS") if r["financial_year"] == year
                 and f"scope {scope}" in str(r["scope_type"]).lower()]
        check(f"CLIMATE_EMISSIONS FY{year} scope {scope}", match[0]["emissions_value"] if match else None, want)
    check("SUPPLIERS (EMPTY) has no rows", len(rows("SUPPLIERS")), 0)
    statuses = [r["verification_status"] for s in ("FINANCIALS", "DIRECTORS", "CLIMATE_EMISSIONS", "WORKFORCE")
                for r in rows(s)]
    print("  verification statuses:", {s: statuses.count(s) for s in set(statuses)})
    print(f"\n{failures} failure(s)")
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
