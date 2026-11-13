"""Offline checks for the supplied extraction schema (no API calls)."""
from pathlib import Path
import unittest

from Annual_Report_Vision_Agent import parse_sheet_spec


class TestSuppliedSchema(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        spec_path = Path(__file__).resolve().parents[1] / "SCHEMA_SPEC.txt"
        cls.schemas = parse_sheet_spec(spec_path.read_text(encoding="utf-8-sig"))

    def test_sheet_count(self):
        self.assertEqual(len(self.schemas), 41)

    def test_header_total(self):
        self.assertEqual(sum(len(s["columns"]) for s in self.schemas.values()), 937)

    def test_empty_sheets_are_header_only(self):
        empty_names = {name for name, schema in self.schemas.items() if schema["empty"]}
        self.assertEqual(empty_names, {
            "FINANCIAL_STATEMENT_LINES", "CLIMATE_ADAPTATION",
            "SUPPLIERS", "SUPPLIER_ASSESSMENTS", "SUPPLY_CHAIN_TRACEABILITY",
        })

    def test_exact_headers_unchanged(self):
        self.assertEqual(self.schemas["EMPLOYEE_TURNOVER"]["columns"][0], "Static Company ID")
        self.assertIn("Reported Turnover Rate (%)", self.schemas["EMPLOYEE_TURNOVER"]["columns"])
        self.assertEqual(list(self.schemas)[0], "COMPANY_PROFILES")
        self.assertEqual(list(self.schemas)[-1], "EMPLOYEE_TURNOVER")


if __name__ == "__main__":
    unittest.main()
