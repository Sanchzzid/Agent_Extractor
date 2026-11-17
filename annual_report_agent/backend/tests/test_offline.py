"""Offline tests: no network. The LLM is replaced by a scripted fake."""
import sys
from pathlib import Path

import pytest
from langchain_core.messages import AIMessage

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from app import config, excel, extraction, llm  # noqa: E402
from app.schema import parse_spec  # noqa: E402

FIX = Path(__file__).parent / "fixtures"
SPEC = (FIX / "test_spec.txt").read_text()


# ------------------------------------------------------------------ schema
def test_spec_parsing_and_classification():
    sheets = parse_spec(SPEC)
    assert [s.name for s in sheets] == ["FINANCIALS", "DIRECTORS", "CLIMATE_EMISSIONS", "WORKFORCE", "SUPPLIERS"]
    fin = sheets[0]
    kinds = {c.name: c.kind for c in fin.columns}
    assert kinds["static_company_id"] == "SYSTEM" and kinds["revenue"] == "MODEL"
    assert next(c for c in fin.columns if c.name == "revenue").money
    assert sheets[4].empty
    assert next(c for c in sheets[1].columns if c.name == "record_key").alias == "row_id"


def test_default_schema_parses():
    sheets = parse_spec((Path(__file__).parents[1] / "default_schema.txt").read_text(encoding="utf-8-sig"))
    assert len(sheets) == 41


@pytest.mark.parametrize("bad", ["FINANCIALS revenue", "1. BAD/NAME: a, b", "1. A: x, x"])
def test_spec_errors(bad):
    with pytest.raises(ValueError):
        parse_spec(bad)


# ------------------------------------------------------------------ JSON parsing
@pytest.mark.parametrize("text", [
    '{"rows": []}',
    '```json\n{"rows": []}\n```',
    'Here you go: {"rows": []} thanks',
    '{"rows": [{"c1": 5,}]}',           # trailing comma -> json_repair
])
def test_parse_json_object_tolerant(text):
    assert "rows" in llm.parse_json_object(text)


def test_parse_json_rejects_non_object():
    with pytest.raises(ValueError):
        llm.parse_json_object("[1, 2]")


# ------------------------------------------------------------------ retries
class FakeChat:
    def __init__(self, replies):
        self.replies = list(replies)
        self.calls = 0

    def invoke(self, messages):
        self.calls += 1
        r = self.replies.pop(0)
        if isinstance(r, Exception):
            raise r
        return AIMessage(content=r, response_metadata={"finish_reason": "stop"})


def _patch(monkeypatch, fake):
    monkeypatch.setattr(llm, "_chat", lambda *a, **k: fake)
    monkeypatch.setattr(llm.time, "sleep", lambda s: None)


def test_reasks_after_invalid_json(monkeypatch):
    fake = FakeChat(["not json at all", '{"pages": "oops"}', '{"pages": [3, 4]}'])
    _patch(monkeypatch, fake)

    def validate(d):
        if not isinstance(d["pages"], list):
            raise ValueError("pages must be a list")
        return d
    assert llm.call_json("x", validate)["pages"] == [3, 4]
    assert fake.calls == 3


def test_network_errors_are_retried(monkeypatch):
    fake = FakeChat([ConnectionError("connection reset"), TimeoutError("timed out"), '{"ok": true}'])
    _patch(monkeypatch, fake)
    assert llm.call_json("x") == {"ok": True}


def test_gives_up_after_retries(monkeypatch):
    fake = FakeChat(["bad"] * (config.LLM_JSON_RETRIES + 1))
    _patch(monkeypatch, fake)
    with pytest.raises(llm.LLMError):
        llm.call_json("x", lambda d: d)


def test_context_error_is_not_retried(monkeypatch):
    fake = FakeChat([RuntimeError("This model's maximum context length is 30000 tokens")])
    _patch(monkeypatch, fake)
    with pytest.raises(llm.ContextTooLarge):
        llm.call_json("x")
    assert fake.calls == 1


# ------------------------------------------------------------------ numbers / verification guard
def test_numbers_and_signs():
    assert extraction.to_decimal("(1,245)") == -1245
    assert extraction.to_decimal("12.5%") == extraction.to_decimal("12.5")
    assert extraction.to_decimal("abc") is None
    assert extraction.same_value(17880, "17,880")
    assert not extraction.same_value(17880, "17,881")


def test_value_in_batch():
    b = extraction.Batch([1], "=== PAGE 1 ===\nRevenue 2,104,876 1,845,220\nChairman Adaeze Okafor")
    assert extraction.value_in_batch(2104876, b) is True
    assert extraction.value_in_batch(-2104876, b) is True
    assert extraction.value_in_batch(9999999, b) is False
    assert extraction.value_in_batch("adaeze okafor", b) is True


def test_verifier_disputed_value_is_not_applied(monkeypatch):
    sheet = parse_spec(SPEC)[0]
    key = next(c.key for c in sheet.model_columns if c.name == "revenue")
    batch = extraction.Batch([6], "=== PAGE 6 ===\nRevenue 2,104,876 1,845,220")
    row = {key: 2104876, "_batch": [6]}
    monkeypatch.setattr(llm, "call_json", lambda *a, **k: {"issues": [
        {"row": 0, "key": key, "verdict": "WRONG", "correct_value": 5555555, "reason": "x"}]})
    log = []
    extraction.verify_batch(sheet, batch, [row], 2023, {}, log)
    assert row[key] == 2104876 and row["_verification"] == "NEEDS_REVIEW"
    assert log[0]["check"] == "DISPUTED"


def test_verifier_not_found_but_printed_is_kept(monkeypatch):
    sheet = parse_spec(SPEC)[0]
    key = next(c.key for c in sheet.model_columns if c.name == "revenue")
    batch = extraction.Batch([6], "=== PAGE 6 ===\nRevenue 2,104,876")
    row = {key: 2104876, "_batch": [6]}
    monkeypatch.setattr(llm, "call_json", lambda *a, **k: {"issues": [
        {"row": 0, "key": key, "verdict": "NOT_FOUND", "correct_value": "N/A", "reason": "x"}]})
    log = []
    extraction.verify_batch(sheet, batch, [row], 2023, {}, log)
    assert row[key] == 2104876 and log[0]["check"] == "VERIFIER OVERRULED"


# ------------------------------------------------------------------ merging rows across page batches
def test_merge_rows_from_different_batches():
    rows = [{"c1": 100, "c2": "N/A", "_batch": [6], "_pages": [6]},
            {"c1": "N/A", "c2": 500, "_batch": [7], "_pages": [7]},
            {"c1": "Jane", "c2": "N/A", "_batch": [3], "_pages": [3]},
            {"c1": "John", "c2": "N/A", "_batch": [3], "_pages": [3]}]
    merged = extraction.merge_batch_rows(rows, ["c1", "c2"])
    assert merged[0]["c1"] == 100 and merged[0]["c2"] == 500 and merged[0]["_pages"] == [6, 7]
    assert len(merged) == 3


# ------------------------------------------------------------------ excel safety
def test_excel_blocks_formula_injection(tmp_path):
    sheets = parse_spec("1. S: static_company_id, note_text")
    merged = {"S": [{"values": {"static_company_id": "X", "note_text": "=HYPERLINK(\"http://evil\")"},
                     "status": "VERIFIED"}]}
    path = tmp_path / "o.xlsx"
    excel.write_workbook(path, sheets, merged, [], [])
    from openpyxl import load_workbook
    ws = load_workbook(path)["S"]
    assert ws["B2"].data_type == "s"


def test_verifier_scale_only_correction_is_ignored(monkeypatch):
    sheet = parse_spec(SPEC)[0]
    key = next(c.key for c in sheet.model_columns if c.name == "revenue")
    batch = extraction.Batch([6], "=== PAGE 6 ===\nRevenue 2,104,876", has_images=True)
    row = {key: 2104876, "_batch": [6]}
    monkeypatch.setattr(llm, "call_json", lambda *a, **k: {"issues": [
        {"row": 0, "key": key, "verdict": "WRONG", "correct_value": "2,104,876,000", "reason": "N'000"}]})
    log = []
    extraction.verify_batch(sheet, batch, [row], 2023, {}, log)
    assert row[key] == 2104876 and row["_verification"] == "VERIFIED" and not log


def test_country_columns_are_text():
    from app.schema import is_numeric
    assert not is_numeric("countries_of_operation") and not is_numeric("country")
    assert is_numeric("employee_count") and is_numeric("revenue")
