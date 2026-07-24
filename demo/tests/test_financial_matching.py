import json
from pathlib import Path

import pytest

from demo.domain.field_validation import require_financial_fields
from demo.domain.financial_matching import (
    match_financial_table,
    normalize_label,
    parse_number,
    parse_period,
)


def test_same_values_match_when_columns_and_units_change():
    cells = [
        {"row": 1, "column": 1, "text": "项目"},
        {"row": 1, "column": 2, "text": "2025年6月30日"},
        {"row": 1, "column": 3, "text": "2024年12月31日"},
        {"row": 2, "column": 1, "text": "资产合计"},
        {"row": 2, "column": 2, "text": "16371.913179"},
        {"row": 2, "column": 3, "text": "18082.424617"},
    ]

    result = match_financial_table(
        cells,
        aliases={"total_assets": ["总资产", "资产合计"]},
        unit="万元",
    )

    assert result["total_assets"]["2025-06-30"] == 163_719_131.79
    assert result["total_assets"]["2024-12-31"] == 180_824_246.17


def test_fixture_cases_cover_layout_period_unit_missing_and_conflict_rules():
    cases = json.loads(Path("demo/fixtures/ocr_cases.yaml").read_text(encoding="utf-8"))
    expected = json.loads(Path("demo/expected/ocr_cases.yaml").read_text(encoding="utf-8"))

    assert len(cases) >= 10
    for case in cases:
        if case["operation"] == "match":
            actual = match_financial_table(**case["input"])
        elif case["operation"] == "validate":
            actual = require_financial_fields(**case["input"])
        else:
            raise AssertionError(f"未知样例操作：{case['operation']}")
        assert actual == expected[case["id"]], case["id"]


@pytest.mark.parametrize(
    ("raw", "expected"),
    [
        ("  一、资产 总计：", "资产总计"),
        ("2025年1-6月", "2025-06-30"),
        ("（ 1，234.50 ）", -1234.5),
    ],
)
def test_normalizers(raw, expected):
    if isinstance(expected, str) and expected.startswith("2025-"):
        assert parse_period(raw) == expected
    elif isinstance(expected, str):
        assert normalize_label(raw) == expected
    else:
        assert parse_number(raw) == expected
