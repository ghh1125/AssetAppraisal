import json
from pathlib import Path

from demo.domain import calculations, registry
from demo.domain.comment_mapping import align_locations_to_output_template
from demo.domain.field_validation import (
    normalize_narrative_modules,
    normalize_valuation_methods,
)
from demo.domain.narrative_policy import select_narrative_fields
from demo.domain.generation_issues import missing_marker


def test_ten_representative_business_examples():
    cases = json.loads(Path("demo/fixtures/cases.yaml").read_text(encoding="utf-8"))
    expected = json.loads(Path("demo/expected/cases.yaml").read_text(encoding="utf-8"))
    assert len(cases) == 10
    actual = {}
    for case in cases:
        op = case["op"]
        if op == "resolve":
            actual[case["id"]] = registry.resolve_candidate(case["candidates"], case["priority"], case["field"])
        elif op == "split_date":
            actual[case["id"]] = calculations.split_date(case["value"])
        elif op == "convert_amount":
            actual[case["id"]] = calculations.convert_amount(case["value"], case["from_unit"], case["to_unit"])
        elif op == "increment":
            actual[case["id"]] = calculations.increment(case["appraised"], case["book"])
        elif op == "increment_rate":
            actual[case["id"]] = calculations.increment_rate(case["appraised"], case["book"])
        elif op == "validity":
            actual[case["id"]] = calculations.validity_period(case["value"])
        elif op == "company_name":
            actual[case["id"]] = calculations.normalize_company_name(case["value"])
        elif op == "methods":
            actual[case["id"]] = calculations.format_methods(case["value"])
        elif op == "human_fill":
            actual[case["id"]] = registry.human_fill(case["value"])
    assert actual == expected


def test_workflow_policy_examples():
    cases = json.loads(
        Path("demo/fixtures/workflow_cases.yaml").read_text(encoding="utf-8")
    )
    expected = json.loads(
        Path("demo/expected/workflow_cases.yaml").read_text(encoding="utf-8")
    )
    assert len(cases) >= 10
    actual = {}
    for case in cases:
        if case["op"] == "methods":
            actual[case["id"]] = normalize_valuation_methods(case["value"])
        elif case["op"] == "modules":
            actual[case["id"]] = normalize_narrative_modules(case["value"])
        elif case["op"] == "narrative_fields":
            actual[case["id"]] = sorted(
                select_narrative_fields(set(case["routed"]), case["selected"])
            )
        elif case["op"] == "missing_marker":
            actual[case["id"]] = missing_marker(case["location"])
        else:
            raise AssertionError(f"未知样例操作：{case['op']}")
    assert actual == expected


def test_domain_has_no_infrastructure_imports():
    banned = ("httpx", "openpyxl", "docx", "streamlit", "fastapi", "pathlib", "import os")
    for path in Path("demo/domain").glob("*.py"):
        text = path.read_text(encoding="utf-8")
        assert not any(word in text for word in banned)


def test_derives_report_and_validity_fields_from_dates():
    fields = {
        "valuation_date_year": "2025年06月30日",
        "valuation_date_month": "2025年06月30日",
        "valuation_date_day": "2025年06月30日",
    }
    actual = calculations.derive_system_fields(fields, report_date="2026-07-22")
    assert actual["report_date_year"] == 2026
    assert actual["report_date_month"] == 7
    assert actual["report_date_day"] == 22
    assert actual["report_number_year"] == 2026
    assert actual["validity_start_year"] == 2025
    assert actual["validity_start_month"] == 6
    assert actual["validity_start_day"] == 30
    assert actual["validity_end_year"] == 2026
    assert actual["validity_end_month"] == 6
    assert actual["validity_end_day"] == 29


def test_derives_template_dates_from_the_manual_valuation_base_date():
    actual = calculations.derive_system_fields(
        {"valuation_base_date": "2025-06-30"},
        report_date="2026-08-18",
    )

    assert actual["valuation_date_year"] == 2025
    assert actual["valuation_date_month"] == 6
    assert actual["valuation_date_day"] == 30
    assert actual["validity_start_year"] == 2025
    assert actual["validity_end_year"] == 2026
    assert actual["validity_end_month"] == 6
    assert actual["validity_end_day"] == 29


def test_aligns_comment_derived_field_to_clean_template_after_paragraph_shift():
    clean_locations = [
        {
            "location_id": "DOCUMENT-P0686-X01",
            "context": "有效期（即20XX年XX月XX日至20XX年XX月XX日）",
            "marker": "XX",
            "occurrence_index": 1,
            "part": "word/document.xml",
            "paragraph_index": 686,
            "record_type": "占位符",
        }
    ]
    annotated_locations = [
        {
            "location_id": "DOCUMENT-P0687-X01",
            "context": "有效期（即20XX年XX月XX日至20XX年XX月XX日）",
            "marker": "XX",
            "occurrence_index": 1,
            "part": "word/document.xml",
            "paragraph_index": 687,
            "record_type": "占位符",
            "field_key": "validity_start_year",
            "comment_texts": ["数据来源：人工基础信息（评估基准日）"],
        }
    ]

    aligned = align_locations_to_output_template(clean_locations, annotated_locations)

    assert aligned[0]["location_id"] == "DOCUMENT-P0686-X01"
    assert aligned[0]["field_key"] == "validity_start_year"


def test_alignment_does_not_create_a_mapping_for_an_unannotated_output_placeholder():
    aligned = align_locations_to_output_template(
        [
            {
                "location_id": "DOCUMENT-P0100-X01",
                "context": "未映射的XXX",
                "marker": "XXX",
                "occurrence_index": 1,
                "part": "word/document.xml",
                "paragraph_index": 100,
                "record_type": "占位符",
            }
        ],
        [],
    )

    assert aligned == []
