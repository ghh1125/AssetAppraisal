import json
from pathlib import Path

from demo.domain.pdf_ocr_fields import resolve_configured_ocr_fields
from demo.pipeline import _keep_unresolved_ocr_issues


def test_semantic_material_value_suppresses_stale_ocr_missing_issue():
    issues = [
        "major_long_term_assets：PDF OCR 表格单元格缺失：electronics@semantic",
        "tax_rates：PDF OCR 表格单元格缺失：vat@semantic",
    ]

    assert _keep_unresolved_ocr_issues(
        issues,
        {"major_long_term_assets": "固定资产账面价值1元。"},
    ) == ["tax_rates：PDF OCR 表格单元格缺失：vat@semantic"]


def test_semantic_ocr_locator_ignores_page_and_table_numbers():
    normalized = {
        "text_blocks": [
            {"page_number": 17, "text": "7、固定资产"},
        ],
        "table_cells": [
            {"page_number": 17, "table_id": "p17-t9", "row": 1, "column": 1, "text": "项目"},
            {"page_number": 17, "table_id": "p17-t9", "row": 1, "column": 2, "text": "机器设备"},
            {"page_number": 17, "table_id": "p17-t9", "row": 1, "column": 3, "text": "合计"},
            {"page_number": 17, "table_id": "p17-t9", "row": 8, "column": 1, "text": "2026.3.31账面价值"},
            {"page_number": 17, "table_id": "p17-t9", "row": 8, "column": 2, "text": "12,345.67"},
            {"page_number": 17, "table_id": "p17-t9", "row": 8, "column": 3, "text": "12,345.67"},
        ],
        "financial_data": [],
        "issues": [],
    }
    values, issues = resolve_configured_ocr_fields(
        normalized,
        {
            "ocr_field_rules": [
                {
                    "field_key": "major_long_term_assets",
                    "template": "固定资产{fixed_assets}元。",
                    "inputs": {
                        "fixed_assets": {
                            "kind": "semantic",
                            "page_markers": ["固定资产"],
                            "row_aliases": ["账面价值"],
                            "column_aliases": ["合计"],
                            "format": "amount",
                        }
                    },
                }
            ]
        },
    )
    assert values == {"major_long_term_assets": "固定资产12,345.67元。"}
    assert issues == []


def test_semantic_locator_uses_row_and_column_structure_to_choose_table_on_same_page():
    normalized = {
        "text_blocks": [{"page_number": 43, "text": "6、长期股权投资 7、固定资产"}],
        "table_cells": [
            {
                "page_number": 43,
                "table_id": "wrong",
                "row": 1,
                "column": 1,
                "text": "项目",
            },
            {
                "page_number": 43,
                "table_id": "wrong",
                "row": 1,
                "column": 2,
                "text": "2025年6月30日余额",
            },
            {
                "page_number": 43,
                "table_id": "wrong",
                "row": 2,
                "column": 1,
                "text": "合计",
            },
            {
                "page_number": 43,
                "table_id": "wrong",
                "row": 2,
                "column": 2,
                "text": "8,400,000.00",
            },
            {
                "page_number": 43,
                "table_id": "fixed-assets",
                "row": 1,
                "column": 1,
                "text": "项目",
            },
            {
                "page_number": 43,
                "table_id": "fixed-assets",
                "row": 1,
                "column": 2,
                "text": "机器设备",
            },
            {
                "page_number": 43,
                "table_id": "fixed-assets",
                "row": 1,
                "column": 3,
                "text": "合计",
            },
            {
                "page_number": 43,
                "table_id": "fixed-assets",
                "row": 25,
                "column": 1,
                "text": "3、2025.6.30账面价值",
            },
            {
                "page_number": 43,
                "table_id": "fixed-assets",
                "row": 25,
                "column": 2,
                "text": "4,466,842.92",
            },
            {
                "page_number": 43,
                "table_id": "fixed-assets",
                "row": 25,
                "column": 3,
                "text": "5,050,511.04",
            },
        ],
        "financial_data": [],
        "issues": [],
    }

    values, issues = resolve_configured_ocr_fields(
        normalized,
        {
            "ocr_field_rules": [
                {
                    "field_key": "major_long_term_assets",
                    "template": "固定资产{fixed_assets}元。",
                    "inputs": {
                        "fixed_assets": {
                            "kind": "semantic",
                            "page_markers": ["固定资产"],
                            "row_aliases": ["账面价值"],
                            "column_aliases": ["合计"],
                            "format": "amount",
                        }
                    },
                }
            ]
        },
    )

    assert values == {"major_long_term_assets": "固定资产5,050,511.04元。"}
    assert issues == []


def test_semantic_locator_does_not_fall_back_to_an_unrelated_numeric_column():
    normalized = {
        "text_blocks": [{"page_number": 43, "text": "7、固定资产"}],
        "table_cells": [
            {
                "page_number": 43,
                "table_id": "fixed-assets",
                "row": 1,
                "column": 1,
                "text": "项目",
            },
            {
                "page_number": 43,
                "table_id": "fixed-assets",
                "row": 1,
                "column": 2,
                "text": "机器设备",
            },
            {
                "page_number": 43,
                "table_id": "fixed-assets",
                "row": 1,
                "column": 3,
                "text": "电子设备",
            },
            {
                "page_number": 43,
                "table_id": "fixed-assets",
                "row": 1,
                "column": 4,
                "text": "合计",
            },
            {
                "page_number": 43,
                "table_id": "fixed-assets",
                "row": 25,
                "column": 1,
                "text": "3、2025.6.30账面价值",
            },
            {
                "page_number": 43,
                "table_id": "fixed-assets",
                "row": 25,
                "column": 2,
                "text": "4,466,842.92",
            },
            {
                "page_number": 43,
                "table_id": "fixed-assets",
                "row": 25,
                "column": 3,
                "text": "",
            },
            {
                "page_number": 43,
                "table_id": "fixed-assets",
                "row": 25,
                "column": 4,
                "text": "5,050,511.04",
            },
        ],
        "financial_data": [],
        "issues": [],
    }

    values, issues = resolve_configured_ocr_fields(
        normalized,
        {
            "ocr_field_rules": [
                {
                    "field_key": "major_long_term_assets",
                    "template": "电子设备{electronics}元。",
                    "inputs": {
                        "electronics": {
                            "kind": "semantic",
                            "page_markers": ["固定资产"],
                            "row_aliases": ["账面价值"],
                            "column_aliases": ["电子设备"],
                            "format": "amount",
                        }
                    },
                }
            ]
        },
    )

    assert values == {}
    assert issues == ["major_long_term_assets：PDF OCR 表格单元格缺失：electronics@semantic"]


def test_tongfu_electronics_locator_requires_an_explicit_electronics_column():
    config = json.loads(
        Path("demo/projects/tongfu.yaml").read_text(encoding="utf-8")
    )
    narrative_locator = config["ocr_field_rules"][0]["inputs"]["electronics"]
    table_locator = config["ocr_aux_fields"]["long_term_electronics"]

    assert narrative_locator["column_aliases"] == ["电子设备"]
    assert table_locator["column_aliases"] == ["电子设备"]
