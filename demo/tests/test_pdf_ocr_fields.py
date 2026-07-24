from demo.domain.pdf_ocr_fields import resolve_configured_ocr_fields


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
