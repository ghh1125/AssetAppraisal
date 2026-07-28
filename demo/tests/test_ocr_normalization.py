import json

from demo.domain.ocr_normalization import normalize_ocr_pages
from demo.schemas import OcrDocument


def test_normalizes_text_and_table_cells_without_sdk_objects():
    pages = [
        {
            "page_number": 3,
            "page_count": 8,
            "blocks": [
                {
                    "block_id": "p3-b1",
                    "block_type": "text",
                    "text": " 资产\n总计 ",
                    "confidence": 0.98,
                    "bbox": [1, 2, 3, 4],
                }
            ],
            "tables": [
                {
                    "table_id": "p3-t1",
                    "cells": [
                        {
                            "row": 1,
                            "column": 2,
                            "row_span": 2,
                            "column_span": 3,
                            "text": "1,234.50",
                            "confidence": 0.97,
                            "bbox": [5, 6, 7, 8],
                        }
                    ],
                }
            ],
        }
    ]

    result = normalize_ocr_pages(pages)

    assert result["text_blocks"][0]["text"] == "资产 总计"
    assert result["text_blocks"][0]["evidence_id"] == "pdf:p3:b1"
    assert result["table_cells"][0]["text"] == "1,234.50"
    assert result["table_cells"][0]["row_span"] == 2
    assert result["table_cells"][0]["column_span"] == 3
    assert result["table_cells"][0]["evidence_id"] == "pdf:p3:t1:r1:c2"
    json.dumps(result, ensure_ascii=False)


def test_ocr_document_contract_round_trips_plain_json():
    document = OcrDocument.model_validate(
        {
            "source_file": "审计报告.pdf",
            "pages": [
                {
                    "page_number": 1,
                    "page_count": 1,
                    "blocks": [],
                    "tables": [],
                }
            ],
            "issues": [],
        }
    )

    assert OcrDocument.model_validate_json(document.model_dump_json()) == document
