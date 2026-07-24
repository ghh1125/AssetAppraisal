from __future__ import annotations

import re
from typing import Any

from demo.schemas import OcrPage


def _clean_text(value: str) -> str:
    return re.sub(r"\s+", " ", value).strip()


def _local_id(value: str, page_number: int) -> str:
    return re.sub(rf"^p{page_number}-", "", value, flags=re.IGNORECASE)


def normalize_ocr_pages(pages: list[dict[str, Any]]) -> dict[str, list[dict[str, Any]]]:
    """把任意 OCR 适配器输出归一为可审计、可 JSON 序列化的行记录。"""

    text_blocks: list[dict[str, Any]] = []
    table_cells: list[dict[str, Any]] = []
    for raw_page in pages:
        page = OcrPage.model_validate(raw_page)
        for block in page.blocks:
            block_id = _local_id(block.block_id, page.page_number)
            text_blocks.append(
                {
                    "page_number": page.page_number,
                    "page_count": page.page_count,
                    "block_id": block.block_id,
                    "block_type": block.block_type,
                    "text": _clean_text(block.text),
                    "confidence": block.confidence,
                    "bbox": list(block.bbox),
                    "evidence_id": f"pdf:p{page.page_number}:{block_id}",
                }
            )
        for table in page.tables:
            table_id = _local_id(table.table_id, page.page_number)
            for cell in table.cells:
                table_cells.append(
                    {
                        "page_number": page.page_number,
                        "page_count": page.page_count,
                        "table_id": table.table_id,
                        "row": cell.row,
                        "column": cell.column,
                        "text": _clean_text(cell.text),
                        "confidence": cell.confidence,
                        "bbox": list(cell.bbox),
                        "evidence_id": (
                            f"pdf:p{page.page_number}:{table_id}:r{cell.row}:c{cell.column}"
                        ),
                    }
                )
    return {
        "text_blocks": text_blocks,
        "table_cells": table_cells,
        "financial_data": [],
        "issues": [],
    }
