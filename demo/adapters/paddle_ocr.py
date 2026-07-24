from __future__ import annotations

from pathlib import Path
from typing import Any


def _plain(value: Any) -> Any:
    if isinstance(value, dict):
        return {str(key): _plain(item) for key, item in value.items()}
    if isinstance(value, (list, tuple)):
        return [_plain(item) for item in value]
    if hasattr(value, "tolist"):
        return _plain(value.tolist())
    if isinstance(value, (str, int, float, bool)) or value is None:
        return value
    return str(value)


def _bbox(value: Any) -> list[float]:
    box = _plain(value) or []
    if len(box) == 4 and all(isinstance(item, (int, float)) for item in box):
        return [float(item) for item in box]
    points = [item for item in box if isinstance(item, list) and len(item) >= 2]
    if points:
        xs = [float(item[0]) for item in points]
        ys = [float(item[1]) for item in points]
        return [min(xs), min(ys), max(xs), max(ys)]
    return []


def _records(ocr_result: dict[str, Any]) -> list[tuple[str, float | None, list[float]]]:
    texts = _plain(ocr_result.get("rec_texts", []))
    scores = _plain(ocr_result.get("rec_scores", []))
    boxes = _plain(ocr_result.get("rec_boxes") or ocr_result.get("rec_polys") or [])
    records = []
    for index, text in enumerate(texts):
        score = float(scores[index]) if index < len(scores) and scores[index] is not None else None
        box = _bbox(boxes[index]) if index < len(boxes) else []
        records.append((str(text), score, box))
    return records


def _table_cells(ocr_result: dict[str, Any]) -> list[dict[str, Any]]:
    records = _records(ocr_result)
    positioned = []
    for index, (text, confidence, box) in enumerate(records):
        center_x = (box[0] + box[2]) / 2 if len(box) == 4 else float(index)
        center_y = (box[1] + box[3]) / 2 if len(box) == 4 else 0.0
        height = box[3] - box[1] if len(box) == 4 else 1.0
        positioned.append((center_y, center_x, height, text, confidence, box))
    positioned.sort(key=lambda item: (item[0], item[1]))

    rows: list[list[tuple[float, float, float, str, float | None, list[float]]]] = []
    for item in positioned:
        if not rows:
            rows.append([item])
            continue
        row_center = sum(cell[0] for cell in rows[-1]) / len(rows[-1])
        row_height = max(cell[2] for cell in rows[-1])
        if abs(item[0] - row_center) <= max(2.0, 0.6 * max(row_height, item[2])):
            rows[-1].append(item)
        else:
            rows.append([item])

    cells = []
    for row_index, row in enumerate(rows, 1):
        for column_index, item in enumerate(sorted(row, key=lambda cell: cell[1]), 1):
            cells.append(
                {
                    "row": row_index,
                    "column": column_index,
                    "text": item[3],
                    "confidence": item[4],
                    "bbox": item[5],
                }
            )
    return cells


class PaddleStructureOcrAdapter:
    """将 PP-StructureV3 结果转换为与 SDK 无关的普通字典。"""

    def __init__(self, pipeline: Any):
        self.pipeline = pipeline

    def extract(self, pdf_path: Path) -> tuple[list[dict[str, Any]], list[str]]:
        try:
            results = self.pipeline.predict(input=str(pdf_path))
            pages = []
            for fallback_index, result in enumerate(results):
                payload = result.json if not callable(result.json) else result.json()
                payload = _plain(payload)
                data = payload.get("res", payload)
                page_number = int(data.get("page_index", fallback_index)) + 1
                page_count = int(data.get("page_count") or page_number)
                blocks = [
                    {
                        "block_id": f"p{page_number}-b{index}",
                        "block_type": "text",
                        "text": text,
                        "confidence": confidence,
                        "bbox": box,
                    }
                    for index, (text, confidence, box) in enumerate(
                        _records(data.get("overall_ocr_res", {})), 1
                    )
                ]
                tables = []
                for index, table in enumerate(data.get("table_res_list", []), 1):
                    table_ocr = table.get("table_ocr_pred") or table.get("overall_ocr_res") or {}
                    tables.append(
                        {
                            "table_id": f"p{page_number}-t{index}",
                            "cells": _table_cells(table_ocr),
                        }
                    )
                pages.append(
                    {
                        "page_number": page_number,
                        "page_count": page_count,
                        "blocks": blocks,
                        "tables": tables,
                    }
                )
            return pages, []
        except Exception as exc:
            return [], [f"PaddleOCR 失败：{exc}"]


def create_local_pipeline(**overrides: Any) -> Any:
    from paddleocr import PPStructureV3

    options = {
        "use_doc_orientation_classify": False,
        "use_doc_unwarping": False,
        "use_textline_orientation": False,
        "use_formula_recognition": False,
        "use_table_recognition": True,
    }
    options.update(overrides)
    return PPStructureV3(**options)
