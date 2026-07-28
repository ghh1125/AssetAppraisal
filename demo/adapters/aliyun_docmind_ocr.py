from __future__ import annotations

import re
import time
from pathlib import Path
from typing import Any, Callable


def _plain(value: Any) -> Any:
    if isinstance(value, dict):
        return {str(key): _plain(item) for key, item in value.items()}
    if isinstance(value, (list, tuple)):
        return [_plain(item) for item in value]
    if hasattr(value, "to_map"):
        return _plain(value.to_map())
    if isinstance(value, (str, int, float, bool)) or value is None:
        return value
    return str(value)


def _bbox(pos: Any) -> list[float]:
    points = _plain(pos) or []
    if len(points) == 4 and all(isinstance(item, (int, float)) for item in points):
        return [float(item) for item in points]
    coordinates: list[tuple[float, float]] = []
    for point in points:
        if isinstance(point, dict) and point.get("x") is not None and point.get("y") is not None:
            coordinates.append((float(point["x"]), float(point["y"])))
        elif isinstance(point, list) and len(point) >= 2:
            coordinates.append((float(point[0]), float(point[1])))
    if not coordinates:
        return []
    xs = [point[0] for point in coordinates]
    ys = [point[1] for point in coordinates]
    return [min(xs), min(ys), max(xs), max(ys)]


def _markdown_text(value: Any) -> str:
    text = str(value or "").strip()
    text = re.sub(r"^#{1,6}\s*", "", text)
    text = re.sub(r"\s{2,}\n", "\n", text)
    return text.strip()


def _cell_text(cell: dict[str, Any]) -> str:
    direct = str(cell.get("text") or "").strip()
    if direct:
        return direct
    parts = [
        str(layout.get("text") or layout.get("markdownContent") or "").strip()
        for layout in cell.get("layouts", [])
        if isinstance(layout, dict)
    ]
    return "".join(part for part in parts if part)


def _table_cells(layout: dict[str, Any]) -> list[dict[str, Any]]:
    cells = []
    for cell in layout.get("cells", []):
        if not isinstance(cell, dict):
            continue
        row_start = int(cell.get("ysc") or 0)
        row_end = int(cell.get("yec") if cell.get("yec") is not None else row_start)
        column_start = int(cell.get("xsc") or 0)
        column_end = int(cell.get("xec") if cell.get("xec") is not None else column_start)
        cells.append(
            {
                "row": row_start + 1,
                "column": column_start + 1,
                "row_span": max(1, row_end - row_start + 1),
                "column_span": max(1, column_end - column_start + 1),
                "text": _cell_text(cell),
                "confidence": cell.get("layoutConf"),
                "bbox": _bbox(cell.get("pos")),
            }
        )
    return cells


def layouts_to_pages(layouts: list[dict[str, Any]]) -> list[dict[str, Any]]:
    normalized = [_plain(layout) for layout in layouts if isinstance(layout, dict)]
    if not normalized:
        return []
    page_count = max(int(layout.get("pageNum") or 0) for layout in normalized) + 1
    pages = {
        page_number: {
            "page_number": page_number,
            "page_count": page_count,
            "blocks": [],
            "tables": [],
        }
        for page_number in range(1, page_count + 1)
    }
    for fallback_index, layout in enumerate(normalized, 1):
        page_number = int(layout.get("pageNum") or 0) + 1
        unique_id = str(layout.get("uniqueId") or f"p{page_number}-l{fallback_index}")
        if layout.get("type") == "table":
            pages[page_number]["tables"].append(
                {
                    "table_id": unique_id,
                    "cells": _table_cells(layout),
                }
            )
            continue
        text = str(layout.get("text") or "").strip()
        if not text:
            text = _markdown_text(layout.get("markdownContent"))
        pages[page_number]["blocks"].append(
            {
                "block_id": unique_id,
                "block_type": str(layout.get("type") or "text"),
                "text": text,
                "confidence": layout.get("layoutConf"),
                "bbox": _bbox(layout.get("pos")),
            }
        )
    return list(pages.values())


def _mapping(value: Any) -> dict[str, Any]:
    plain = _plain(value)
    if not isinstance(plain, dict):
        return {}
    data = plain.get("Data") or plain.get("data")
    if isinstance(data, dict):
        return data
    return plain


class AliyunDocMindOcrAdapter:
    def __init__(
        self,
        client: Any,
        *,
        vlm: bool = False,
        poll_interval_seconds: float = 5,
        timeout_seconds: float = 900,
        layout_step_size: int = 3000,
        sleep: Callable[[float], None] = time.sleep,
        clock: Callable[[], float] = time.monotonic,
        redact_values: tuple[str, ...] = (),
    ):
        self.client = client
        self.vlm = vlm
        self.poll_interval_seconds = poll_interval_seconds
        self.timeout_seconds = timeout_seconds
        self.layout_step_size = layout_step_size
        self.sleep = sleep
        self.clock = clock
        self.redact_values = tuple(value for value in redact_values if value)

    def _safe_text(self, value: object) -> str:
        message = str(value)
        for redact_value in self.redact_values:
            message = message.replace(redact_value, "***")
        return message

    def _safe_error(self, error: Exception) -> str:
        return self._safe_text(error)

    def extract(self, pdf_path: Path) -> tuple[list[dict[str, Any]], list[str]]:
        try:
            task_id = self.client.submit(pdf_path, vlm=self.vlm)
        except Exception as exc:
            return [], [f"阿里云 OCR 提交失败：{self._safe_error(exc)}"]

        started_at = self.clock()
        while True:
            try:
                status_payload = _mapping(self.client.status(task_id))
            except Exception as exc:
                return [], [
                    f"阿里云 OCR 状态查询失败（{task_id}）：{self._safe_error(exc)}"
                ]
            status = str(
                status_payload.get("Status") or status_payload.get("status") or ""
            ).lower()
            if status == "success":
                break
            if status in {"fail", "failed"}:
                code = self._safe_text(
                    status_payload.get("Code")
                    or status_payload.get("code")
                    or "unknown"
                )
                message = self._safe_text(
                    status_payload.get("Message")
                    or status_payload.get("message")
                    or "任务处理失败"
                )
                return [], [
                    f"阿里云 OCR 任务失败（{task_id}，{code}）：{message}"
                ]
            if self.clock() - started_at >= self.timeout_seconds:
                return [], [
                    f"阿里云 OCR 轮询超时（{task_id}，{self.timeout_seconds:g} 秒）"
                ]
            self.sleep(self.poll_interval_seconds)

        layouts: list[dict[str, Any]] = []
        layout_num = 0
        while True:
            try:
                result_payload = _mapping(
                    self.client.result(
                        task_id,
                        layout_num=layout_num,
                        layout_step_size=self.layout_step_size,
                    )
                )
            except Exception as exc:
                return [], [
                    f"阿里云 OCR 结果获取失败（{task_id}，起始块 {layout_num}）："
                    f"{self._safe_error(exc)}"
                ]
            batch = result_payload.get("layouts") or []
            if not isinstance(batch, list):
                return [], [f"阿里云 OCR 返回结构异常（{task_id}）：layouts 不是列表"]
            layouts.extend(item for item in batch if isinstance(item, dict))
            if len(batch) < self.layout_step_size:
                break
            layout_num += len(batch)

        pages = layouts_to_pages(layouts)
        if not pages:
            return [], ["阿里云 OCR 返回成功但没有可用文本或表格"]
        return pages, []
