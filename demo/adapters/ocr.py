from __future__ import annotations

from pathlib import Path
from typing import Any, Callable


class OcrAdapter:
    def __init__(self, engine: Callable[[Path], list[dict[str, Any]]] | None = None):
        self.engine = engine

    def extract(self, pdf_path: Path) -> tuple[list[dict[str, Any]], list[str]]:
        if self.engine is None:
            return [], ["OCR 未配置，相关字段按规则留空"]
        try:
            return self.engine(pdf_path), []
        except Exception as exc:
            return [], [f"OCR 失败：{exc}"]


def create_paddle_engine(language: str = "ch"):
    from demo.adapters.paddle_ocr import PaddleStructureOcrAdapter, create_local_pipeline

    adapter = PaddleStructureOcrAdapter(create_local_pipeline(lang=language))

    def extract(path: Path) -> list[dict[str, Any]]:
        pages, issues = adapter.extract(path)
        if issues:
            raise RuntimeError("；".join(issues))
        return pages

    return extract
