"""Closed-world contract for LLM-assisted PDF page localization."""

from __future__ import annotations

from collections import defaultdict
from typing import Any


def _page_number(item: dict[str, Any]) -> int:
    try:
        return int(item.get("page_number") or item.get("page") or 0)
    except (TypeError, ValueError):
        return 0


def build_pdf_page_locator_request(
    normalized: dict[str, Any],
    field_names: dict[str, str],
    field_keys: list[str],
) -> dict[str, Any]:
    """Build a compact page-numbered OCR request for the locator model."""
    pages: dict[int, list[str]] = defaultdict(list)
    for item in [*(normalized.get("text_blocks", []) or []), *(normalized.get("table_cells", []) or [])]:
        if not isinstance(item, dict):
            continue
        page = _page_number(item)
        if page <= 0:
            continue
        text = str(item.get("text") or "").strip()
        if text:
            pages[page].append(text)
    return {
        "version": "pdf_page_locator.v1",
        "fields": [
            {"field_key": key, "field_name": field_names.get(key, key)}
            for key in field_keys
            if key
        ],
        "pages": [
            {
                "page_number": page,
                "page_label": f"第{page}页",
                "ocr_text": " ".join(parts)[:5000],
            }
            for page, parts in sorted(pages.items())
        ],
    }


def validate_pdf_page_locator_response(
    payload: dict[str, Any],
    *,
    allowed_field_keys: list[str],
    available_pages: set[int],
) -> tuple[dict[str, int], list[str]]:
    allowed = set(allowed_field_keys)
    locations: dict[str, int] = {}
    issues: list[str] = []
    raw = payload.get("locations", []) if isinstance(payload, dict) else []
    if not isinstance(raw, list):
        return {}, ["PDF页码定位返回结构无效：locations 必须是数组"]
    for item in raw:
        if not isinstance(item, dict):
            issues.append("PDF页码定位返回了非对象条目，已丢弃")
            continue
        key = str(item.get("field_key") or "")
        if key not in allowed:
            issues.append(f"PDF页码定位返回未授权字段，已丢弃：{key}")
            continue
        try:
            page = int(item.get("page_number") or 0)
        except (TypeError, ValueError):
            page = 0
        if page not in available_pages:
            issues.append(f"PDF页码定位返回不存在的页码，已丢弃：{key}={page}")
            continue
        if key in locations:
            issues.append(f"PDF页码定位字段重复，已保留首项：{key}")
            continue
        locations[key] = page
    return locations, issues
