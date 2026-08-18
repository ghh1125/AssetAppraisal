"""Human-readable source labels for Word comments and LLM review prompts."""

from __future__ import annotations

import re
from collections import defaultdict
from typing import Any

from demo.domain.financial_matching import normalize_label


PDF_FIELD_MARKERS: dict[str, tuple[str, ...]] = {
    "historical_balance_sheet_table": ("资产负债表", "总资产", "负债", "所有者权益", "资产总计"),
    "historical_income_statement_table": ("利润表", "营业收入", "营业利润", "净利润", "营业成本"),
    "asset_scope_summary_table": ("资产负债范围", "流动资产", "非流动资产", "资产合计", "资产总计", "负债合计"),
    "long_term_assets_table": ("长期资产", "电子设备", "无形资产", "长期待摊费用", "账面净值", "账面金额"),
    "major_long_term_assets": ("长期资产", "电子设备", "无形资产", "长期待摊费用", "账面净值", "账面金额"),
    "book_net_assets": ("资产总计", "负债合计", "所有者权益"),
    "tax_rates": ("增值税", "所得税", "税率"),
    "valuation_scope": ("评估对象", "评估范围"),
}


def _clean_name(source_file: str) -> str:
    name = re.split(r"[\\/]", str(source_file or ""))[-1]
    return name or "上传材料"


def _sheet_name(locator: str) -> str:
    text = str(locator or "")
    match = re.search(r"([^!；;，,\s]+)!", text)
    return match.group(1) if match else ""


def human_source_locator(source_file: str, locator: Any, field_name: str) -> str:
    """Render source evidence without internal keys or long raw cell lists."""
    text = str(locator or "").strip()
    display_field = str(field_name or "").strip() or "对应字段"
    if text.startswith("审计 PDF"):
        return text
    if text.startswith("第") and "页" in text:
        return f"审计 PDF {text}：{display_field}"
    sheet = _sheet_name(text)
    if sheet:
        return f"《{_clean_name(source_file)}》工作表“{sheet}”（{display_field}）"
    internal_labels = {
        "asset_scope_summary_table": "资产负债范围表",
        "long_term_assets_table": "主要长期资产账面记录表",
        "major_long_term_assets": "主要长期资产账面记录表",
        "historical_balance_sheet_table": "历史资产负债表",
        "historical_income_statement_table": "历史利润表",
        "book_net_assets": "审计后账面净资产",
    }
    for key, label in internal_labels.items():
        if text == key or text.endswith(f"/ {key}") or text.endswith(f"\\ {key}"):
            if str(source_file or "").lower().endswith(".pdf"):
                # Do not expose an internal OCR key or the former
                # “页码未定位” implementation detail to reviewers.  The
                # page resolver replaces this with a concrete page whenever
                # OCR/LLM evidence contains one.
                return f"审计 PDF：{label}"
            return f"{label}（结构化表）"
    if text:
        return text if len(text) <= 80 else f"《{_clean_name(source_file)}》（{display_field}）"
    return display_field


def locate_pdf_field_page(
    normalized: dict[str, Any] | None,
    field_key: str,
    field_name: str,
) -> str:
    """Find the likely PDF page for a field from OCR text/table markers."""
    page_from_evidence = _page_from_financial_evidence(normalized or {}, field_key)
    if page_from_evidence:
        return f"审计 PDF 第{page_from_evidence}页：{field_name}"
    markers = PDF_FIELD_MARKERS.get(field_key, ())
    if not markers:
        return f"审计 PDF：{field_name}"
    page_text: dict[int, list[str]] = defaultdict(list)
    data = normalized or {}
    for block in data.get("text_blocks", []) or []:
        if not isinstance(block, dict):
            continue
        page = block.get("page_number") or block.get("page")
        try:
            page_number = int(page)
        except (TypeError, ValueError):
            continue
        page_text[page_number].append(str(block.get("text", "")))
    for cell in data.get("table_cells", []) or []:
        if not isinstance(cell, dict):
            continue
        page = cell.get("page_number") or cell.get("page")
        try:
            page_number = int(page)
        except (TypeError, ValueError):
            continue
        page_text[page_number].append(str(cell.get("text", "")))
    if not page_text:
        return f"审计 PDF：{field_name}"
    normalized_markers = [normalize_label(item) for item in markers]
    best_page = 0
    best_score = 0
    for page_number, parts in page_text.items():
        text = normalize_label("".join(parts))
        score = sum(1 for marker in normalized_markers if marker and marker in text)
        if score > best_score:
            best_score = score
            best_page = page_number
    # A single strong table marker is enough when the OCR table itself has
    # multiple cells/rows (for example a long-term-assets table whose title
    # was split across OCR cells).  Requiring two markers caused valid PDF
    # pages to be shown as unlocated.
    if best_page and best_score >= 1:
        return f"审计 PDF 第{best_page}页：{field_name}"
    return f"审计 PDF：{field_name}"


def _page_from_financial_evidence(data: dict[str, Any], field_key: str) -> int:
    table_pages: dict[str, int] = {}
    for cell in data.get("table_cells", []) or []:
        if not isinstance(cell, dict):
            continue
        table_id = str(cell.get("table_id") or "")
        if not table_id:
            continue
        try:
            page = int(cell.get("page_number") or cell.get("page") or 0)
        except (TypeError, ValueError):
            continue
        if page > 0:
            table_pages[table_id] = page
    for record in data.get("financial_data", []) or []:
        if not isinstance(record, dict) or str(record.get("field_key") or "") != field_key:
            continue
        evidence_id = str(record.get("evidence_id") or "")
        page_number = record.get("page_number") or record.get("page")
        try:
            direct_page = int(page_number or 0)
        except (TypeError, ValueError):
            direct_page = 0
        if direct_page > 0:
            return direct_page
        for table_id, page in table_pages.items():
            if table_id and table_id in evidence_id:
                return page
        match = re.search(r"(?:^|[^\d])p(?:age)?[^\d]?(\d{1,4})", evidence_id, re.IGNORECASE)
        if match:
            return int(match.group(1))
    return 0


def pdf_field_locators(
    normalized: dict[str, Any] | None,
    field_keys: set[str],
    field_names: dict[str, str],
) -> dict[str, str]:
    return {
        field_key: locate_pdf_field_page(
            normalized,
            field_key,
            field_names.get(field_key, field_key),
        )
        for field_key in field_keys
    }
