"""Map comment-annotated Word locations to the existing field contract.

The communication template originally used yellow runs as instructions.  The
new template keeps the same placeholders but moves the instructions into Word
comments.  This module is deliberately pure: it only consumes inventory
records and the validated legacy mapping, so c2m can replace the file adapter
without changing business rules.
"""

from __future__ import annotations

import re
from copy import deepcopy
from typing import Any, Iterable


def _location_parts(location_id: str) -> tuple[int, int] | None:
    match = re.search(r"-P(\d+)-X(\d+)$", str(location_id))
    return (int(match.group(1)), int(match.group(2))) if match else None


def _normalize_context(value: Any) -> str:
    text = str(value or "")
    text = re.sub(r"[（(][^）)]*[）)]", "", text)
    return re.sub(r"[^\u4e00-\u9fffA-Za-z0-9X%]", "", text)


def _field_from_comment(text: str) -> str | None:
    if not text or "暂时不做填充" in text or "未确认" in text or "不确定" in text:
        return None
    if "评估结论采用方法" in text:
        return "final_valuation_method"
    if "评估方法" in text:
        return "selected_valuation_method"
    if "报告生成" in text or "系统时间" in text or "报告日期" in text:
        return "report_date"
    if "报告编号" in text or "流水号" in text:
        return "report_serial"
    if "委托类型" in text:
        return "transaction_type"
    if "委托方简称" in text:
        return "commissioning_party_short_name"
    if "委托方" in text:
        return "commissioning_party_name"
    if "评估主体简称" in text:
        return "target_company_short_name"
    if "评估主体" in text or "被评估公司" in text:
        return "target_company_name"
    if "评估对象" in text:
        return "valuation_subject_type"
    if "所有者权益" in text or "净资产" in text:
        return "book_net_assets"
    if "增值率" in text:
        return "appraisal_increment_rate"
    if "增值" in text:
        return "appraisal_increment"
    if "收益法" in text and "评估值" in text:
        return "income_approach_value"
    if "资产基础法" in text and "估值" in text:
        return "asset_approach_value"
    if "审计报告文件名称" in text:
        return "audit_report_name"
    if "资产负债状况" in text:
        return "historical_balance_sheet_table"
    if "经营状况" in text or "利润表" in text:
        return "historical_income_statement_table"
    if "长期资产" in text or "货币资金" in text:
        return "major_long_term_assets"
    if "软件著作权" in text:
        return "software_copyrights"
    if "账外无形资产" in text:
        return "unrecorded_intangibles"
    if "税率" in text:
        return "tax_rates"
    if "股权结构" in text or "历史沿革" in text:
        return "ownership_history"
    if "股东及股权" in text:
        return "ownership_at_valuation_date"
    if "评估范围" in text:
        return "valuation_scope"
    if "模块信息" in text or "大模型" in text:
        return "company_profile_text"
    return None


def comment_field_candidates(comment_texts: Iterable[str]) -> list[str]:
    result: list[str] = []
    for text in comment_texts:
        field = _field_from_comment(text)
        if field and field not in result:
            result.append(field)
    return result


def _comment_field_sequence(text: str) -> list[str]:
    """Return an ordered field sequence when a comment describes a group.

    Word comments are attached to a paragraph, not to an individual ``XXX``.
    For those paragraphs, the placeholder order is the only reliable way to
    distinguish repeated values.  These rules use the explicit wording in the
    annotation (not filenames or coordinates) and are intentionally narrow.
    """
    text = str(text or "")
    if ("评估结论采用方法" in text or "评估结论方法采用" in text) and "金额数据" in text:
        return [
            "final_valuation_method", "final_appraisal_value",
            "final_value_chinese", "book_net_assets", "appraisal_increment",
            "appraisal_increment_rate",
        ]
    if "评估主体全称" in text and ("所有者权益" in text or "净资产" in text) and "收益法" in text:
        return [
            "target_company_name", "book_net_assets", "income_approach_value",
            "income_increment", "income_increment_rate",
        ]
    if "评估主体全称" in text and ("所有者权益" in text or "净资产" in text) and "资产基础法" in text:
        return [
            "target_company_name", "book_net_assets", "asset_approach_value",
            "asset_increment", "asset_increment_rate",
        ]
    if "报告生成" in text or "系统时间" in text or "报告日期" in text:
        return ["report_date_year", "report_date_month", "report_date_day"]
    if "评估基准日~" in text or "评估基准日加一年" in text:
        return [
            "validity_start_year", "validity_start_month", "validity_start_day",
            "validity_end_year", "validity_end_month", "validity_end_day",
        ]
    if "PDF/财务表格识别获取基准日" in text:
        return [
            "target_company_name", "transaction_type",
            "validity_start_year", "validity_start_month", "validity_start_day",
            "validity_end_year", "validity_end_month", "validity_end_day",
        ]
    if "评估方法" in text and "评估对象" in text and "评估基准日" in text:
        return [
            "selected_valuation_method", "target_company_name", "transaction_type",
            "valuation_subject_type", "valuation_date_year", "valuation_date_month",
            "valuation_date_day",
        ]
    return []


def build_comment_aware_locations(
    template_locations: list[dict[str, Any]],
    base_locations: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    """Return the same mapping contract with IDs from the current template.

    Exact paragraph/occurrence matches are preferred.  When a comment-based
    template removed the old yellow helper text and shifted paragraph numbers,
    normalized context matching recovers the old field definition.  Comment
    text is the final fallback and never overwrites a stronger exact mapping.
    """
    if not any(item.get("comment_texts") for item in template_locations):
        return base_locations
    by_position: dict[tuple[int, int], list[dict[str, Any]]] = {}
    by_context: dict[str, list[dict[str, Any]]] = {}
    context_counts: dict[str, int] = {}
    for item in base_locations:
        parts = _location_parts(str(item.get("location_id", "")))
        if parts:
            by_position.setdefault(parts, []).append(item)
        by_context.setdefault(_normalize_context(item.get("context", "")), []).append(item)
    for item in template_locations:
        normalized = _normalize_context(item.get("context", ""))
        context_counts[normalized] = context_counts.get(normalized, 0) + 1
    result: list[dict[str, Any]] = []
    for item in template_locations:
        parts = _location_parts(str(item.get("location_id", "")))
        candidates = by_position.get(parts, []) if parts else []
        normalized = _normalize_context(item.get("context", ""))
        if not candidates:
            context_candidates = by_context.get(normalized, [])
            # Paragraph numbers can shift after a comment is added/removed.
            # Pair placeholders by their occurrence within the same semantic
            # paragraph instead of reusing the first field for every XXX.
            current_occurrence = int(item.get("occurrence_index", 1))
            candidates = [
                candidate for candidate in context_candidates
                if (
                    int(candidate.get("occurrence_index", 1))
                    if candidate.get("occurrence_index") is not None
                    else (_location_parts(str(candidate.get("location_id", ""))) or (0, 1))[1]
                ) == current_occurrence
            ] or context_candidates[:1]
        selected = deepcopy(candidates[0]) if candidates else {
            "field_key": "",
            "field_name": "",
            "source_kind": "",
            "source_file": "",
            "source_locator": "",
            "record_type": item.get("record_type", "占位符"),
        }
        selected.update({
            "location_id": item["location_id"],
            "context": item.get("context", selected.get("context", "")),
            "marker": item.get("marker", selected.get("marker", "XXX")),
            "record_type": item.get("record_type", selected.get("record_type", "占位符")),
            "part": item.get("part", selected.get("part", "word/document.xml")),
            "paragraph_index": item.get("paragraph_index", selected.get("paragraph_index", 0)),
            "occurrence_index": item.get("occurrence_index", selected.get("occurrence_index", 1)),
            "in_table": item.get("in_table", selected.get("in_table", False)),
            "comment_ids": item.get("comment_ids", []),
            "comment_texts": item.get("comment_texts", []),
        })
        comment_text = " ".join(item.get("comment_texts", []))
        comment_keys = _comment_field_sequence(comment_text)
        if comment_keys:
            occurrence = int(item.get("occurrence_index", 1))
            selected["field_key"] = comment_keys[min(max(occurrence - 1, 0), len(comment_keys) - 1)]
        else:
            comment_keys = comment_field_candidates(item.get("comment_texts", []))
            # A single annotated placeholder is authoritative even when the
            # legacy mapping had a stale field for that paragraph.  For a
            # multi-placeholder paragraph, retain the ordered base mapping
            # unless a dedicated sequence rule above applies.
            if len(comment_keys) == 1 and context_counts.get(normalized, 0) == 1:
                selected["field_key"] = comment_keys[0]
        if not comment_keys and "评估基准日为20XX年XX月XX日" in str(item.get("context", "")):
            occurrence = int(item.get("occurrence_index", 1))
            comment_keys = [
                ["valuation_date_year", "valuation_date_month", "valuation_date_day"][
                    min(max(occurrence - 1, 0), 2)
                ]
            ]
        if not selected.get("field_key") and comment_keys:
            selected["field_key"] = comment_keys[0]
        # A comment can describe several placeholders in one paragraph.  If
        # the base mapping supplied an ordered sequence, retain it; otherwise
        # use the corresponding unique comment key.
        if comment_keys and len(comment_keys) == 1 and not candidates:
            selected["field_key"] = comment_keys[0]
        result.append(selected)
    return result
