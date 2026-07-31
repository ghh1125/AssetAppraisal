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


def infer_comment_source(comment_texts: Iterable[str]) -> dict[str, str]:
    """Classify the source instruction written in the latest Word comments.

    The annotation is the business source-of-truth.  This intentionally does
    not infer a provider from a filename or a paragraph number.  A comment can
    name more than one source (for example, API facts used as LLM evidence),
    so ``source_kind`` remains ``mixed`` in that case and the full instruction
    is retained for the audit trail.
    """
    instruction = " | ".join(str(item or "").strip() for item in comment_texts if str(item or "").strip())
    if not instruction:
        return {"source_kind": "", "source_instruction": ""}
    if any(token in instruction for token in ("暂时不做填充", "获取不到", "未确认", "无填充时高亮")):
        return {"source_kind": "unresolved_manual", "source_instruction": instruction}
    if "已有前序数据插入" in instruction:
        return {"source_kind": "derived", "source_instruction": instruction}
    has_manual = "人工基础信息" in instruction or "人工输入" in instruction or "用户" in instruction
    has_qcc = "企查查" in instruction or "天眼查" in instruction or "API" in instruction
    has_material = any(token in instruction for token in ("PDF", "财务表格", "表格", "审计报告", "Excel"))
    has_llm = "大模型" in instruction or "LLM" in instruction
    has_system = "系统时间" in instruction or "系统日期" in instruction or "报告生成" in instruction
    sources = sum(bool(item) for item in (has_manual, has_qcc, has_material, has_llm, has_system))
    if sources > 1:
        kind = "mixed"
    elif has_qcc:
        kind = "qichacha_api"
    elif has_llm:
        kind = "bailian_glm"
    elif has_material:
        kind = "pdf_ocr_xlsx"
    elif has_manual:
        kind = "node_input"
    elif has_system:
        kind = "system"
    else:
        kind = "unclassified"
    return {"source_kind": kind, "source_instruction": instruction}


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
    if "模块信息" in text or "大模型" in text:
        return ["company_profile_text"]
    if ("评估结论采用方法" in text or "评估结论方法采用" in text) and "金额数据" in text:
        return [
            # The first marker is in “被评估单位XXX价值”; the comment
            # describes the source of the following result, not that marker.
            "valuation_subject_type", "final_appraisal_value",
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
    if "已有前序数据插入" in text:
        return [
            "final_valuation_method", "final_appraisal_value", "final_value_chinese_wan",
            "appraisal_increment", "appraisal_increment_rate",
        ]
    return []


def _comment_context_field(item: dict[str, Any], source_kind: str) -> str | None:
    """Map comment-only anchors (headings/table lead-ins) to a business key."""
    context = str(item.get("context", ""))
    instruction = " ".join(str(text) for text in item.get("comment_texts", []))
    combined = f"{context} {instruction}"
    if source_kind == "qichacha_api":
        if "委托方" in combined:
            return "commissioning_party_profile"
        if "评估主体" in combined:
            return "target_company_profile"
        if "股东及股权" in combined or "股东" in context:
            return "ownership_at_valuation_date"
        if "软件著作权" in combined:
            return "software_copyrights"
        if "账外无形资产" in combined:
            return "unrecorded_intangibles"
    if source_kind == "pdf_ocr_xlsx":
        if "评估对象和评估范围" in context or "评估范围" in combined:
            return "valuation_scope"
        if "长期资产" in combined or "账面记录情况" in context:
            return "major_long_term_assets"
        if "资产基础法估值" in combined:
            return "asset_approach_result_section"
    return None


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
        comment_texts = item.get("comment_texts", [])
        marker_count = len(list(re.finditer(r"X{2,}", str(item.get("context", "")), re.I)))
        source_comment_texts = comment_texts
        if len(comment_texts) == marker_count and marker_count > 1:
            source_comment_texts = [
                comment_texts[min(max(int(item.get("occurrence_index", 1)) - 1, 0), len(comment_texts) - 1)]
            ]
        comment_source = infer_comment_source(source_comment_texts)
        selected["comment_source_kind"] = comment_source["source_kind"]
        selected["comment_source_instruction"] = comment_source["source_instruction"]
        # For an explicit manual-input comment, the annotation must override
        # any stale Excel locator retained in the legacy mapping.  Other
        # source classes stay as lineage metadata because a PDF comment may
        # legitimately be supported by an OCR workbook or an Excel cross-check.
        if comment_source["source_kind"] == "node_input":
            selected["source_kind"] = "人工输入"
            selected["source_file"] = "人工基础信息"
            selected["source_locator"] = "人工基础信息"
        unresolved_positions = [
            index
            for index, text in enumerate(source_comment_texts, 1)
            if any(token in str(text) for token in ("暂时不做填充", "获取不到", "未确认", "无填充时高亮"))
        ]
        # A paragraph can contain two separately anchored comments.  Word's
        # inventory exposes them in anchor order; apply an unresolved comment
        # only to its corresponding marker (not to every marker in the
        # paragraph, e.g. P0481's first XX is unresolved while the second is
        # the selected method).
        if unresolved_positions and (
            len(source_comment_texts) == 1
            or int(item.get("occurrence_index", 1)) <= unresolved_positions[0]
        ):
            selected["force_unresolved"] = True
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
        contextual_key = _comment_context_field(item, comment_source["source_kind"])
        if contextual_key and item.get("record_type") == "批注内容块":
            # A comment-only heading/table lead-in has no marker occurrence;
            # its source instruction is authoritative over any legacy context
            # candidate paired by normalized paragraph text.
            selected["field_key"] = contextual_key
        elif not selected.get("field_key") and contextual_key:
            selected["field_key"] = contextual_key
        result.append(selected)
    return result
