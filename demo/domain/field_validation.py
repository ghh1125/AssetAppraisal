from __future__ import annotations

import re
from typing import Any


VALUATION_SUBJECT_TYPES = (
    "股东全部权益价值",
    "股东部分权益价值",
    "企业整体价值",
    "资产组价值",
)

VALUATION_METHODS = ("资产基础法", "收益法", "市场法")
TRANSACTION_TYPES = ("转让", "收购", "增资", "减资")
NARRATIVE_MODULES = (
    "industry_overview",
    "business_and_segments",
    "main_products",
    "customers_suppliers",
    "profit_model_swot",
    "comparable_list",
)


def normalize_report_serial(value: Any) -> str:
    """Normalize the user-facing report serial to the template's XXX slot.

    The communication template already contains ``银信评报字（年份）第`` and
    ``号`` around the placeholder.  Users often paste a complete report number
    (for example ``苏正评报字（2025）第001号``), which used to produce a
    duplicated prefix/suffix in the generated document.  Keep only the serial
    portion for that template slot while accepting a plain ``001`` as well.
    """
    text = re.sub(r"\s+", "", str(value or ""))
    if not text:
        return ""
    if "第" in text:
        text = text.rsplit("第", 1)[1]
    text = re.sub(r"号$", "", text)
    return text


def validate_report_serial_input(value: Any) -> str:
    """Validate the image-defined non-negative integer report serial."""
    text = normalize_report_serial(value)
    if not text or not re.fullmatch(r"\d+", text):
        raise ValueError("评估报告编号流水号必须是大于等于零的整数")
    return text


def validate_required_text(value: Any, label: str, max_length: int) -> str:
    text = str(value or "").strip()
    if not text:
        raise ValueError(f"{label}不能为空")
    if len(text) > max_length:
        raise ValueError(f"{label}不能超过{max_length}个字符")
    return text


def report_number_year(value: Any) -> str:
    """Extract the report-number year when the user pasted a full number."""
    match = re.search(r"(?<!\d)((?:19|20)\d{2})(?!\d)", str(value or ""))
    return match.group(1) if match else ""


def validate_valuation_subject_type(value: Any) -> str:
    """Validate the controlled vocabulary used by the valuation-object input."""
    text = str(value or "").strip()
    if text not in VALUATION_SUBJECT_TYPES:
        allowed = "、".join(VALUATION_SUBJECT_TYPES)
        raise ValueError(f"评估对象必须是以下选项之一：{allowed}")
    return text


def normalize_valuation_methods(value: Any) -> str:
    """Normalize the image-defined multi-select valuation methods."""
    if isinstance(value, str):
        values = re.split(r"[、,，/和\s]+", value.strip()) if value.strip() else []
    elif isinstance(value, (list, tuple, set)):
        values = [str(item).strip() for item in value]
    else:
        values = []
    aliases = {"资产评估法": "资产基础法", "资产基础法": "资产基础法"}
    normalized: list[str] = []
    for item in values:
        item = aliases.get(item, item)
        if item and item not in normalized:
            normalized.append(item)
    unknown = [item for item in normalized if item not in VALUATION_METHODS]
    if unknown:
        raise ValueError("评估方法只能选择：" + "、".join(VALUATION_METHODS))
    if not normalized:
        raise ValueError("评估方法至少选择一种")
    return "、".join(normalized)


def validate_transaction_type(value: Any) -> str:
    text = str(value or "").strip()
    if text not in TRANSACTION_TYPES:
        raise ValueError("委托类型只能选择：" + "、".join(TRANSACTION_TYPES))
    return text


def validate_final_valuation_method(value: Any) -> str:
    text = str(value or "").strip()
    aliases = {"资产评估法": "资产基础法"}
    text = aliases.get(text, text)
    if text not in VALUATION_METHODS:
        raise ValueError("评估结论采用方法只能选择：" + "、".join(VALUATION_METHODS))
    return text


def normalize_narrative_modules(value: Any) -> list[str]:
    if value in (None, ""):
        return list(NARRATIVE_MODULES)
    if value == []:
        # Node 1 deliberately has no narrative choice.  Node 2 generates all
        # candidates, and the reviewer may choose none when filling Word.
        return []
    values = [value] if isinstance(value, str) else list(value) if isinstance(value, (list, tuple, set)) else []
    selected = [item for item in values if item in NARRATIVE_MODULES]
    if not selected:
        raise ValueError("主体概况模块至少选择一个")
    return selected


def require_financial_fields(
    data: dict[str, Any], required_fields: list[str]
) -> dict[str, Any]:
    missing = sorted(
        field
        for field in required_fields
        if field not in data or data[field] in (None, "", [], {})
    )
    conflicts = list(data.get("_conflicts", []))
    return {"valid": not missing and not conflicts, "missing_fields": missing, "conflicts": conflicts}


def apply_missing_field_policy(
    fields: dict[str, Any],
    evidence: dict[str, dict[str, Any]],
    required_fields: list[str],
    label: str,
) -> dict[str, Any]:
    """Keep unmatched values empty and mark their evidence as missing.

    The function is intentionally side-effect free so the same business rule
    can be reused by the CLI, web service, and a future c2m integration.
    """
    updated_fields = dict(fields)
    updated_evidence = {
        key: dict(value) if isinstance(value, dict) else value
        for key, value in evidence.items()
    }
    missing = sorted(
        field
        for field in required_fields
        if field not in updated_fields
        or updated_fields[field] in (None, "", [], {})
    )
    for field in missing:
        updated_fields[field] = ""
        if (
            not isinstance(updated_evidence.get(field), dict)
            or updated_evidence[field].get("kind")
            != "unfinished_appraisal"
        ):
            updated_evidence[field] = {
                "kind": "missing",
                "file": "",
                "locator": "指定来源未匹配到值",
            }
    return {
        "valid": not missing,
        "missing_fields": missing,
        "fields": updated_fields,
        "evidence": updated_evidence,
        "issues": [
            f"高优先级：{label}未匹配到，Word已保留黄色占位符：{field}"
            for field in missing
        ],
    }
