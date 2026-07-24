from __future__ import annotations

import re
from typing import Any


VALUATION_SUBJECT_TYPES = (
    "股东全部权益价值",
    "股东部分权益价值",
    "企业整体价值",
    "资产组价值",
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
