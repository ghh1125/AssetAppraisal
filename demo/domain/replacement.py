from __future__ import annotations

from decimal import Decimal, InvalidOperation
from typing import Any
import re

from .calculations import flexible_date_parts
from .field_validation import normalize_report_serial
from .generation_issues import missing_marker


PLACEHOLDER = re.compile(r"X{2,}", re.I)


def _trim_repeated_suffix(item: dict[str, Any], value: Any) -> Any:
    if not isinstance(value, str) or str(value).startswith("【待人工补充："):
        return value
    occurrence_match = re.search(r"-X(\d+)$", item.get("location_id", ""))
    if not occurrence_match:
        return value
    markers = list(PLACEHOLDER.finditer(item.get("context", "")))
    occurrence = int(occurrence_match.group(1))
    if occurrence > len(markers):
        return value
    tail = item.get("context", "")[markers[occurrence - 1].end() :]
    tail = re.sub(r"^(?:[（(][^）)]*[）)])+", "", tail)
    if tail.startswith(("有限责任公司", "有限公司")):
        for suffix in ("有限责任公司", "有限公司"):
            if value.endswith(suffix):
                return value[: -len(suffix)]
    for suffix in ("价值", "法"):
        if tail.startswith(suffix) and value.endswith(suffix):
            return value[: -len(suffix)]
    return value


def _display_value(item: dict[str, Any], value: Any) -> Any:
    key = item["field_key"]
    if str(value).startswith("【待人工补充："):
        return value
    if key == "report_serial":
        return normalize_report_serial(value)
    part = next((name for name in ("year", "month", "day") if key.endswith(f"_{name}")), None)
    if part is None:
        if "20XX" in item.get("context", "") and re.fullmatch(r"(?:19|20)\d{2}", str(value)):
            return str(value)[-2:]
        if item.get("unit_scope") in {"万元", "%"}:
            try:
                number = Decimal(str(value).replace(",", ""))
            except InvalidOperation:
                return value
            if item.get("unit_scope") == "万元":
                return f"{number:,.2f}"
            formatted = f"{number:.2f}"
            # Some annotated templates put the percent sign after the
            # placeholder, while newer review comments require the mapper to
            # supply it.  Keep exactly one sign in the final document.
            marker = str(item.get("marker", ""))
            context = str(item.get("context", ""))
            marker_pos = context.find(marker) if marker else -1
            suffix = context[marker_pos + len(marker):] if marker_pos >= 0 else ""
            return formatted if suffix.lstrip().startswith("%") else f"{formatted}%"
        return value
    parts = flexible_date_parts(value)
    if parts:
        value = parts[part]
    elif re.fullmatch(r"\d+(?:\.0+)?", str(value)):
        value = int(float(value))
    else:
        return value
    if part == "year" and "20XX" in item.get("context", ""):
        return f"{int(value) % 100:02d}"
    if part in {"month", "day"}:
        return f"{int(value):02d}"
    return str(value)


def build_replacements(locations: list[dict[str, Any]], fields: dict[str, Any]) -> dict[str, str]:
    result = {}
    for item in locations:
        if item.get("force_unresolved"):
            result[item["location_id"]] = missing_marker(item)
            continue
        value = fields.get(item["field_key"])
        if value in (None, "", []):
            value = missing_marker(item)
        if isinstance(value, dict) and "caption" in value:
            value = value["caption"]
        value = _display_value(item, value)
        value = _trim_repeated_suffix(item, value)
        if isinstance(value, (dict, list)):
            if isinstance(value, list):
                value = "；".join(str(x) for x in value)
            else:
                value = "；".join(f"{k}：{v}" for k, v in value.items())
        result[item["location_id"]] = str(value)
    return result
