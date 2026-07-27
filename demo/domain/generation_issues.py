from __future__ import annotations

import hashlib
from typing import Any


def missing_marker(location: dict[str, Any]) -> str:
    marker = str(location.get("marker", ""))
    return marker if "XX" in marker.upper() else "XXX"


def issues_for_missing_locations(
    locations: list[dict[str, Any]],
    fields: dict[str, Any],
) -> list[dict[str, Any]]:
    issues: list[dict[str, Any]] = []
    for location in locations:
        key = str(location["field_key"])
        if fields.get(key) not in (None, "", [], {}):
            continue
        location_id = str(location["location_id"])
        marker = missing_marker(location)
        digest = hashlib.sha1(
            f"{location_id}:{key}".encode("utf-8")
        ).hexdigest()[:10]
        issues.append(
            {
                "issue_id": f"GEN-{digest}",
                "priority": "高",
                "category": "missing_field",
                "page_number": "",
                "page_basis": "unavailable",
                "location_id": location_id,
                "location_type": "段落",
                "location_description": str(location.get("context", "")),
                "field_key": key,
                "field_name": str(location.get("field_name", key)),
                "current_text": marker,
                "problem": "指定来源未匹配到可用值",
                "expected_source": str(location.get("source_kind", "")),
                "source_file": str(location.get("source_file", "")),
                "source_locator": str(location.get("source_locator", "")),
                "suggestion": "补充对应材料或人工确认后替换黄色占位符",
                "status": "待人工处理",
            }
        )
    return issues
