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


def issues_from_word_findings(
    findings: list[dict[str, Any]],
    locations: list[dict[str, Any]],
    fields: dict[str, Any],
    evidence: dict[str, dict[str, Any]],
) -> list[dict[str, Any]]:
    mapped_by_paragraph: dict[tuple[str, int], list[dict[str, Any]]] = {}
    for location in locations:
        key = (
            str(location.get("part", "")),
            int(location.get("paragraph_index", 0)),
        )
        mapped_by_paragraph.setdefault(key, []).append(location)

    issues: list[dict[str, Any]] = []
    for finding in findings:
        key = (
            str(finding.get("part", "")),
            int(finding.get("paragraph_index", 0)),
        )
        candidates = mapped_by_paragraph.get(key, [])
        mapped = next(
            (
                item
                for item in candidates
                if fields.get(str(item.get("field_key", "")))
                in (None, "", [], {})
            ),
            candidates[0] if candidates else None,
        )
        field_key = (
            str(mapped.get("field_key", ""))
            if mapped
            else "word_table_placeholder"
        )
        source = evidence.get(field_key, {})
        location_id = str(finding["location_id"])
        digest = hashlib.sha1(location_id.encode("utf-8")).hexdigest()[:10]
        issues.append(
            {
                "issue_id": f"GEN-{digest}",
                "priority": "高",
                "category": (
                    "missing_field"
                    if mapped
                    else "unmapped_placeholder"
                ),
                "page_number": "",
                "page_basis": "unavailable",
                "location_id": location_id,
                "location_type": str(
                    finding.get("location_type", "段落")
                ),
                "location_description": str(
                    finding.get("context", "")
                ),
                "field_key": field_key,
                "field_name": (
                    str(mapped.get("field_name", field_key))
                    if mapped
                    else "未映射的 Word 占位符"
                ),
                "current_text": str(
                    finding.get("current_text", "XXX")
                ),
                "problem": "生成后仍存在未解析占位符",
                "expected_source": (
                    str(mapped.get("source_kind", ""))
                    if mapped
                    else "项目配置、上传材料或人工输入"
                ),
                "source_file": str(source.get("file", "")),
                "source_locator": str(
                    source.get("locator", "")
                    or (
                        mapped.get("source_locator", "")
                        if mapped
                        else ""
                    )
                ),
                "suggestion": "补充对应材料或人工确认后替换黄色占位符",
                "status": "待人工处理",
                "part": str(finding.get("part", "")),
                "paragraph_index": int(
                    finding.get("paragraph_index", 0)
                ),
            }
        )
    return issues


def apply_page_locations(
    issues: list[dict[str, Any]],
    generated_pages: dict[str, int | str],
    template_pages: dict[str, int | str],
) -> list[dict[str, Any]]:
    result: list[dict[str, Any]] = []
    for issue in issues:
        item = dict(issue)
        location_id = str(item.get("location_id", ""))
        if location_id in generated_pages:
            item["page_number"] = generated_pages[location_id]
            item["page_basis"] = "generated_report"
        elif location_id in template_pages:
            item["page_number"] = template_pages[location_id]
            item["page_basis"] = "template"
        else:
            item["page_number"] = ""
            item["page_basis"] = "unavailable"
        item.pop("part", None)
        item.pop("paragraph_index", None)
        result.append(item)
    return result
