from __future__ import annotations

import hashlib
import re
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
    mapped_by_id = {
        str(location.get("location_id", "")): location
        for location in locations
    }
    mapped_by_paragraph: dict[tuple[str, int], list[dict[str, Any]]] = {}
    for location in locations:
        location_id = str(location.get("location_id", ""))
        match = re.search(r"^(.+)-P(\d+)-[XH]\d+$", location_id)
        key = (
            str(location.get("part", ""))
            or (match.group(1) if match else ""),
            int(location.get("paragraph_index", 0))
            or (int(match.group(2)) if match else 0),
        )
        mapped_by_paragraph.setdefault(key, []).append(location)

    issues: list[dict[str, Any]] = []
    for finding in findings:
        part = str(finding.get("part", ""))
        short_part = (
            part.rsplit("/", 1)[-1].removesuffix(".xml").upper()
        )
        key = (
            short_part,
            int(finding.get("paragraph_index", 0)),
        )
        candidates = mapped_by_paragraph.get(key, [])
        occurrence = int(finding.get("occurrence_index", 1))
        mapped = mapped_by_id.get(str(finding.get("location_id", "")))
        if mapped is None:
            mapped = next(
                (
                    item
                    for item in candidates
                    if int(item.get("occurrence_index", 1))
                    == occurrence
                    if fields.get(str(item.get("field_key", "")))
                    in (None, "", [], {})
                ),
                next(
                    (
                        item
                        for item in candidates
                        if int(item.get("occurrence_index", 1))
                        == occurrence
                    ),
                    None,
                ),
            )
        field_key = (
            str(mapped.get("field_key", ""))
            if mapped
            else "word_table_placeholder"
        )
        source = evidence.get(field_key, {})
        unfinished_appraisal = (
            str(source.get("kind", "")) == "unfinished_appraisal"
        )
        location_id = str(finding["location_id"])
        digest = hashlib.sha1(location_id.encode("utf-8")).hexdigest()[:10]
        issues.append(
            {
                "issue_id": f"GEN-{digest}",
                "priority": "高",
                "category": (
                    "unfinished_appraisal"
                    if unfinished_appraisal
                    else "missing_field"
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
                "problem": (
                    "疑似尚未完成评估，评估列为空或全零"
                    if unfinished_appraisal
                    else "生成后仍存在未解析占位符"
                ),
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
                "suggestion": (
                    "确认评估工作簿是否已完成；完成后重新上传，"
                    "未完成则保持黄色 XXX"
                    if unfinished_appraisal
                    else "补充对应材料或人工确认后替换黄色占位符"
                ),
                "status": "待人工处理",
                "part": str(finding.get("part", "")),
                "paragraph_index": int(
                    finding.get("paragraph_index", 0)
                ),
            }
        )
    return issues


def organize_generation_issues(
    issues: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    """Prepare unresolved items for page-by-page business review."""
    priority_rank = {"高": 0, "中": 1, "低": 2}
    organized: list[dict[str, Any]] = []
    for issue in issues:
        item = dict(issue)
        page = item.get("page_number", "")
        description = str(
            item.get("location_description", "")
            or item.get("field_name", "")
            or item.get("location_id", "")
        )
        item["review_location"] = (
            f"第{page}页｜{description}"
            if page not in (None, "")
            else f"页码待确认｜{description}"
        )
        item["review_action"] = str(
            item.get("suggestion", "") or "人工核对并更新"
        )
        organized.append(item)

    def sort_key(item: dict[str, Any]) -> tuple[Any, ...]:
        page = item.get("page_number", "")
        try:
            page_number = int(page)
            page_missing = 0
        except (TypeError, ValueError):
            page_number = 10**9
            page_missing = 1
        return (
            page_missing,
            page_number,
            priority_rank.get(str(item.get("priority", "")), 9),
            str(item.get("location_id", "")),
        )

    return sorted(organized, key=sort_key)


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
