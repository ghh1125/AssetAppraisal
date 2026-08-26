"""Closed-world contract for LLM-written Word provenance comments.

The model may rewrite an existing evidence summary for readability, but it
cannot change the status, source, value, or target Word location selected by
the deterministic pipeline.
"""

from __future__ import annotations

import re
from typing import Any


TRACE_TITLES = {
    "verified": "【来源已核验】",
    "fallback": "【来源待补充核验】",
    "review": "【数据冲突，需人工复核】",
    "missing": "【未找到数据】",
}


def _without_title(value: Any) -> str:
    return re.sub(r"^【[^】]+】\s*", "", str(value or "").strip())


def build_traceability_comment_request(
    annotations: list[dict[str, Any]],
) -> tuple[dict[str, Any], dict[str, list[int]]]:
    """Group repeated table-cell annotations into compact LLM requests."""
    grouped: dict[tuple[str, str], dict[str, Any]] = {}
    indexes: dict[tuple[str, str], list[int]] = {}
    for index, item in enumerate(annotations):
        status = str(item.get("status", ""))
        if status not in TRACE_TITLES:
            continue
        field_key = str(item.get("field_key") or "unresolved_word_location")
        key = (field_key, status)
        indexes.setdefault(key, []).append(index)
        if key in grouped:
            continue
        grouped[key] = {
            "comment_id": f"trace-{len(grouped) + 1}",
            "field_name": str(item.get("field_name") or "该字段"),
            "status": status,
            "required_title": TRACE_TITLES[status],
            "evidence_summary": _without_title(item.get("comment")),
        }
    groups = list(grouped.values())
    index_map = {
        groups[position]["comment_id"]: indexes[key]
        for position, key in enumerate(grouped)
    }
    return {"version": "traceability_comments.v1", "comments": groups}, index_map


def validate_traceability_comment_response(
    payload: dict[str, Any],
    *,
    allowed_comment_ids: set[str],
) -> tuple[dict[str, str], list[str]]:
    """Accept only known IDs and non-empty human-facing comment bodies."""
    issues: list[str] = []
    result: dict[str, str] = {}
    raw = payload.get("comments", []) if isinstance(payload, dict) else []
    if not isinstance(raw, list):
        return {}, ["LLM 来源批注返回结构无效：comments 必须是数组"]
    for item in raw:
        if not isinstance(item, dict):
            issues.append("LLM 来源批注返回了非对象条目，已丢弃")
            continue
        comment_id = str(item.get("comment_id", ""))
        body = _without_title(item.get("comment"))
        if comment_id not in allowed_comment_ids:
            issues.append(f"LLM 来源批注返回未知编号，已丢弃：{comment_id}")
            continue
        if not body:
            issues.append(f"LLM 来源批注正文为空，已使用规则兜底：{comment_id}")
            continue
        if comment_id in result:
            issues.append(f"LLM 来源批注编号重复，已保留首项：{comment_id}")
            continue
        result[comment_id] = body[:500]
    return result, issues
