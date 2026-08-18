"""Guardrails for the LLM-assisted Word mapping agent.

The agent is deliberately a *selector*, not a writer: it may associate an
unmapped Word location with one of the already approved business fields, but
may not create a field, alter a value, or change a comment-declared source.
"""

from __future__ import annotations

from typing import Any, Iterable


def mapping_agent_request(
    locations: Iterable[dict[str, Any]],
    allowed_fields: Iterable[str],
) -> dict[str, Any]:
    fields = sorted({str(item) for item in allowed_fields if str(item)})
    return {
        "allowed_field_keys": fields,
        "locations": [
            {
                "location_id": str(item.get("location_id", "")),
                "context": str(item.get("context", ""))[:800],
                "marker": str(item.get("marker", "")),
                "comment_instruction": " | ".join(
                    str(comment) for comment in item.get("comment_texts", [])
                )[:1200],
            }
            for item in locations
            if not str(item.get("field_key", ""))
        ],
    }


def validate_mapping_agent_response(
    payload: Any,
    *,
    unresolved_location_ids: Iterable[str],
    allowed_fields: Iterable[str],
) -> tuple[dict[str, str], list[str]]:
    """Accept only closed-world ``location_id -> field_key`` selections."""
    unresolved = {str(item) for item in unresolved_location_ids}
    fields = {str(item) for item in allowed_fields}
    raw = payload.get("mappings", payload) if isinstance(payload, dict) else {}
    if not isinstance(raw, dict):
        return {}, ["LLM 映射代理返回结构无效，已忽略"]
    accepted: dict[str, str] = {}
    issues: list[str] = []
    for location_id, field_key in raw.items():
        location_id, field_key = str(location_id), str(field_key)
        if location_id not in unresolved:
            issues.append(f"LLM 映射代理尝试覆盖已确认位置，已忽略：{location_id}")
            continue
        if field_key not in fields:
            issues.append(f"LLM 映射代理返回未授权字段，已忽略：{location_id}->{field_key}")
            continue
        accepted[location_id] = field_key
    return accepted, issues
