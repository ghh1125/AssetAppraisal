"""Closed-world request/response contract for LLM evidence review.

The model reviews an already selected value and its source candidates.  It
never receives permission to supply a new figure or change the source chosen
by the field-level reconciliation rule.
"""

from __future__ import annotations

from typing import Any

from demo.domain.source_labels import human_source_locator


REVIEW_STATUSES = frozenset({"accept", "needs_review", "conflict", "missing"})


def needs_review_fallback(
    field: dict[str, Any],
    *,
    reason: str = "LLM未返回有效复核结论，已转为人工复核。",
) -> dict[str, str]:
    """Return a closed-world review when the model response is unusable.

    The fallback deliberately preserves the already selected value and source.
    It only guarantees that every submitted evidence item has a controlled
    review status and a human-readable review action.
    """
    field_key = str(field.get("field_key", ""))
    field_name = str(field.get("field_name") or field_key or "该字段")
    selected = field.get("selected", {})
    selected = selected if isinstance(selected, dict) else {}
    source_file = str(selected.get("source_file", "")).strip()
    source_locator = str(selected.get("source_locator", "")).strip()
    if source_file and source_locator:
        source_text = f"当前值来源于《{source_file}》“{source_locator}”"
    elif source_file:
        source_text = f"当前值来源于《{source_file}》"
    elif source_locator:
        source_text = f"当前值定位为“{source_locator}”"
    else:
        source_text = "当前值缺少可核验的来源定位"
    return {
        "field_key": field_key,
        "status": "needs_review",
        "reason": reason,
        "comment": (
            f"{field_name}的自动取数复核未返回有效结论。{source_text}。"
            "请人工核对原始材料中的科目、期间、金额单位及单体/合并口径后确认。"
        ),
    }


def _review_locator(source_file: Any, locator: Any, field_name: str) -> str:
    """Expose a reviewer-facing table name instead of an internal field key."""
    return human_source_locator(str(source_file or ""), locator, field_name)


def build_evidence_review_request(
    reconciliation_rows: list[dict[str, Any]],
    field_names: dict[str, str],
) -> dict[str, Any]:
    fields: list[dict[str, Any]] = []
    for row in reconciliation_rows:
        selected = row.get("selected")
        if not isinstance(selected, dict) or selected.get("value") in (None, "", [], {}):
            continue
        field_key = str(row.get("field_key", ""))
        if not field_key:
            continue
        candidates = [
            {
                key: item.get(key, "")
                for key in ("value", "source_kind", "source_file", "source_locator")
            }
            for item in row.get("candidates", [])
            if isinstance(item, dict)
        ]
        candidate_values = {str(item.get("value", "")).strip() for item in candidates}
        selected_value = str(selected.get("value", "")).strip()
        fields.append(
            {
                "field_key": field_key,
                "field_name": field_names.get(field_key, field_key),
                "selected": {
                    key: selected.get(key, "")
                    for key in ("value", "source_kind", "source_file", "source_locator")
                },
                "candidates": candidates,
                "review_required": bool(
                    row.get("discrepancies")
                    or any(value != selected_value for value in candidate_values)
                    or (
                        isinstance(selected, dict)
                        and selected.get("source_kind") not in {"pdf_ocr", "pdf_ocr_xlsx"}
                    )
                ),
            }
        )
        # Keep the model's input readable too.  The machine field key remains
        # in the closed-world contract, but the location shown to the model is
        # a concrete worksheet/table label whenever one is available.
        fields[-1]["selected"]["source_locator"] = _review_locator(
            fields[-1]["selected"].get("source_file"),
            fields[-1]["selected"].get("source_locator"),
            fields[-1]["field_name"],
        )
        for candidate in fields[-1]["candidates"]:
            candidate["source_locator"] = _review_locator(
                candidate.get("source_file"),
                candidate.get("source_locator"),
                fields[-1]["field_name"],
            )
    return {"version": "evidence_review.v1", "fields": fields}


def validate_evidence_review_response(
    payload: dict[str, Any],
    *,
    allowed_field_keys: list[str],
) -> tuple[list[dict[str, str]], list[str]]:
    allowed = set(allowed_field_keys)
    reviews: list[dict[str, str]] = []
    issues: list[str] = []
    raw_reviews = payload.get("reviews", []) if isinstance(payload, dict) else []
    if not isinstance(raw_reviews, list):
        return [], ["LLM 取数复核返回结构无效：reviews 必须是数组"]
    seen: set[str] = set()
    for item in raw_reviews:
        if not isinstance(item, dict):
            issues.append("LLM 取数复核返回了非对象条目，已丢弃")
            continue
        field_key = str(item.get("field_key", ""))
        status = str(item.get("status", ""))
        reason = str(item.get("reason", "")).strip()
        if field_key not in allowed:
            issues.append(f"LLM 取数复核返回未授权字段，已丢弃：{field_key}")
            continue
        if status not in REVIEW_STATUSES:
            issues.append(f"LLM 取数复核字段 {field_key} 状态无效，已丢弃：{status}")
            continue
        if field_key in seen:
            issues.append(f"LLM 取数复核字段 {field_key} 重复，已保留首项")
            continue
        comment = str(item.get("comment", "")).strip()
        review = {"field_key": field_key, "status": status, "reason": reason}
        if comment:
            review["comment"] = comment
        reviews.append(review)
        seen.add(field_key)
    return reviews, issues
