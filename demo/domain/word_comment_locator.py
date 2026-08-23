"""Closed-world validation for LLM-assisted Word comment placement."""

from __future__ import annotations

from typing import Any, Iterable


def build_word_comment_locator_request(
    review_item: dict[str, Any],
    candidates: Iterable[dict[str, Any]],
    recommended_candidate_ids: Iterable[str] = (),
) -> dict[str, Any]:
    """Expose only real Word positions; the model cannot invent a location."""
    candidate_rows = []
    for item in candidates:
        candidate_id = str(item.get("candidate_id", "")).strip()
        if not candidate_id:
            continue
        candidate_rows.append(
            {
                "candidate_id": candidate_id,
                "part": str(item.get("part", "")),
                "paragraph_text": str(item.get("paragraph_text", ""))[:800],
                "table_index": item.get("table_index", ""),
                "row_index": item.get("row_index", ""),
                "column_index": item.get("column_index", ""),
                "row_text": str(item.get("row_text", ""))[:500],
                "period_text": str(item.get("period_text", ""))[:300],
                "table_context": str(item.get("table_context", ""))[:500],
                "matched_value": str(item.get("matched_value", "")),
            }
        )
    allowed_ids = {item["candidate_id"] for item in candidate_rows}
    recommended = [
        str(item)
        for item in recommended_candidate_ids
        if str(item) in allowed_ids
    ]
    return {
        "review_target": {
            "field_key": str(review_item.get("field_key", "")),
            "field_name": str(review_item.get("field_name", "")),
            "target_value": str(
                review_item.get("excel_value")
                if review_item.get("review_kind") in {"excel_fallback", "llm_review"}
                else review_item.get("pdf_value", "")
            ),
            "expected_table_index": review_item.get("word_table_index", ""),
            "expected_row_label": str(review_item.get("word_context_hint", "")),
            "expected_period_label": str(review_item.get("word_period_hint", "")),
            "mapped_word_anchors": review_item.get("word_anchor_hints", []),
        },
        "recommended_candidate_ids": recommended,
        "candidates": candidate_rows,
    }


def validate_word_comment_locator_response(
    payload: Any,
    *,
    allowed_candidate_ids: Iterable[str],
) -> str | None:
    """Accept exactly one ID from the supplied candidate set, or no match."""
    allowed = {str(item) for item in allowed_candidate_ids if str(item)}
    if not isinstance(payload, dict):
        return None
    candidate_id = str(payload.get("candidate_id", "")).strip()
    if not candidate_id:
        # A valid empty choice means the model actively rejected every real
        # candidate.  Keep it distinct from a failed/unparseable request.
        return ""
    if candidate_id not in allowed:
        return None
    return candidate_id
