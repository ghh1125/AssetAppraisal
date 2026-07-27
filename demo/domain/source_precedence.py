from __future__ import annotations

from copy import deepcopy
from typing import Any


def _has_value(value: Any) -> bool:
    return value not in (None, "", [], {})


def prefer_semantic_result(
    *,
    fixed_value: Any,
    fixed_evidence: dict[str, Any] | None,
    semantic_value: Any,
    semantic_evidence: dict[str, Any] | None,
) -> dict[str, Any]:
    """Use an explainable semantic result before a fixed-coordinate fallback."""
    if _has_value(semantic_value):
        return {
            "value": deepcopy(semantic_value),
            "evidence": deepcopy(semantic_evidence or {}),
        }
    return {
        "value": deepcopy(fixed_value),
        "evidence": deepcopy(fixed_evidence or {}),
    }
