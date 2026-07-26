from __future__ import annotations

from collections.abc import Mapping
from typing import Any


def select_narrative_fields(
    routed_fields: set[str],
    selected_modules: list[str],
) -> set[str]:
    """Keep the mandatory profile plus only user-selected routed modules."""
    selected = set(selected_modules)
    return {
        field_key
        for field_key in routed_fields
        if field_key == "company_profile_section" or field_key in selected
    }


def should_create_candidate_report(
    reviews: Mapping[str, Mapping[str, Any]],
    *,
    financial_fields_complete: bool = True,
) -> bool:
    """Create a candidate only after review and complete financial inputs."""
    return financial_fields_complete and any(
        review.get("status") in {"completed", "completed_with_issues"}
        for review in reviews.values()
    )
