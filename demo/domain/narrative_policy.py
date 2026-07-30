from __future__ import annotations

from collections.abc import Mapping
from typing import Any


LLM_TEMPLATE_FIELDS = (
    "company_profile_section",
    "industry_overview",
    "business_and_segments",
    "main_products",
    "customers_suppliers",
    "profit_model_swot",
    "comparable_list",
)


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

def select_llm_candidates(
    candidates: Mapping[str, Any],
    selected_fields: list[str] | tuple[str, ...] | set[str],
) -> dict[str, Any]:
    """Return only candidate text that has a fixed LLM slot in the template.

    Every slot is optional at the fill checkpoint.  Unknown keys are
    deliberately dropped so an LLM cannot invent a Word destination.
    """
    selected = set(selected_fields)
    return {
        key: value
        for key, value in candidates.items()
        if key in LLM_TEMPLATE_FIELDS
        and value not in (None, "", [], {})
        and key in selected
    }
