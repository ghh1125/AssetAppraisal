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

# Company profile is filled automatically when deterministic/API evidence is
# available. These are the six report sections the reviewer chooses in node 2.
SELECTABLE_LLM_TEMPLATE_FIELDS = LLM_TEMPLATE_FIELDS[1:]

NARRATIVE_MODULE_LABELS = {
    "industry_overview": "所处行业及行业介绍",
    "business_and_segments": "业务内容及细分市场",
    "main_products": "主要产品",
    "customers_suppliers": "主要客户及供应商",
    "profit_model_swot": "盈利模式和SWOT分析",
    "comparable_list": "对标上市公司（列表多维度展示）",
}


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
