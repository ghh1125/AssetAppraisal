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


def compose_company_profile_narrative(
    company_profile: str | None,
    selected_modules: Mapping[str, Any],
) -> str:
    """Render selected node-2 modules into the template's one profile body.

    The approved Word template has one writable paragraph below “3、被评估
    单位概述”. Its comment lists six optional modules, but it does not contain
    six independent placeholders. Compose only the reviewer-selected
    candidates into that single body location.
    """

    sections: list[str] = []
    profile = str(company_profile or "").strip()
    if profile:
        sections.append(profile)
    for field_key in SELECTABLE_LLM_TEMPLATE_FIELDS:
        value = str(selected_modules.get(field_key, "") or "").strip()
        if not value:
            continue
        label = NARRATIVE_MODULE_LABELS[field_key]
        value = value.removeprefix(f"{label}：").removeprefix(f"{label}:").strip()
        sections.append(f"{label}：{value}")
    return "\n".join(sections)


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
