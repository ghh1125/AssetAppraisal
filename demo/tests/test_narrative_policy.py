from demo.domain.narrative_policy import (
    select_narrative_fields,
)


def test_company_profile_is_always_allowed_and_only_selected_modules_are_added():
    allowed = select_narrative_fields(
        {"company_profile_section", "industry_overview", "main_products"},
        ["main_products"],
    )

    assert allowed == {"company_profile_section", "main_products"}


def test_unrouted_selected_module_is_not_added():
    allowed = select_narrative_fields(
        {"company_profile_section", "main_products"},
        ["industry_overview"],
    )

    assert allowed == {"company_profile_section"}
