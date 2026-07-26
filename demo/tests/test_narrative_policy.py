from demo.domain.narrative_policy import (
    select_narrative_fields,
    should_create_candidate_report,
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


def test_candidate_report_requires_at_least_one_completed_review():
    assert should_create_candidate_report({}) is False
    assert should_create_candidate_report(
        {"format": {"status": "failed", "findings": []}}
    ) is False
    assert should_create_candidate_report(
        {"format": {"status": "completed", "findings": []}}
    ) is True
    assert should_create_candidate_report(
        {"data": {"status": "completed_with_issues", "findings": [{}]}}
    ) is True
