import pytest

from demo.domain.field_validation import (
    normalize_narrative_modules,
    normalize_valuation_methods,
    validate_transaction_type,
)


def test_valuation_methods_accept_multiple_choices_and_market_method():
    assert normalize_valuation_methods(["收益法", "市场法"]) == "收益法、市场法"
    assert normalize_valuation_methods("资产评估法、收益法") == "资产基础法、收益法"


def test_valuation_methods_require_at_least_one_choice():
    with pytest.raises(ValueError, match="至少选择一种"):
        normalize_valuation_methods([])


def test_transaction_type_uses_the_four_business_options():
    assert validate_transaction_type("增资") == "增资"
    with pytest.raises(ValueError, match="委托类型"):
        validate_transaction_type("清算")


def test_narrative_modules_default_to_all_six_and_filter_unknown_values():
    assert len(normalize_narrative_modules(None)) == 6
    assert normalize_narrative_modules(["main_products", "unknown"]) == ["main_products"]


def test_narrative_modules_allow_empty_selection_before_the_candidate_checkpoint():
    assert normalize_narrative_modules([]) == []
