from demo.domain.financial_table_semantics import (
    canonical_long_term_asset_category,
    canonical_period,
    choose_historical_columns,
    detail_header_role,
)


def test_canonical_period_accepts_dates_and_rejects_rates():
    assert canonical_period("2025年6月30日") == (2025, 6, 30, "2025年6月30日")
    assert canonical_period("2024年度") == (2024, 12, 31, "2024年度")
    assert canonical_period("期末数") == (None, None, None, "期末数")
    assert canonical_period("基准日审定数") == (
        None,
        None,
        None,
        "基准日审定数",
    )
    assert canonical_period("增长率%") is None


def test_historical_columns_prefer_actual_periods_over_growth_and_forecast():
    headers = {
        2: ["2023年度", "实际数"],
        3: ["2024年度", "审定数"],
        4: ["2024年度", "同比增长率"],
        5: ["2025年度", "预测数"],
    }

    assert choose_historical_columns(headers, valuation_year=2024) == [2, 3]


def test_historical_columns_keep_two_sided_period_columns_in_sheet_order():
    headers = {
        2: ["期初数"],
        3: ["期末数"],
        4: ["序号"],
        6: ["期初数"],
        7: ["期末数"],
    }

    assert choose_historical_columns(headers, candidate_columns=[6, 7]) == [6, 7]


def test_historical_columns_keep_relative_valuation_period_after_dated_periods():
    headers = {
        2: ["2023年"],
        3: ["2024年"],
        4: ["评估基准期"],
    }

    assert choose_historical_columns(headers) == [2, 3, 4]


def test_long_term_asset_aliases_and_detail_headers_are_normalized():
    assert canonical_long_term_asset_category("办公电子设备") == "电子设备"
    assert canonical_long_term_asset_category("电子设备类") == "电子设备"
    assert canonical_long_term_asset_category("机器设备") is None
    assert detail_header_role("资产编号") == "asset_id"
    assert detail_header_role("资产类别") == "category"
    assert detail_header_role("账面净值") == "book_net"
    assert detail_header_role("帐面净值") == "book_net"
    assert detail_header_role("评估价值") == "appraised"
