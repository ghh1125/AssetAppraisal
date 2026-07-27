from demo.domain.historical_table_merge import merge_historical_tables


def _table(headers, rows):
    return {
        "caption": "历史财务表",
        "rows": [["项目\\报表日", *headers], *rows],
    }


def test_same_period_tables_fill_only_unresolved_cells():
    existing = _table(
        ["2022年度", "2023年度", "2024年度"],
        [
            ["总资产", "100.00", "200.00", "300.00"],
            ["负债", "XXX", "70.00", "90.00"],
            ["所有者权益", "XXX", "130.00", "210.00"],
        ],
    )
    candidate = _table(
        ["2022年度", "2023年度", "2024年度"],
        [
            ["总资产", "XXX", "XXX", "XXX"],
            ["负债", "40.00", "70.00", "90.00"],
            ["所有者权益", "60.00", "130.00", "210.00"],
        ],
    )

    merged = merge_historical_tables(existing, candidate)

    assert merged["rows"][1] == ["总资产", "100.00", "200.00", "300.00"]
    assert merged["rows"][2] == ["负债", "40.00", "70.00", "90.00"]
    assert merged["rows"][3] == ["所有者权益", "60.00", "130.00", "210.00"]


def test_incompatible_periods_choose_more_complete_dated_table():
    existing = _table(
        ["历史期1", "期初数", "期末数"],
        [
            ["总资产", "XXX", "246.00", "228.00"],
            ["负债", "XXX", "55.00", "12.00"],
            ["所有者权益", "XXX", "190.00", "216.00"],
        ],
    )
    candidate = _table(
        ["2022年度", "2023年度", "2024年度"],
        [
            ["总资产", "246.00", "228.00", "230.00"],
            ["负债", "55.00", "12.00", "10.00"],
            ["所有者权益", "191.00", "216.00", "220.00"],
        ],
    )

    assert merge_historical_tables(existing, candidate) == candidate


def test_incompatible_equal_coverage_prefers_explicit_calendar_periods():
    existing = _table(
        ["历史期1", "期初数", "期末数"],
        [["总资产", "100.00", "200.00", "300.00"]],
    )
    candidate = _table(
        ["2022年度", "2023年度", "2024年度"],
        [["总资产", "100.00", "200.00", "300.00"]],
    )

    assert merge_historical_tables(existing, candidate) == candidate
